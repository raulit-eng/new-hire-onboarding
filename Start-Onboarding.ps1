#Requires -Version 7.0
<#
.SYNOPSIS
    Onboards new hires from an HR spreadsheet: checks each row, creates accounts, adds groups,
    assigns licenses, requests devices and writes an IT checklist for each person.

.DESCRIPTION
    By default this runs against a practice directory (output/practice_directory.json), so it is
    safe to run anywhere. Use -Mode ActiveDirectory to run against a real domain; add -WhatIf
    first to preview every change without making it.

    Rows with problems are never guessed at. They go to output/hr_issues.csv for HR to fix,
    and the rest of the file keeps going. It is safe to run again: people who already have
    an account are skipped.

.EXAMPLE
    ./Start-Onboarding.ps1 -AsOf 2026-10-20 -Reset
    Runs the sample file against a fresh practice directory and builds the dashboard.

.EXAMPLE
    ./Start-Onboarding.ps1 -HRFile '\\hr\exports\new_hires.csv' -Mode ActiveDirectory -WhatIf
    Previews what would change in Active Directory, without changing anything.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$HRFile = (Join-Path $PSScriptRoot 'data/hr_new_hires.csv'),
    [ValidateSet('Simulated', 'ActiveDirectory')][string]$Mode = 'Simulated',
    [datetime]$AsOf = (Get-Date),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config/onboarding.json'),
    [switch]$Reset,          # start the practice directory over from data/directory_seed.json
    [switch]$NoDashboard,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'src/Onboarding.psm1') -Force
Reset-RunLog
Set-LogQuiet $Quiet.IsPresent
Set-PreviewMode ([bool]$WhatIfPreference)
$runStarted = Get-Date

# ---------- 1. read the HR file and settings
Write-Step 'Read HR file' info "Opening $(Split-Path $HRFile -Leaf)"
$records = @(Import-HRFile -Path $HRFile)
Write-Step 'Read HR file' ok "Read $($records.Count) rows from $(Split-Path $HRFile -Leaf)"
$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
Write-Step 'Read HR file' info "Loaded access profiles for $(@($config.departments.PSObject.Properties).Count) departments and $(@($config.sites.PSObject.Properties).Count) sites"

if ($Mode -eq 'ActiveDirectory') {
    if (-not (Get-Module -ListAvailable ActiveDirectory)) { throw 'The ActiveDirectory module (RSAT) is not installed.' }
    Import-Module ActiveDirectory
    $directory = [pscustomobject]@{
        domain = $config.domain
        users  = [System.Collections.Generic.List[object]]@(Get-ADUser -Filter * -Properties EmployeeID, mail | ForEach-Object {
                [pscustomobject]@{ EmployeeID = $_.EmployeeID; Username = $_.SamAccountName; Email = $_.mail; Enabled = $_.Enabled }
            })
    }
    Write-Step 'Read HR file' info "Connected to Active Directory: $($directory.users.Count) existing accounts"
} else {
    $directory = New-SimulatedDirectory -StatePath (Join-Path $OutputPath 'practice_directory.json') `
        -SeedPath (Join-Path $PSScriptRoot 'data/directory_seed.json') -Reset:($Reset -and -not $WhatIfPreference)
    Write-Step 'Read HR file' info "Opened practice directory: $($directory.users.Count) existing accounts"
}
if ($WhatIfPreference) { Write-Step 'Read HR file' warn 'Preview mode (-WhatIf): no accounts or groups will be changed; reports are still written' }

# ---------- 2. check every row, then onboard the good ones
$hires = [System.Collections.Generic.List[object]]::new()
$seen = @{}
$plannedNames = [System.Collections.Generic.List[string]]::new()
$requests = [System.Collections.Generic.List[object]]::new()
$requestNumber = 0

foreach ($r in $records) {
    $who = @{ EmployeeID = $r.EmployeeID; Name = "$($r.FirstName) $($r.LastName)".Trim() }
    $hire = [ordered]@{
        employee_id = $r.EmployeeID; name = $who.Name; department = $r.Department; title = $r.JobTitle; site = $r.Site
        manager = $r.Manager; start_date = $r.StartDate; employment_type = $r.EmploymentType; row = $r.RowNumber
        status = 'pending'; issues = @(); display_name = $null; sam = $null; username = $null; email = $null; groups = @(); license = $null
        devices = @(); requests = @(); checklist = $null; account_expires = $null
    }
    Write-Step 'Validate' info "Checking row $($r.RowNumber): $($r.JobTitle), $($r.Department)" @who

    if ($r.EmployeeID -and $seen.ContainsKey($r.EmployeeID)) {
        $hire.status = 'skipped'; $hire.issues = @("Duplicate of row $($seen[$r.EmployeeID]) in the HR file")
        Write-Step 'Validate' skip "Skipped: duplicate of row $($seen[$r.EmployeeID])" @who
        $hires.Add([pscustomobject]$hire); continue
    }
    if ($r.EmployeeID) { $seen[$r.EmployeeID] = $r.RowNumber }

    $existing = $directory.users | Where-Object { $_.EmployeeID -and $_.EmployeeID -eq $r.EmployeeID } | Select-Object -First 1
    if ($existing -and $existing.Enabled) {
        $hire.status = 'skipped'; $hire.username = $existing.Username; $hire.issues = @("Already has an active account ($($existing.Username))")
        Write-Step 'Validate' skip "Skipped: already onboarded as $($existing.Username)" @who
        $hires.Add([pscustomobject]$hire); continue
    }
    if ($existing) {
        $hire.status = 'needs_hr'; $hire.issues = @("Former employee with a disabled account ($($existing.Username)). Confirm this is a rehire, then reactivate that account instead of creating a new one")
        Write-Step 'Validate' warn "Held for HR: matches disabled account $($existing.Username) (possible rehire)" @who
        $hires.Add([pscustomobject]$hire); continue
    }

    $issues = Test-NewHireRecord -Record $r -Config $config -Directory $directory -AsOf $AsOf
    if ($issues.Count) {
        $hire.status = 'needs_hr'; $hire.issues = $issues
        foreach ($i in $issues) { Write-Step 'Validate' warn "Held for HR: $i" @who }
        $hires.Add([pscustomobject]$hire); continue
    }
    Write-Step 'Validate' ok 'All checks passed' @who

    try {
        $taken = @($directory.users | ForEach-Object { $_.Username; if ($_.PSObject.Properties['SamAccountName']) { $_.SamAccountName } }) + $plannedNames
        $username = New-Username -FirstName $r.FirstName -LastName $r.LastName -Taken $taken
        $first = New-Username -FirstName $r.FirstName -LastName $r.LastName
        if ($username -ne $first) { Write-Step 'Create account' info "Username $first is already taken; using $username" @who }
        $plannedNames.Add($username)

        $plan = Get-OnboardingPlan -Record $r -Config $config -Username $username
        Invoke-ProvisionAccount -Plan $plan -Directory $directory -Mode $Mode -Config $config

        $ids = @(foreach ($device in $plan.Devices) {
            $requestNumber++
            $id = 'REQ-{0}-{1:D3}' -f $AsOf.ToString('yyMMdd'), $requestNumber
            $requests.Add([pscustomobject]@{ RequestID = $id; EmployeeID = $plan.EmployeeID; Name = $plan.DisplayName; Device = $device
                    Site = $plan.Site; NeededBy = $plan.StartDate; Status = 'Open' })
            $id
        })
        $listed = for ($k = 0; $k -lt $ids.Count; $k++) { "$($plan.Devices[$k]) ($($ids[$k]))" }
        $noun = if ($ids.Count -eq 1) { 'request' } else { 'requests' }
        Write-Step 'Request device' ok "Opened $($ids.Count) device ${noun}: $($listed -join ', '), needed by $($plan.StartDate)" @who

        $checklist = New-ITChecklist -Plan $plan -RequestIds $ids -Folder (Join-Path $OutputPath 'checklists')
        Write-Step 'Write checklist' ok "Wrote IT checklist: checklists/$(Split-Path $checklist -Leaf)" @who

        $hire.status = 'onboarded'; $hire.display_name = $plan.DisplayName; $hire.sam = $plan.SamAccountName; $hire.username = $plan.Username; $hire.email = $plan.Email
        $hire.groups = $plan.Groups; $hire.license = $plan.License; $hire.devices = $plan.Devices; $hire.requests = @($ids)
        $hire.checklist = "checklists/$(Split-Path $checklist -Leaf)"; $hire.account_expires = $plan.AccountExpires
    } catch {
        $hire.status = 'error'; $hire.issues = @("$_")
        Write-Step 'Create account' error "Failed: $_" @who
    }
    $hires.Add([pscustomobject]$hire)
}

# ---------- 3. save results
# Reports are written even in preview mode, so you can review what would happen
$PSDefaultParameterValues['Export-Csv:WhatIf'] = $false
$PSDefaultParameterValues['Set-Content:WhatIf'] = $false
New-Item -ItemType Directory -Force -Path $OutputPath -WhatIf:$false | Out-Null
if ($Mode -eq 'Simulated' -and -not $WhatIfPreference) { Save-SimulatedDirectory -Directory $directory }

$hires | Select-Object employee_id, name, status, username, email, department, title, site, start_date, license,
@{ n = 'groups'; e = { $_.groups -join '; ' } }, @{ n = 'requests'; e = { $_.requests -join '; ' } }, account_expires |
    Export-Csv (Join-Path $OutputPath 'onboarding_summary.csv') -NoTypeInformation -Encoding utf8
$hires | Where-Object status -eq 'needs_hr' | Select-Object employee_id, name, row, @{ n = 'issue'; e = { $_.issues -join '; ' } } |
    Export-Csv (Join-Path $OutputPath 'hr_issues.csv') -NoTypeInformation -Encoding utf8
$requests | Export-Csv (Join-Path $OutputPath 'device_requests.csv') -NoTypeInformation -Encoding utf8

$counts = [ordered]@{}
foreach ($s in 'onboarded', 'needs_hr', 'skipped', 'error') { $counts[$s] = @($hires | Where-Object status -eq $s).Count }
Write-Step 'Summary' ok ("Done: {0} onboarded, {1} held for HR, {2} skipped, {3} errors" -f $counts.onboarded, $counts.needs_hr, $counts.skipped, $counts.error)
if ($counts.needs_hr) { Write-Step 'Summary' warn "Sent $($counts.needs_hr) row(s) back to HR: output/hr_issues.csv" }

$run = [ordered]@{
    company = 'Summit Logistics (sample data)'; hr_file = (Split-Path $HRFile -Leaf); mode = $Mode
    what_if = [bool]$WhatIfPreference; as_of = $AsOf.ToString('yyyy-MM-dd'); started = $runStarted.ToString('o')
    finished = (Get-Date).ToString('o'); counts = $counts; device_requests = $requests.Count
    hires = $hires; events = (Get-RunLog)
}
$runLogPath = Join-Path $OutputPath 'run_log.json'
$run | ConvertTo-Json -Depth 6 | Set-Content $runLogPath -Encoding utf8

if (-not $NoDashboard) {
    & (Join-Path $PSScriptRoot 'Build-Dashboard.ps1') -RunLog $runLogPath | Out-Null
    if (-not $Quiet) { Write-Host "Dashboard: docs/index.html" -ForegroundColor Cyan }
}
[pscustomobject]$counts
