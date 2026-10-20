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
    [string]$InventoryFile,  # equipment list; defaults to the practice copy of data/equipment_inventory.csv
    [switch]$Reset,          # start the practice directory, equipment list and ticket numbers over
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
# Equipment: returned devices, devices still with someone, and spares in stock
$seedInventory = Join-Path $PSScriptRoot 'data/equipment_inventory.csv'
if (-not $InventoryFile) {
    $InventoryFile = Join-Path $OutputPath 'practice_inventory.csv'
    if (($Reset -and -not $WhatIfPreference) -or -not (Test-Path $InventoryFile)) {
        New-Item -ItemType Directory -Force -Path $OutputPath -WhatIf:$false | Out-Null
        Copy-Item $seedInventory $InventoryFile -Force -WhatIf:$false
    }
}
$inventory = Import-EquipmentInventory -Path $InventoryFile
$spares = @($inventory | Where-Object Status -eq 'In Stock').Count
Write-Step 'Read HR file' info "Loaded equipment list: $($inventory.Count) devices, $spares spares in stock"

$counterFile = Join-Path $OutputPath 'ticket_counter.txt'
$nextTicket = if (-not $Reset -and (Test-Path $counterFile)) { [int](Get-Content $counterFile -Raw) } else { [int]$config.equipment.firstTicketNumber }
if ($WhatIfPreference) { Write-Step 'Read HR file' warn 'Preview mode (-WhatIf): no accounts or groups will be changed; reports are still written' }

# ---------- 2. check every row, then onboard the good ones
$hires = [System.Collections.Generic.List[object]]::new()
$seen = @{}
$plannedNames = [System.Collections.Generic.List[string]]::new()
$allTickets = [System.Collections.Generic.List[object]]::new()

foreach ($r in $records) {
    $who = @{ EmployeeID = $r.EmployeeID; Name = "$($r.FirstName) $($r.LastName)".Trim() }
    $hire = [ordered]@{
        employee_id = $r.EmployeeID; name = $who.Name; department = $r.Department; title = $r.JobTitle; site = $r.Site
        manager = $r.Manager; start_date = $r.StartDate; employment_type = $r.EmploymentType; row = $r.RowNumber
        status = 'pending'; issues = @(); display_name = $null; sam = $null; username = $null; email = $null; groups = @(); license = $null
        devices = @(); requests = @(); tickets = @(); checklist = $null; checklist_data = $null; account_expires = $null
        replacing = $null
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
    $predecessor = Get-Predecessor -Record $r -Directory $directory -Inventory $inventory -AsOf $AsOf
    if ($predecessor) {
        $left = if ($predecessor.StillWorking) { "still working until $($predecessor.EndDate)" } else { "left $($predecessor.EndDate)" }
        Write-Step 'Validate' info "Replacing $($predecessor.Name), $($predecessor.Title) ($left); $($predecessor.Equipment.Count) device(s) on record" @who
    } else {
        Write-Step 'Validate' info 'New position (not replacing anyone)' @who
    }

    try {
        $taken = @($directory.users | ForEach-Object { $_.Username; if ($_.PSObject.Properties['SamAccountName']) { $_.SamAccountName } }) + $plannedNames
        $username = New-Username -FirstName $r.FirstName -LastName $r.LastName -Taken $taken
        $first = New-Username -FirstName $r.FirstName -LastName $r.LastName
        if ($username -ne $first) { Write-Step 'Create account' info "Username $first is already taken; using $username" @who }
        $plannedNames.Add($username)

        $plan = Get-OnboardingPlan -Record $r -Config $config -Username $username
        $gap = @(Get-AccessGap -Predecessor $predecessor -Plan $plan)
        Invoke-ProvisionAccount -Plan $plan -Directory $directory -Mode $Mode -Config $config -AccessGap $gap -PredecessorName $(if ($predecessor) { $predecessor.Name })

        $counter = [ref]$nextTicket
        $tickets = New-DeviceTickets -Plan $plan -Predecessor $predecessor -Inventory $inventory -Config $config -AsOf $AsOf -NextNumber $counter
        $nextTicket = $counter.Value
        foreach ($t in $tickets) {
            $asset = if ($t.asset_tag) { " $($t.asset_tag)" } else { '' }
            $status = if ($t.priority -in 'Urgent') { 'warn' } else { 'ok' }
            Write-Step 'Create tickets' $status "$($t.id) $($t.type): $($t.device)$asset, due $([datetime]::Parse($t.due).ToString('MMM d')), $($t.priority) priority" @who
            $allTickets.Add($t)
        }
        $checklistData = Get-ChecklistData -Plan $plan -Tickets $tickets -Predecessor $predecessor -AccessGap $gap
        $checklist = New-ITChecklist -Checklist $checklistData -Username $plan.Username -Folder (Join-Path $OutputPath 'checklists')
        $todo = @($checklistData.sections | ForEach-Object { $_.items } | Where-Object { -not $_.done }).Count
        Write-Step 'Write checklist' ok "Wrote IT checklist: checklists/$(Split-Path $checklist -Leaf) ($todo items to do)" @who

        $hire.status = 'onboarded'; $hire.display_name = $plan.DisplayName; $hire.sam = $plan.SamAccountName; $hire.username = $plan.Username; $hire.email = $plan.Email
        $hire.groups = $plan.Groups; $hire.license = $plan.License; $hire.devices = $plan.Devices
        $hire.tickets = @($tickets | ForEach-Object id); $hire.checklist = "checklists/$(Split-Path $checklist -Leaf)"
        $hire.checklist_data = $checklistData; $hire.account_expires = $plan.AccountExpires
        if ($predecessor) {
            $hire.replacing = [ordered]@{
                employee_id = $predecessor.EmployeeID; name = $predecessor.Name; title = $predecessor.Title; site = $predecessor.Site
                end_date = $predecessor.EndDate; still_working = $predecessor.StillWorking; access_gap = @($gap)
                equipment = @($predecessor.Equipment | ForEach-Object {
                        $use = $tickets | Where-Object asset_tag -eq $_.AssetTag | Select-Object -First 1
                        [ordered]@{ asset_tag = $_.AssetTag; type = $_.Type; model = $_.Model; purchased = $_.PurchaseDate; status = $_.Status
                            outcome = if ($use) { $use.type } elseif ($_.Type -in @($config.equipment.personalItems)) { 'Not re-issued (personal item)' } else { 'Not needed' }
                            ticket = if ($use) { $use.id } else { $null } }
                    })
            }
        }
    } catch {
        $hire.status = 'error'; $hire.issues = @("$_")
        Write-Verbose $_.ScriptStackTrace
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
if (-not $WhatIfPreference) {
    # Record which devices are now spoken for, so the next run doesn't hand out the same spare
    foreach ($t in $allTickets | Where-Object { $_.asset_tag -and $_.type -in 'Re-image and re-issue', 'Image from stock', 'Issue from stock' }) {
        $d = $inventory | Where-Object AssetTag -eq $t.asset_tag
        $d.Status = 'Reserved'; $d.AssignedTo = $t.employee_id; $d.Notes = "Reserved for $($t.for) ($($t.id))"
    }
    foreach ($t in $allTickets | Where-Object type -eq 'Retire device') {
        $d = $inventory | Where-Object AssetTag -eq $t.asset_tag
        $d.Status = 'Pending retirement'; $d.Notes = "Retire ($($t.id))"
    }
    $inventory | Select-Object AssetTag, Type, Model, PurchaseDate, Status, AssignedTo, Site, Notes |
        Export-Csv $InventoryFile -NoTypeInformation -Encoding utf8
    Set-Content $counterFile $nextTicket
}

$hires | Select-Object employee_id, name, status, username, email, department, title, site, start_date, license,
@{ n = 'groups'; e = { $_.groups -join '; ' } }, @{ n = 'tickets'; e = { $_.tickets -join '; ' } },
@{ n = 'replacing'; e = { if ($_.replacing) { $_.replacing.name } else { '' } } }, account_expires |
    Export-Csv (Join-Path $OutputPath 'onboarding_summary.csv') -NoTypeInformation -Encoding utf8
$hires | Where-Object status -eq 'needs_hr' | Select-Object employee_id, name, row, @{ n = 'issue'; e = { $_.issues -join '; ' } } |
    Export-Csv (Join-Path $OutputPath 'hr_issues.csv') -NoTypeInformation -Encoding utf8
$allTickets | Select-Object id, type, priority, due, employee_id, for, device, asset_tag, model, source, site, team, status, summary,
@{ n = 'steps'; e = { $_.steps -join ' | ' } } | Export-Csv (Join-Path $OutputPath 'tickets.csv') -NoTypeInformation -Encoding utf8
ConvertTo-Json @($allTickets) -Depth 5 | Set-Content (Join-Path $OutputPath 'tickets.json') -Encoding utf8

$counts = [ordered]@{}
foreach ($s in 'onboarded', 'needs_hr', 'skipped', 'error') { $counts[$s] = @($hires | Where-Object status -eq $s).Count }
Write-Step 'Summary' ok ("Done: {0} onboarded, {1} held for HR, {2} skipped, {3} errors" -f $counts.onboarded, $counts.needs_hr, $counts.skipped, $counts.error)
Write-Step 'Summary' info ("Opened {0} IT tickets: {1}" -f $allTickets.Count, (($allTickets | Group-Object type | ForEach-Object { "$($_.Count) $($_.Name.ToLower())" }) -join ', '))
if ($counts.needs_hr) { Write-Step 'Summary' warn "Sent $($counts.needs_hr) row(s) back to HR: output/hr_issues.csv" }

$run = [ordered]@{
    company = 'Summit Logistics (sample data)'; hr_file = (Split-Path $HRFile -Leaf); mode = $Mode
    what_if = [bool]$WhatIfPreference; as_of = $AsOf.ToString('yyyy-MM-dd'); started = $runStarted.ToString('o')
    finished = (Get-Date).ToString('o'); counts = $counts
    hires = $hires; tickets = @($allTickets); events = (Get-RunLog)
}
$runLogPath = Join-Path $OutputPath 'run_log.json'
$run | ConvertTo-Json -Depth 10 | Set-Content $runLogPath -Encoding utf8

if (-not $NoDashboard) {
    & (Join-Path $PSScriptRoot 'Build-Dashboard.ps1') -RunLog $runLogPath | Out-Null
    if (-not $Quiet) { Write-Host "Dashboard: docs/index.html" -ForegroundColor Cyan }
}
[pscustomobject]$counts
