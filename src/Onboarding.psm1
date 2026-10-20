#Requires -Version 7.0
<#
.SYNOPSIS
    Building blocks for new-hire onboarding: reading the HR file, checking each row,
    building usernames, planning access, and carrying out the plan against a directory.

    Two directory "providers" do the actual work:
      * Simulated       - a JSON practice directory (default; safe to run anywhere)
      * ActiveDirectory - real AD cmdlets (needs the RSAT ActiveDirectory module; supports -WhatIf)
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------- run log
$script:Events = [System.Collections.Generic.List[object]]::new()
$script:Quiet = $false
$script:Preview = $false

function Reset-RunLog { $script:Events.Clear() }
function Get-RunLog { , $script:Events.ToArray() }
function Set-LogQuiet([bool]$Quiet) { $script:Quiet = $Quiet }
function Set-PreviewMode([bool]$Preview) { $script:Preview = $Preview }

function Write-Step {
    <# Records one step in the run log (which feeds the dashboard) and prints it to the console. #>
    param(
        [Parameter(Mandatory)][string]$Stage,
        [Parameter(Mandatory)][ValidateSet('info', 'ok', 'warn', 'error', 'skip')][string]$Status,
        [Parameter(Mandatory)][string]$Message,
        [string]$EmployeeID,
        [string]$Name
    )
    if ($script:Preview -and $Status -eq 'ok' -and $Stage -notin 'Read HR file', 'Validate', 'Summary') { $Message = "Preview, not done: $Message" }
    $evt = [ordered]@{
        seq = $script:Events.Count + 1; time = (Get-Date).ToString('o'); stage = $Stage; status = $Status
        employee_id = $EmployeeID; name = $Name; message = $Message
    }
    $script:Events.Add([pscustomobject]$evt)
    if (-not $script:Quiet) {
        $color = @{ info = 'Gray'; ok = 'Green'; warn = 'Yellow'; error = 'Red'; skip = 'DarkGray' }[$Status]
        $who = if ($Name) { "[$Name] " } else { '' }
        Write-Host ("{0,-16} {1}{2}" -f $Stage, $who, $Message) -ForegroundColor $color
    }
}

# ---------------------------------------------------------------- reading and cleaning HR data
function ConvertTo-PlainAscii {
    <# 'Hernández' -> 'Hernandez', "O'Brien" -> 'OBrien'. Usernames can't safely hold accents or apostrophes. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $decomposed = $Text.Normalize([Text.NormalizationForm]::FormD)
    $chars = $decomposed.ToCharArray() | Where-Object {
        [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark
    }
    (-join $chars) -replace "[^A-Za-z0-9\-]", ''
}

function ConvertTo-NameCase {
    <# 'grace' -> 'Grace', 'KIM' -> 'Kim', "o'brien" -> "O'Brien", 'delgado-villanueva' -> 'Delgado-Villanueva' #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)
    if (-not $Name) { return $Name }
    $lower = $Name.Trim().ToLowerInvariant()
    [regex]::Replace($lower, "(^|[\s\-'])(\p{L})", { param($m) $m.Groups[1].Value + $m.Groups[2].Value.ToUpperInvariant() })
}

function Import-HRFile {
    <# Reads the HR export and trims stray spaces from every field. #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { throw "HR file not found: $Path" }
    $rows = @(Import-Csv -Path $Path -Encoding utf8)
    $line = 1
    foreach ($row in $rows) {
        $line++
        $clean = [ordered]@{ RowNumber = $line }
        foreach ($p in $row.PSObject.Properties) { $clean[$p.Name] = ($p.Value ?? '').Trim() }
        $clean.FirstName = ConvertTo-NameCase $clean.FirstName
        $clean.LastName = ConvertTo-NameCase $clean.LastName
        [pscustomobject]$clean
    }
}

# ---------------------------------------------------------------- checks
function Test-NewHireRecord {
    <#
    .SYNOPSIS
        Returns the problems with one HR row. An empty list means it's ready to onboard.
        Anything here goes back to HR rather than being guessed at.
    #>
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)]$Directory,
        [Parameter(Mandatory)][datetime]$AsOf
    )
    $issues = [System.Collections.Generic.List[string]]::new()
    foreach ($field in 'EmployeeID', 'FirstName', 'LastName', 'Department', 'JobTitle', 'Site', 'Manager', 'StartDate') {
        if (-not $Record.$field) { $issues.Add("Missing $field") }
    }
    if ($Record.EmployeeID -and $Record.EmployeeID -notmatch '^E\d{5}$') { $issues.Add("Employee ID '$($Record.EmployeeID)' is not in the E12345 format") }
    if ($Record.Department -and -not $Config.departments.PSObject.Properties[$Record.Department]) {
        $issues.Add("Department '$($Record.Department)' has no access profile in config/onboarding.json")
    }
    if ($Record.Site -and -not $Config.sites.PSObject.Properties[$Record.Site]) { $issues.Add("Unknown site '$($Record.Site)'") }
    if ($Record.Manager -and -not ($Directory.users | Where-Object { $_.Email -eq $Record.Manager -and $_.Enabled })) {
        $issues.Add("Manager '$($Record.Manager)' was not found in the directory")
    }
    if ($Record.StartDate) {
        $start = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($Record.StartDate, 'yyyy-MM-dd', $null, 'None', [ref]$start)) {
            $issues.Add("Start date '$($Record.StartDate)' is not a valid date (use YYYY-MM-DD)")
        } elseif ($start -lt $AsOf.Date) {
            $issues.Add("Start date $($Record.StartDate) has already passed")
        } elseif ($start -gt $AsOf.Date.AddDays($Config.maxDaysAhead)) {
            $issues.Add("Start date $($Record.StartDate) is more than $($Config.maxDaysAhead) days away")
        }
    }
    if ($Record.EmploymentType -and $Record.EmploymentType -notin 'Full-time', 'Part-time', 'Contractor') {
        $issues.Add("Unknown employment type '$($Record.EmploymentType)'")
    }
    , $issues.ToArray()
}

# ---------------------------------------------------------------- planning
function New-Username {
    <#
    .SYNOPSIS
        first.last, made unique against the directory and anyone already planned in this run.
        'Ana Smith' with an existing ana.smith becomes ana.smith2. This is used for the
        sign-in name and email; see Get-SamAccountName for the 20-character AD logon name.
    #>
    param(
        [Parameter(Mandatory)][string]$FirstName,
        [Parameter(Mandatory)][string]$LastName,
        [string[]]$Taken = @()
    )
    $base = ("{0}.{1}" -f (ConvertTo-PlainAscii $FirstName), (ConvertTo-PlainAscii $LastName)).ToLowerInvariant()
    # (a separate name matters: PowerShell variables ignore case, so $taken would overwrite the [string[]] $Taken parameter)
    $inUse = [System.Collections.Generic.HashSet[string]]::new([string[]]$Taken, [StringComparer]::OrdinalIgnoreCase)
    $n = 1
    while ($true) {
        $suffix = if ($n -eq 1) { '' } else { "$n" }
        $candidate = $base + $suffix
        if (-not $inUse.Contains($candidate)) { return $candidate }
        $n++
    }
}

function Get-SamAccountName {
    <# The pre-Windows 2000 logon name can only be 20 characters. Longer usernames are shortened, keeping any number suffix. #>
    param([Parameter(Mandatory)][string]$Username)
    if ($Username.Length -le 20) { return $Username }
    $suffix = if ($Username -match '(\d+)$') { $Matches[1] } else { '' }
    $core = $Username.Substring(0, $Username.Length - $suffix.Length)
    $core.Substring(0, 20 - $suffix.Length).TrimEnd('.', '-') + $suffix
}

function New-TempPassword {
    <# A random 16-character temporary password, returned as a SecureString. It is never written to a log or file. #>
    $sets = 'ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!@#$%*?-'
    $all = -join $sets
    $chars = foreach ($set in $sets) { $set[[Security.Cryptography.RandomNumberGenerator]::GetInt32($set.Length)] }
    $chars += 1..12 | ForEach-Object { $all[[Security.Cryptography.RandomNumberGenerator]::GetInt32($all.Length)] }
    $shuffled = $chars | Sort-Object { [Security.Cryptography.RandomNumberGenerator]::GetInt32([int]::MaxValue) }
    $secure = [securestring]::new()
    foreach ($c in $shuffled) { $secure.AppendChar($c) }
    $secure.MakeReadOnly()
    $secure
}

function Get-OnboardingPlan {
    <# Works out everything a new hire needs: username, email, groups, license and devices. #>
    param(
        [Parameter(Mandatory)]$Record,
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Username
    )
    $dept = $Config.departments.($Record.Department)
    $site = $Config.sites.($Record.Site)
    $contractor = $Record.EmploymentType -eq 'Contractor'
    $groups = [System.Collections.Generic.List[string]]::new()
    $Config.everyoneGroups + $dept.groups + $site.group | ForEach-Object { if (-not $groups.Contains($_)) { $groups.Add($_) } }
    $license = $dept.license
    if ($contractor) {
        $Config.contractor.removeGroups | ForEach-Object { [void]$groups.Remove($_) }
        $Config.contractor.addGroups | ForEach-Object { if (-not $groups.Contains($_)) { $groups.Add($_) } }
        $license = $Config.contractor.license
    }
    $display = if ($Record.PreferredName) { "$($Record.PreferredName) $($Record.LastName)" } else { "$($Record.FirstName) $($Record.LastName)" }
    $start = [datetime]::ParseExact($Record.StartDate, 'yyyy-MM-dd', $null)
    [pscustomobject][ordered]@{
        EmployeeID     = $Record.EmployeeID
        DisplayName    = $display
        FirstName      = $Record.FirstName
        LastName       = $Record.LastName
        Username       = $Username
        SamAccountName = Get-SamAccountName $Username
        Email          = "$Username@$($Config.domain)"
        Department     = $Record.Department
        JobTitle       = $Record.JobTitle
        Site           = $Record.Site
        SiteCode       = $site.code
        Manager        = $Record.Manager
        StartDate      = $Record.StartDate
        EmploymentType = $Record.EmploymentType
        Groups         = $groups.ToArray()
        License        = $license
        Devices        = @($dept.devices)
        Apps           = @($dept.apps)
        AccountExpires = if ($contractor) { $start.AddDays($Config.contractorAccountDays).ToString('yyyy-MM-dd') } else { $null }
    }
}

# ---------------------------------------------------------------- providers
function New-SimulatedDirectory {
    <# Opens the JSON practice directory, creating it from the seed file the first time. #>
    param([Parameter(Mandatory)][string]$StatePath, [Parameter(Mandatory)][string]$SeedPath, [switch]$Reset)
    if ($Reset -or -not (Test-Path $StatePath)) {
        New-Item -ItemType Directory -Force -Path (Split-Path $StatePath) -WhatIf:$false | Out-Null
        Copy-Item $SeedPath $StatePath -Force -WhatIf:$false
    }
    $dir = Get-Content $StatePath -Raw | ConvertFrom-Json -AsHashtable
    $dir.users = [System.Collections.Generic.List[object]]@($dir.users | ForEach-Object { [pscustomobject]$_ })
    $dir.StatePath = $StatePath
    [pscustomobject]$dir
}

function Save-SimulatedDirectory {
    param([Parameter(Mandatory)]$Directory)
    $out = [ordered]@{ domain = $Directory.domain; users = $Directory.users; groups = $Directory.groups }
    $out | ConvertTo-Json -Depth 6 | Set-Content -Path $Directory.StatePath -Encoding utf8
}

function Invoke-ProvisionAccount {
    <#
    .SYNOPSIS
        Creates the account, adds groups (licenses are assigned by group), and returns what was done.
        Simulated mode writes to the practice directory; ActiveDirectory mode calls the real cmdlets.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)]$Directory,
        [Parameter(Mandatory)][ValidateSet('Simulated', 'ActiveDirectory')][string]$Mode,
        [Parameter(Mandatory)]$Config
    )
    $who = @{ EmployeeID = $Plan.EmployeeID; Name = "$($Plan.FirstName) $($Plan.LastName)" }
    $password = New-TempPassword   # generated for the account, then discarded; never logged

    if ($Mode -eq 'ActiveDirectory') {
        $manager = Get-ADUser -Filter "mail -eq '$($Plan.Manager)'" -ErrorAction Stop
        $params = @{
            Name = $Plan.DisplayName; GivenName = $Plan.FirstName; Surname = $Plan.LastName; DisplayName = $Plan.DisplayName
            SamAccountName = $Plan.SamAccountName; UserPrincipalName = $Plan.Email; EmailAddress = $Plan.Email
            EmployeeID = $Plan.EmployeeID; Title = $Plan.JobTitle; Department = $Plan.Department; Office = $Plan.Site
            Manager = $manager; Path = $Config.ouPath; AccountPassword = $password; ChangePasswordAtLogon = $true; Enabled = $true
        }
        if ($Plan.AccountExpires) { $params.AccountExpirationDate = [datetime]$Plan.AccountExpires }
        if ($PSCmdlet.ShouldProcess($Plan.Username, 'Create AD user')) { New-ADUser @params -ErrorAction Stop }
    } else {
        $Directory.users.Add([pscustomobject][ordered]@{
                EmployeeID = $Plan.EmployeeID; Username = $Plan.Username; SamAccountName = $Plan.SamAccountName; DisplayName = $Plan.DisplayName; Email = $Plan.Email
                Department = $Plan.Department; JobTitle = $Plan.JobTitle; Site = $Plan.Site; Manager = $Plan.Manager
                Enabled = $true; MustChangePassword = $true; AccountExpires = $Plan.AccountExpires; Created = (Get-Date).ToString('o')
            })
    }
    Write-Step 'Create account' ok "Created $($Plan.Email) for $($Plan.DisplayName); temporary password set, must change at first sign-in" @who
    if ($Plan.SamAccountName -ne $Plan.Username) {
        Write-Step 'Create account' info "Windows logon name shortened to the 20-character limit: $($Plan.SamAccountName)" @who
    }
    if ($Plan.AccountExpires) { Write-Step 'Create account' info "Contractor account set to expire $($Plan.AccountExpires)" @who }

    foreach ($group in $Plan.Groups) {
        if ($Mode -eq 'ActiveDirectory') {
            if ($PSCmdlet.ShouldProcess($group, "Add $($Plan.Username)")) { Add-ADGroupMember -Identity $group -Members $Plan.SamAccountName -ErrorAction Stop }
        } else {
            if (-not $Directory.groups.ContainsKey($group)) { $Directory.groups[$group] = [System.Collections.Generic.List[string]]::new() }
            elseif ($Directory.groups[$group] -isnot [System.Collections.Generic.List[string]]) {
                $Directory.groups[$group] = [System.Collections.Generic.List[string]]@($Directory.groups[$group])
            }
            $Directory.groups[$group].Add($Plan.Username)
        }
    }
    Write-Step 'Add to groups' ok "Added to $($Plan.Groups.Count) groups: $($Plan.Groups -join ', ')" @who

    # Licenses are assigned through license groups, the same way Entra ID group-based licensing works
    if ($Mode -eq 'ActiveDirectory') {
        if ($PSCmdlet.ShouldProcess($Plan.License, "Add $($Plan.Username)")) { Add-ADGroupMember -Identity $Plan.License -Members $Plan.SamAccountName -ErrorAction Stop }
    } else {
        if (-not $Directory.groups.ContainsKey($Plan.License)) { $Directory.groups[$Plan.License] = [System.Collections.Generic.List[string]]::new() }
        elseif ($Directory.groups[$Plan.License] -isnot [System.Collections.Generic.List[string]]) {
            $Directory.groups[$Plan.License] = [System.Collections.Generic.List[string]]@($Directory.groups[$Plan.License])
        }
        $Directory.groups[$Plan.License].Add($Plan.Username)
    }
    $licenseName = $Plan.License -replace '^LIC-', '' -replace '-', ' '
    Write-Step 'Assign license' ok "Assigned $licenseName license (via $($Plan.License) group)" @who
}

# ---------------------------------------------------------------- outputs
function New-ITChecklist {
    <# Writes a Markdown checklist for IT, the manager and the new hire. #>
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string[]]$RequestIds, [Parameter(Mandatory)][string]$Folder)
    New-Item -ItemType Directory -Force -Path $Folder -WhatIf:$false | Out-Null
    $devices = for ($i = 0; $i -lt $Plan.Devices.Count; $i++) { "- [ ] $($Plan.Devices[$i]) prepared and labeled (request $($RequestIds[$i]))" }
    $mdm = @($Plan.Devices | Where-Object { $_ -in 'Laptop', 'Desktop', 'iPhone', 'iPad' }) | ForEach-Object {
        if ($_ -in 'iPhone', 'iPad') { "- [ ] $_ enrolled in Jamf through Apple Business Manager" } else { "- [ ] $_ enrolled in Intune, BitLocker encryption confirmed" }
    }
    $lines = @(
        "# IT onboarding checklist: $($Plan.DisplayName)", '',
        "| | |", "|---|---|",
        "| Employee ID | $($Plan.EmployeeID) |", "| Title | $($Plan.JobTitle), $($Plan.Department) |",
        "| Site | $($Plan.Site) |", "| Manager | $($Plan.Manager) |", "| Start date | $($Plan.StartDate) |",
        "| Sign-in | $($Plan.Email) |", "| Windows logon | $($Plan.SamAccountName) |",
        $(if ($Plan.AccountExpires) { "| Account expires | $($Plan.AccountExpires) (contractor) |" }), '',
        '## Done by the script',
        "- [x] Account created with a temporary password (must change at first sign-in)",
        "- [x] Added to groups: $($Plan.Groups -join ', ')",
        "- [x] License assigned: $($Plan.License)",
        "- [x] Device requests opened: $($RequestIds -join ', ')", '',
        '## Before day one (IT)', $devices, $mdm,
        "- [ ] Give the temporary password to $($Plan.Manager) through the approved password-delivery process",
        "- [ ] Confirm app access: $($Plan.Apps -join ', ')", '',
        '## Day one (with the new hire)',
        '- [ ] Hand off devices and have the new hire sign the equipment form',
        '- [ ] Sign in and change the temporary password',
        '- [ ] Enroll in multi-factor authentication (Duo)',
        '- [ ] Check Outlook, Teams and the nearest printer',
        '- [ ] Walk through how to open an IT ticket', '',
        '## First week',
        '- [ ] Security awareness training assigned and completed',
        '- [ ] Manager confirms all access is correct'
    ) | Where-Object { $null -ne $_ }
    $path = Join-Path $Folder "$($Plan.Username).md"
    ($lines -join "`n") + "`n" | Set-Content -Path $path -Encoding utf8 -WhatIf:$false
    $path
}

Export-ModuleMember -Function Write-Step, Reset-RunLog, Get-RunLog, Set-LogQuiet, Set-PreviewMode, ConvertTo-PlainAscii, ConvertTo-NameCase,
Import-HRFile, Test-NewHireRecord, New-Username, Get-SamAccountName, New-TempPassword, Get-OnboardingPlan, New-SimulatedDirectory,
Save-SimulatedDirectory, Invoke-ProvisionAccount, New-ITChecklist
