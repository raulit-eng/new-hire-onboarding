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
    $replacing = Get-ReplacingID $Record
    if ($replacing -and -not ($Directory.users | Where-Object { $_.EmployeeID -eq $replacing })) {
        $issues.Add("Replacing employee '$replacing' was not found in the directory")
    }
    , $issues.ToArray()
}

function Get-ReplacingID($Record) {
    <# The ReplacingEmployeeID column is optional; blank means a new position. #>
    $p = $Record.PSObject.Properties['ReplacingEmployeeID']
    if ($p -and $p.Value) { "$($p.Value)".Trim() } else { $null }
}

function Get-Predecessor {
    <# The person a new hire is replacing, with their access and the equipment they had. #>
    param([Parameter(Mandatory)]$Record, [Parameter(Mandatory)]$Directory, $Inventory = @(), [Parameter(Mandatory)][datetime]$AsOf)
    $id = Get-ReplacingID $Record
    if (-not $id) { return $null }
    $u = $Directory.users | Where-Object { $_.EmployeeID -eq $id } | Select-Object -First 1
    if (-not $u) { return $null }
    $prop = { param($o, $n) if ($o.PSObject.Properties[$n]) { $o.$n } else { $null } }
    $end = & $prop $u 'EndDate'
    [pscustomobject][ordered]@{
        EmployeeID = $u.EmployeeID; Name = $u.DisplayName; Username = $u.Username; Title = & $prop $u 'JobTitle'
        Site = & $prop $u 'Site'; EndDate = $end
        StillWorking = [bool]($end -and [datetime]$end -ge $AsOf.Date)
        Groups = @(& $prop $u 'Groups' | Where-Object { $_ })
        Equipment = @($Inventory | Where-Object { $_.AssignedTo -eq $u.EmployeeID })
    }
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

# ---------------------------------------------------------------- equipment and IT tickets
function Import-EquipmentInventory {
    <# Reads the equipment list: returned devices, devices still with someone, and spares in stock. #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return [System.Collections.Generic.List[object]]::new() }
    [System.Collections.Generic.List[object]]@(Import-Csv $Path -Encoding utf8 | ForEach-Object {
            # Pool: Ready (cleaned and re-imaged), New (in box), or Loaner. Older files have no Pool column.
            if (-not $_.PSObject.Properties['Pool']) { $_ | Add-Member -NotePropertyName Pool -NotePropertyValue $(if ($_.Status -eq 'In Stock') { 'New' } else { '' }) }
            $_ | Add-Member -NotePropertyName Reserved -NotePropertyValue $false -PassThru
        })
}

function Get-BusinessDaysBefore {
    <# The date N working days (Mon-Fri) before $Date, but never before $NotBefore. #>
    param([datetime]$Date, [int]$Days, [datetime]$NotBefore)
    $d = $Date.Date
    while ($Days -gt 0) { $d = $d.AddDays(-1); if ($d.DayOfWeek -notin 'Saturday', 'Sunday') { $Days-- } }
    if ($d -lt $NotBefore.Date) { $NotBefore.Date } else { $d }
}

function Get-BusinessDaysAfter {
    <# The date N working days (Mon-Fri) after $Date. #>
    param([datetime]$Date, [int]$Days)
    $d = $Date.Date
    while ($Days -gt 0) { $d = $d.AddDays(1); if ($d.DayOfWeek -notin 'Saturday', 'Sunday') { $Days-- } }
    $d
}

function Get-DeviceAge([string]$PurchaseDate, [datetime]$AsOf) {
    [Math]::Round(($AsOf - [datetime]$PurchaseDate).TotalDays / 365.25, 1)
}

function New-DeviceTickets {
    <#
    .SYNOPSIS
        Decides how each device the new hire needs will be covered, and returns IT tickets for it.

        For a replacement hire, the previous person's equipment is reused when it has been
        returned and is still within its lifecycle (re-image and re-issue). Equipment past its
        lifecycle gets a retire ticket. Equipment that hasn't come back gets a recover ticket.
        Anything not covered that way comes from a spare: the site's own stock first, then the
        central depot (HQ), and Ready (cleaned and re-imaged) devices before New ones. If there
        is no spare, a new device is ordered and a loaner is issued for day one. Devices are
        reserved as they're assigned, so two hires never get the same one.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        $Predecessor,
        [Parameter(Mandatory)][AllowEmptyCollection()]$Inventory,
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][datetime]$AsOf,
        [Parameter(Mandatory)][ref]$NextNumber
    )
    $eq = $Config.equipment
    $start = [datetime]::ParseExact($Plan.StartDate, 'yyyy-MM-dd', $null)
    $due = Get-BusinessDaysBefore -Date $start -Days $eq.readyBusinessDaysBeforeStart -NotBefore $AsOf
    $daysToStart = ($start - $AsOf.Date).Days
    $tickets = [System.Collections.Generic.List[object]]::new()
    $team = "$($eq.supportTeam) - $($Plan.SiteCode)"

    $new = {
        param($Type, $Device, $Asset, $Model, $Summary, $Steps, $DueDate, $Priority, $From)
        $id = '{0}{1}' -f $eq.ticketPrefix, $NextNumber.Value
        $NextNumber.Value++
        [pscustomobject][ordered]@{
            id = $id; type = $Type; employee_id = $Plan.EmployeeID; for = $Plan.DisplayName; device = $Device
            asset_tag = $Asset; model = $Model; source = $From; site = $Plan.Site; team = $team
            due = $DueDate.ToString('yyyy-MM-dd'); priority = $Priority; status = 'Open'; summary = $Summary; steps = @($Steps)
            replacing = if ($Predecessor) { $Predecessor.Name } else { $null }
        }
    }
    $priority = if ($daysToStart -le 7) { 'High' } else { 'Normal' }
    $mdmStep = {
        param($device, $model)
        if ($device -in 'iPhone', 'iPad' -or $model -like 'MacBook*') { 'Erase and re-enroll in Jamf through Apple Business Manager; assign VPP apps' }
        else { 'Reset with Windows Autopilot and confirm Intune enrollment and BitLocker' }
    }

    foreach ($device in $Plan.Devices) {
        $covered = $false
        $life = [double]($eq.lifecycleYears.$device ?? 4)

        $personal = $device -in @($eq.personalItems)   # e.g. headsets: always a fresh one, never re-issued
        # 1. The previous person's equipment of this type
        $old = if ($Predecessor -and -not $personal) { $Predecessor.Equipment | Where-Object { $_.Type -eq $device -and -not $_.Reserved } | Select-Object -First 1 }
        if ($old) {
            $old.Reserved = $true
            $age = Get-DeviceAge $old.PurchaseDate $AsOf
            $tooOld = $age -ge $life
            if ($old.Status -ne 'Returned') {
                # Still with the previous person (or never returned): get it back first
                $when = if ($Predecessor.StillWorking) { "after their last day, $($Predecessor.EndDate)" } else { "now; it was due back $($Predecessor.EndDate)" }
                $recoverDue = if ($Predecessor.StillWorking) { Get-BusinessDaysAfter -Date ([datetime]$Predecessor.EndDate) -Days 1 } else { $AsOf.Date }
                $recoverPriority = if ($Predecessor.StillWorking) { 'Normal' } else { 'High' }
                $after = if ($tooOld) { "Retire it once it's back: it is $age years old, past its $life-year lifecycle" } else { 'Check it in, wipe it, and mark it In Stock for the next hire' }
                $tickets.Add((& $new 'Recover equipment' $device $old.AssetTag $old.Model `
                            "Collect $($Predecessor.Name)'s $device ($($old.AssetTag)) $when." `
                            @("Contact $($Predecessor.Name)'s manager to arrange the return", 'If it is not returned within 5 business days, lock and wipe it remotely through Intune/Jamf',
                        $after) $recoverDue $recoverPriority $Predecessor.Name))
            } elseif ($tooOld) {
                $tickets.Add((& $new 'Retire device' $device $old.AssetTag $old.Model `
                            "$($old.AssetTag) ($($old.Model), $age years old) is past its $life-year lifecycle. Retire it instead of re-issuing." `
                            @("Wipe $($old.AssetTag) and remove it from Intune/Jamf and Apple Business Manager", 'Mark retired in the asset inventory',
                        'Send to e-waste recycling and keep the certificate of destruction') $AsOf.Date.AddDays(30) 'Low' $Predecessor.Name))
            } else {
                $tickets.Add((& $new 'Re-image and re-issue' $device $old.AssetTag $old.Model `
                            "Re-issue $($Predecessor.Name)'s $device ($($old.AssetTag), $($old.Model), $age years old) to $($Plan.DisplayName)." `
                            @("Pull $($old.AssetTag) from the returned-equipment shelf at $($old.Site)", (& $mdmStep $device $old.Model),
                        "Install role apps: $($Plan.Apps -join ', ')", "Update the asset record: assigned to $($Plan.DisplayName) ($($Plan.EmployeeID))",
                        "Deliver to $($Plan.Site) by $($due.ToString('MMM d'))") $due $priority $Predecessor.Name))
                $covered = $true
            }
        }
        if ($covered) { continue }

        # 2. A spare: this site's stock, then the central depot; Ready (re-imaged) before New
        $pools = @($eq.stockPreference)
        $spare = $Inventory | Where-Object {
            $_.Status -eq 'In Stock' -and $_.Type -eq $device -and -not $_.Reserved -and $_.Pool -in $pools -and
            ($_.Site -eq $Plan.Site -or $_.Site -eq $eq.depotSite)
        } | Sort-Object @{ e = { $_.Site -ne $Plan.Site } }, @{ e = { [array]::IndexOf($pools, $_.Pool) } }, AssetTag | Select-Object -First 1
        if ($spare) {
            $spare.Reserved = $true
            $ship = if ($spare.Site -ne $Plan.Site) { "Ship from $($spare.Site) to $($Plan.Site)" } else { $null }
            $steps = @("Pull spare $($spare.AssetTag) ($($spare.Model)) at $($spare.Site)", (& $mdmStep $device $spare.Model),
                "Install role apps: $($Plan.Apps -join ', ')", "Update the asset record: assigned to $($Plan.DisplayName) ($($Plan.EmployeeID))") + @($ship | Where-Object { $_ }) +
            @("Deliver to $($Plan.Site) by $($due.ToString('MMM d'))")
            $where = if ($ship) { " (transfer from $($spare.Site))" } else { '' }
            if ($personal) {
                $steps = @("Pull new $device $($spare.AssetTag) ($($spare.Model)) at $($spare.Site)", "Update the asset record: assigned to $($Plan.DisplayName) ($($Plan.EmployeeID))") +
                @($ship | Where-Object { $_ }) + @("Deliver to $($Plan.Site) by $($due.ToString('MMM d'))")
                $tickets.Add((& $new 'Issue from stock' $device $spare.AssetTag $spare.Model "Issue new $device $($spare.AssetTag)$where to $($Plan.DisplayName)." $steps $due $priority 'Stock'))
                $tickets[-1] | Add-Member -NotePropertyName pool -NotePropertyValue $spare.Pool
            } else {
                $condition = if ($spare.Pool -eq 'Ready') { 're-imaged, ready to go' } else { 'new in box' }
            $tickets.Add((& $new 'Image from stock' $device $spare.AssetTag $spare.Model "Set up $device $($spare.AssetTag) ($condition)$where for $($Plan.DisplayName)." $steps $due $priority 'Stock'))
            $tickets[-1] | Add-Member -NotePropertyName pool -NotePropertyValue $spare.Pool
            }
            continue
        }

        # 3. Nothing available: order one, and lend a loaner so the new hire isn't empty-handed on day one
        $orderPriority = if ($daysToStart -le 7) { 'Urgent' } else { 'High' }
        $loaner = $Inventory | Where-Object {
            $_.Status -eq 'In Stock' -and $_.Pool -eq 'Loaner' -and $_.Type -eq $device -and -not $_.Reserved -and
            ($_.Site -eq $Plan.Site -or $_.Site -eq $eq.depotSite)
        } | Sort-Object @{ e = { $_.Site -ne $Plan.Site } }, AssetTag | Select-Object -First 1
        $steps = @("Request a quote for a standard $device and get manager approval", 'Place the order with expedited shipping',
            'On arrival: tag it, add it to the asset inventory, and enroll it in Intune/Jamf')
        if ($loaner) {
            $loaner.Reserved = $true
            $steps = @("Issue loaner $($loaner.AssetTag) ($($loaner.Model), from $($loaner.Site)) for day one") + $steps +
            @("Swap the new $device for loaner $($loaner.AssetTag), then re-image the loaner and return it to the loaner pool")
            $summary = "No $device in stock for $($Plan.DisplayName). Order one now, and lend loaner $($loaner.AssetTag) until it arrives."
        } else {
            $steps += "Deliver to $($Plan.Site) by $($start.ToString('MMM d')); no loaner is free, so escalate if it will be late"
            $summary = "No $device in stock and no loaner free for $($Plan.DisplayName). Order one now; it's needed by $($start.ToString('MMM d'))."
        }
        $tickets.Add((& $new 'Order new device' $device $null $null $summary $steps $due $orderPriority 'Purchase'))
        $tickets[-1] | Add-Member -NotePropertyName loaner_tag -NotePropertyValue $(if ($loaner) { $loaner.AssetTag } else { $null })
        $tickets[-1] | Add-Member -NotePropertyName loaner_model -NotePropertyValue $(if ($loaner) { $loaner.Model } else { $null })
    }
    , $tickets.ToArray()
}

function Get-AccessGap {
    <# Groups the previous person had that the new hire won't get automatically. The manager decides on these. #>
    param($Predecessor, [Parameter(Mandatory)]$Plan)
    # Callers wrap the result in @() because PowerShell turns an empty list into $null on return
    if (-not $Predecessor) { return }
    $planned = @($Plan.Groups) + $Plan.License
    $Predecessor.Groups | Where-Object { $_ -notin $planned -and $_ -notlike 'LIC-*' -and $_ -notlike 'Site-*' }
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
        [Parameter(Mandatory)]$Config,
        [string[]]$AccessGap = @(),
        [string]$PredecessorName
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
    if ($AccessGap.Count) {
        Write-Step 'Add to groups' warn "Not added automatically: $($AccessGap -join ', ') ($PredecessorName had it). Added to the checklist for the manager to approve" @who
    }

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
function Get-ChecklistData {
    <#
    .SYNOPSIS
        The checklist for one new hire as structured data: details at the top, then sections of
        items marked done (by the script) or still to do. The dashboard and the Markdown file
        both come from this.
    #>
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][AllowEmptyCollection()]$Tickets, $Predecessor, [string[]]$AccessGap = @())
    $item = { param($Text, $Done = $false) [pscustomobject][ordered]@{ text = $Text; done = [bool]$Done } }
    $details = [ordered]@{
        'Employee ID' = $Plan.EmployeeID; 'Title' = "$($Plan.JobTitle), $($Plan.Department)"; 'Site' = $Plan.Site
        'Manager' = $Plan.Manager; 'Start date' = $Plan.StartDate; 'Sign-in' = $Plan.Email; 'Windows logon' = $Plan.SamAccountName
    }
    if ($Plan.AccountExpires) { $details['Account expires'] = "$($Plan.AccountExpires) (contractor)" }
    $details['Position'] = if ($Predecessor) { "Replacing $($Predecessor.Name) ($($Predecessor.Title))" } else { 'New position' }

    $done = @(
        (& $item 'Account created with a temporary password (must change at first sign-in)' $true)
        (& $item "Added to groups: $($Plan.Groups -join ', ')" $true)
        (& $item "License assigned: $($Plan.License)" $true)
        (& $item "IT tickets opened: $(($Tickets | ForEach-Object { "$($_.id) ($($_.type.ToLower()), $($_.device))" }) -join '; ')" $true)
    )
    $before = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $Tickets | Where-Object { $_.type -in 'Re-image and re-issue', 'Image from stock', 'Issue from stock', 'Order new device' }) {
        $what = switch ($t.type) {
            'Re-image and re-issue' { "$($t.device) $($t.asset_tag) re-imaged and ready (previously $($t.replacing)'s)" }
            'Image from stock' { "$($t.device) $($t.asset_tag) imaged from stock and ready" }
            'Issue from stock' { "New $($t.device) $($t.asset_tag) ready" }
            'Order new device' {
                if ($t.PSObject.Properties['loaner_tag'] -and $t.loaner_tag) { "Loaner $($t.device) $($t.loaner_tag) ready for day one; new $($t.device) ordered" }
                else { "New $($t.device) ordered, received and set up (no loaner free: escalate if it will be late)" }
            }
        }
        $before.Add((& $item "$what - $($t.id), due $($t.due)"))
    }
    $recover = @($Tickets | Where-Object type -eq 'Recover equipment' | ForEach-Object {
            & $item "Recover $($t = $_; $t.replacing)'s $($t.device) $($t.asset_tag) by $($t.due) and return it to stock - $($t.id)"
        })
    $before.Add((& $item "Give the temporary password to $($Plan.Manager) through the approved password-delivery process"))
    $before.Add((& $item "Confirm app access: $($Plan.Apps -join ', ')"))
    if ($AccessGap.Count) {
        $before.Add((& $item "Ask $($Plan.Manager) whether $($Plan.DisplayName) also needs access $($Predecessor.Name) had: $($AccessGap -join ', ')"))
    }
    $sections = @(
        [pscustomobject]@{ title = 'Done by the script'; items = $done }
        [pscustomobject]@{ title = 'Before day one (IT)'; items = $before.ToArray() }
        if ($recover.Count) { [pscustomobject]@{ title = "Equipment to recover from $($Predecessor.Name)"; items = $recover } }
        [pscustomobject]@{ title = 'Day one (with the new hire)'; items = @(
                (& $item 'Hand off devices and have the new hire sign the equipment form')
                (& $item 'Sign in and change the temporary password')
                (& $item 'Enroll in multi-factor authentication (Duo)')
                (& $item 'Check Outlook, Teams and the nearest printer')
                (& $item 'Walk through how to open an IT ticket')) }
        [pscustomobject]@{ title = 'First week'; items = @(
                (& $item 'Security awareness training assigned and completed')
                (& $item 'Manager confirms all access is correct')
                @($Tickets | Where-Object { $_.PSObject.Properties['loaner_tag'] -and $_.loaner_tag } | ForEach-Object {
                        & $item "When the new $($_.device) arrives, swap it for loaner $($_.loaner_tag) and return the loaner - $($_.id)"
                    })) }
    )
    [pscustomobject]@{ title = "IT onboarding checklist: $($Plan.DisplayName)"; details = [pscustomobject]$details; sections = $sections }
}

function New-ITChecklist {
    <# Writes the checklist to a Markdown file and returns its path. #>
    param([Parameter(Mandatory)]$Checklist, [Parameter(Mandatory)][string]$Username, [Parameter(Mandatory)][string]$Folder)
    New-Item -ItemType Directory -Force -Path $Folder -WhatIf:$false | Out-Null
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# $($Checklist.title)"); $lines.Add(''); $lines.Add('| | |'); $lines.Add('|---|---|')
    foreach ($p in $Checklist.details.PSObject.Properties) { $lines.Add("| $($p.Name) | $($p.Value) |") }
    foreach ($section in $Checklist.sections) {
        $lines.Add(''); $lines.Add("## $($section.title)")
        foreach ($i in $section.items) { $lines.Add("- [$(if ($i.done) { 'x' } else { ' ' })] $($i.text)") }
    }
    $path = Join-Path $Folder "$Username.md"
    ($lines -join "`n") + "`n" | Set-Content -Path $path -Encoding utf8 -WhatIf:$false
    $path
}

Export-ModuleMember -Function Write-Step, Reset-RunLog, Get-RunLog, Set-LogQuiet, Set-PreviewMode, ConvertTo-PlainAscii, ConvertTo-NameCase,
Import-HRFile, Test-NewHireRecord, New-Username, Get-SamAccountName, New-TempPassword, Get-OnboardingPlan, New-SimulatedDirectory,
Save-SimulatedDirectory, Invoke-ProvisionAccount, New-ITChecklist, Get-ChecklistData, Get-ReplacingID, Get-Predecessor,
Import-EquipmentInventory, Get-BusinessDaysBefore, New-DeviceTickets, Get-AccessGap
