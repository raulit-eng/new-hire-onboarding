# Pester tests. Run with:  Invoke-Pester ./tests

BeforeAll {
    $root = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $root 'src/Onboarding.psm1') -Force
    $config = Get-Content (Join-Path $root 'config/onboarding.json') -Raw | ConvertFrom-Json
    $asOf = [datetime]'2026-10-20'
}

Describe 'Usernames' {
    It 'uses first.last in lowercase' {
        New-Username -FirstName 'Maria' -LastName 'Gonzalez' | Should -Be 'maria.gonzalez'
    }
    It 'removes accents and apostrophes' {
        New-Username -FirstName 'Luis' -LastName 'Hernández' | Should -Be 'luis.hernandez'
        New-Username -FirstName 'Sofia' -LastName "O'Brien" | Should -Be 'sofia.obrien'
    }
    It 'adds a number when the name is taken' {
        New-Username -FirstName 'Ana' -LastName 'Smith' -Taken 'ana.smith' | Should -Be 'ana.smith2'
        New-Username -FirstName 'Ana' -LastName 'Smith' -Taken 'ana.smith', 'ANA.SMITH2' | Should -Be 'ana.smith3'
    }
    It 'shortens the Windows logon name to 20 characters' {
        Get-SamAccountName 'rosa.delgado-villanueva' | Should -Be 'rosa.delgado-villanu'
        Get-SamAccountName 'rosa.delgado-villanueva2' | Should -Be 'rosa.delgado-villan2'
        (Get-SamAccountName 'maria.gonzalez') | Should -Be 'maria.gonzalez'
    }
}

Describe 'Cleaning HR data' {
    It 'fixes name capitalization' {
        ConvertTo-NameCase 'KIM' | Should -Be 'Kim'
        ConvertTo-NameCase 'grace' | Should -Be 'Grace'
        ConvertTo-NameCase "o'brien" | Should -Be "O'Brien"
        ConvertTo-NameCase 'delgado-villanueva' | Should -Be 'Delgado-Villanueva'
    }
}

Describe 'Temporary passwords' {
    It 'are 16 characters with every character type' {
        $plain = [Net.NetworkCredential]::new('', (New-TempPassword)).Password
        $plain.Length | Should -Be 16
        $plain | Should -MatchExactly '[A-Z]'
        $plain | Should -MatchExactly '[a-z]'
        $plain | Should -Match '[0-9]'
        $plain | Should -Match '[!@#$%*?\-]'
    }
    It 'are different every time' {
        $a = [Net.NetworkCredential]::new('', (New-TempPassword)).Password
        $b = [Net.NetworkCredential]::new('', (New-TempPassword)).Password
        $a | Should -Not -Be $b
    }
}

Describe 'Checking HR rows' {
    BeforeAll {
        $directory = [pscustomobject]@{ users = @([pscustomobject]@{ Email = 'boss@summitlogistics.example'; Enabled = $true }) }
        function New-Row([hashtable]$Change = @{}) {
            $row = [ordered]@{ EmployeeID = 'E20001'; FirstName = 'Test'; LastName = 'Person'; Department = 'Sales'
                JobTitle = 'Rep'; Site = 'Burbank'; Manager = 'boss@summitlogistics.example'; StartDate = '2026-10-26'; EmploymentType = 'Full-time' }
            foreach ($k in $Change.Keys) { $row[$k] = $Change[$k] }
            [pscustomobject]$row
        }
    }
    It 'passes a good row' {
        Test-NewHireRecord -Record (New-Row) -Config $config -Directory $directory -AsOf $asOf | Should -BeNullOrEmpty
    }
    It 'flags <Problem>' -ForEach @(
        @{ Problem = 'a missing manager'; Change = @{ Manager = '' }; Expect = 'Missing Manager' }
        @{ Problem = 'an unknown manager'; Change = @{ Manager = 'nobody@summitlogistics.example' }; Expect = 'not found in the directory' }
        @{ Problem = 'an unknown department'; Change = @{ Department = 'Legal' }; Expect = "Department 'Legal'" }
        @{ Problem = 'an unknown site'; Change = @{ Site = 'Mars' }; Expect = "Unknown site 'Mars'" }
        @{ Problem = 'a start date in the past'; Change = @{ StartDate = '2026-10-01' }; Expect = 'already passed' }
        @{ Problem = 'a start date too far away'; Change = @{ StartDate = '2027-06-01' }; Expect = 'more than 60 days' }
        @{ Problem = 'a badly formatted date'; Change = @{ StartDate = '10/26/2026' }; Expect = 'not a valid date' }
        @{ Problem = 'a bad employee ID'; Change = @{ EmployeeID = '12345' }; Expect = 'E12345 format' }
    ) {
        $issues = Test-NewHireRecord -Record (New-Row $Change) -Config $config -Directory $directory -AsOf $asOf
        ($issues -join '; ') | Should -BeLike "*$Expect*"
    }
}

Describe 'Access plans' {
    BeforeAll {
        $base = @{ EmployeeID = 'E20001'; FirstName = 'Test'; LastName = 'Person'; PreferredName = ''; Department = 'Sales'; JobTitle = 'Rep'
            Site = 'Burbank'; Manager = 'boss@x'; StartDate = '2026-10-26'; EmploymentType = 'Full-time' }
    }
    It 'combines company, department and site groups' {
        $plan = Get-OnboardingPlan -Record ([pscustomobject]$base) -Config $config -Username 'test.person'
        $plan.Groups | Should -Contain 'All-Staff'
        $plan.Groups | Should -Contain 'Sales-Team'
        $plan.Groups | Should -Contain 'Site-BUR'
        $plan.License | Should -Be 'LIC-M365-E3'
        $plan.Devices | Should -Be @('Laptop', 'iPhone')
    }
    It 'gives contractors limited access and an expiring account' {
        $row = $base.Clone(); $row.EmploymentType = 'Contractor'
        $plan = Get-OnboardingPlan -Record ([pscustomobject]$row) -Config $config -Username 'test.person'
        $plan.Groups | Should -Contain 'Contractors'
        $plan.Groups | Should -Not -Contain 'VPN-Users'
        $plan.License | Should -Be 'LIC-M365-F3'
        $plan.AccountExpires | Should -Be '2027-01-24'
    }
}

Describe 'Full run on the sample HR file' {
    BeforeAll {
        $out = Join-Path ([IO.Path]::GetTempPath()) "onboarding-test-$([guid]::NewGuid())"
        $script = Join-Path $root 'Start-Onboarding.ps1'
        $first = & $script -AsOf $asOf -OutputPath $out -Reset -NoDashboard -Quiet
        $log = Get-Content (Join-Path $out 'run_log.json') -Raw | ConvertFrom-Json
    }
    AfterAll { Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue }

    It 'onboards the good rows and holds the rest for HR' {
        $first.onboarded | Should -Be 11
        $first.needs_hr | Should -Be 5
        $first.skipped | Should -Be 1
        $first.error | Should -Be 0
    }
    It 'gives Ana Smith a unique username' {
        ($log.hires | Where-Object employee_id -eq 'E10423').username | Should -Be 'ana.smith2'
    }
    It 'flags the rehire instead of creating a second account' {
        $jasmine = $log.hires | Where-Object employee_id -eq 'E10435'
        $jasmine.status | Should -Be 'needs_hr'
        $jasmine.issues[0] | Should -BeLike '*disabled account*'
    }
    It 'writes a checklist for everyone onboarded' {
        @(Get-ChildItem (Join-Path $out 'checklists') -Filter *.md).Count | Should -Be 11
    }
    It 'never writes a password to any output file' {
        foreach ($file in Get-ChildItem $out -Recurse -File) {
            Get-Content $file.FullName -Raw | Should -Not -Match '(?i)password\s*[:=]'
        }
    }
    It 'is safe to run again: nobody is onboarded twice' {
        $second = & $script -AsOf $asOf -OutputPath $out -NoDashboard -Quiet
        $second.onboarded | Should -Be 0
        $second.skipped | Should -Be 12
    }
    It 'changes nothing in preview mode' {
        $preview = Join-Path ([IO.Path]::GetTempPath()) "onboarding-preview-$([guid]::NewGuid())"
        & $script -AsOf $asOf -OutputPath $preview -WhatIf -NoDashboard -Quiet | Out-Null
        $dir = Get-Content (Join-Path $preview 'practice_directory.json') -Raw | ConvertFrom-Json
        $seed = Get-Content (Join-Path $root 'data/directory_seed.json') -Raw | ConvertFrom-Json
        @($dir.users).Count | Should -Be @($seed.users).Count
        Test-Path (Join-Path $preview 'ticket_counter.txt') | Should -BeFalse
        Remove-Item $preview -Recurse -Force
    }
}

Describe 'Active Directory mode (with the AD cmdlets mocked)' {
    BeforeAll {
        # Stand-ins so the test runs on machines without RSAT; Pester replaces them with mocks
        function global:New-ADUser { param($Name, $GivenName, $Surname, $DisplayName, $SamAccountName, $UserPrincipalName, $EmailAddress,
                $EmployeeID, $Title, $Department, $Office, $Manager, $Path, $AccountPassword, $ChangePasswordAtLogon, $Enabled, $AccountExpirationDate) }
        function global:Add-ADGroupMember { param($Identity, $Members) }
        function global:Get-ADUser { param($Filter) }
        Mock -ModuleName Onboarding Get-ADUser { [pscustomobject]@{ SamAccountName = 'boss' } }
        Mock -ModuleName Onboarding New-ADUser { }
        Mock -ModuleName Onboarding Add-ADGroupMember { }
        Set-LogQuiet $true
        $row = [pscustomobject]@{ EmployeeID = 'E20001'; FirstName = 'Rosa'; LastName = 'Delgado-Villanueva'; PreferredName = ''; Department = 'HR'
            JobTitle = 'HR Generalist'; Site = 'Burbank'; Manager = 'boss@x'; StartDate = '2026-10-26'; EmploymentType = 'Full-time' }
        $plan = Get-OnboardingPlan -Record $row -Config $config -Username 'rosa.delgado-villanueva'
    }
    AfterAll { Remove-Item function:global:New-ADUser, function:global:Add-ADGroupMember, function:global:Get-ADUser -ErrorAction SilentlyContinue }

    It 'creates the user with a 20-character logon name and a password that must be changed' {
        Invoke-ProvisionAccount -Plan $plan -Directory ([pscustomobject]@{}) -Mode ActiveDirectory -Config $config
        Should -Invoke -ModuleName Onboarding New-ADUser -Times 1 -ParameterFilter {
            $SamAccountName -eq 'rosa.delgado-villanu' -and $UserPrincipalName -eq 'rosa.delgado-villanueva@summitlogistics.example' -and
            $ChangePasswordAtLogon -eq $true -and $AccountPassword -is [securestring]
        }
        Should -Invoke -ModuleName Onboarding Add-ADGroupMember -Times ($plan.Groups.Count + 1) -Exactly
    }
    It 'makes no changes with -WhatIf' {
        Invoke-ProvisionAccount -Plan $plan -Directory ([pscustomobject]@{}) -Mode ActiveDirectory -Config $config -WhatIf
        Should -Invoke -ModuleName Onboarding New-ADUser -Times 0 -Exactly
        Should -Invoke -ModuleName Onboarding Add-ADGroupMember -Times 0 -Exactly
    }
}

Describe 'Replacements and IT tickets' {
    BeforeAll {
        $out = Join-Path ([IO.Path]::GetTempPath()) "onboarding-tickets-$([guid]::NewGuid())"
        & (Join-Path $root 'Start-Onboarding.ps1') -AsOf $asOf -OutputPath $out -Reset -NoDashboard -Quiet | Out-Null
        $log = Get-Content (Join-Path $out 'run_log.json') -Raw | ConvertFrom-Json
        $tickets = @($log.tickets)
        function TicketsFor($id) { @($tickets | Where-Object employee_id -eq $id) }
    }
    AfterAll { Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue }

    It "re-issues the previous person's returned equipment when it's still young enough" {
        $t = TicketsFor 'E10421'
        ($t | Where-Object type -eq 'Re-image and re-issue').asset_tag | Should -Be @('AT-002101', 'AT-002102')
    }
    It 'retires equipment past its lifecycle and sets up a spare instead' {
        $t = TicketsFor 'E10422'
        ($t | Where-Object type -eq 'Retire device').asset_tag | Should -Be 'AT-001877'
        $spare = $t | Where-Object type -eq 'Image from stock'
        $spare.asset_tag | Should -Be 'AT-002301'
        $spare.steps -join ' ' | Should -BeLike '*Ship from Los Angeles HQ to Long Beach*'
    }
    It 'recovers equipment from someone still working, after their last day' {
        $recover = TicketsFor 'E10423' | Where-Object type -eq 'Recover equipment'
        $recover.asset_tag | Should -Be 'AT-002233'
        $recover.due | Should -Be '2026-11-02'    # first business day after Oct 30
        (TicketsFor 'E10423' | Where-Object type -eq 'Image from stock').asset_tag | Should -Be 'AT-002320'
    }
    It 'flags unreturned equipment from someone who already left as high priority' {
        $recover = TicketsFor 'E10425' | Where-Object type -eq 'Recover equipment'
        $recover.priority | Should -Be 'High'
        $recover.due | Should -Be '2026-10-20'
    }
    It 'never re-issues personal items like headsets' {
        ($tickets | Where-Object { $_.device -eq 'Headset' }).type | Sort-Object -Unique | Should -Be 'Issue from stock'
    }
    It 'orders a new device when nothing is available' {
        (TicketsFor 'E10428').type | Should -Be 'Order new device'
    }
    It 'never gives the same device to two people' {
        $given = $tickets | Where-Object { $_.type -in 'Re-image and re-issue', 'Image from stock', 'Issue from stock' } | ForEach-Object asset_tag
        ($given | Sort-Object -Unique).Count | Should -Be $given.Count
    }
    It 'lists access the previous person had that the new hire does not get' {
        ($log.hires | Where-Object employee_id -eq 'E10422').replacing.access_gap | Should -Be @('AP-Approvers', 'NetSuite-Admins')
    }
    It 'sends a replacement that does not exist back to HR' {
        ($log.hires | Where-Object employee_id -eq 'E10436').issues -join ' ' | Should -BeLike "*Replacing employee 'E10125' was not found*"
    }
    It 'puts every ticket on the checklist' {
        $h = $log.hires | Where-Object employee_id -eq 'E10430'
        $text = ($h.checklist_data.sections.items.text) -join ' '
        foreach ($id in $h.tickets) { $text | Should -BeLike "*$id*" }
    }
    It 'reserves devices so the next run cannot hand them out again' {
        $inv = Import-Csv (Join-Path $out 'practice_inventory.csv')
        ($inv | Where-Object AssetTag -eq 'AT-002301').Status | Should -Be 'Reserved'
        ($inv | Where-Object AssetTag -eq 'AT-001877').Status | Should -Be 'Pending retirement'
    }
    It 'counts business days, skipping weekends' {
        Get-BusinessDaysBefore -Date '2026-10-26' -Days 2 -NotBefore '2026-10-01' | Should -Be ([datetime]'2026-10-22')
    }
}
