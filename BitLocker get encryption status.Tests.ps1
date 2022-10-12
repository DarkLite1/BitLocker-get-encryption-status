#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    Get-Job | Remove-Job -Force -EA Ignore
    
    $realCmdLet = @{
        InvokeCommand = Get-Command Invoke-Command
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = 'TestDrive:/log'
    }

    Mock Get-ADComputer
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Test-ADOuExistsHC { $true }
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams -EA ignore

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
    }
    Context 'the ImportFile' {
        BeforeAll {
            Function Remove-HashTablePropertyHC {
                Param (
                    [Parameter(Mandatory)]
                    [HashTable]$HashTable,
                    [Parameter(Mandatory)]
                    [String[]]$PropertyName
                )
            
                $PropertyName | ForEach-Object {
                    $testHashTable = $HashTable
                    $testPath = $_
                    do {
                        $keys = $testPath -split '\.', 2
                                
                        if ($keys.Count -eq 1) {
                            $testHashTable.Remove($keys[0])
                        }
                        else {
                            $testPath = $keys[1]
                            $testHashTable = $testHashTable[$keys[0]]
                        }
                    
                    } while (
                        $keys.Count -ne 1
                    )
                }        
            }
        }
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'is missing property <_>' -ForEach @(
            'AD.OU', 
            'SendMail.To',
            'SendMail.When'
        ) {
            $testJsonFile = @{
                AD       = @{
                    Property = @{
                        ToMonitor = @('Office') 
                        InReport  = @('SamAccountName', 'Office', 'Title')
                    }
                    OU       = @{
                        Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                    }
                }
                SendMail = @{
                    When = 'Always'
                    To   = 'bob@contoso.com'
                }
            }

            Remove-HashTablePropertyHC -HashTable $testJsonFile -PropertyName $_

            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property '$_' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'AD.OU contains a non existing OU' {
            $testJsonFile = @{
                AD       = @{
                    OU = @('OU=Wrong,DC=contoso,DC=com')
                }
                SendMail = @{
                    When = 'Always'
                    To   = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams
            
            Mock Test-ADOuExistsHC { $false } -ParameterFilter {
                $Name -eq $testJsonFile.AD.OU
            }
            
            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*OU 'OU=Wrong,DC=contoso,DC=com' defined in 'AD.OU' does not exist*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
    It 'no computers found in OU' {
        $testJsonFile = @{
            AD       = @{
                OU = @('OU=BEL,OU=EU,DC=contoso,DC=com')
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        . $testScript @testParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*No computers found in any of the active directory organizational units*')
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
    }
}
Describe 'when the script runs for the first time' {
    BeforeAll {
        $testData = @(
            [PSCustomObject]@{
                ComputerName = 'PC1'
                Tpm          = @{
                    TpmActivated = $true
                    TpmPresent   = $true
                    TpmEnabled   = $true
                    TpmReady     = $true
                    TpmOwned     = $true
                }
                BitLocker    = @{
                    Volumes  = @(
                        @{
                            MountPoint           = 'C:'
                            CapacityGB           = 237.0482
                            EncryptionPercentage = 100
                            VolumeStatus         = 'FullyEncrypted'
                            ProtectionStatus     = 'on'
                            LockStatus           = 'Unlocked'
                        }
                    )
                    Recovery = @(
                        @{
                            MountPoint       = 'C:'
                            ProtectorType    = 'Tpm'
                            RecoveryPassword = $null
                        },
                        @{
                            MountPoint       = 'C:'
                            ProtectorType    = 'RecoveryPassword'
                            RecoveryPassword = 'abc'
                        }
                    )
                }
                Error        = $null
                Date         = Get-Date
            }
        )
        Mock Get-ADComputer {
            [PSCustomObject]@{
                Name = $testData[0].ComputerName
            }
        }
        Mock Invoke-Command {
            & $realCmdLet.InvokeCommand -Scriptblock { 
                $using:testData
            } -AsJob -ComputerName $env:COMPUTERNAME
        }

        $testJsonFile = @{
            AD       = @{
                OU = @(
                    'OU=BEL,OU=EU,DC=contoso,DC=com',
                    'OU=NLD,OU=EU,DC=contoso,DC=com'
                )
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams -Verbose
    }
    Context 'collect all BitLocker volumes' {
        It 'call Get-ADComputer with the correct arguments' {
            Should -Invoke Get-ADComputer -Scope Describe -Times 2 -Exactly 

            Should -Invoke Get-ADComputer -Scope Describe -Times 1 -Exactly -ParameterFilter {
                ($SearchBase -eq $testJsonFile.AD.OU[0])
            }

            Should -Invoke Get-ADComputer -Scope Describe -Times 1 -Exactly -ParameterFilter {
                ($SearchBase -eq $testJsonFile.AD.OU[1])
            }
        }
        It 'call Invoke-Command with the correct arguments' {
            Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly 

            Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly -ParameterFilter {
                ($ScriptBlock) -and
                ($ComputerName -eq $testData[0].ComputerName) -and
                ($AsJob)
            }
        }
    }
    Context 'export an Excel file' {
        Context "with worksheet 'BitLockerVolumes'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        ComputerName = $testData[0].ComputerName
                        Date         = $testData[0].Date
                        Drive        = $testData[0].BitLocker.Volumes[0].MountPoint
                        Size         = '237 GB'
                        Encrypted    = '100 %'
                        VolumeStatus = $testData[0].BitLocker.Volumes[0].VolumeStatus
                        Status       = 'Protection ON (Unlocked)'
                        KeyProtector = 'Tpm, RecoveryPassword: abc'
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx'
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'BitLockerVolumes'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.ComputerName -eq $testRow.ComputerName
                    }
                    $actualRow.Drive | Should -Be $testRow.Drive
                    $actualRow.Date.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.Date.ToString('yyyyMMdd HHmm')
                    $actualRow.Size | Should -Be $testRow.Size
                    $actualRow.Encrypted | Should -Be $testRow.Encrypted
                    $actualRow.VolumeStatus | Should -Be $testRow.VolumeStatus
                    $actualRow.Status | Should -Be $testRow.Status
                    $actualRow.KeyProtector | Should -Be $testRow.KeyProtector
                }
            }
        }
        Context "with worksheet 'TPM'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        ComputerName = $testData[0].ComputerName
                        Date         = $testData[0].Date
                        Activated    = $testData[0].Tpm.TpmActivated
                        Present      = $testData[0].Tpm.TpmPresent
                        Enabled      = $testData[0].Tpm.TpmEnabled
                        Ready        = $testData[0].Tpm.TpmReady
                        Owned        = $testData[0].Tpm.TpmOwned
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx'
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'TpmStatuses'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.ComputerName -eq $testRow.ComputerName
                    }
                    $actualRow.Activated | Should -Be $testRow.Activated
                    $actualRow.Date.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.Date.ToString('yyyyMMdd HHmm')
                    $actualRow.Present | Should -Be $testRow.Present
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Ready | Should -Be $testRow.Ready
                    $actualRow.Owned | Should -Be $testRow.Owned
                }
            }
        } -Tag test
    }
    Context 'no e-mail or further action is taken' {
        It 'because there are no previous BitLocker volumes available in a previously exported Excel file' {
            Should -Not -Invoke Send-MailHC -Scope Describe 
            Should -Invoke Write-EventLog -Scope Describe -Times 1 -Exactly -ParameterFilter {
                $Message -like '*No comparison possible*'
            }
        }
    }
}
Describe 'when the script runs after a snapshot was created' {
    BeforeAll {
        $testAdUser = @(
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(1)
                CanonicalName         = 'OU=Texas,OU=USA,DC=contoso,DC=net'
                Co                    = 'USA'
                Company               = 'US Government'
                Department            = 'Texas rangers'
                Description           = 'Ranger'
                DisplayName           = 'Chuck Norris'
                DistinguishedName     = 'dis chuck'
                EmailAddress          = 'gmail@chuck.norris'
                EmployeeID            = '1'
                EmployeeType          = 'Special'
                Enabled               = $true
                ExtensionAttribute8   = '3'
                Fax                   = '2'
                GivenName             = 'Chuck'
                HomePhone             = '4'
                HomeDirectory         = 'c:\chuck'
                Info                  = "best`nguy`never"
                IpPhone               = '5'
                Surname               = 'Norris'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'President'
                MobilePhone           = '6'
                Name                  = 'Chuck Norris'
                Office                = 'Texas'
                OfficePhone           = '7'
                Pager                 = '9'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'cnorris'
                ScriptPath            = 'c:\cnorris\script.ps1'
                Title                 = 'Texas lead ranger'
                UserPrincipalName     = 'norris@world'
                WhenChanged           = (Get-Date).AddDays(-5)
                WhenCreated           = (Get-Date).AddYears(-3)
            }
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(2)
                CanonicalName         = 'OU=Tennessee,OU=USA,DC=contoso,DC=net'
                Co                    = 'America'
                Company               = 'Retired'
                Department            = 'US Army snipers'
                Description           = 'Sniper'
                DisplayName           = 'Bob Lee Swagger'
                DistinguishedName     = 'dis bob'
                EmailAddress          = 'bl@tenessee.com'
                EmployeeID            = '9'
                EmployeeType          = 'Sniper'
                Enabled               = $true
                ExtensionAttribute8   = '11'
                Fax                   = '10'
                GivenName             = 'Bob Lee'
                HomePhone             = '12'
                HomeDirectory         = 'c:\swagger'
                Info                  = "best`nsniper`nin`nthe`nworld"
                IpPhone               = '13'
                Surname               = 'Swagger'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'US President'
                MobilePhone           = '14'
                Name                  = 'Bob Lee Swagger'
                Office                = 'Tennessee'
                OfficePhone           = '15'
                Pager                 = '16'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'lswagger'
                ScriptPath            = 'c:\swagger\script.ps1'
                Title                 = 'Corporal'
                UserPrincipalName     = 'swagger@world'
                WhenChanged           = (Get-Date).AddDays(-7)
                WhenCreated           = (Get-Date).AddYears(-30)
            }
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(2)
                CanonicalName         = 'OU=London,OU=GBR,DC=contoso,DC=net'
                Co                    = 'United Kingdom'
                Company               = 'MI6'
                Department            = 'Special agent'
                Description           = 'agent 007'
                DisplayName           = 'James Bond'
                DistinguishedName     = 'dis bond'
                EmailAddress          = '007@mi6.com'
                EmployeeID            = '17'
                EmployeeType          = 'Agent'
                Enabled               = $true
                ExtensionAttribute8   = '18'
                Fax                   = '19'
                GivenName             = 'James'
                HomePhone             = '20'
                HomeDirectory         = 'c:\bond'
                Info                  = "best agent"
                IpPhone               = '21'
                Surname               = 'Bond'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'M'
                MobilePhone           = '22'
                Name                  = 'James Bond'
                Office                = 'London'
                OfficePhone           = '23'
                Pager                 = '24'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'jbond'
                ScriptPath            = 'c:\bond\script.ps1'
                Title                 = 'Commander at sea'
                UserPrincipalName     = 'bond@world'
                WhenChanged           = (Get-Date).AddDays(-90)
                WhenCreated           = (Get-Date).AddYears(-10)
            }
        )
        Mock Get-ADComputer {
            $testAdUser[0..1]
        }

        $testJsonFile = @{
            AD       = @{
                Property = @{
                    ToMonitor = @('Description', 'Title')
                    InReport  = @('*')
                }
                OU       = @{
                    Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                }
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'and a user account is removed from AD' {
        BeforeAll {
            Mock Get-ADComputer {
                $testAdUser[0]
            }

            .$testScript @testParams
        }
        Context 'export an Excel file with all current BitLocker volumes' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        AccountExpirationDate     = $testAdUser[0].AccountExpirationDate
                        Country                   = $testAdUser[0].Co
                        Company                   = $testAdUser[0].Company
                        Department                = $testAdUser[0].Department
                        Description               = $testAdUser[0].Description
                        DisplayName               = $testAdUser[0].DisplayName
                        EmailAddress              = $testAdUser[0].EmailAddress
                        EmployeeID                = $testAdUser[0].EmployeeID
                        EmployeeType              = $testAdUser[0].EmployeeType
                        Enabled                   = $testAdUser[0].Enabled
                        Fax                       = $testAdUser[0].Fax
                        FirstName                 = $testAdUser[0].GivenName
                        HeidelbergCementBillingID = $testAdUser[0].extensionAttribute8
                        HomePhone                 = $testAdUser[0].HomePhone
                        HomeDirectory             = $testAdUser[0].HomeDirectory
                        IpPhone                   = $testAdUser[0].IpPhone
                        LastName                  = $testAdUser[0].Surname
                        LastLogonDate             = $testAdUser[0].LastLogonDate
                        LockedOut                 = $testAdUser[0].LockedOut
                        Manager                   = 'manager chuck'
                        MobilePhone               = $testAdUser[0].MobilePhone
                        Name                      = $testAdUser[0].Name
                        Notes                     = 'best guy ever'
                        Office                    = $testAdUser[0].Office
                        OfficePhone               = $testAdUser[0].OfficePhone
                        OU                        = 'OU chuck'
                        Pager                     = $testAdUser[0].Pager
                        PasswordExpired           = $testAdUser[0].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[0].PasswordNeverExpires
                        SamAccountName            = $testAdUser[0].SamAccountName
                        LogonScript               = $testAdUser[0].scriptPath
                        Title                     = $testAdUser[0].Title
                        TSAllowLogon              = 'TS AllowLogon chuck'
                        TSHomeDirectory           = 'TS HomeDirectory chuck'
                        TSHomeDrive               = 'TS HomeDrive chuck'
                        TSUserProfile             = 'TS UserProfile chuck'
                        UserPrincipalName         = $testAdUser[0].UserPrincipalName
                        WhenChanged               = $testAdUser[0].WhenChanged
                        WhenCreated               = $testAdUser[0].WhenCreated
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx' | 
                Sort-Object 'CreationTime' | Select-Object -Last 1
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'AllUsers'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
        Context 'export an Excel file with the differences' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Status                    = 'REMOVED'
                        UpdatedFields             = ''
                        AccountExpirationDate     = $testAdUser[1].AccountExpirationDate
                        Country                   = $testAdUser[1].Co
                        Company                   = $testAdUser[1].Company
                        Department                = $testAdUser[1].Department
                        Description               = $testAdUser[1].Description
                        DisplayName               = $testAdUser[1].DisplayName
                        EmailAddress              = $testAdUser[1].EmailAddress
                        EmployeeID                = $testAdUser[1].EmployeeID
                        EmployeeType              = $testAdUser[1].EmployeeType
                        Enabled                   = $testAdUser[1].Enabled
                        Fax                       = $testAdUser[1].Fax
                        FirstName                 = $testAdUser[1].GivenName
                        HeidelbergCementBillingID = $testAdUser[1].extensionAttribute8
                        HomePhone                 = $testAdUser[1].HomePhone
                        HomeDirectory             = $testAdUser[1].HomeDirectory
                        IpPhone                   = $testAdUser[1].IpPhone
                        LastName                  = $testAdUser[1].Surname
                        LastLogonDate             = $testAdUser[1].LastLogonDate
                        LockedOut                 = $testAdUser[1].LockedOut
                        Manager                   = 'manager bob'
                        MobilePhone               = $testAdUser[1].MobilePhone
                        Name                      = $testAdUser[1].Name
                        Notes                     = 'best sniper in the world'
                        Office                    = $testAdUser[1].Office
                        OfficePhone               = $testAdUser[1].OfficePhone
                        OU                        = 'OU bob'
                        Pager                     = $testAdUser[1].Pager
                        PasswordExpired           = $testAdUser[1].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[1].PasswordNeverExpires
                        SamAccountName            = $testAdUser[1].SamAccountName
                        LogonScript               = $testAdUser[1].scriptPath
                        Title                     = $testAdUser[1].Title
                        TSAllowLogon              = 'TS AllowLogon bob'
                        TSHomeDirectory           = 'TS HomeDirectory bob'
                        TSHomeDrive               = 'TS HomeDrive bob'
                        TSUserProfile             = 'TS UserProfile bob'
                        UserPrincipalName         = $testAdUser[1].UserPrincipalName
                        WhenChanged               = $testAdUser[1].WhenChanged
                        WhenCreated               = $testAdUser[1].WhenCreated
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Differences.xlsx'
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Differences'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.Status | Should -Be $testRow.Status
                    $actualRow.UpdatedFields | Should -Be $testRow.UpdatedFields
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
    }
    Context 'and a user account is added to AD' {
        BeforeAll {
            Mock Get-ADComputer {
                $testAdUser[0..2]
            }

            .$testScript @testParams
        }
        Context 'export an Excel file with all current BitLocker volumes' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        AccountExpirationDate     = $testAdUser[0].AccountExpirationDate
                        Country                   = $testAdUser[0].Co
                        Company                   = $testAdUser[0].Company
                        Department                = $testAdUser[0].Department
                        Description               = $testAdUser[0].Description
                        DisplayName               = $testAdUser[0].DisplayName
                        EmailAddress              = $testAdUser[0].EmailAddress
                        EmployeeID                = $testAdUser[0].EmployeeID
                        EmployeeType              = $testAdUser[0].EmployeeType
                        Enabled                   = $testAdUser[0].Enabled
                        Fax                       = $testAdUser[0].Fax
                        FirstName                 = $testAdUser[0].GivenName
                        HeidelbergCementBillingID = $testAdUser[0].extensionAttribute8
                        HomePhone                 = $testAdUser[0].HomePhone
                        HomeDirectory             = $testAdUser[0].HomeDirectory
                        IpPhone                   = $testAdUser[0].IpPhone
                        LastName                  = $testAdUser[0].Surname
                        LastLogonDate             = $testAdUser[0].LastLogonDate
                        LockedOut                 = $testAdUser[0].LockedOut
                        Manager                   = 'manager chuck'
                        MobilePhone               = $testAdUser[0].MobilePhone
                        Name                      = $testAdUser[0].Name
                        Notes                     = 'best guy ever'
                        Office                    = $testAdUser[0].Office
                        OfficePhone               = $testAdUser[0].OfficePhone
                        OU                        = 'OU chuck'
                        Pager                     = $testAdUser[0].Pager
                        PasswordExpired           = $testAdUser[0].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[0].PasswordNeverExpires
                        SamAccountName            = $testAdUser[0].SamAccountName
                        LogonScript               = $testAdUser[0].scriptPath
                        Title                     = $testAdUser[0].Title
                        TSAllowLogon              = 'TS AllowLogon chuck'
                        TSHomeDirectory           = 'TS HomeDirectory chuck'
                        TSHomeDrive               = 'TS HomeDrive chuck'
                        TSUserProfile             = 'TS UserProfile chuck'
                        UserPrincipalName         = $testAdUser[0].UserPrincipalName
                        WhenChanged               = $testAdUser[0].WhenChanged
                        WhenCreated               = $testAdUser[0].WhenCreated
                    }
                    @{
                        AccountExpirationDate     = $testAdUser[1].AccountExpirationDate
                        Country                   = $testAdUser[1].Co
                        Company                   = $testAdUser[1].Company
                        Department                = $testAdUser[1].Department
                        Description               = $testAdUser[1].Description
                        DisplayName               = $testAdUser[1].DisplayName
                        EmailAddress              = $testAdUser[1].EmailAddress
                        EmployeeID                = $testAdUser[1].EmployeeID
                        EmployeeType              = $testAdUser[1].EmployeeType
                        Enabled                   = $testAdUser[1].Enabled
                        Fax                       = $testAdUser[1].Fax
                        FirstName                 = $testAdUser[1].GivenName
                        HeidelbergCementBillingID = $testAdUser[1].extensionAttribute8
                        HomePhone                 = $testAdUser[1].HomePhone
                        HomeDirectory             = $testAdUser[1].HomeDirectory
                        IpPhone                   = $testAdUser[1].IpPhone
                        LastName                  = $testAdUser[1].Surname
                        LastLogonDate             = $testAdUser[1].LastLogonDate
                        LockedOut                 = $testAdUser[1].LockedOut
                        Manager                   = 'manager bob'
                        MobilePhone               = $testAdUser[1].MobilePhone
                        Name                      = $testAdUser[1].Name
                        Notes                     = 'best sniper in the world'
                        Office                    = $testAdUser[1].Office
                        OfficePhone               = $testAdUser[1].OfficePhone
                        OU                        = 'OU bob'
                        Pager                     = $testAdUser[1].Pager
                        PasswordExpired           = $testAdUser[1].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[1].PasswordNeverExpires
                        SamAccountName            = $testAdUser[1].SamAccountName
                        LogonScript               = $testAdUser[1].scriptPath
                        Title                     = $testAdUser[1].Title
                        TSAllowLogon              = 'TS AllowLogon bob'
                        TSHomeDirectory           = 'TS HomeDirectory bob'
                        TSHomeDrive               = 'TS HomeDrive bob'
                        TSUserProfile             = 'TS UserProfile bob'
                        UserPrincipalName         = $testAdUser[1].UserPrincipalName
                        WhenChanged               = $testAdUser[1].WhenChanged
                        WhenCreated               = $testAdUser[1].WhenCreated
                    }
                    @{
                        AccountExpirationDate     = $testAdUser[2].AccountExpirationDate
                        Country                   = $testAdUser[2].Co
                        Company                   = $testAdUser[2].Company
                        Department                = $testAdUser[2].Department
                        Description               = $testAdUser[2].Description
                        DisplayName               = $testAdUser[2].DisplayName
                        EmailAddress              = $testAdUser[2].EmailAddress
                        EmployeeID                = $testAdUser[2].EmployeeID
                        EmployeeType              = $testAdUser[2].EmployeeType
                        Enabled                   = $testAdUser[2].Enabled
                        Fax                       = $testAdUser[2].Fax
                        FirstName                 = $testAdUser[2].GivenName
                        HeidelbergCementBillingID = $testAdUser[2].extensionAttribute8
                        HomePhone                 = $testAdUser[2].HomePhone
                        HomeDirectory             = $testAdUser[2].HomeDirectory
                        IpPhone                   = $testAdUser[2].IpPhone
                        LastName                  = $testAdUser[2].Surname
                        LastLogonDate             = $testAdUser[2].LastLogonDate
                        LockedOut                 = $testAdUser[2].LockedOut
                        Manager                   = 'manager bond'
                        MobilePhone               = $testAdUser[2].MobilePhone
                        Name                      = $testAdUser[2].Name
                        Notes                     = 'best agent'
                        Office                    = $testAdUser[2].Office
                        OfficePhone               = $testAdUser[2].OfficePhone
                        OU                        = 'OU bond'
                        Pager                     = $testAdUser[2].Pager
                        PasswordExpired           = $testAdUser[2].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[2].PasswordNeverExpires
                        SamAccountName            = $testAdUser[2].SamAccountName
                        LogonScript               = $testAdUser[2].scriptPath
                        Title                     = $testAdUser[2].Title
                        TSAllowLogon              = 'TS AllowLogon bond'
                        TSHomeDirectory           = 'TS HomeDirectory bond'
                        TSHomeDrive               = 'TS HomeDrive bond'
                        TSUserProfile             = 'TS UserProfile bond'
                        UserPrincipalName         = $testAdUser[2].UserPrincipalName
                        WhenChanged               = $testAdUser[2].WhenChanged
                        WhenCreated               = $testAdUser[2].WhenCreated
                    }
                )    
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx' | 
                Sort-Object 'CreationTime' | Select-Object -Last 1
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'AllUsers'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
        Context 'export an Excel file with the differences' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Status                    = 'ADDED'
                        UpdatedFields             = ''
                        AccountExpirationDate     = $testAdUser[2].AccountExpirationDate
                        Country                   = $testAdUser[2].Co
                        Company                   = $testAdUser[2].Company
                        Department                = $testAdUser[2].Department
                        Description               = $testAdUser[2].Description
                        DisplayName               = $testAdUser[2].DisplayName
                        EmailAddress              = $testAdUser[2].EmailAddress
                        EmployeeID                = $testAdUser[2].EmployeeID
                        EmployeeType              = $testAdUser[2].EmployeeType
                        Enabled                   = $testAdUser[2].Enabled
                        Fax                       = $testAdUser[2].Fax
                        FirstName                 = $testAdUser[2].GivenName
                        HeidelbergCementBillingID = $testAdUser[2].extensionAttribute8
                        HomePhone                 = $testAdUser[2].HomePhone
                        HomeDirectory             = $testAdUser[2].HomeDirectory
                        IpPhone                   = $testAdUser[2].IpPhone
                        LastName                  = $testAdUser[2].Surname
                        LastLogonDate             = $testAdUser[2].LastLogonDate
                        LockedOut                 = $testAdUser[2].LockedOut
                        Manager                   = 'manager bond'
                        MobilePhone               = $testAdUser[2].MobilePhone
                        Name                      = $testAdUser[2].Name
                        Notes                     = 'best agent'
                        Office                    = $testAdUser[2].Office
                        OfficePhone               = $testAdUser[2].OfficePhone
                        OU                        = 'OU bond'
                        Pager                     = $testAdUser[2].Pager
                        PasswordExpired           = $testAdUser[2].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[2].PasswordNeverExpires
                        SamAccountName            = $testAdUser[2].SamAccountName
                        LogonScript               = $testAdUser[2].scriptPath
                        Title                     = $testAdUser[2].Title
                        TSAllowLogon              = 'TS AllowLogon bond'
                        TSHomeDirectory           = 'TS HomeDirectory bond'
                        TSHomeDrive               = 'TS HomeDrive bond'
                        TSUserProfile             = 'TS UserProfile bond'
                        UserPrincipalName         = $testAdUser[2].UserPrincipalName
                        WhenChanged               = $testAdUser[2].WhenChanged
                        WhenCreated               = $testAdUser[2].WhenCreated
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Differences.xlsx'
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Differences'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.Status | Should -Be $testRow.Status
                    $actualRow.UpdatedFields | Should -Be $testRow.UpdatedFields
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
    }
    Context 'and a user account is updated in AD' {
        BeforeAll {
            $testOriginalValue = @{
                Description = $testAdUser[0].Description
                Title       = $testAdUser[0].Title
            }

            $testAdUser[0].Description = 'changed description'
            $testAdUser[0].Title = 'changed title'

            Mock Get-ADComputer {
                $testAdUser[0..1]
            }

            .$testScript @testParams
        }
        Context 'export an Excel file with all current BitLocker volumes' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        AccountExpirationDate     = $testAdUser[0].AccountExpirationDate
                        Country                   = $testAdUser[0].Co
                        Company                   = $testAdUser[0].Company
                        Department                = $testAdUser[0].Department
                        Description               = $testAdUser[0].Description
                        DisplayName               = $testAdUser[0].DisplayName
                        EmailAddress              = $testAdUser[0].EmailAddress
                        EmployeeID                = $testAdUser[0].EmployeeID
                        EmployeeType              = $testAdUser[0].EmployeeType
                        Enabled                   = $testAdUser[0].Enabled
                        Fax                       = $testAdUser[0].Fax
                        FirstName                 = $testAdUser[0].GivenName
                        HeidelbergCementBillingID = $testAdUser[0].extensionAttribute8
                        HomePhone                 = $testAdUser[0].HomePhone
                        HomeDirectory             = $testAdUser[0].HomeDirectory
                        IpPhone                   = $testAdUser[0].IpPhone
                        LastName                  = $testAdUser[0].Surname
                        LastLogonDate             = $testAdUser[0].LastLogonDate
                        LockedOut                 = $testAdUser[0].LockedOut
                        Manager                   = 'manager chuck'
                        MobilePhone               = $testAdUser[0].MobilePhone
                        Name                      = $testAdUser[0].Name
                        Notes                     = 'best guy ever'
                        Office                    = $testAdUser[0].Office
                        OfficePhone               = $testAdUser[0].OfficePhone
                        OU                        = 'OU chuck'
                        Pager                     = $testAdUser[0].Pager
                        PasswordExpired           = $testAdUser[0].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[0].PasswordNeverExpires
                        SamAccountName            = $testAdUser[0].SamAccountName
                        LogonScript               = $testAdUser[0].scriptPath
                        Title                     = $testAdUser[0].Title
                        TSAllowLogon              = 'TS AllowLogon chuck'
                        TSHomeDirectory           = 'TS HomeDirectory chuck'
                        TSHomeDrive               = 'TS HomeDrive chuck'
                        TSUserProfile             = 'TS UserProfile chuck'
                        UserPrincipalName         = $testAdUser[0].UserPrincipalName
                        WhenChanged               = $testAdUser[0].WhenChanged
                        WhenCreated               = $testAdUser[0].WhenCreated
                    }
                    @{
                        AccountExpirationDate     = $testAdUser[1].AccountExpirationDate
                        Country                   = $testAdUser[1].Co
                        Company                   = $testAdUser[1].Company
                        Department                = $testAdUser[1].Department
                        Description               = $testAdUser[1].Description
                        DisplayName               = $testAdUser[1].DisplayName
                        EmailAddress              = $testAdUser[1].EmailAddress
                        EmployeeID                = $testAdUser[1].EmployeeID
                        EmployeeType              = $testAdUser[1].EmployeeType
                        Enabled                   = $testAdUser[1].Enabled
                        Fax                       = $testAdUser[1].Fax
                        FirstName                 = $testAdUser[1].GivenName
                        HeidelbergCementBillingID = $testAdUser[1].extensionAttribute8
                        HomePhone                 = $testAdUser[1].HomePhone
                        HomeDirectory             = $testAdUser[1].HomeDirectory
                        IpPhone                   = $testAdUser[1].IpPhone
                        LastName                  = $testAdUser[1].Surname
                        LastLogonDate             = $testAdUser[1].LastLogonDate
                        LockedOut                 = $testAdUser[1].LockedOut
                        Manager                   = 'manager bob'
                        MobilePhone               = $testAdUser[1].MobilePhone
                        Name                      = $testAdUser[1].Name
                        Notes                     = 'best sniper in the world'
                        Office                    = $testAdUser[1].Office
                        OfficePhone               = $testAdUser[1].OfficePhone
                        OU                        = 'OU bob'
                        Pager                     = $testAdUser[1].Pager
                        PasswordExpired           = $testAdUser[1].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[1].PasswordNeverExpires
                        SamAccountName            = $testAdUser[1].SamAccountName
                        LogonScript               = $testAdUser[1].scriptPath
                        Title                     = $testAdUser[1].Title
                        TSAllowLogon              = 'TS AllowLogon bob'
                        TSHomeDirectory           = 'TS HomeDirectory bob'
                        TSHomeDrive               = 'TS HomeDrive bob'
                        TSUserProfile             = 'TS UserProfile bob'
                        UserPrincipalName         = $testAdUser[1].UserPrincipalName
                        WhenChanged               = $testAdUser[1].WhenChanged
                        WhenCreated               = $testAdUser[1].WhenCreated
                    }
                )    
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx' | 
                Sort-Object 'CreationTime' | Select-Object -Last 1
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'AllUsers'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
        Context 'export an Excel file with the differences' {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Status                    = 'UPDATED_AFTER'
                        UpdatedFields             = 'Description, Title'
                        AccountExpirationDate     = $testAdUser[0].AccountExpirationDate
                        Country                   = $testAdUser[0].Co
                        Company                   = $testAdUser[0].Company
                        Department                = $testAdUser[0].Department
                        Description               = $testAdUser[0].Description
                        DisplayName               = $testAdUser[0].DisplayName
                        EmailAddress              = $testAdUser[0].EmailAddress
                        EmployeeID                = $testAdUser[0].EmployeeID
                        EmployeeType              = $testAdUser[0].EmployeeType
                        Enabled                   = $testAdUser[0].Enabled
                        Fax                       = $testAdUser[0].Fax
                        FirstName                 = $testAdUser[0].GivenName
                        HeidelbergCementBillingID = $testAdUser[0].extensionAttribute8
                        HomePhone                 = $testAdUser[0].HomePhone
                        HomeDirectory             = $testAdUser[0].HomeDirectory
                        IpPhone                   = $testAdUser[0].IpPhone
                        LastName                  = $testAdUser[0].Surname
                        LastLogonDate             = $testAdUser[0].LastLogonDate
                        LockedOut                 = $testAdUser[0].LockedOut
                        Manager                   = 'manager chuck'
                        MobilePhone               = $testAdUser[0].MobilePhone
                        Name                      = $testAdUser[0].Name
                        Notes                     = 'best guy ever'
                        Office                    = $testAdUser[0].Office
                        OfficePhone               = $testAdUser[0].OfficePhone
                        OU                        = 'OU chuck'
                        Pager                     = $testAdUser[0].Pager
                        PasswordExpired           = $testAdUser[0].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[0].PasswordNeverExpires
                        SamAccountName            = $testAdUser[0].SamAccountName
                        LogonScript               = $testAdUser[0].scriptPath
                        Title                     = $testAdUser[0].Title
                        TSAllowLogon              = 'TS AllowLogon chuck'
                        TSHomeDirectory           = 'TS HomeDirectory chuck'
                        TSHomeDrive               = 'TS HomeDrive chuck'
                        TSUserProfile             = 'TS UserProfile chuck'
                        UserPrincipalName         = $testAdUser[0].UserPrincipalName
                        WhenChanged               = $testAdUser[0].WhenChanged
                        WhenCreated               = $testAdUser[0].WhenCreated
                    }
                    @{
                        Status                    = 'UPDATED_BEFORE'
                        UpdatedFields             = 'Description, Title'
                        AccountExpirationDate     = $testAdUser[0].AccountExpirationDate
                        Country                   = $testAdUser[0].Co
                        Company                   = $testAdUser[0].Company
                        Department                = $testAdUser[0].Department
                        Description               = $testOriginalValue.Description
                        DisplayName               = $testAdUser[0].DisplayName
                        EmailAddress              = $testAdUser[0].EmailAddress
                        EmployeeID                = $testAdUser[0].EmployeeID
                        EmployeeType              = $testAdUser[0].EmployeeType
                        Enabled                   = $testAdUser[0].Enabled
                        Fax                       = $testAdUser[0].Fax
                        FirstName                 = $testAdUser[0].GivenName
                        HeidelbergCementBillingID = $testAdUser[0].extensionAttribute8
                        HomePhone                 = $testAdUser[0].HomePhone
                        HomeDirectory             = $testAdUser[0].HomeDirectory
                        IpPhone                   = $testAdUser[0].IpPhone
                        LastName                  = $testAdUser[0].Surname
                        LastLogonDate             = $testAdUser[0].LastLogonDate
                        LockedOut                 = $testAdUser[0].LockedOut
                        Manager                   = 'manager chuck'
                        MobilePhone               = $testAdUser[0].MobilePhone
                        Name                      = $testAdUser[0].Name
                        Notes                     = 'best guy ever'
                        Office                    = $testAdUser[0].Office
                        OfficePhone               = $testAdUser[0].OfficePhone
                        OU                        = 'OU chuck'
                        Pager                     = $testAdUser[0].Pager
                        PasswordExpired           = $testAdUser[0].PasswordExpired
                        PasswordNeverExpires      = $testAdUser[0].PasswordNeverExpires
                        SamAccountName            = $testAdUser[0].SamAccountName
                        LogonScript               = $testAdUser[0].scriptPath
                        Title                     = $testOriginalValue.Title
                        TSAllowLogon              = 'TS AllowLogon chuck'
                        TSHomeDirectory           = 'TS HomeDirectory chuck'
                        TSHomeDrive               = 'TS HomeDrive chuck'
                        TSUserProfile             = 'TS UserProfile chuck'
                        UserPrincipalName         = $testAdUser[0].UserPrincipalName
                        WhenChanged               = $testAdUser[0].WhenChanged
                        WhenCreated               = $testAdUser[0].WhenCreated
                    }
                )
    
                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Differences.xlsx'
    
                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Differences'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.Status -eq $testRow.Status
                    }
                    $actualRow.SamAccountName | 
                    Should -Be $testRow.SamAccountName
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmm')
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.FirstName | Should -Be $testRow.FirstName
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.IpPhone | Should -Be $testRow.IpPhone
                    $actualRow.LastName | Should -Be $testRow.LastName
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmm')
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.Name | Should -Be $testRow.Name
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.OU | Should -Be $testRow.OU
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | 
                    Should -Be $testRow.PasswordNeverExpires
                    $actualRow.SamAccountName | Should -Be $testRow.SamAccountName
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.UserPrincipalName | 
                    Should -Be $testRow.UserPrincipalName
                    $actualRow.WhenChanged.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenChanged.ToString('yyyyMMdd HHmm')
                    $actualRow.WhenCreated.ToString('yyyyMMdd HHmm') | 
                    Should -Be $testRow.WhenCreated.ToString('yyyyMMdd HHmm')
                }
            }
        }
    }
}
Describe 'monitor only the requested AD properties' {
    BeforeAll {
        $testAdUser = @(
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(1)
                CanonicalName         = 'OU=Texas,OU=USA,DC=contoso,DC=net'
                Co                    = 'USA'
                Company               = 'US Government'
                Department            = 'Texas rangers'
                Description           = 'Ranger'
                DisplayName           = 'Chuck Norris'
                DistinguishedName     = 'dis chuck'
                EmailAddress          = 'gmail@chuck.norris'
                EmployeeID            = '1'
                EmployeeType          = 'Special'
                Enabled               = $true
                ExtensionAttribute8   = '3'
                Fax                   = '2'
                GivenName             = 'Chuck'
                HomePhone             = '4'
                HomeDirectory         = 'c:\chuck'
                Info                  = "best`nguy`never"
                IpPhone               = '5'
                Surname               = 'Norris'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'President'
                MobilePhone           = '6'
                Name                  = 'Chuck Norris'
                Office                = 'Texas'
                OfficePhone           = '7'
                Pager                 = '9'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'cnorris'
                ScriptPath            = 'c:\cnorris\script.ps1'
                Title                 = 'Texas lead ranger'
                UserPrincipalName     = 'norris@world'
                WhenChanged           = (Get-Date).AddDays(-5)
                WhenCreated           = (Get-Date).AddYears(-3)
            }
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(2)
                CanonicalName         = 'OU=Tennessee,OU=USA,DC=contoso,DC=net'
                Co                    = 'America'
                Company               = 'Retired'
                Department            = 'US Army snipers'
                Description           = 'Sniper'
                DisplayName           = 'Bob Lee Swagger'
                DistinguishedName     = 'dis bob'
                EmailAddress          = 'bl@tenessee.com'
                EmployeeID            = '9'
                EmployeeType          = 'Sniper'
                Enabled               = $true
                ExtensionAttribute8   = '11'
                Fax                   = '10'
                GivenName             = 'Bob Lee'
                HomePhone             = '12'
                HomeDirectory         = 'c:\swagger'
                Info                  = "best`nsniper`nin`nthe`nworld"
                IpPhone               = '13'
                Surname               = 'Swagger'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'US President'
                MobilePhone           = '14'
                Name                  = 'Bob Lee Swagger'
                Office                = 'Tennessee'
                OfficePhone           = '15'
                Pager                 = '16'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'lswagger'
                ScriptPath            = 'c:\swagger\script.ps1'
                Title                 = 'Corporal'
                UserPrincipalName     = 'swagger@world'
                WhenChanged           = (Get-Date).AddDays(-7)
                WhenCreated           = (Get-Date).AddYears(-30)
            }
        )
        
        Mock Get-ADComputer {
            $testAdUser
        }

        $testJsonFile = @{
            AD       = @{
                Property = @{
                    ToMonitor = @('Description')
                    InReport  = @('Office', 'Title')
                }
                OU       = @{
                    Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                }
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'to the Excel file with the differences' {
        BeforeAll {
            $testOriginalValue = @{
                Description = $testAdUser[0].Description
            }

            $testAdUser[0].Description = 'changed description'
            $testAdUser[1].Title = 'changed title'

            Mock Get-ADComputer {
                $testAdUser
            }

            .$testScript @testParams

            $testExportedExcelRows = @(
                @{
                    Status         = 'UPDATED_AFTER'
                    UpdatedFields  = 'Description'
                    Description    = $testAdUser[0].Description
                    Office         = $testAdUser[0].Office
                    SamAccountName = $testAdUser[0].SamAccountName
                    Title          = $testAdUser[0].Title
                }
                @{
                    Status         = 'UPDATED_BEFORE'
                    UpdatedFields  = 'Description'
                    Description    = $testOriginalValue.Description
                    Office         = $testAdUser[0].Office
                    SamAccountName = $testAdUser[0].SamAccountName
                    Title          = $testAdUser[0].Title
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Differences.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Differences'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Status -eq $testRow.Status
                }
                $actualRow.SamAccountName | 
                Should -Be $testRow.SamAccountName
                $actualRow.Description | Should -Be $testRow.Description
                $actualRow.Office | Should -Be $testRow.Office
                $actualRow.Title | Should -Be $testRow.Title
                $actualRow.UpdatedFields | Should -Be $testRow.UpdatedFields

                foreach (
                    $testProp in 
                    $actualRow.PSObject.Properties.Name 
                ) {
                    @(
                        'SamAccountName', 'Status', 'UpdatedFields',
                        'Description', 'Title', 'Office'
                    ) | 
                    Should -Contain $testProp
                }
            }
        }
    }
}
Describe 'export only the requested AD properties' {
    BeforeAll {
        $testAdUser = @(
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(1)
                CanonicalName         = 'OU=Texas,OU=USA,DC=contoso,DC=net'
                Co                    = 'USA'
                Company               = 'US Government'
                Department            = 'Texas rangers'
                Description           = 'Ranger'
                DisplayName           = 'Chuck Norris'
                DistinguishedName     = 'dis chuck'
                EmailAddress          = 'gmail@chuck.norris'
                EmployeeID            = '1'
                EmployeeType          = 'Special'
                Enabled               = $true
                ExtensionAttribute8   = '3'
                Fax                   = '2'
                GivenName             = 'Chuck'
                HomePhone             = '4'
                HomeDirectory         = 'c:\chuck'
                Info                  = "best`nguy`never"
                IpPhone               = '5'
                Surname               = 'Norris'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'President'
                MobilePhone           = '6'
                Name                  = 'Chuck Norris'
                Office                = 'Texas'
                OfficePhone           = '7'
                Pager                 = '9'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'cnorris'
                ScriptPath            = 'c:\cnorris\script.ps1'
                Title                 = 'Texas lead ranger'
                UserPrincipalName     = 'norris@world'
                WhenChanged           = (Get-Date).AddDays(-5)
                WhenCreated           = (Get-Date).AddYears(-3)
            }
        )
        
        Mock Get-ADComputer {
            $testAdUser
        }

        $testJsonFile = @{
            AD       = @{
                Property = @{
                    ToMonitor = @('Description', 'Title')
                    InReport  = @('Office')
                }
                OU       = @{
                    Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                }
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'to the Excel file with the differences' {
        BeforeAll {
            $testOriginalValue = @{
                Description = $testAdUser[0].Description
                Title       = $testAdUser[0].Title
            }

            $testAdUser[0].Description = 'changed description'
            $testAdUser[0].Title = 'changed title'

            Mock Get-ADComputer {
                $testAdUser
            }

            .$testScript @testParams

            $testExportedExcelRows = @(
                @{
                    Status         = 'UPDATED_AFTER'
                    UpdatedFields  = 'Description, Title'
                    Description    = $testAdUser[0].Description
                    Office         = $testAdUser[0].Office
                    SamAccountName = $testAdUser[0].SamAccountName
                    Title          = $testAdUser[0].Title
                }
                @{
                    Status         = 'UPDATED_BEFORE'
                    UpdatedFields  = 'Description, Title'
                    Description    = $testOriginalValue.Description
                    Office         = $testAdUser[0].Office
                    SamAccountName = $testAdUser[0].SamAccountName
                    Title          = $testOriginalValue.Title
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Differences.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Differences'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Status -eq $testRow.Status
                }
                $actualRow.SamAccountName | 
                Should -Be $testRow.SamAccountName
                $actualRow.Description | Should -Be $testRow.Description
                $actualRow.Office | Should -Be $testRow.Office
                $actualRow.Title | Should -Be $testRow.Title
                $actualRow.UpdatedFields | Should -Be $testRow.UpdatedFields

                foreach (
                    $testProp in 
                    $actualRow.PSObject.Properties.Name 
                ) {
                    @(
                        'SamAccountName', 'Status', 'UpdatedFields',
                        'Description', 'Title', 'Office'
                    ) | 
                    Should -Contain $testProp
                }
            }
        }
    }
}
Describe 'send a mail with SendMail.When set to Always when' {
    BeforeAll {
        $testAdUser = @(
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(1)
                CanonicalName         = 'OU=Texas,OU=USA,DC=contoso,DC=net'
                Co                    = 'USA'
                Company               = 'US Government'
                Department            = 'Texas rangers'
                Description           = 'Ranger'
                DisplayName           = 'Chuck Norris'
                DistinguishedName     = 'dis chuck'
                EmailAddress          = 'gmail@chuck.norris'
                EmployeeID            = '1'
                EmployeeType          = 'Special'
                Enabled               = $true
                ExtensionAttribute8   = '3'
                Fax                   = '2'
                GivenName             = 'Chuck'
                HomePhone             = '4'
                HomeDirectory         = 'c:\chuck'
                Info                  = "best`nguy`never"
                IpPhone               = '5'
                Surname               = 'Norris'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'President'
                MobilePhone           = '6'
                Name                  = 'Chuck Norris'
                Office                = 'Texas'
                OfficePhone           = '7'
                Pager                 = '9'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'cnorris'
                ScriptPath            = 'c:\cnorris\script.ps1'
                Title                 = 'Texas lead ranger'
                UserPrincipalName     = 'norris@world'
                WhenChanged           = (Get-Date).AddDays(-5)
                WhenCreated           = (Get-Date).AddYears(-3)
            }
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(2)
                CanonicalName         = 'OU=Tennessee,OU=USA,DC=contoso,DC=net'
                Co                    = 'America'
                Company               = 'Retired'
                Department            = 'US Army snipers'
                Description           = 'Sniper'
                DisplayName           = 'Bob Lee Swagger'
                DistinguishedName     = 'dis bob'
                EmailAddress          = 'bl@tenessee.com'
                EmployeeID            = '9'
                EmployeeType          = 'Sniper'
                Enabled               = $true
                ExtensionAttribute8   = '11'
                Fax                   = '10'
                GivenName             = 'Bob Lee'
                HomePhone             = '12'
                HomeDirectory         = 'c:\swagger'
                Info                  = "best`nsniper`nin`nthe`nworld"
                IpPhone               = '13'
                Surname               = 'Swagger'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'US President'
                MobilePhone           = '14'
                Name                  = 'Bob Lee Swagger'
                Office                = 'Tennessee'
                OfficePhone           = '15'
                Pager                 = '16'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'lswagger'
                ScriptPath            = 'c:\swagger\script.ps1'
                Title                 = 'Corporal'
                UserPrincipalName     = 'swagger@world'
                WhenChanged           = (Get-Date).AddDays(-7)
                WhenCreated           = (Get-Date).AddYears(-30)
            }
        )
        
        Mock Get-ADComputer {
            $testAdUser[0]
        }

        $testJsonFile = @{
            AD       = @{
                Property = @{
                    ToMonitor = @('Description', 'Title')
                    InReport  = @('Office')
                }
                OU       = @{
                    Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                }
            }
            SendMail = @{
                When = 'Always'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'no changes are detected' {
        BeforeAll {
            .$testScript @testParams

            $testMail = @{
                To       = $testJsonFile.SendMail.To
                Bcc      = $ScriptAdmin
                Priority = 'Normal'
                Subject  = 'No changes detected'
                Message  = "*<p>BitLocker volumes:*"
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeNullOrEmpty
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq $testMail.To) -and
            ($Bcc -eq $testMail.Bcc) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            (-not $Attachments) -and
            ($Message -like $testMail.Message)
            }
        }
    }
    Context 'changes are detected' {
        BeforeAll {
            Mock Get-ADComputer {
                $testAdUser[0..1]
            }

            .$testScript @testParams
            
            $currentAdUsers | Should -HaveCount 2
            $previousAdUsers | Should -HaveCount 1

            $testMail = @{
                To          = $testJsonFile.SendMail.To
                Bcc         = $ScriptAdmin
                Priority    = 'Normal'
                Subject     = '1 added, 0 updated, 0 removed'
                Message     = "*<p>BitLocker volumes:</p>*Check the attachment for details*"
                Attachments = '* - Differences.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq $testMail.To) -and
            ($Bcc -eq $testMail.Bcc) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like $testMail.Attachments) -and
            ($Message -like $testMail.Message)
            }
        }
    }
}
Describe 'with SendMail.When set to OnlyWhenResultsAreFound' {
    BeforeAll {
        $testAdUser = @(
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(1)
                CanonicalName         = 'OU=Texas,OU=USA,DC=contoso,DC=net'
                Co                    = 'USA'
                Company               = 'US Government'
                Department            = 'Texas rangers'
                Description           = 'Ranger'
                DisplayName           = 'Chuck Norris'
                DistinguishedName     = 'dis chuck'
                EmailAddress          = 'gmail@chuck.norris'
                EmployeeID            = '1'
                EmployeeType          = 'Special'
                Enabled               = $true
                ExtensionAttribute8   = '3'
                Fax                   = '2'
                GivenName             = 'Chuck'
                HomePhone             = '4'
                HomeDirectory         = 'c:\chuck'
                Info                  = "best`nguy`never"
                IpPhone               = '5'
                Surname               = 'Norris'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'President'
                MobilePhone           = '6'
                Name                  = 'Chuck Norris'
                Office                = 'Texas'
                OfficePhone           = '7'
                Pager                 = '9'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'cnorris'
                ScriptPath            = 'c:\cnorris\script.ps1'
                Title                 = 'Texas lead ranger'
                UserPrincipalName     = 'norris@world'
                WhenChanged           = (Get-Date).AddDays(-5)
                WhenCreated           = (Get-Date).AddYears(-3)
            }
            [PSCustomObject]@{
                AccountExpirationDate = (Get-Date).AddYears(2)
                CanonicalName         = 'OU=Tennessee,OU=USA,DC=contoso,DC=net'
                Co                    = 'America'
                Company               = 'Retired'
                Department            = 'US Army snipers'
                Description           = 'Sniper'
                DisplayName           = 'Bob Lee Swagger'
                DistinguishedName     = 'dis bob'
                EmailAddress          = 'bl@tenessee.com'
                EmployeeID            = '9'
                EmployeeType          = 'Sniper'
                Enabled               = $true
                ExtensionAttribute8   = '11'
                Fax                   = '10'
                GivenName             = 'Bob Lee'
                HomePhone             = '12'
                HomeDirectory         = 'c:\swagger'
                Info                  = "best`nsniper`nin`nthe`nworld"
                IpPhone               = '13'
                Surname               = 'Swagger'
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                Manager               = 'US President'
                MobilePhone           = '14'
                Name                  = 'Bob Lee Swagger'
                Office                = 'Tennessee'
                OfficePhone           = '15'
                Pager                 = '16'
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
                SamAccountName        = 'lswagger'
                ScriptPath            = 'c:\swagger\script.ps1'
                Title                 = 'Corporal'
                UserPrincipalName     = 'swagger@world'
                WhenChanged           = (Get-Date).AddDays(-7)
                WhenCreated           = (Get-Date).AddYears(-30)
            }
        )
        
        Mock Get-ADComputer {
            $testAdUser[0]
        }

        $testJsonFile = @{
            AD       = @{
                Property = @{
                    ToMonitor = @('Description', 'Title')
                    InReport  = @('Office')
                }
                OU       = @{
                    Include = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                }
            }
            SendMail = @{
                When = 'OnlyWhenResultsAreFound'
                To   = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'send no mail when there are no changes' {
        BeforeAll {
            .$testScript @testParams
        }
        It 'Send-MailHC is not called' {
            Should -Not -Invoke Send-MailHC  -Scope Context
        }
    }
    Context 'send a mail when there are changes' {
        BeforeAll {
            Mock Get-ADComputer {
                $testAdUser[0..1]
            }

            .$testScript @testParams

            $testMail = @{
                To          = $testJsonFile.SendMail.To
                Bcc         = $ScriptAdmin
                Priority    = 'Normal'
                Subject     = '1 added, 0 updated, 0 removed'
                Message     = "*<p>BitLocker volumes:</p>*Check the attachment for details*"
                Attachments = '* - Differences.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq $testMail.To) -and
            ($Bcc -eq $testMail.Bcc) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like $testMail.Attachments) -and
            ($Message -like $testMail.Message)
            }
        }
    }
}