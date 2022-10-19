#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
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
            'SendMail.To'
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

            .$testScript @testParams -SendMail
                        
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
        It 'Jobs.MaxConcurrent is not a number' {
            Mock Test-ADOuExistsHC { $true }

            $testJsonFile = @{
                AD       = @{
                    OU = @('OU=EU,DC=contoso,DC=com')
                }
                SendMail = @{
                    When = 'Always'
                    To   = 'bob@contoso.com'
                }
                Jobs     = @{
                    MaxConcurrent = 'a'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams
            
            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*The value 'a' in 'Jobs.MaxConcurrent' is not a number*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        It 'Jobs.TimeOutInMinutes is not a number' {
            Mock Test-ADOuExistsHC { $true }

            $testJsonFile = @{
                AD       = @{
                    OU = @('OU=EU,DC=contoso,DC=com')
                }
                SendMail = @{
                    When = 'Always'
                    To   = 'bob@contoso.com'
                }
                Jobs     = @{
                    MaxConcurrent    = 5
                    TimeOutInMinutes = 'a'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams
            
            .$testScript @testParams
                        
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*The value 'a' in 'Jobs.TimeOutInMinutes' is not a number*")
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
            $testData
        }

        $testJsonFile = @{
            AD       = @{
                OU = @(
                    'OU=BEL,OU=EU,DC=contoso,DC=com'
                )
            }
            SendMail = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams -SendMail
    }
    Context 'collect all BitLocker volumes' {
        It 'call Get-ADComputer with the correct arguments' {
            Should -Invoke Get-ADComputer -Scope Describe -Times 1 -Exactly 

            Should -Invoke Get-ADComputer -Scope Describe -Times 1 -Exactly -ParameterFilter {
                ($SearchBase -eq $testJsonFile.AD.OU[0])
            }
        }
        It 'call Invoke-Command with the correct arguments' {
            Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly 

            Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly -ParameterFilter {
                ($ScriptBlock) -and
                ($ComputerName -eq $testData[0].ComputerName) -and
                ($ErrorAction -eq 'Ignore')
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
                        Size         = '237'
                        Encrypted    = '1' 
                        # Excel stores percentages divided by 100 '100 %'
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
        }
    }
    Context "an e-mail is sent when the switch 'SendMail' is used" {
        BeforeAll {
            $testMail = @{
                Header      = $testParams.ScriptName
                To          = $testJsonFile.SendMail.To
                Bcc         = $ScriptAdmin
                Priority    = 'Normal'
                Subject     = '1 BitLocker volume'
                Message     = "*<p>Scan the hard drives of computers in active directory for their BitLocker and TPM status.</p>*"
                Attachments = '*.xlsx'
            }
        }
        It 'Send-MailHC is called with the correct arguments' {
            $mailParams.Header | Should -Be $testMail.Header
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called once' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Header -eq $testMail.Header) -and
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
Describe 'when the script' {
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
            $testData
        }

        $testJsonFile = @{
            AD       = @{
                OU = @(
                    'OU=BEL,OU=EU,DC=contoso,DC=com'
                )
            }
            SendMail = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        .$testScript @testParams
    }
    It "no e-mail is sent when the switch 'SendMail' is not used" {
        Should -Not -Invoke Send-MailHC -Exactly 1 -Scope Describe
    }
    Context 'runs again after the first run' {
        BeforeAll {
            $testDataNew = @(
                [PSCustomObject]@{
                    ComputerName = 'PC2'
                    Tpm          = @{
                        TpmActivated = $true
                        TpmPresent   = $false
                        TpmEnabled   = $true
                        TpmReady     = $true
                        TpmOwned     = $true
                    }
                    BitLocker    = @{
                        Volumes  = @(
                            @{
                                MountPoint           = 'D:'
                                CapacityGB           = 200
                                EncryptionPercentage = 50
                                VolumeStatus         = 'FullyEncrypted'
                                ProtectionStatus     = 'on'
                                LockStatus           = 'Unlocked'
                            }
                        )
                        Recovery = @(
                            @{
                                MountPoint       = 'D:'
                                ProtectorType    = 'Tpm'
                                RecoveryPassword = $null
                            },
                            @{
                                MountPoint       = 'D:'
                                ProtectorType    = 'RecoveryPassword'
                                RecoveryPassword = 'xyz'
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
                [PSCustomObject]@{
                    Name = $testDataNew[0].ComputerName
                }
            }
            Mock Invoke-Command {
                $testDataNew
            }

            $testJsonFile = @{
                AD       = @{
                    OU = @(
                        'OU=BEL,OU=EU,DC=contoso,DC=com'
                    )
                }
                SendMail = @{
                    To = 'bob@contoso.com'
                }
            }
            $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

            .$testScript @testParams -SendMail
            
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - State.xlsx' | 
            Sort-Object 'CreationTime' | Select-Object -Last 1
        }
        Context 'collect all BitLocker volumes' {
            It 'call Get-ADComputer with the correct arguments' {
                Should -Invoke Get-ADComputer -Scope Describe -Times 2 -Exactly 
    
                Should -Invoke Get-ADComputer -Scope Describe -Times 2 -Exactly -ParameterFilter {
                    ($SearchBase -eq $testJsonFile.AD.OU[0])
                }
            }
            It 'call Invoke-Command with the correct arguments' {
                Should -Invoke Invoke-Command -Scope Describe -Times 2 -Exactly 
    
                Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly -ParameterFilter {
                    ($ScriptBlock) -and
                    ($ComputerName[0] -eq $testData[0].ComputerName) -and
                    (-not $ComputerName[1]) -and
                    ($ErrorAction -eq 'Ignore')
                }
    
                Should -Invoke Invoke-Command -Scope Describe -Times 1 -Exactly -ParameterFilter {
                    ($ScriptBlock) -and
                    ($ComputerName[0] -eq $testData[0].ComputerName) -and
                    ($ComputerName[1] -eq $testDataNew[0].ComputerName) -and
                    ($ErrorAction -eq 'Ignore')
                }
            }
        }
        Context 'export a merged Excel file with previous and new data' {
            Context "with worksheet 'BitLockerVolumes'" {
                BeforeAll {
                    $testExportedExcelRows = @(
                        @{
                            ComputerName = $testData[0].ComputerName
                            Date         = $testData[0].Date
                            Drive        = $testData[0].BitLocker.Volumes[0].MountPoint
                            Size         = '237'
                            Encrypted    = '1' 
                            # Excel stores percentages divided by 100 '100 %'
                            VolumeStatus = $testData[0].BitLocker.Volumes[0].VolumeStatus
                            Status       = 'Protection ON (Unlocked)'
                            KeyProtector = 'Tpm, RecoveryPassword: abc'
                        },
                        @{
                            ComputerName = $testDataNew[0].ComputerName
                            Date         = $testDataNew[0].Date
                            Drive        = $testDataNew[0].BitLocker.Volumes[0].MountPoint
                            Size         = '200'
                            Encrypted    = '0.5'
                            # Excel stores percentages divided by 100 '50 %'
                            VolumeStatus = $testDataNew[0].BitLocker.Volumes[0].VolumeStatus
                            Status       = 'Protection ON (Unlocked)'
                            KeyProtector = 'Tpm, RecoveryPassword: xyz'
                        }
                    )
        
                    $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'BitLockerVolumes'
                }
                It 'to the log folder' {
                    $testExcelLogFile | Should -Not -BeNullOrEmpty
                }
                It 'with the correct total rows' {
                    $actual | Should -HaveCount $testExportedExcelRows.Count
                } -Tag test
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
                        },
                        @{
                            ComputerName = $testDataNew[0].ComputerName
                            Date         = $testDataNew[0].Date
                            Activated    = $testDataNew[0].Tpm.TpmActivated
                            Present      = $testDataNew[0].Tpm.TpmPresent
                            Enabled      = $testDataNew[0].Tpm.TpmEnabled
                            Ready        = $testDataNew[0].Tpm.TpmReady
                            Owned        = $testDataNew[0].Tpm.TpmOwned
                        }
                    )
        
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
            }
        }
        Context "an e-mail is sent when the switch 'SendMail' is used" {
            BeforeAll {
                $testMail = @{
                    Header      = $testParams.ScriptName
                    To          = $testJsonFile.SendMail.To
                    Bcc         = $ScriptAdmin
                    Priority    = 'Normal'
                    Subject     = '2 BitLocker volumes'
                    Message     = "*<p>Scan the hard drives of computers in active directory for their BitLocker and TPM status.</p>*
                    *<th*>BitLocker volumes</th>*
                    *<td>Total</td>*2*
                    *<td>Previous export</td>*1*
                    *<th*>TPM statuses</th>*
                    *<td>Total</td>*2*
                    *<td>Previous export</td>*1*
                    *<th*>Errors</th>*
                    *<td>Total</td>*0*Check the attachment for details*"
                    Attachments = $testExcelLogFile.FullName
                }
            }
            It 'Send-MailHC is called with the correct arguments' {
                $mailParams.Header | Should -Be $testMail.Header
                $mailParams.To | Should -Be $testMail.To
                $mailParams.Bcc | Should -Be $testMail.Bcc
                $mailParams.Priority | Should -Be $testMail.Priority
                $mailParams.Subject | Should -Be $testMail.Subject
                $mailParams.Message | Should -BeLike $testMail.Message
                $mailParams.Attachments | Should -Be $testMail.Attachments
            }
            It 'Send-MailHC is called once' {
                Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                    ($Header -eq $testMail.Header) -and
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
} -Tag test
