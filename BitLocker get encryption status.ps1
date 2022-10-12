#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Scan computers to get their BitLocker encryption status.

    .DESCRIPTION
        The active directory is scanned for computer names and each computer is
        queried for its BitLocker encryption status. This data is then stored in
        an Excel file, combined with the previously exported data, and emailed to the user.

        Combining data from the last run with the current run allows the script
        to collect data when clients are online and consolidate it with 
        previously gathered data.

        Computers that the script is unable to query (offline, permissions, ..) 
        are disregarded and not included in the report.

    .PARAMETER ImportFile
        Contains all the required parameters to run the script. These parameters
        are explained below and an example can be found in file 'Example.json'.

    .PARAMETER AD.OU
        Collection of organizational units in active directory where to search 
        for computer objects.

    .PARAMETER SendMail.Header
        The header to use in the e-mail sent to the users. If SendMail.Header
        is not provided the ScriptName will be used.

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER SendMail.When
        Determines when an e-mail is sent to the end user.
        Valid options:
        - Never   : No e-mail is sent. Useful for collecting data with 
                    scheduled tasks and not spamming the user. 
        - Always  : always sent an e-mail to the user
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Hashtable]$ExcelWorksheetName = @{
        Volumes = 'BitLockerVolumes'
        Errors  = 'Errors'
        Tpm     = 'TpmStatuses'
    },
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\BitLocker\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    $scriptBlock = {
        try {
            $result = [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                Tpm          = $null
                BitLocker    = @{
                    Volumes  = @()
                    Recovery = @()
                }
                Error        = $null
                Date         = Get-Date
            }
    
            $result.Tpm = Get-Tpm -ErrorAction Ignore
            
            $result.BitLocker.Volumes += Get-BitLockerVolume
            
            $result.BitLocker.Recovery += Foreach (
                $volume in
                $result.BitLocker.Volumes
            ) {
            (Get-BitLockerVolume -MountPoint $volume.MountPoint).KeyProtector |
                ForEach-Object {
                    [PSCustomObject]@{
                        MountPoint       = $volume.MountPoint
                        ProtectorType    = $_.KeyProtectorType
                        RecoveryPassword = $_.RecoveryPassword
                    }
                }
            }
        }
        catch {
            $result.Error = $_
            $Error.RemoveAt(0)
        }
        finally {
            $result
        }
    }

    Try {
        $now = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                Format       = 'yyyy-MM-dd HHmmss (DayOfWeek)'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @logParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        try {
            #region Import .json file
            $M = "Import .json file '$ImportFile'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
            #endregion

            #region Test .json file
            if (-not ($adOus = $file.AD.OU)) {
                throw "Property 'AD.OU' not found."
            }
            $adOus | Where-Object { -not (Test-ADOuExistsHC -Name $_) } | 
            ForEach-Object {
                throw "OU '$_' defined in 'AD.OU' does not exist"
            }
            if (-not ($mailTo = $file.SendMail.To)) {
                throw "Property 'SendMail.To' not found."
            }
            if (-not ($mailWhen = $file.SendMail.When)) {
                throw "Property 'SendMail.When' not found."
            }
            if (-not ($sendMailHeader = $SendMail.Header)) {
                $SendMailHeader = $ScriptName
            }
            if (
                $mailWhen -notMatch '^Always$|^Never$'
            ) {
                throw "The value '$mailWhen' in 'SendMail.When' is not supported. Only the value 'Always' or 'Never' can be used."
            }
            if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
                $MaxConcurrentJobs = 30
            }
            if ($MaxConcurrentJobs -isNot [int]) {
                throw "The value '$MaxConcurrentJobs' in 'MaxConcurrentJobs' is not a number."
            }
            #endregion
        }
        catch {
            throw "Failed to import file '$ImportFile': $_"
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
Process {
    Try {
        $data = @{
            Errors           = @()
            BitLockerVolumes = @{
                Previous = @()
                Current  = @()
                Updated  = @()
            }
            TpmStatuses      = @{
                Previous = @()
                Current  = @()
                Updated  = @()
            }
        }

        $excelParams = @{
            Path          = "$logFile - State.xlsx"
            WorksheetName = $null
            TableName     = $null
            AutoSize      = $true
            FreezeTopRow  = $true
            Verbose       = $false
        }

        $mailParams = @{
            To        = $mailTo
            Bcc       = $ScriptAdmin
            Priority  = 'Normal'
            LogFolder = $logParams.LogFolder
            Header    = $sendMailHeader
            Save      = "$logFile - Mail.html"
        }

        #region Get AD computers
        [array]$computers = foreach ($ou in $adOus) {
            [array]$tmpComputers = Get-ADComputer -SearchBase $ou -Filter *

            $tmpComputers

            $M = "Found {0} computer{1} in OU '{2}'" -f 
            $tmpComputers.Count,
            $(if ($tmpComputers.Count -ne 1) { 's' }),
            $ou
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        }

        if (-not $computers) {
            throw "No computers found in any of the active directory organizational units: $adOus"
        }
        #endregion

        #region Get previously exported data from Excel file
        $M = "Get previously exported data from the latest Excel file in folder '{0}'" -f $logParams.LogFolder
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            LiteralPath = $logParams.LogFolder
            Filter      = '* - State.xlsx'
            File        = $true
        }
        $lastExcelFile = Get-ChildItem @params | 
        Sort-Object 'CreationTime' | Select-Object -Last 1

        if ($lastExcelFile) {
            $worksheets = Get-ExcelSheetInfo -Path $lastExcelFile.FullName
            
            #region previously exported BitLocker volumes                
            if ($worksheets.Name -contains $ExcelWorksheetName.Volumes) {
                $params = @{
                    Path          = $lastExcelFile.FullName
                    WorksheetName = $ExcelWorksheetName.Volumes
                    ErrorAction   = 'Stop'
                }
                $data.BitLockerVolumes.Previous += Import-Excel @params
    
                $M = "Found {0} BitLocker volume{1} in Excel file '{2}'" -f 
                $data.BitLockerVolumes.Previous.Count, 
                $(if ($data.BitLockerVolumes.Previous.Count -ne 1) { 's' }),
                $lastExcelFile.FullName
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M    
            }
            #endregion

            #region previously exported Tpm statuses
            if ($worksheets.Name -contains $ExcelWorksheetName.Tpm) {
                $params = @{
                    Path          = $lastExcelFile.FullName
                    WorksheetName = $ExcelWorksheetName.Tpm
                    ErrorAction   = 'Stop'
                }
                $data.TpmStatuses.Previous += Import-Excel @params
    
                $M = "Found {0} Tpm {1} in Excel file '{2}'" -f 
                $data.TpmStatuses.Previous.Count, 
                $(
                    if ($data.TpmStatuses.Previous.Count -ne 1) 
                    { 'statuses' } else { 'status' }
                ),
                $lastExcelFile.FullName
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M    
            }
            #endregion
        }
        #endregion

        #region Get current BitLocker volumes and Tpm status
        $params = @{
            ScriptBlock   = $scriptBlock
            ComputerName  = $computers.Name
            ThrottleLimit = $MaxConcurrentJobs
            AsJob         = $true
        }
        $jobs = Invoke-Command @params

        $jobResults = $jobs | Wait-Job | Receive-Job
        #endregion

        #region Remove errors

        # not interested in unhandled errors as they are connection errors
        # of clients that are offline or where we have no permissions
        $Error.Clear()
        #endregion

        #region Create Excel sheet 'Errors'
        $data.Errors += $jobResults | Where-Object { $_.Error } | 
        Select-Object -Property @{
            Name       = 'ComputerName';
            Expression = { $_.ComputerName }
        },
        'Error'

        $M = 'Found {0} error{1} querying BitLocker volumes' -f 
        $data.Errors.Count,
        $(if ($data.Errors.Count -ne 1) { 's' })
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
        if ($data.Errors) {
            $excelParams.WorksheetName = $excelParams.TableName = $ExcelWorksheetName.Errors
            
            $M = "Export {0} row{1} to Excel file '{2}' worksheet '{3}'" -f 
            $data.Errors.Count, 
            $(if ($data.Errors.Count -ne 1) { 's' }), 
            $excelParams.Path,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            
            $data.Errors | Export-Excel @excelParams
         
            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region BitLocker volumes

        #region Convert job objects
        $data.BitLockerVolumes.Current += foreach ($jobResult in $jobResults) {
            $jobResult.BitLocker.Volumes |
            Select-Object -Property @{
                Name       = 'ComputerName';
                Expression = { $jobResult.ComputerName }
            },
            @{
                Name       = 'Date';
                Expression = { $jobResult.Date }
            },
            @{
                Name       = 'Drive';
                Expression = { $_.MountPoint }
            },
            @{
                Name       = 'Size';
                Expression = { '{0} {1}' -f [math]::Round($_.CapacityGB), 'GB' }
            },
            @{
                Name       = 'Encrypted';
                Expression = { '{0} {1}' -f $_.EncryptionPercentage, '%' }
            },
            @{
                Name       = 'VolumeStatus';
                Expression = { $_.VolumeStatus }
            },
            @{
                Name       = 'Status';
                Expression = {
                    'Protection {0}{1}' -f $_.ProtectionStatus.ToUpper(), $(
                        if ($_.ProtectionStatus -eq 'On') {
                            ' ({0})' -f $_.LockStatus
                        }
                    )
                }
            },
            @{
                Name       = 'KeyProtector';
                Expression = {
                    $mountPoint = $_.MountPoint
                ($jobResult.BitLocker.Recovery | Where-Object {
                        $_.MountPoint -eq $mountPoint
                    } | ForEach-Object {
                        '{0}{1}' -f $_.ProtectorType, $(
                            if ($_.RecoveryPassword) {
                                ': {0}' -f $_.RecoveryPassword
                            }
                        )
                    }) -join ', '
                }
            }
        }
        
        $M = 'Found {0} BitLocker volume{1} on {2} computer{3}' -f 
        $data.BitLockerVolumes.Current.Count,
        $(if ($data.BitLockerVolumes.Current.Count -ne 1) { 's' }),
        $computers.Count,
        $(if ($computers.Count -ne 1) { 's' })
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Merge old and new data
        $data.BitLockerVolumes.Updated += $data.BitLockerVolumes.Current

        $data.BitLockerVolumes.Updated += $data.BitLockerVolumes.Previous.Where(
            { $data.BitLockerVolumes.Updated.ComputerName -notContains $_.ComputerName }
        )

        # remove PC's that are no longer in the OU's
        $data.BitLockerVolumes.Updated = $data.BitLockerVolumes.Updated.Where(
            { $computers.Name -contains $_.ComputerName }
        )

        $M = "BitLocker volumes:`r`n- Current: {0}`r`n- Previous: {1}`r`n- Updated: {2}" -f 
        $data.BitLockerVolumes.Current.Count,
        $data.BitLockerVolumes.Previous.Count,
        $data.BitLockerVolumes.Updated.Count
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Create updated Excel sheet
        if ($data.BitLockerVolumes.Updated) {
            $excelParams.WorksheetName = $excelParams.TableName = $ExcelWorksheetName.Volumes
            
            $M = "Export {0} row{1} to Excel file '{2}' worksheet '{3}'" -f 
            $data.BitLockerVolumes.Updated.Count, 
            $(if ($data.BitLockerVolumes.Updated.Count -ne 1) { 's' }), 
            $excelParams.Path,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $data.BitLockerVolumes.Updated | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #endregion

        #region TPM statuses
        
        #region Convert job objects
        $data.TpmStatuses.Current += foreach (
            $jobResult in 
            $jobResults | Where-Object { $_.Tpm } 
        ) {
            $jobResult | Select-Object -Property @{
                Name       = 'ComputerName';
                Expression = { $_.ComputerName }
            },
            @{
                Name       = 'Date';
                Expression = { $jobResult.Date }
            },
            @{
                Name       = 'Activated'
                Expression = { $_.Tpm.TpmActivated }
            },
            @{
                Name       = 'Present'
                Expression = { $_.Tpm.TpmPresent }
            },
            @{
                Name       = 'Enabled'
                Expression = { $_.Tpm.TpmEnabled }
            },
            @{
                Name       = 'Ready'
                Expression = { $_.Tpm.TpmReady }
            },
            @{
                Name       = 'Owned'
                Expression = { $_.Tpm.TpmOwned }
            }
        }
        
        $M = 'Found {0} TPM {1} on {2} computer{3}' -f 
        $data.TpmStatuses.Current.Count,
        $(
            if ($data.TpmStatuses.Current.Count -ne 1) 
            { 'statuses' }else { 'status' }
        ),
        $computers.Count,
        $(if ($computers.Count -ne 1) { 's' })
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion        

        #region Merge old and new data
        $data.TpmStatuses.Updated += $data.TpmStatuses.Current

        $data.TpmStatuses.Updated += $data.TpmStatuses.Previous.Where(
            { $data.TpmStatuses.Updated.ComputerName -notContains $_.ComputerName }
        )

        # remove PC's that are no longer in the OU's
        $data.TpmStatuses.Updated = $data.TpmStatuses.Updated.Where(
            { $computers.Name -contains $_.ComputerName }
        )

        $M = "TPM statuses:`r`n- Current: {0}`r`n- Previous: {1}`r`n- Updated: {2}" -f 
        $data.TpmStatuses.Current.Count,
        $data.TpmStatuses.Previous.Count,
        $data.TpmStatuses.Updated.Count
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Create updated Excel sheet
        if ($data.TpmStatuses.Updated) {
            $excelParams.WorksheetName = $excelParams.TableName = $ExcelWorksheetName.Tpm
            
            $M = "Export {0} row{1} to Excel file '{2}' worksheet '{3}'" -f 
            $data.TpmStatuses.Updated.Count, 
            $(if ($data.TpmStatuses.Updated.Count -ne 1) { 's' }), 
            $excelParams.Path,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $data.TpmStatuses.Updated | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
End {
    Try {
        if ($mailWhen -eq 'Always') {
            $counter = @{
                currentUsers  = $data.BitLockerVolumes.Current.Count
                previousUsers = $data.BitLockerVolumes.Previous.Count
                updatedUsers  = $data.BitLockerVolumes.Updated.Count
                errors        = $data.Errors.Count
            }

            #region Subject and Priority
            $mailParams.Subject = '{0} BitLocker volume{1}' -f
            $data.BitLockerVolumes.Updated.Count,
            $(
                if ($data.BitLockerVolumes.Updated.Count -ne 1) { 's' }
            )

            if ($data.Errors.Count) {
                $mailParams.Priority = 'High'
                $mailParams.Subject += ', {0} error{1}' -f 
                $data.Errors.Count, 
                $(
                    if ($data.Errors.Count -ne 1) { 's' }
                )
            }
            #endregion

            #region Create HTML table
            $htmlTable = "
            <table>
                <tr colspan=`"2`">
                    <th>BitLocker volumes</th>
                </tr>
                <tr>
                    <td>Total</td>
                    <td>$($data.BitLockerVolumes.Updated.Count)</td>
                </tr>
                <tr>
                    <td>Previous export</td>
                    <td>$($data.BitLockerVolumes.Previous.Count)</td>
                </tr>
                <tr colspan=`"2`">
                    <th>TPM statuses</th>
                </tr>
                <tr>
                    <td>Total</td>
                    <td>$($data.TpmStatuses.Updated.Count)</td>
                </tr>
                <tr>
                    <td>Previous export</td>
                    <td>$($data.TpmStatuses.Previous.Count)</td>
                </tr>
                <tr colspan=`"2`">
                    <th>Errors</th>
                </tr>
                <tr>
                    <td>Total</td>
                    <td>$($data.Errors.Count)</td>
                </tr>
            </table>" 
            #endregion

            #region Send mail
            $mailParams.Message = "
            <p>Scan the hard drives of computers in active directory for their BitLocker and TPM status.</p>
            $htmlTable
            {0}{1}" -f 
            $(
                if ($mailParams.Attachments) {
                    '<p><i>* Check the attachment for details</i></p>'
                }
            ),
            $(
                $adOus | ConvertTo-OuNameHC -OU | Sort-Object |
                ConvertTo-HtmlListHC -Header 'Organizational units:'
            )

            $M = "Send mail`r`n- Header:`t{0}`r`n- To:`t`t{1}`r`n- Subject:`t{2}" -f 
            $mailParams.Header, $($mailParams.To -join ','), $mailParams.Subject
            Write-Verbose $M
            
            Get-ScriptRuntimeHC -Stop
            Send-MailHC @mailParams
            #endregion
        }
        else {
            $M = "No e-mail is sent because 'Mail.When = $mailWhen'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}