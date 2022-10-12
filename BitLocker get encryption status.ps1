#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Scan computers to get their BitLocker encryption status.

    .DESCRIPTION
        The active directory is scanned for computer names and each computer is
        queried for its BitLocker encryption status. This data is then stored in
        an Excel file and emailed to the user.

        Computers that the script is unable to query (offline, permissions, ..) 
        are disregarded and not included in the report.

    .PARAMETER ImportFile
        Contains all the required parameters to run the script. These parameters
        are explained below and an example can be found in file 'Example.json'.

    .PARAMETER AD.OU
        Collection of organizational units in active directory where to search 
        for computer objects.

    .PARAMETER SendMail.Header
        The header to use in the e-mail sent to the end user.

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.

    .PARAMETER SendMail.When
        Determines when an e-mail is sent to the end user.
        Valid options:
        - OnlyWhenResultsAreFound : when no results are found no e-mail is sent
        - Always                  : always sent an e-mail, even when no results 
                                    are found
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
            if (
                $mailWhen -notMatch '^Always$|^OnlyWhenResultsAreFound$'
            ) {
                throw "The value '$mailWhen' in 'SendMail.When' is not supported. Only the value 'Always' or 'OnlyWhenResultsAreFound' can be used."
            }
            #endregion

            $mailParams = @{
                To        = $mailTo
                Bcc       = $ScriptAdmin
                Priority  = 'Normal'
                LogFolder = $logParams.LogFolder
                Header    = $ScriptName 
                Save      = "$logFile - Mail.html"
            }
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

        $data = @{
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

        $previousExport = @{
            BitLockerVolumes = @()
            TpmStatuses      = @()
        }

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
            ThrottleLimit = 30
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

        $excelParams = @{
            Path          = "$logFile - State.xlsx"
            WorksheetName = $null
            TableName     = $null
            AutoSize      = $true
            FreezeTopRow  = $true
            Verbose       = $false
        }

        #region Create Excel sheet 'Errors'
        [array]$bitLockerErrors = $jobResults | Where-Object { $_.Error } | 
        Select-Object -Property @{
            Name       = 'ComputerName';
            Expression = { $_.ComputerName }
        },
        'Error'

        $M = 'Found {0} error{1} querying BitLocker volumes' -f $bitLockerErrors.Count,
        $(if ($bitLockerErrors.Count -ne 1) { 's' })
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        
        if ($bitLockerErrors) {
            $excelParams.WorksheetName = $excelParams.TableName = $ExcelWorksheetName.Errors
            
            $M = "Export {0} row{1} to Excel file '{2}' worksheet '{3}'" -f 
            $bitLockerErrors.Count, 
            $(if ($bitLockerErrors.Count -ne 1) { 's' }), 
            $excelParams.Path,
            $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            
            $bitLockerErrors | Export-Excel @excelParams
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
        if (($mailWhen -eq 'Always') -or ($differencesAdUsers)) {
            $counter = @{
                currentUsers  = $data.BitLockerVolumes.Current.Count
                previousUsers = $data.BitLockerVolumes.Previous.Count
                updatedUsers  = $data.BitLockerVolumes.Updated.Count
                errors        = $Error.Count
            }

            #region Subject and Priority
            $mailParams.Subject = if (
                (
                    $counter.updatedUsers + 
                    $counter.removedUsers + 
                    $counter.addedUsers
                ) -eq 0
            ) {
                'No changes detected'
            }
            else {
                '{0} added, {1} updated, {2} removed' -f $counter.addedUsers,
                $counter.updatedUsers, $counter.removedUsers
            }

            if ($counter.errors) {
                $mailParams.Priority = 'High'
                $mailParams.Subject += ', {0} error{1}' -f $counter.errors, $(
                    if ($counter.errors -ne 1) { 's' }
                )
            }
            #endregion

            #region Create html lists
            $htmlErrorList = if ($counter.errors) {
                "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.errors, 
                $(
                    if ($counter.errors -ne 1) { 's' }
                ),
                $(
                    $Error.Exception.Message | Where-Object { $_ } | 
                    ConvertTo-HtmlListHC
                )
            }
            #endregion

            #region Send mail
            $htmlTable = "
            <table>
                <tr>
                    <th>{0}</th>
                    <td>{1}</td>
                </tr>
                <tr>
                    <th>{2}</th>
                    <td>{3}</td>
                </tr>
                <tr>
                    <th>Added</th>
                    <td>{4}</td>
                </tr>
                <tr>
                    <th>Updated</th>
                    <td>{5}</td>
                </tr>
                <tr>
                    <th>Removed</th>
                    <td>{6}</td>
                </tr>
            </table>" -f 
            $now.ToString('dd/MM/yyyy HH:mm'), $counter.currentUsers, 
            $lastExcelFile.CreationTime.ToString('dd/MM/yyyy HH:mm'), 
            $counter.previousUsers, $counter.addedUsers, $counter.updatedUsers, 
            $counter.removedUsers

            $mailParams.Message = "
            $htmlErrorList
            <p>BitLocker volumes:</p>
            $htmlTable
            {0}" -f $(
                if ($mailParams.Attachments) {
                    '<p><i>* Check the attachment for details</i></p>'
                }
            )
            
            $M = "Send mail`r`n- Header:`t{0}`r`n- To:`t`t{1}`r`n- Subject:`t{2}" -f 
            $mailParams.Header, $($mailParams.To -join ','), $mailParams.Subject
            Write-Verbose $M
            
            Get-ScriptRuntimeHC -Stop
            Send-MailHC @mailParams
            #endregion
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