#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Create a list of all the user accounts found in a specific OU within AD.

    .DESCRIPTION
        Report all the users accounts found in an organizational unit within
        active directory. Check if the found user is a member of the group names
        in the import file. Send the result by mail in an Excel sheet.

    .PARAMETER OU
        One or more organizational units in the active directory.

    .PARAMETER GroupName
        One or more active directory group names. Every user account
        will be checked for group membership and an extra column will be added
        to the Excel sheet with the group name and a true/false value.

    .PARAMETER ImportFile
        A .json file containing the script arguments.

    .PARAMETER LogFolder
        Location for the log files.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Users all\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($adOUs = $File.AD.OU)) {
            throw "Input file '$ImportFile': No 'AD.OU' found."
        }

        if (-not ($adGroupNames = $File.AD.GroupName)) {
            throw "Input file '$ImportFile': No 'AD.GroupName' found."
        }
        #endregion

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
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
        #region Get group members
        $groupMember = @{}

        foreach (
            $groupName in
            ($adGroupNames | Sort-Object -Unique)
        ) {
            $M = "Get group members '$groupName'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $groupMember[$groupName] = Get-ADGroupMember -Identity $groupName -Recursive -EA Stop |
            Select-Object -ExpandProperty SamAccountName
        }
        #endregion

        #region Get users
        $M = "Get users in OU '$adOUs'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $adUsers = Get-ADUserHC -OU $adOUs
        #endregion

        #region Add group membership
        $M = "Add group membership"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        foreach ($user in $adUsers) {
            $groupMember.GetEnumerator().ForEach(
                {
                    $params = @{
                        InputObject       = $user
                        NotePropertyName  = $_.Name
                        NotePropertyValue = $_.Value -contains $user.'Logon name'
                    }
                    Add-Member @params
                }
            )
        }
        #endregion

        #region Export to Excel
        $excelParams = @{
            Path               = $logFile + ' - Result.xlsx'
            AutoSize           = $true
            BoldTopRow         = $true
            FreezeTopRow       = $true
            WorkSheetName      = 'Users'
            TableName          = 'Users'
            NoNumberConversion = @(
                'Employee ID', 'OfficePhone', 'HomePhone',
                'MobilePhone', 'ipPhone', 'Fax', 'Pager'
            )
            ErrorAction        = 'Stop'
        }

        $M = "Export users to Excel file '$($excelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        Remove-Item $excelParams.Path -Force -EA Ignore

        $adUsers | Select-Object -ExcludeProperty 'SmtpAddresses' -Property *,
        @{
            Name       = 'SmtpAddresses'
            Expression = {
                $_.SmtpAddresses -join ', '
            }
        } | Export-Excel @excelParams

        $mailParams.Attachments = $excelParams.Path
        #endregion

        $mailParams.Message = "A total of <b>$($adUsers.Count) user accounts</b> have been found. <p><i>* Check the attachment for details </i></p>
            $($adOUs | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:')"

        $mailParams.Subject = "$(@($adUsers).count) user accounts"

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}