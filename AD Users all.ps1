﻿#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Retrieve all the user accounts found in a specific OU within AD.

    .DESCRIPTION
        Retrieve all active directory user accounts found in a specific 
        organizational unit within the active directory.

    .PARAMETER OU
        One or more organizational units in the active directory.

    .PARAMETER GroupName
        One or more active directory group object names. Every user account 
        will be checked for group membership and an extra column will be added
        to the Excel sheet with the group name and a true/false value.

    .PARAMETER MailTo
        One or more e-mail addresses.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$OU,
    [Parameter(Mandatory)]
    [String[]]$GroupName,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Users all\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
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

        $MailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

Process {
    Try {
        #region Get group members
        foreach ($G in ($GroupName | Sort-Object -Unique)) {
            $M = "Get group members '$G'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $GroupMember = @{}
            $GroupMember[$G] = Get-ADGroupMember -Identity $G -Recursive -EA Stop |
            Select-Object -ExpandProperty SamAccountName
        }
        #endregion

        #region Get users
        $M = "Get users in OU '$OU'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        $Users = Get-ADUserHC -OU $OU
        #endregion

        #region Add group membership
        $M = "Add group membership"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        foreach ($U in $Users) {
            $GroupMember.GetEnumerator().ForEach( {
                    $AddMemberParams = @{
                        InputObject       = $U
                        NotePropertyName  = $_.Name
                        NotePropertyValue = $_.Value -contains $U.'Logon name'
                    }
                    Add-Member @AddMemberParams
                })
        }
        #endregion

        #region Export to Excel
        $ExcelParams = @{
            Path               = $LogFile + ' - Result.xlsx'
            AutoSize           = $true
            BoldTopRow         = $true
            FreezeTopRow       = $true
            WorkSheetName      = 'Users'
            TableName          = 'Users'
            NoNumberConversion = 'Employee ID', 'OfficePhone', 'HomePhone', 
            'MobilePhone', 'ipPhone', 'Fax', 'Pager'
            ErrorAction        = 'Stop'
        }

        $M = "Export users to Excel file '$($ExcelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        Remove-Item $ExcelParams.Path -Force -EA Ignore
        $Users | Select-Object -ExcludeProperty 'SmtpAddresses' -Property *, @{
            Name       = 'SmtpAddresses'
            Expression = {
                $_.SmtpAddresses -join ', '
            }
        } | Export-Excel @ExcelParams

        $MailParams.Attachments = $ExcelParams.Path
        #endregion

        $MailParams.Message = "A total of <b>$(@($Users).count) user accounts</b> have been found. <p><i>* Check the attachment for details </i></p>
            $($OU | ConvertTo-OuNameHC -OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:')"

        $MailParams.Subject = "$(@($Users).count) user accounts"

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}