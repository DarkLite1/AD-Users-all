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

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$OU,
    [Parameter(Mandatory)]
    [String[]]$GroupName,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        $LogParams = @{
            LogFolder = New-FolderHC -Path $LogFolder -ChildPath "AD Reports\AD users all\$ScriptName"
            Name      = $ScriptName
            Date      = 'ScriptStartTime'
        }
        $LogFile = New-LogFileNameHC @LogParams

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
        $GroupMember = @{}

        foreach ($G in ($GroupName | Sort-Object -Unique)) {
            $GroupMember[$G] = Get-ADGroupMember -Identity $G -Recursive -EA Stop |
            Select-Object -ExpandProperty SamAccountName
        }
        #endregion

        $Users = Get-ADUserHC -OU $OU

        #region Add group membership
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
        Remove-Item $ExcelParams.Path -Force -EA Ignore
        $Users | Export-Excel @ExcelParams

        $MailParams.Attachments = $ExcelParams.Path

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