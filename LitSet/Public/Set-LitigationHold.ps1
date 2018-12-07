function Set-LitigationHold {
    <#
    .SYNOPSIS
    Sets litigation hold for all mailboxes

    .DESCRIPTION
    Sets litigation hold for all mailboxes and reports on results

    .PARAMETER LogFile
    Log file created during the process
    Once the process is complete it is renamed with a time stamp and moved to the Archive Log Path

    .PARAMETER LogFilePath
    The location of the log file created when the script runs.  This is just the path e.g. c:\results

    .PARAMETER ArchiveLogPath
    Where the log files are moved.  This is just the path e.g. c:\results\archive

    .PARAMETER Owner
    This is for logging within 365. The litigation hold will be stamped with the owner you specify here.
    Use email address format e.g. boss@contoso.com

    .PARAMETER ReplaceCredentials
    Use this switch to replace the office 365 credentials stored to connect and run this script
    You will be prompted to enter username and password.  Use this command to do so as an example:
    Set-LitigationHold -ReplaceCredentials -LogFilePath C:\results

    .EXAMPLE
    Set-LitigationHold -LogFile litlog.csv -LogFilePath C:\results -ArchiveLogPath C:\results\archive -Owner boss@contoso.com -Verbose

    .NOTES
    It is necessary to create the directory structure. For the examples mentioned here, create these 3 directories:
    c:\results
    c:\results\archive
    c:\results\SS

    #>
    param
    (

        [Parameter()]
        [string] $LogFile,

        [Parameter(Mandatory)]
        [string] $LogFilePath,

        [Parameter()]
        [string] $ArchiveLogPath,

        [Parameter()]
        [string] $Owner,

        [Parameter()]
        [switch] $ReplaceCredentials

    )

    $CurrentErrorActionPref = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

    $Time = Get-Date -Format "yyyy-MM-dd-HHmm"

    $LogFilePath = $LogFilePath.Trim('\')

    if (-not $ReplaceCredentials) {
        $User = Get-Content -Path ("{0}\SS\User.txt" -f $LogFilePath)
        $Pass = Get-Content -Path ("{0}\SS\Pass.txt" -f $LogFilePath) | ConvertTo-SecureString
        $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $Pass
    }
    else {
        $Credential = Get-Credential -Message "ENTER USERNAME & PASSWORD FOR OFFICE 365/AZURE AD"
        $Credential.UserName | Out-File ("{0}\SS\User.txt" -f $LogFilePath) -Force
        $Credential.Password | ConvertFrom-SecureString | Out-File ("{0}\SS\Pass.txt" -f $LogFilePath) -Force
        break
    }

    $ErrorLogFile = ('Error_{0}') -f $LogFile

    $Log = Join-Path $LogFilePath $LogFile
    $ErrorLog = Join-Path $LogFilePath $ErrorLogFile

    Write-Log -Log $Log -AddToLog ("Script executed at {0} " -f $Time)
    Start-Transcript -Path ("{0}\Transcript\Transcript_{1:yyyyMMddhhmm}.log" -f $LogFilePath, $Time)

    $Connect = @{
        Name              = "LitScript"
        ConfigurationName = "Microsoft.Exchange"
        ConnectionUri     = "https://outlook.office365.com/powershell"
        Credential        = $Cred
        Authentication    = "Basic"
        AllowRedirection  = $true
        Verbose           = $false
    }

    $EXOSession = New-PSSession @Connect
    Import-Module (Import-PSSession $EXOSession -AllowClobber -WarningAction SilentlyContinue -Verbose:$false) -Global -Verbose:$false | Out-Null

    Connect-MsolService -Credential $Cred -Verbose:$false

    $HardSplat = @{
        ResultSize           = "Unlimited"
        RecipientTypeDetails = 'UserMailbox'
        ErrorAction          = 'Stop'
    }
    $SoftSplat = @{
        ResultSize           = "Unlimited"
        RecipientTypeDetails = 'UserMailbox'
        SoftDeletedMailbox   = $true
        ErrorAction          = 'Stop'
    }

    try {
        $Hard = Get-Mailbox @HardSplat
    }
    catch {
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: $Hard = Get-Mailbox @HardSplat  (Full Error Below)'
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
    }

    try {
        $Soft = Get-Mailbox @SoftSplat
    }
    catch {
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: $Soft = Get-Mailbox @SoftSplat  (Full Error Below)'
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
    }


    Write-Log -Log $Log -AddToLog '=========================================================================================='
    Write-Log -Log $Log -AddToLog "Report: All mailboxes (UserMailbox, SharedMailbox, RoomMailbox & EquipmentMailbox)"

    $SharedSplat = @{
        ResultSize           = "Unlimited"
        RecipientTypeDetails = 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox'
        ErrorAction          = 'Stop'
    }

    try {
        $Shared = Get-Mailbox @SharedSplat
        $SharedLicensed = $Shared | Where-Object {
            (Get-MsolUser -ObjectId $_.ExternalDirectoryObjectId).IsLicensed -eq $True
        }
    }
    catch {
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: $Shared = Get-Mailbox @SharedSplat (Full Error Below)'
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
    }
    try {
        $SharedNoLicense = $Shared | Where-Object {
            (Get-MsolUser -ObjectId $_.ExternalDirectoryObjectId).IsLicensed -eq $False
        }
        $SharedNoLicenseCount = $SharedNoLicense.guid.count
    }
    catch {
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: (Get-MsolUser -ObjectId $_.ExternalDirectoryObjectId).IsLicensed -eq $False (Full Error Below)'
        Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
        Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
    }

    $All = $Hard + $Soft
    $AllLitCount = ($All.LitigationHoldEnabled -eq $true).count
    $AllNoLitCount = ($All.LitigationHoldEnabled -eq $false).count
    $AllCount = $All.count + $Shared.count

    $SharedProps = $SharedLicensed | select PrimarySmtpAddress, guid, LitigationHoldEnabled, LitigationHoldDuration
    $SharedLit = $SharedLicensed | Where-Object {$_.LitigationHoldEnabled -eq "$true"}
    $SharedLitCount = $SharedLit.guid.count

    $SharedNoLit = $SharedLicensed | Where-Object {$_.LitigationHoldEnabled -eq "$false"}
    $SharedNoLitCount = $SharedNoLit.guid.count

    Write-Log -Log $Log -AddToLog ("`tFound {0} Mailboxes - Total" -f $AllCount)
    Write-Log -Log $Log -AddToLog ("`tFound {0} Mailboxes - (not soft-deleted)" -f $Hard.Count)
    Write-Log -Log $Log -AddToLog ("`tFound {0} Mailboxes - (soft-deleted)" -f $Soft.Count)
    Write-Log -Log $Log -AddToLog ("`tFound {0} Mailboxes - (shared/resource)" -f $Shared.Count)
    Write-Log -Log $Log -AddToLog '=========================================================================================='

    Write-Log -Log $Log -AddToLog "Report: Litigation Hold on Licensed Mailboxes"
    Write-Log -Log $Log -AddToLog ("`tFound {0} litigation hold enabled" -f ($AllLitCount + $SharedLitCount))
    Write-Log -Log $Log -AddToLog ("`tFound {0} litigation hold disabled" -f (($AllNoLitCount + $SharedNoLitCount) - $SharedNoLicenseCount))
    Write-Log -Log $Log -AddToLog '=========================================================================================='
    Write-Log -Log $Log -AddToLog 'Action: Set litigation hold and/or set duration to unlimited'

    $SetHard = $Hard | Where-Object {
        $_.LitigationHoldEnabled -eq $false -or
        $_.LitigationHoldDuration -ne "Unlimited"
    }

    $SetSoft = $Soft | Where-Object {
        $_.LitigationHoldEnabled -eq $false -or
        $_.LitigationHoldDuration -ne "Unlimited"
    }

    $SetShared = $SharedProps | Where-Object {
        $_.LitigationHoldEnabled -eq $false -or
        $_.LitigationHoldDuration -ne "Unlimited"
    }

    $SetHardSplat = @{
        LitigationHoldEnabled  = $True
        LitigationHoldDuration = "Unlimited"
        LitigationHoldOwner    = $Owner
        ErrorAction            = 'Stop'
    }

    $SetSoftSplat = @{
        LitigationHoldEnabled  = $True
        LitigationHoldDuration = "Unlimited"
        InactiveMailbox        = $True
        ErrorAction            = 'Stop'
    }

    foreach ($CurHard in $SetHard) {
        try {
            $HardMail = $CurHard.PrimarySmtpAddress
            Set-Mailbox @SetHardSplat -Identity $CurHard.guid.guid
            Write-Log -Log $Log -AddToLog ("`tSuccessfully applied litigation hold to Mailbox {0}" -f $HardMail)
        }
        catch {
            Write-Log -Log $Log -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $HardMail)
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $HardMail)
            Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: Set-Mailbox @SetHardSplat -Identity $CurHard.guid.guid (Full Error Below)'
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
        }
    }

    foreach ($CurSoft in $SetSoft) {
        try {
            $SoftMail = $CurSoft.PrimarySmtpAddress
            Set-Mailbox @SetSoftSplat -Identity $CurSoft.guid.guid
            Write-Log -Log $Log -AddToLog ("`tSuccessfully applied litigation hold to Mailbox {0}" -f $SoftMail)
        }
        catch {
            Write-Log -Log $Log -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $SoftMail)
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $SoftMail)
            Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: Set-Mailbox @SetSoftSplat -Identity $CurSoft.guid.guid (Full Error Below)'
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
        }
    }

    foreach ($CurShared in $SetShared) {
        try {
            $SharedMail = $CurShared.PrimarySmtpAddress
            Set-Mailbox @SetHardSplat -Identity $CurShared.guid.guid
            Write-Log -Log $Log -AddToLog ("`tSuccessfully applied litigation hold to Mailbox {0}" -f $SharedMail)
        }
        catch {
            Write-Log -Log $Log -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $SharedMail)
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog ("`tFAILED to apply litigation hold to Mailbox {0}" -f $SharedMail)
            Write-Log -Log $ErrorLog -AddToLog 'Error executing this line of code: Set-Mailbox @SetHardSplat -Identity $CurShared.guid.guid (Full Error Below)'
            Write-Log -Log $ErrorLog -AddToLog '=========================================================================================='
            Write-Log -Log $ErrorLog -AddToLog $_.Exception.Message
        }
    }

    Write-Log -Log $Log -AddToLog '=========================================================================================='

    $HardSplat.Filter = 'LitigationHoldEnabled -eq $true'
    $SoftSplat.Filter = 'LitigationHoldEnabled -eq $true'
    $SharedSplat.Filter = 'LitigationHoldEnabled -eq $true'
    $HardCheck = Get-Mailbox @HardSplat
    $SoftCheck = Get-Mailbox @SoftSplat
    $SharedCheck = Get-Mailbox @SharedSplat
    $SharedLicenseCheck = $SharedCheck | Where-Object {
        (Get-MsolUser -ObjectId $_.ExternalDirectoryObjectId).IsLicensed -eq $True
    }

    $AllCheckLit = $HardCheck + $SoftCheck
    $TotalLitCount = $AllCheckLit.count + $SharedLicenseCheck.guid.count

    $SharedDur = $SharedLicenseCheck | Where-Object {$_.LitigationHoldDuration -eq "Unlimited"}
    $SharedDurCount = $SharedDur.guid.count
    $AllCheckDur = $AllCheckLit | Where-Object {$_.LitigationHoldDuration -eq "Unlimited"}
    $TotalDurCount = $AllCheckDur.count + $SharedDurCount

    Write-Log -Log $Log -AddToLog ('{0} mailboxes with litigation hold enabled' -f $TotalLitCount)
    Write-Log -Log $Log -AddToLog ('{0} mailboxes with litigation hold enabled and litigation hold duration is unlimited' -f $TotalDurCount)

    $AllCheckMail = ($AllCheckLit.PrimarySmtpAddress + $SharedLicenseCheck.PrimarySmtpAddress) | Sort-Object

    foreach ($CurCheck in $AllCheckMail) {
        Write-Log -Log $Log -AddToLog ("`t{0}" -f $CurCheck)
    }

    Move-Item -Path $Log -Destination ("{0}\Archive\{1:yyyyMMddhhmm}.log" -f $LogFilePath, $Time)
    Move-Item -Path $ErrorLog -Destination ("{0}\Archive\Error_{1:yyyyMMddhhmm}.log" -f $LogFilePath, $Time) -ErrorAction SilentlyContinue

    Get-PSSession | Remove-PSSession

    Stop-Transcript

    $ErrorActionPreference = $CurrentErrorActionPref
}
