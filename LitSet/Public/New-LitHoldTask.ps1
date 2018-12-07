function New-LitHoldTask {

    $TaskSplat = @{
        TaskName      = "Office 365 Litigation Hold"
        User          = "computer\kevin"
        Executable    = "PowerShell.exe"
        Argument      = '-ExecutionPolicy RemoteSigned -Command Set-LitigationHold -LogFilePath c:\scripts\lit -LogFile LitLog.txt -Owner admin@contoso.onmicrosoft.com'
        At            = "3:27pm"
        DaysOfWeek    = "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"
        WeeksInterval = 1
    }

    Add-TaskWeekly @TaskSplat

}