# Check server uptime
$os = Get-WmiObject -Class Win32_OperatingSystem
$uptime = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
Write-Host "Server Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"

# Get disk storage information
$disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
Write-Host "Storage Space Information:"
foreach ($disk in $disks) {
    $freeSpace = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalSpace = [math]::Round($disk.Size / 1GB, 2)
    Write-Host "Drive $($disk.DeviceID): $freeSpace GB free out of $totalSpace GB total"
}

# Get last installed updates
$lastUpdate = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
Write-Host "Last Installed Update: $($lastUpdate.Description) - $($lastUpdate.HotFixID) on $($lastUpdate.InstalledOn)"

# Check application and system event logs
$logs = @("Application", "System")
$startDate = (Get-Date).AddDays(-30)
foreach ($log in $logs) {
    Write-Host "Event Log: $log"
    $events = Get-WinEvent -LogName $log -FilterXPath "*[System[TimeCreated[@SystemTime>='$($startDate.ToUniversalTime().ToString('o'))'] and (Level=2 or Level=3)]]" | Group-Object -Property ID
    $eventInfo = @()
    foreach ($event in $events) {
        $firstOccurrence = $event.Group | Sort-Object TimeCreated | Select-Object -First 1
        $lastOccurrence = $event.Group | Sort-Object TimeCreated -Descending | Select-Object -First 1
        $eventInfo += [PSCustomObject]@{
            'ID'              = $event.Name
            'Level'           = $firstOccurrence.LevelDisplayName
            'First Occurrence'= $firstOccurrence.TimeCreated
            'Last Occurrence' = $lastOccurrence.TimeCreated
            'Total'           = $event.Count
            'Message'         = $firstOccurrence.Message.Split([Environment]::NewLine, 2)[0]
        }
    }
    $eventInfo | Format-Table -AutoSize
}

Read-Host -Prompt "Done. Press any key to exit"