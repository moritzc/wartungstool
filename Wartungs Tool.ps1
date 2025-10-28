#Requires -RunAsAdministrator

# Self-elevation for non-admin sessions
param (
    [string]$Beta = ''
)

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $admRequest = Read-Host -Prompt "You didn't run this script as an Administrator. Enter [Y] to execute as Admin"
    if ($admRequest -eq 'Y') {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
}

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.UpdateServices.Administration')

Set-StrictMode -Version Latest

# region Environment discovery
$script:HostInfo = [ordered]@{}
$script:HostInfo.HostName = $env:COMPUTERNAME
try {
    $script:HostInfo.Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name
} catch {
    $script:HostInfo.Domain = 'Workgroup'
}
$script:HostInfo.BuildNumber = [int](Get-WmiObject -Class Win32_OperatingSystem).BuildNumber
$script:HostInfo.WsusInstalled = $false
$script:HostInfo.ADInstalled = $false
$script:HostInfo.HVInstalled = $false

$featureCommand = Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue
if ($featureCommand) {
    try {
        if (-not (Get-Module -Name ServerManager)) {
            Import-Module -Name ServerManager -ErrorAction Stop
        }
        $features = Get-WindowsFeature
        $script:HostInfo.WsusInstalled = ($null -ne ($features | Where-Object { $_.Name -eq 'UpdateServices' -and $_.InstallState -eq 'Installed' }))
        $script:HostInfo.ADInstalled = ($null -ne ($features | Where-Object { $_.Name -eq 'AD-Domain-Services' -and $_.InstallState -eq 'Installed' }))
        $script:HostInfo.HVInstalled = ($null -ne ($features | Where-Object { $_.Name -eq 'Hyper-V' -and $_.InstallState -eq 'Installed' }))
    } catch {
        Write-Verbose "Unable to query Windows Features: $($_.Exception.Message)"
    }
} else {
    Write-Verbose 'Get-WindowsFeature not available; assuming WSUS, AD DS, and Hyper-V roles are not installed.'
}
# endregion

# region Output helpers
$script:OutputControl = $null

function Write-AppOutput {
    param(
        [Parameter(Mandatory)] [string]$Text,
        [object]$Color = [System.Drawing.Color]::Black,
        [switch]$Bold,
        [switch]$NewLine
    )
    if (-not $script:OutputControl) {
        return
    }

    $resolvedColor = if ($Color -is [System.Drawing.Color]) {
        $Color
    } elseif ($Color -is [string] -and $Color) {
        $named = [System.Drawing.Color]::FromName($Color)
        if ($named.IsKnownColor -or $named.IsNamedColor -or $named.IsSystemColor) {
            $named
        } else {
            [System.Drawing.Color]::Black
        }
    } else {
        [System.Drawing.Color]::Black
    }

    $fontName = $script:OutputControl.Font.Name
    $fontSize = $script:OutputControl.Font.Size
    $fontStyle = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = New-Object System.Drawing.Font($fontName, $fontSize, $fontStyle)

    $script:OutputControl.Invoke({
        param($rtb, $text, $color, $font, $appendNewLine)
        $rtb.SelectionStart = $rtb.TextLength
        $rtb.SelectionLength = 0
        $rtb.SelectionColor = $color
        $rtb.SelectionFont = $font
        $rtb.AppendText($text)
        if ($appendNewLine) {
            $rtb.AppendText([Environment]::NewLine)
        }
        $rtb.SelectionColor = $rtb.ForeColor
        $rtb.SelectionFont = $rtb.Font
        $rtb.ScrollToCaret()
    }, $script:OutputControl, $Text, $resolvedColor, $font, [bool]$NewLine)
}

function Write-AppLine {
    param(
        [Parameter(Mandatory)][string]$Text,
        [object]$Color = [System.Drawing.Color]::Black,
        [switch]$Bold
    )
    Write-AppOutput -Text $Text -Color $Color -Bold:$Bold -NewLine
}

function Write-AppError {
    param(
        [Parameter(Mandatory)][string]$Message
    )
    Write-AppLine -Text $Message -Color ([System.Drawing.Color]::Red) -Bold
}
# endregion

# region Feature functions
function Get-WSUSContentReport {
    if (-not $script:HostInfo.WsusInstalled) {
        throw 'WSUS is not installed on this server.'
    }

    $regPath = 'HKLM:\SOFTWARE\Microsoft\Update Services\Server\Setup'
    $contentFolder = $null
    if ($script:HostInfo.BuildNumber -lt 9601) {
        if ($script:HostInfo.BuildNumber -gt 9599) {
            $contentFolder = (Get-ItemProperty -Path $regPath).ContentDir
        } else {
            throw 'Only supported on Windows Server 2012 R2 or later.'
        }
    } else {
        $contentFolder = Get-ItemPropertyValue -Path $regPath -Name ContentDir
    }

    $size = 0
    if (Test-Path -LiteralPath (Join-Path $contentFolder 'WsusContent')) {
        $size = (Get-ChildItem -LiteralPath (Join-Path $contentFolder 'WsusContent') -Recurse -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum / 1GB
    }

    return [pscustomobject]@{
        ContentFolder = $contentFolder
        SizeGB        = [Math]::Round($size, 2)
    }
}

function Get-SystemDriveFreeSpace {
    param(
        [string]$DriveLetter = 'C'
    )
    $volume = Get-Volume -DriveLetter $DriveLetter
    return [Math]::Round($volume.SizeRemaining / 1GB, 2)
}

function Get-SystemUptime {
    $lastBoot = (Get-WmiObject -Class Win32_OperatingSystem).LastBootUpTime
    $uptime = (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($lastBoot)
    return $uptime
}

function Get-ADComputers {
    Get-ADComputer -Filter { OperatingSystem -Like 'Windows *' } -Property OperatingSystem, OperatingSystemVersion |
        Sort-Object -Property OperatingSystem -Descending
}

function Get-ServerUptimes {
    $servers = Get-ADComputer -Filter { OperatingSystem -Like '*Server*' } | Sort-Object OperatingSystemVersion -Descending
    foreach ($server in $servers) {
        if (Test-Connection -ComputerName $server.Name -Count 1 -TimeToLive 1 -Quiet) {
            $bootTime = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server.Name).LastBootUpTime
            $uptime = (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($bootTime)
            [pscustomobject]@{
                Name        = $server.Name
                Reachable   = $true
                LastBoot    = $bootTime
                DaysUp      = [int]$uptime.Days
            }
        } else {
            [pscustomobject]@{
                Name        = $server.Name
                Reachable   = $false
                LastBoot    = $null
                DaysUp      = $null
            }
        }
    }
}

function Get-ClientUptimes {
    $clients = Get-ADComputer -Filter { OperatingSystem -Like 'Windows*' } -Properties *
    foreach ($client in $clients) {
        if (Test-Connection -ComputerName $client.Name -Count 1 -TimeToLive 1 -Quiet) {
            $bootTime = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $client.Name).LastBootUpTime
            $uptime = (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($bootTime)
            [pscustomobject]@{
                Name      = $client.Name
                Reachable = $true
                DaysUp    = [int]$uptime.Days
            }
        } else {
            [pscustomobject]@{
                Name      = $client.Name
                Reachable = $false
                DaysUp    = $null
            }
        }
    }
}

function Get-RecentHotFixes {
    Get-HotFix | Where-Object { $_.InstalledOn -gt ((Get-Date).AddDays(-40)) } | Sort-Object -Property InstalledOn -Descending
}

function Get-WSUSSyncFailures {
    if (-not $script:HostInfo.WsusInstalled) {
        throw 'WSUS is not installed on this server.'
    }
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
    $subscription = $wsus.GetSubscription()
    $subscription.GetSynchronizationHistory() | Where-Object { $_.Result -eq 'Failed' }
}

function Get-EventOverview {
    param(
        [int]$Days = 30
    )
    $logs = @('Application', 'System')
    $startDate = (Get-Date).AddDays(-$Days)
    foreach ($log in $logs) {
        $filterTime = $startDate.ToUniversalTime().ToString('o')
        $rawEvents = Get-WinEvent -LogName $log -FilterXPath "*[System[TimeCreated[@SystemTime>='$filterTime'] and (Level=1 or Level=2 or Level=3)]]"
        $groups = $rawEvents | Group-Object -Property Id
        [pscustomobject]@{
            LogName = $log
            Total   = $rawEvents.Count
            Events  = $groups | ForEach-Object {
                $ordered = $_.Group | Sort-Object -Property TimeCreated
                $first = $ordered[0]
                $last = $ordered[-1]
                [pscustomobject]@{
                    Id              = $_.Name
                    Level           = $first.LevelDisplayName
                    FirstOccurrence = $first.TimeCreated
                    LastOccurrence  = $last.TimeCreated
                    Count           = $_.Count
                    Message         = ($first.Message -split [Environment]::NewLine)[0]
                }
            } | Sort-Object -Property Count -Descending
        }
    }
}

function Get-HyperVSnapshotReport {
    if (-not $script:HostInfo.HVInstalled) {
        throw 'Hyper-V role is not installed on this server.'
    }

    $snapshots = Get-VM | Get-VMSnapshot
    $vmHostPath = (Get-VMHost | Select-Object -ExpandProperty VirtualMachinePath).TrimEnd('\\')
    $avhdxInfo = @()
    if (Test-Path -LiteralPath $vmHostPath) {
        $vmFolders = Get-ChildItem -Path $vmHostPath -Recurse -Directory -ErrorAction Stop
        foreach ($folder in $vmFolders) {
            $files = Get-ChildItem -Path $folder.FullName -Recurse -Filter '*.avhdx' -File -ErrorAction Stop
            foreach ($file in $files) {
                $avhdxInfo += [pscustomobject]@{ Folder = $folder.FullName; File = $file.FullName }
            }
        }
    }

    [pscustomobject]@{
        Snapshots = $snapshots
        Avhdx     = $avhdxInfo
    }
}

function Get-DFSRHealth {
    $eventIds = 2212, 4012
    $currentDate = Get-Date
    $startDate = $currentDate.AddDays(-180)
    $events = @()
    foreach ($eventId in $eventIds) {
        $events += Get-WinEvent -FilterHashtable @{ LogName = 'DFS Replication'; Id = $eventId; StartTime = $startDate; EndTime = $currentDate } -ErrorAction SilentlyContinue
    }
    return $events
}

function Invoke-WSUSCleanup {
    if (-not $script:HostInfo.WsusInstalled) {
        throw 'WSUS is not installed on this server.'
    }
    $superseded = Get-WsusUpdate -Classification All -Approval Approved -Status InstalledOrNotApplicable | Where-Object { $_.Update.IsSuperseded }
    return $superseded
}

function Get-WSUSErrors {
    if (-not $script:HostInfo.WsusInstalled) {
        throw 'WSUS is not installed on this server.'
    }
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
    $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $summariesComputerFailed = $wsus.GetSummariesPerComputerTarget($updateScope, $computerScope) | Where-Object { $_.FailedCount -ne 0 } | Sort-Object FailedCount, UnknownCount, NotInstalledCount -Descending
    $computers = Get-WsusComputer
    $computerFailures = foreach ($computerFailed in $summariesComputerFailed) {
        $computer = $computers | Where-Object { $_.Id -eq $computerFailed.ComputerTargetId }
        $failedUpdates = ($wsus.GetComputerTargets($computerScope) | Where-Object { $_.Id -eq $computerFailed.ComputerTargetId }).GetUpdateInstallationInfoPerUpdate($updateScope) | Where-Object { $_.UpdateInstallationState -eq 'Failed' }
        $failedUpdateDetails = foreach ($failed in $failedUpdates) {
            $update = $wsus.GetUpdate($failed.UpdateId)
            [pscustomobject]@{
                Title  = $update.Title
                Update = $update
            }
        }
        [pscustomobject]@{
            Computer      = $computer
            Summary       = $computerFailed
            FailedUpdates = $failedUpdateDetails
        }
    }
    return [pscustomobject]@{
        ServerName = $wsus.ServerName
        Failures   = $computerFailures
    }
}

function Show-WSUSQuickReport {
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
    $computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $computers = Get-WsusComputer

    $reportForm = New-Object System.Windows.Forms.Form
    $reportForm.Text = 'WSUS Quick Report'
    $reportForm.Size = New-Object System.Drawing.Size(880,710)
    $reportForm.StartPosition = 'CenterParent'

    $comboComputers = New-Object System.Windows.Forms.ComboBox
    $comboComputers.Size = New-Object System.Drawing.Size(400,20)
    $comboComputers.Location = New-Object System.Drawing.Point(10,10)
    foreach ($computer in $computers) {
        [void]$comboComputers.Items.Add($computer.FullDomainName)
    }

    $checkInstalled = New-Object System.Windows.Forms.CheckBox
    $checkInstalled.Text = 'Installed'
    $checkInstalled.Location = New-Object System.Drawing.Point(10,40)

    $checkFailed = New-Object System.Windows.Forms.CheckBox
    $checkFailed.Text = 'Failed'
    $checkFailed.Location = New-Object System.Drawing.Point(10,65)

    $actionButton = New-Object System.Windows.Forms.Button
    $actionButton.Text = 'Action'
    $actionButton.Size = New-Object System.Drawing.Size(150,40)
    $actionButton.Location = New-Object System.Drawing.Point(10,95)

    $reportBox = New-Object System.Windows.Forms.RichTextBox
    $reportBox.Location = New-Object System.Drawing.Point(10,150)
    $reportBox.Size = New-Object System.Drawing.Size(840,500)
    $reportBox.ReadOnly = $true

    $actionButton.Add_Click({
        $reportBox.Clear()
        $selected = $comboComputers.SelectedItem
        if (-not $selected) {
            $reportBox.AppendText("Select a computer first." + [Environment]::NewLine)
            return
        }
        $target = Get-WsusComputer | Where-Object { $_.FullDomainName -eq $selected }
        if (-not $target) {
            $reportBox.AppendText("Unable to load computer details." + [Environment]::NewLine)
            return
        }
        $reportBox.AppendText("$($target.FullDomainName)`r`n$($target.Make)`r`n$($target.OSDescription)`r`n")
        $installations = $target.GetUpdateInstallationInfoPerUpdate($updateScope)
        if ($checkInstalled.Checked) {
            $installed = $installations | Where-Object { $_.UpdateInstallationState -eq 'Installed' }
            $reportBox.AppendText("`r`nInstalled Updates:`r`n")
            foreach ($item in $installed) {
                $update = $wsus.GetUpdate($item.UpdateId)
                $reportBox.AppendText("- $($update.Title)`r`n")
            }
        }
        if ($checkFailed.Checked) {
            $failed = $installations | Where-Object { $_.UpdateInstallationState -eq 'Failed' }
            $reportBox.AppendText("`r`nFailed Updates:`r`n")
            foreach ($item in $failed) {
                $update = $wsus.GetUpdate($item.UpdateId)
                $reportBox.AppendText("- $($update.Title)`r`n")
            }
        }
    })

    $reportForm.Controls.Add($comboComputers)
    $reportForm.Controls.Add($checkInstalled)
    $reportForm.Controls.Add($checkFailed)
    $reportForm.Controls.Add($actionButton)
    $reportForm.Controls.Add($reportBox)

    [void]$reportForm.ShowDialog()
}
# endregion

# region GUI creation
function New-FeatureButton {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Container,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][scriptblock]$OnClick
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Size = New-Object System.Drawing.Size(150,40)
    $button.Text = $Text
    $button.Margin = New-Object System.Windows.Forms.Padding(5)
    $button.Add_Click($OnClick)
    $Container.Controls.Add($button)
    return $button
}

function Initialize-App {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($Beta -eq 'beta') { 'Wartungstool - Beta' } else { 'Wartungstool v0.9.1.1' }
    $form.Size = New-Object System.Drawing.Size(1000, 720)
    $form.MinimumSize = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = 'CenterScreen'

    $headerLabel = New-Object System.Windows.Forms.Label
    $headerLabel.Text = "Wartungs Toolbox auf $($script:HostInfo.HostName) in $($script:HostInfo.Domain)"
    $headerLabel.Font = New-Object System.Drawing.Font('Verdana', 9)
    $headerLabel.AutoSize = $true
    $headerLabel.Dock = 'Top'
    $headerLabel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 5)

    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = 'Top'
    $buttonPanel.AutoSize = $true
    $buttonPanel.WrapContents = $true
    $buttonPanel.AutoSizeMode = 'GrowAndShrink'
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(10)

    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = 'Top'
    $topPanel.AutoSize = $true
    $topPanel.AutoSizeMode = 'GrowAndShrink'
    $topPanel.Controls.Add($buttonPanel)
    $topPanel.Controls.Add($headerLabel)

    $outputBox = New-Object System.Windows.Forms.RichTextBox
    $outputBox.Dock = 'Fill'
    $outputBox.Font = New-Object System.Drawing.Font('Verdana', 9)
    $outputBox.Multiline = $true
    $outputBox.ReadOnly = $true
    $outputBox.BackColor = [System.Drawing.Color]::White
    $outputBox.HideSelection = $false
    $outputBox.ScrollBars = 'Vertical'

    $script:OutputControl = $outputBox

    $form.Controls.Add($outputBox)
    $form.Controls.Add($topPanel)

    # Buttons
    [void](New-FeatureButton -Container $buttonPanel -Text 'Export to wtlog.txt' -OnClick {
        try {
            $now = Get-Date -Format 'dd.MM.yyyy_HH-mm'
            $wtLog = "wtlog_$now.txt"
            $script:OutputControl.Text | Out-File -FilePath $wtLog -Encoding UTF8
            $notification = New-Object -ComObject Wscript.Shell
            $notification.Popup("Ausgabe in $wtLog gespeichert!") | Out-Null
        } catch {
            Write-AppError "Export failed: $($_.Exception.Message)"
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'C: Free' -OnClick {
        try {
            $free = Get-SystemDriveFreeSpace -DriveLetter 'C'
            Write-AppLine ("C: Free(GB): {0:N2}" -f $free)
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Uptime' -OnClick {
        try {
            $uptime = Get-SystemUptime
            Write-AppLine "Uptime: $($uptime.Days) Tage $($uptime.Hours) Stunden"
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Clear' -OnClick {
        $script:OutputControl.Clear()
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Get AD Computers' -OnClick {
        try {
            $computers = Get-ADComputers
            foreach ($pc in $computers) {
                Write-AppLine ("{0}`t{1}`tName: {2}" -f $pc.OperatingSystem, $pc.OperatingSystemVersion, $pc.Name)
            }
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Get Server Uptimes' -OnClick {
        try {
            foreach ($server in Get-ServerUptimes) {
                if ($server.Reachable) {
                    Write-AppLine ("{0}`tLast Boot: {1} Uptime: {2} Tag(e)" -f $server.Name, $server.LastBoot, $server.DaysUp)
                } else {
                    Write-AppError ("{0}`t nicht erreichbar" -f $server.Name)
                }
            }
            Write-AppLine '------------------------------------'
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Get Client Uptimes' -OnClick {
        try {
            foreach ($client in Get-ClientUptimes) {
                if ($client.Reachable) {
                    Write-AppLine ("{0}`tUptime: {1} Tag(e)" -f $client.Name, $client.DaysUp)
                } else {
                    Write-AppError ("{0}`t nicht erreichbar" -f $client.Name)
                }
            }
            Write-AppLine '------------------------------------'
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Get Last updates' -OnClick {
        try {
            foreach ($update in Get-RecentHotFixes) {
                Write-AppLine ("Installed on: {0:d} ID: {1} Type: {2}" -f $update.InstalledOn, $update.HotFixID, $update.Description)
            }
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Eventlog' -OnClick {
        try {
            Show-EventLog
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    [void](New-FeatureButton -Container $buttonPanel -Text 'Event Overview' -OnClick {
        try {
            foreach ($overview in Get-EventOverview -Days 30) {
                Write-AppLine "Event Log: $($overview.LogName)" -Bold
                Write-AppLine ("Total Events: {0}" -f $overview.Total)
                foreach ($event in $overview.Events) {
                    Write-AppLine '--------------------------'
                    $details = "ID: {0}, Level: {1}, First Occurrence: {2}, Last Occurrence: {3}, Total: {4}" -f $event.Id, $event.Level, $event.FirstOccurrence, $event.LastOccurrence, $event.Count
                    $message = "Message: {0}" -f $event.Message
                    switch ($event.Level) {
                        'Error' { $color = [System.Drawing.Color]::Red; $bold = $true }
                        'Critical' { $color = [System.Drawing.Color]::Red; $bold = $true }
                        'Warning' { $color = [System.Drawing.Color]::Orange; $bold = $false }
                        default { $color = [System.Drawing.Color]::Black; $bold = $false }
                    }
                    Write-AppLine $details -Color $color -Bold:$bold
                    Write-AppLine $message -Color $color -Bold:$bold
                }
                Write-AppLine ''
            }
        } catch {
            Write-AppError $_.Exception.Message
        }
    })

    if ($script:HostInfo.HVInstalled) {
        [void](New-FeatureButton -Container $buttonPanel -Text 'Check-Snapshots' -OnClick {
            try {
                $report = Get-HyperVSnapshotReport
                if ($report.Snapshots) {
                    Write-AppLine 'Existing Snapshots:' -Bold
                    foreach ($snapshot in $report.Snapshots) {
                        Write-AppLine $snapshot.ToString()
                    }
                } else {
                    Write-AppLine 'No Snapshots found'
                }
                if ($report.Avhdx.Count -gt 0) {
                    Write-AppLine 'AVHDX files detected:' -Bold
                    foreach ($entry in $report.Avhdx) {
                        Write-AppError "Folder: $($entry.Folder)"
                        Write-AppError "File: $($entry.File)"
                    }
                } else {
                    Write-AppLine 'No AVHDX files found in any VM folder.'
                }
            } catch {
                Write-AppError $_.Exception.Message
            }
        })
    }

    if ($script:HostInfo.WsusInstalled) {
        [void](New-FeatureButton -Container $buttonPanel -Text 'WSUS Content' -OnClick {
            try {
                $info = Get-WSUSContentReport
                Write-AppLine "ContentFolder: $($info.ContentFolder)"
                Write-AppLine ("WSUSContent Size (GB): {0:N2}" -f $info.SizeGB)
            } catch {
                Write-AppError $_.Exception.Message
            }
        })

        [void](New-FeatureButton -Container $buttonPanel -Text 'WSUS Errors (Admin)' -OnClick {
            try {
                $report = Get-WSUSErrors
                if (-not $report.Failures -or $report.Failures.Count -eq 0) {
                    Write-AppLine "No computers were found on the WSUS server ($($report.ServerName)) with updates in error!"
                    return
                }
                Write-AppLine "Computers were found on the WSUS server ($($report.ServerName)) with failed updates!"
                foreach ($failure in $report.Failures) {
                    $computer = $failure.Computer
                    $summary = $failure.Summary
                    Write-AppLine ("`r`n{0} (IP:{1} - Wsus Id:{2})" -f $computer.FullDomainName, $computer.IPAddress, $summary.ComputerTargetId)
                    Write-AppLine ("Make".PadRight(20) + ": " + $computer.Make)
                    Write-AppLine ("Model".PadRight(20) + ": " + $computer.Model)
                    Write-AppLine ("OS".PadRight(20) + ": " + $computer.OSDescription)
                    Write-AppLine ("Last update".PadRight(20) + ": " + $summary.LastUpdated)
                    $failedUpdates = $failure.FailedUpdates
                    if ($failedUpdates) {
                        Write-AppLine ' Failed updates:'
                        foreach ($update in $failedUpdates) {
                            Write-AppError "- $($update.Title)"
                        }
                    }
                }
            } catch {
                Write-AppError $_.Exception.Message
            }
        })

        [void](New-FeatureButton -Container $buttonPanel -Text 'Shrink WSUS Content' -OnClick {
            try {
                $superseded = Invoke-WSUSCleanup
                $count = $superseded.Count
                Write-AppLine "Suche Superseded Updates"
                $message = "$count superseded Updates gefunden. Updates ablehnen?"
                $result = [System.Windows.Forms.MessageBox]::Show($message, 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-AppLine 'Declining superseded updates'
                    $counter = 1
                    foreach ($update in $superseded) {
                        $update | Deny-WsusUpdate
                        Write-Progress -Activity 'Declining Updates' -CurrentOperation $counter -PercentComplete (($counter / $count) * 100)
                        $counter++
                    }
                    Write-AppLine "$count Superseded Updates abgelehnt."
                    Write-AppLine 'Unneeded Content Files werden entfernt'
                    Get-WsusServer | Invoke-WsusServerCleanup -CleanupUnneededContentFiles
                    Write-AppLine 'WSUS Content verkleinert.'
                } else {
                    Write-AppLine 'Updates wurden nicht abgelehnt.'
                }
            } catch {
                Write-AppError $_.Exception.Message
            }
        })

        [void](New-FeatureButton -Container $buttonPanel -Text 'WSUS Sync Errors' -OnClick {
            try {
                Write-AppLine 'Failed Synchronizations:'
                foreach ($fail in Get-WSUSSyncFailures) {
                    Write-AppLine '--------------------------'
                    Write-AppLine "ID: $($fail.Id)"
                    Write-AppLine "End Time: $($fail.EndTime)"
                    Write-AppError "Error: $($fail.Error)"
                }
            } catch {
                Write-AppError $_.Exception.Message
            }
        })
    }

    if ($script:HostInfo.ADInstalled) {
        [void](New-FeatureButton -Container $buttonPanel -Text 'Check DFSR Error' -OnClick {
            try {
                $events = Get-DFSRHealth
                if ($events.Count -eq 0) {
                    Write-AppLine 'No DFS Replication events with IDs 2212 or 4012 were found in the last 180 days.'
                } else {
                    Write-AppLine 'DFS Replication errors found!' -Bold -Color ([System.Drawing.Color]::OrangeRed)
                    $occurrences = $events | Group-Object -Property Id | ForEach-Object {
                        $sorted = $_.Group | Sort-Object -Property TimeCreated
                        [pscustomobject]@{
                            Id = $_.Name
                            Count = $_.Count
                            FirstOccurrence = $sorted[0].TimeCreated
                            LastOccurrence  = $sorted[-1].TimeCreated
                        }
                    }
                    foreach ($occurrence in $occurrences) {
                        Write-AppLine ("ID: {0} Count: {1} First: {2} Last: {3}" -f $occurrence.Id, $occurrence.Count, $occurrence.FirstOccurrence, $occurrence.LastOccurrence) -Color ([System.Drawing.Color]::OrangeRed)
                    }
                }
            } catch {
                Write-AppError $_.Exception.Message
            }
        })
    }

    if ($Beta -eq 'beta' -and $script:HostInfo.WsusInstalled) {
        [void](New-FeatureButton -Container $buttonPanel -Text 'WSUS Quick Report' -OnClick {
            try {
                Show-WSUSQuickReport
            } catch {
                Write-AppError $_.Exception.Message
            }
        })
    }

    return $form
}
# endregion

$form = Initialize-App
[void]$form.ShowDialog()
