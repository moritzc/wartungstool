#Self-Elevate
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    $ADMRequest = Read-Host -Prompt "You didn't run this script as an Administrator. Enter [Y] to execute as Admin"
    If ($ADMRequest -eq "Y" )
	{	Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
    Exit}
}



#Load System Windows Forms (PreReqs)  https://prosystech.nl/powershell-service-desk-ict-tool-gui/
[reflection.assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null


#Übergabeparameter beim Ausführen
$beta = $args[0]

#für WSUS 
Set-StrictMode -Version Latest

#[int], weil wir Strings nicht vergleichen können.
$winversion = [int](Get-WmiObject -class Win32_OperatingSystem).BuildNumber

#wsuscheck
$wsuscheck = Get-WindowsFeature | Where-Object {$_.name -eq "UpdateServices"}
$adcheck = Get-WindowsFeature | Where-Object {$_.name -eq "AD-Domain-Services"}

$network = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain()
$hostname = hostname


#ProgrammHeader und Oberer Text
$guiForm = New-Object System.Windows.Forms.Form
$guiForm.Text = "Wartungstool"
if($beta -eq 'beta')
{
$guiForm.Text = "Wartungstool - Beta-Args"
}
$guiForm.Size = New-Object System.Drawing.Size (880,710)

$guiLabel = New-Object System.Windows.Forms.Label
$guiLabel.Location = New-Object System.Drawing.Size (10,10)
$guiLabel.Size = New-Object System.Drawing.Size (800,30)
$guiLabel.Text = "Wartungs Toolbox auf " + $hostname +" in "+ $network.name
$guiLabel.Font = New-Object System.Drawing.Font ("Verdana",9)


#Textbox
$guiLabel3 = New-Object system.windows.Forms.TextBox
$guiLabel3.Location = New-Object System.Drawing.Size (10,400)
$guiLabel3.Size = New-Object System.Drawing.size (850,450)
$guiLabel3.Location = '10,200'
$guiLabel3.ScrollBars = "Vertical"
$guiLabel3.Multiline = $true
$guiLabel3.Text = ""
$guiLabel3.Font = New-Object System.Drawing.Font ("Verdana",9)

#Buttons
	
$selectCMD = New-Object System.Windows.Forms.Button
$selectCMD.Size = New-Object System.Drawing.Size (150,40)
$selectCMD.Text = 'Export to wtlog.txt'
$selectCMD.Location = '10,40'
$selectCMD.Add_Click({
	$now = date
	#$filename = $now.day+" wtlog.txt"
	#+now.month+"."+now.year+" "+$now.hour+":"+$now.minute+":"+$now.second
	$now = Get-Date -Format dd.MM.yyyy_HH-mm #| ForEach-Object { $_ -replace ":", "." }
    $WTlog = "wtlog_$now.txt"
	$guilabel3.Text | Out-File $WTlog
	$notif = New-Object -ComObject Wscript.Shell
	$notif.Popup("Ausgabe in $WTLog Gespeichert!")
    }
)


$selectContentsize = New-Object System.Windows.Forms.Button
$selectContentsize.Size = New-Object System.Drawing.Size (150,40)
$selectContentsize.Text = 'WSUS Content'
$selectContentsize.Location = '10,90'

$selectContentsize.Add_Click({
if($winversion -lt 9601){
	if($winversion -gt 9599)
	{
		$location = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Update Services\Server\Setup').ContentDir
	}
	else {guilabel3.Text += "Nur Kompatibel für Server 2012 R2+" + "`r`n"}
}
else
{
		$location = Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Update Services\Server\Setup' -Name ContentDir
		
}
		$guiLabel3.Text += "ContentFolder: " + $location + "`r`n"
		$contentsize = (gci $location\WsusContent -Recurse| measure Length -s).sum / 1Gb
		$guiLabel3.Text += "WSUSContent Size (GB): " + $contentsize.ToString('N2') + " `r`n"
}
)


$selectCfree = New-Object System.Windows.Forms.Button
$selectCfree.Size = New-Object System.Drawing.Size (150,40)
$selectCfree.Text = 'C: Free'
$selectCfree.Location = '170,40'
$selectCfree.Add_Click({
	    $guiLabel3.Text += "C: Free(GB): " + ((Get-Volume C).SizeRemaining / 1Gb).ToString('N2') + "`r`n"
		
    }
)

$selectUptime = New-Object System.Windows.Forms.Button
$selectUptime.Size = New-Object System.Drawing.Size (150,40)
$selectUptime.Text = 'Uptime'
$selectUptime.Location = '170,90'
$selectUptime.Add_Click({
	$lastboottime = (Get-WMIObject -Class Win32_OperatingSystem).LastBootUpTime
	
	$sysuptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($lastboottime)
    $guiLabel3.Text += "Uptime: " + $sysuptime.days + " Tage " + $sysuptime.hours + " Stunden " + "`r`n"
    }
)

$selectClear = New-Object System.Windows.Forms.Button
$selectClear.Size = New-Object System.Drawing.Size (150,40)
$selectClear.Text = 'Clear'
$selectClear.Location = '170,140'
$selectClear.Add_Click(
{
	$guilabel3.Text = ''
}
)


#Alle Anderen Properties können auch angezeigt werden.

$selectGETAD = New-Object System.Windows.Forms.Button
$selectGETAD.Size = New-Object System.Drawing.Size (150,40)
$selectGETAD.Text = 'Get AD Computers'
$selectGETAD.Location = '330,40'
$selectGETAD.Add_click({
	$adcomputers = Get-ADComputer -Filter {OperatingSystem -Like "Windows 10*"} -Property * | Sort-Object OperatingSystemVersion -Descending
	ForEach($pc in $adcomputers)
	{
		$guilabel3.Text += $pc.OperatingSystem + " " + $pc.OperatingSystemVersion + "`t" + "Name: " + $pc.Name  +  "`r`n" 
	}
}
)

$selectSRVUP  = New-Object System.Windows.Forms.Button
$selectSRVUP.Size = New-Object System.Drawing.Size (150,40)
$selectSRVUP.Text = 'Get Server Uptimes'
$selectSRVUP.Location = '330,140'
$selectSRVUP.Add_click({
		$servernames = Get-ADComputer -Filter {OperatingSystem -Like "*Server*"} | Sort-Object OperatingSystemVersion -Descending
		foreach($server in $servernames){
		if(Test-Connection -ComputerName $server.name -Count 1 -TimeToLive 1 -Quiet)
		{
		$Bootupdate = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server.name).LastBootUpTime
		$Bootuptime = (Get-WMIObject -Class Win32_OperatingSystem -ComputerName $server.Name).LastBootUpTime
		$uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootuptime)
		$guilabel3.Text += $server.Name + "`t" + " Last Boot: " + $Bootupdate + " Uptime: " + $uptime.days + " Tag(e)" + "`r`n" 
		}
		else{$guilabel3.Text += $server.Name + "`t"  + "nicht erreichbar" + "`r`n" }
}
$guilabel3.Text += "`r`n" + "------------------------------------" + "`r`n"}
)

$selectClientUP  = New-Object System.Windows.Forms.Button
$selectClientUP.Size = New-Object System.Drawing.Size (150,40)
$selectClientUP.Text = 'Get Client Uptimes'
$selectClientUP.Location = '490,140'
$selectClientUP.Add_click({
		$clients = Get-ADComputer -Filter {OperatingSystem -Like "Windows*"} -Properties *
		#ComputerName in PS5.1 evtl. Targetname für ältere
		foreach($client in $clients){
		if(Test-Connection -ComputerName $client.name -Count 1 -TimeToLive 1 -Quiet)
			{
				$Bootupdate = (gwmi win32_operatingsystem -computer $client.name).LastBootUpTime
				$Bootuptime = (Get-WMIObject -Class Win32_OperatingSystem -ComputerName $client.Name).LastBootUpTime
				$uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($Bootuptime)
				$guilabel3.Text += $Client.Name + "`t"  + "Uptime: " + $uptime.days + " Tag(e)" + "`r`n" 
			}
		else{$guilabel3.Text += $Client.Name + "`t"  + "nicht erreichbar" + "`r`n" }
		
}
$guilabel3.Text += "`r`n" + "------------------------------------" + "`r`n"
}
)	

$selectUpdatetimes = New-Object System.Windows.Forms.Button
$selectUpdatetimes.Size = New-Object System.Drawing.Size (150,40)
$selectUpdatetimes.Text = 'Get Last updates'
$selectUpdatetimes.Location = '330,90'
$selectUpdatetimes.Add_click({
		$lastupdates = Get-HotFix | ?{$_.InstalledOn -gt ((Get-Date).AddDays(-40))} | sort installedon -desc
		Foreach($updt in $lastupdates)
		{
			$guilabel3.Text += " Installdate: " + $updt.InstalledOn + " ID: " + $updt.HotFixID + " Type: " + $updt.Description + "`r`n"
		}
}
)

$selectWSUSSync = New-Object System.Windows.Forms.Button
$selectWSUSSync.Size = New-Object System.Drawing.Size (150,40)
$selectWSUSSync.Text = 'WSUS Sync Errors'
$selectWSUSSync.Location = '490,90'
$selectWSUSSync.Add_click(
{
	[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer()
	$wsus
	$sub=$wsus.GetSubscription()
	$failedsync = $sub.GetSynchronizationHistory() | Where-Object Result -eq 'Failed'
	$guilabel3.Text += "Failed Synchronizations:" + "`r`n"
	
	foreach($Fail in $failedsync)
	{
			$guilabel3.Text += "`r`n"+"ID: " + $Fail.Id +"`r`n"+ "End Time: " + $Fail.EndTime + "`r`n"+"Error: " + $Fail.Error + "`r`n"
	}
	$guilabel3.Text += "`r`n" + "--------------------" + "`r`n"
})


$selectEventlog = New-Object System.Windows.Forms.Button
$selectEventlog.Size = New-Object System.Drawing.Size (150,40)
$selectEventlog.Text = 'Eventlog'
$selectEventlog.Location = '490,40'
$selectEventlog.Add_click(
{
	Show-EventLog
})




#TEMPORÄR
$selectWSUSErrors = New-Object System.Windows.Forms.Button
$selectWSUSErrors.Size = New-Object System.Drawing.Size (150,40)
$selectWSUSErrors.Text = 'WSUS Errors (Admin)'
$selectWSUSErrors.Location = '10,140'
$selectWSUSErrors.Add_Click(
{	
	$HeaderChars = 0
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
	$computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
	$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
	$summariesComputerFailed = $wsus.GetSummariesPerComputerTarget($updateScope,$computerScope) | Where-Object FailedCount -NE 0 | Sort-Object FailedCount, UnknownCount, NotInstalledCount -Descending
	$computers = Get-WsusComputer
	$computersErrorEvents = $wsus.GetUpdateEventHistory([System.DateTime]::Today.AddDays(-7), [System.DateTime]::Today) | Where-Object ComputerId -ne Guid.Empty | Where-Object IsError -eq True
	
	If ($summariesComputerFailed -EQ 0 -or $summariesComputerFailed -EQ $null){
		$guiLabel3.Text += "No computers were found on the WSUS server (" + $wsus.ServerName + ") with updates in error!" + "`r`n"
    }
	
	Else {
	$guiLabel3.Text += "Computers were found on the WSUS server (" + $wsus.ServerName + ") with failed updates!" +"`r`n"
	}
	
	ForEach ($computerFailed In $summariesComputerFailed) {
  $computer = $computers | Where-Object Id -eq $computerFailed.ComputerTargetId

  
  # FullDomainName e IP
  $outputText = $computer.FullDomainName + " (IP:" + $computer.IPAddress + " - Wsus Id:" + $computerFailed.ComputerTargetId + ")"
  $guiLabel3.Text += ("`r`n" + $outputText)

  # Hardware info
  $outputText = " Hardware info".PadRight($HeaderChars) + ": " + $computer.Make + " " + $computer.Model
  $guiLabel3.Text += $outputText + "`r`n"
 

  # Operating system
  $outputText = " Operating system".PadRight($HeaderChars) + ": " + $computer.OSDescription
  $guiLabel3.Text += $outputText + "`r`n"
 


  # Update failed
  $outputText = " Update failed".PadRight($HeaderChars) + ": " + $computerFailed.FailedCount
  $guiLabel3.Text += $outputText + "`r`n"
 

  # Update unknown
  $outputText = " Update unknown".PadRight($HeaderChars) + ": " + $computerFailed.UnknownCount
  $guiLabel3.Text += $outputText + "`r`n"
 

  # Update not installed
  $outputText = " Update not installed".PadRight($HeaderChars) + ": " + $computerFailed.NotInstalledCount
  $guiLabel3.Text += $outputText + "`r`n"
 

  # Update installed pending reboot
  $outputText = " Update installed pending reboot".PadRight($HeaderChars) + ": " + $computerFailed.InstalledPendingRebootCount
  $guiLabel3.Text += $outputText + "`r`n"
 

  # Last sync result
  $outputText = " Last sync result".PadRight($HeaderChars) + ": " + $computer.LastSyncResult
  $guiLabel3.Text += $outputText + "`r`n"
  

  # Last sync time
  $outputText = " Last sync time".PadRight($HeaderChars) + ": " + ($computer.LastSyncTime).ToString()
  If ($computer.LastSyncTime -LE [System.DateTime]::Today.AddDays(-7)){
      $guiLabel3.Text += $outputText + "`r`n"
  }
  Else {
    $guiLabel3.Text += $outputText + "`r`n"
  }


  # Last updated
  $outputText = " Last update".PadRight($HeaderChars) + ": " + ($computerFailed.LastUpdated).ToString()
  $guiLabel3.Text += $outputText + "`r`n"

  # Failed Updates
  $computerUpdatesFailed = ($wsus.GetComputerTargets($computerScope) | Where-Object Id -EQ $computerFailed.ComputerTargetId).GetUpdateInstallationInfoPerUpdate($updateScope) | Where UpdateInstallationState -EQ Failed

  $computerUpdateFailedIndex=0
  ForEach ($update In $computerUpdatesFailed) {
    If ($computerUpdateFailedIndex -EQ 0){
      $outputText = " Failed updates".PadRight($HeaderChars) + ": " + "`r`n"
    }
    Else{
      $outputText = "".PadRight($HeaderChars+2)
    }

    $outputText += "-" + $wsus.GetUpdate($update.UpdateId).Title + "`r`n"
    $guiLabel3.Text += $outputText
 
    $computerUpdateFailedIndex += 1
  }


}	}
)


##Experimental
$selectWsusreport= New-Object System.Windows.Forms.Button
$selectWsusreport.Size = New-Object System.Drawing.Size (150,40)
$selectWsusreport.Text = 'WSUS Quick Report'
$selectWsusreport.Location = '650,90'
$selectWsusreport.Add_click({
		#Report Popup
$guiForm2 = New-Object System.Windows.Forms.Form
$guiForm2.Text = "WSUS Quick report"
$guiForm2.Size = New-Object System.Drawing.Size (880,710)

#Textbox
$guireportbox = New-Object system.windows.Forms.TextBox
$guireportbox.Location = New-Object System.Drawing.Size (10,400)
$guireportbox.Size = New-Object System.Drawing.size (740,400)
$guireportbox.Location = '10,200'
$guireportbox.ScrollBars = "Vertical"
$guireportbox.Multiline = $true
$guireportbox.Text = ""
$guireportbox.Font = New-Object System.Drawing.Font ("Verdana",9)

#Get Objects
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
	$computerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
	$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
	$summariesComputer = $wsus.GetSummariesPerComputerTarget($updateScope,$computerScope) | Sort-Object FailedCount, UnknownCount, NotInstalledCount -Descending
	$computers = Get-WsusComputer
	$computersErrorEvents = $wsus.GetUpdateEventHistory([System.DateTime]::Today.AddDays(-7), [System.DateTime]::Today) | Where-Object ComputerId -ne Guid.Empty | Where-Object IsError -eq True

		
$comboBox1= New-Object System.Windows.Forms.ComboBox	
$comboBox2 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Size = New-Object System.Drawing.Size (400,20)
#$comboBox11.Location = '10,20'
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	$target = [PSCustomObject]@{
		ID = ''
		Name = ''
	}
    foreach($computer in $computers)
	{
	$comboBox1.Items.Add($computer.FullDomainName)
	$target.ID = $computer.ID
	$target.Name = $computer.FullDomainName
	}

#$selection = $wsus.GetUpdateEventHistory([System.DateTime]::Today.AddDays(-7), [System.DateTime]::Today) | Where-Object ID -eq $combobox1.SelectedItem 

$checkbox1 = New-Object System.Windows.Forms.Checkbox 
$checkbox1.Location = New-Object System.Drawing.Size(10,30) 
$checkbox1.Size = New-Object System.Drawing.Size(500,20)
$checkbox1.Text = "Installed"
$checkbox1.TabIndex = 4


$checkbox2 = New-Object System.Windows.Forms.Checkbox 
$checkbox2.Location = New-Object System.Drawing.Size(10,50) 
$checkbox2.Size = New-Object System.Drawing.Size(500,20)
$checkbox2.Text = "Failed"
$checkbox2.TabIndex = 3



$selectdo = New-Object System.Windows.Forms.Button
$selectdo.Size = New-Object System.Drawing.Size (150,40)
$selectdo.Text = 'Action'
$selectdo.Location = '10,70'
$selectdo.Add_Click({
	$preselect = Get-WsusComputer | Where-Object FullDomainName -eq $comboBox1.SelectedItem
	$guireportbox.Text += $preselect.Make + "`r`n" + $preselect.OSDescription + "`r`n" 
	#$preselect = $target | Where-Object FullDomainName -eq $combobox1.SelectedItem
	#$select = $summariesComputer | Where-Object Id -eq $preselect.ID
	
	#$guireportbox.Text += $select.Make + "`r`n" + $select.OSDescription + "`r`n" 
	#$guireportbox.Text += "Update failed " + $select.FailedCount
	#$notif = New-Object -ComObject Wscript.Shell
	#$notif.Popup("Test")
    }
)

	
$GuiForm2.Controls.Add($comboBox1)
$GuiForm2.Controls.Add($guireportbox)
$GuiForm2.Controls.Add($comboBox2)

$guiform2.Controls.Add($selectdo)
$guiform2.Controls.Add($checkbox2)
$guiform2.Controls.Add($checkbox1)

$guiForm2.ShowDialog() 
}
)


	
	
#Buttons platzieren
$guiForm.Controls.Add($guiLabel)
$guiForm.Controls.Add($guiLabel3)
$guiForm.Controls.Add($selectCMD)
$guiForm.Controls.Add($selectCfree)
$guiForm.Controls.Add($selectUptime)
$guiForm.Controls.Add($selectEventlog)
#WSUS Features nur für WSUS Server
if($wsuscheck.InstallState -eq "Installed")
{
$guiForm.Controls.Add($selectContentsize)
$guiForm.Controls.Add($selectWSUSErrors)
$guiForm.Controls.Add($selectWSUSSync)
}



$guiForm.Controls.Add($selectClear)
#AD Features nur für AD Server
if($adcheck.InstallState -eq "Installed")
{
$guiForm.Controls.Add($selectGETAD)
$guiForm.Controls.Add($selectSRVUP)
$guiForm.Controls.Add($selectClientUP)
}
$guiForm.Controls.Add($selectUpdatetimes)

if($beta -eq 'beta')
{
$guiForm.Controls.Add($selectWsusreport)

$guiForm.Controls.Add($selectContentsize)
$guiForm.Controls.Add($selectWSUSErrors)
$guiForm.Controls.Add($selectWSUSSync)
$guiForm.Controls.Add($selectGETAD)
$guiForm.Controls.Add($selectSRVUP)
$guiForm.Controls.Add($selectClientUP)
}
#Starting the GUI
$guiForm.ShowDialog() 