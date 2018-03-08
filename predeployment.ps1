#1. Start PowerShell as Administrator.
#2. Powershell.exe -version 2
#3. Set-ExecutionPolicy Unrestricted.
#4. Close PowerShell.
#5. Open a new PowerShell session unelevated in order to run script as the logged in user.
#6. Run script ./predeployment.ps1
######################################################
#Enumerate which drive letter the deployment stick is.
#
Write-Host "Beginning Data Gathering..."
Start-Sleep -s 2
$overlaydrive=(Get-WMIObject Win32_LogicalDisk | Where-Object{$_.VolumeName -eq 'F50 PIVOT'} | ForEach-Object{$_.DeviceID})
#
#Set the location to the "Refresh Screenshots" directory on the deployment stick.
#
Set-Location $overlaydrive\"Refresh Screenshots"
#
#Create a directory for the pre-deployment information for the current user.
#
$refreshfoldername = New-Item -Name "$env:USERNAME" -type "Directory"
#
#Gather general information from systeminfo
systeminfo.exe | out-file -FilePath "$refreshfoldername\systeminfo_$env:USERNAME.txt"
#Gather ipconfigall information.
#
ipconfig /all | Out-File -FilePath "$refreshfoldername\ipconfigall_$env:USERNAME.txt"
#
#Call the registry to get the list of installed programs.
#
Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | 
Export-Csv -Path "$refreshfoldername\installedsoftware_$env:USERNAME.csv"
#
#Copy Chrome bookmarks to Refresh Screenshots directory.
#
Copy-Item "$Env:USERPROFILE\AppData\Local\Google\Chrome\User Data\Default\Bookmarks" -Destination "$refreshfoldername\Chrome_Bookmarks_$env:USERNAME"
#
#Gather printer information and output to Refresh Screenshots directory.
#
Get-WMIObject Win32_Printer -ComputerName $env:COMPUTERNAME | Select-Object Name, PortName, Path | Export-Csv -Path "$refreshfoldername\printer_config_$env:USERNAME.csv"
#
#Gather network drive paths and output to Refresh Screenshots directory.
#
##Try this command to see if this is more reliable.
Get-WmiObject -Class Win32_MappedLogicalDisk | select Name, ProviderName | format-table -autosize | out-file -filepath "$refreshfoldername\mapped_drives_$env:USERNAME.txt"
#$MappedDrives = @{}
#Get-WmiObject win32_logicaldisk -Filter 'drivetype=4' | Foreach { $MappedDrives.($_.deviceID) = $_.ProviderName }
#$MappedDrives | Out-File -FilePath "$refreshfoldername\mapped_drives_$env:USERNAME.txt"
#
#Gather admin group members and output to Refresh Screenshots directory.\
#
net localgroup administrators | Out-File "$refreshfoldername\admin_group_$env:USERNAME.txt"
#
#Gather Outlook PST locations and output to Refresh Screenshots directory.
#
$outlook = New-Object -comObject Outlook.Application 
$outlook.Session.Stores | where { ($_.FilePath -like '*.PST') } | format-table DisplayName, FilePath -autosize |
Out-File -FilePath "$refreshfoldername\outlook_archives_$env:USERNAME.txt"
Write-Host "Data Gathering Complete."
Start-Sleep -s 2
