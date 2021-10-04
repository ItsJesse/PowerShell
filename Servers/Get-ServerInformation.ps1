<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.153
	 Created on:   	4/7/2020 11:50 AM
	 Created by:   	jrs
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

<# *Chassis type,	*Classification,	*CPU core thread,	*CPU count,	*CPU manufacturer,	
*CPU name,	*CPU speed (MHz),	*CPU type,	*Default Gateway,	*Disk space (GB),	*DNS Domain,	*Fault count,	
*Firewall status,	*Fully qualified domain name,	--Hardware Status,	*Host name,	*IP Address,	*Location,	*Manufacturer,	
*Model ID,	*Model number,	*Monitor	Name,	*Operational status,	*Operating System,	*OS Address Width (bits),	
*OS Service Pack,	*OS Version,	*RAM (MB),	*Serial number,
--Support group,	Supported by,	Backed Up By,	Monitored By,	Patched By,	Used for,	Is Virtual
#>

# Hypnos and Moros just static place ipaddress

param(
	[array]$Servers = (Get-Content -Path "$PSScriptRoot\ServerList.txt"),
	[string]$OutputCsvPath,
	[switch]$OutGridView
)


$Date = Get-Date -Format yyyyMMdd

#$OutputCsvPath = "$PSScriptRoot\ServerLists_$Date.csv"

#$Servers = Get-Content -Path "$PSScriptRoot\ServerList.txt" # Make text import of servers Get-File -Path <Server Text Documents>
#$Servers = ("PSHELL", "LJSAVAGE")

[System.Collections.ArrayList]$sysCollection = New-Object System.Collections.ArrayList($null)

foreach ($s in $Servers)
{
	Write-Host "Checking Server: $s" -ForegroundColor Cyan
	
	try
	{
		$CPU = Get-WmiObject Win32_Processor -ComputerName $s
		$System = Get-WmiObject Win32_ComputerSystem -ComputerName $s -ErrorAction Stop
		$BIOS = Get-WmiObject Win32_BIOS -ComputerName $s
		$OS = Get-WmiObject Win32_OperatingSystem -ComputerName $s
		[array]$Disks = Get-WmiObject Win32_LogicalDisk -ComputerName $s -Filter "Drivetype='3'"
		[array]$Network = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $s | where { $_.IPAddress -ne $null -and $_.Description -notmatch "LoopBack" }
		#$RAM = Get-WmiObject Win32_PhysicalMemory
		
		$Info = @{
			"Hostname" = $s
			"FQDN"	   = $s + ".bbdes.org"
			"CPU Core Thread" = $CPU.ThreadCount -join ', '
			"CPU Count" = $System.NumberOfProcessors
			"CPU Manufacturer" = $CPU.Manufacturer -join ', '
			"CPU Name" = $CPU.Name -join ', '
			"CPU Speed (MHz)" = $CPU.MaxClockSpeed -join ', '
			"CPU Type" = $CPU.SocketDesignation -join ', '
			"RAM (MB)" = "{0:n2}" -f ($system.TotalPhysicalMemory / 1MB)
			"IPAddress" = ($Network.IPAddress | where { $_ -like "*.*.*.*" }) -join ', ' #Get IPv4 only
			"Default Gateway" = $Network.DefaultIPGateway -join ', '
			"DNS Domain" = $Network.DNSServerSearchOrder -join ', '
			"Manufacturer" = $System.Manufacturer
			"Model ID" = $System.Model
			"Model Number" = $System.SystemSKUNumber
			"Chassis Type" = $System.ChassisSKUNumber
			"Serial Number" = $BIOS.SerialNumber
			"Operating System" = $OS.Name.Substring(0, $os.Name.IndexOf("|") - 1)
			"OS Address Width (bits)" = $OS.OSArchitecture
			"OS Service Pack" = ("SP {0}" -f [string]$OS.ServicePackMajorVersion)
			"OS Version" = $OS.Version
			"Location" = (&{
					if ($Network.IPAddress -match "10.0.13" -or $Network.IPAddress -match "10.0.12" -or $Network.IPAddress -match "10.0.10" -or $Network.IPAddress -match "205.1.10") { "CHR" }
					elseif ($Network.IPAddress -match "10.0.17") { "CON" }
					elseif ($Network.IPAddress -match "205.1.12") { "DOV" }
					elseif ($Network.IPAddress -match "205.1.14") { "SAL" }
					elseif ($Network.IPAddress -match "10.10.10") { "DEPOT" }
					else { "Unknown" }
				}) #Get location of device based on IP address
			"Classification" = $null
			"Firewall Status" = $null
			"Fault Count" = $null
			"Monitor Name" = $null
			"Operational Status" = $null
			"Support Group" = $null
			"Supported By" = $null
			"Backed Up By" = $null
			"Monitored By" = $null
			"Patched By" = $null
			"Used For" = $null
			"Is Virtual" = (&{
					if ($System.Model -eq "Virtual Machine") { "Yes" }
					else {"No"}
				})
			}
			
			$Disks | foreach-object { $Info."Drive$($_.Name -replace ':', '')" = "$([string]([System.Math]::Round($_.Size/1gb, 2))) GB" }
	}
	catch [System.Exception]
	{
		Write-Host "Error communicating with $s, skipping to next" -ForegroundColor Red
		$Info = @{
			"Hostname" = [string]$s
			ErrorMessage = [string]$_.Exception.Message
			ErrorItem = [string]$_.Exception.ItemName
		}
		Continue
	}
	finally
	{
		[void]$sysCollection.Add((New-Object PSObject -Property $Info))
	}
}

if ($OutputCsvPath)
{
	$sysCollection `
	| select-object "Hostname", "FQDN", "CPU Core Thread", "CPU Count", "CPU Manufacturer", "CPU Name", "CPU Speed (MHz)", "CPU Type", `
					"RAM (MB)", "IPAddress", "Default Gateway", "DNS Domain", "Manufacturer", "Model ID", "Model Number", "Chassis Type", `
					"Serial Number", "Operating System", "OS Address Width (bits)", "OS Service Pack", "OS Version", "Location", `
					"DriveC", "DriveD", "DriveE", "DriveF", "Classification", "Firewall Status", "Fault Count", "Monitor Name", `
					"Operational Status", "Support Group", "Supported By", "Backed Up By", "Monitored By", "Patched By", `
					"Used For", "Is Virtual", ErrorMessage, ErrorItem `
	| sort -Property "Hostname" `
	| Export-CSV -path $OutputCsvPath -NoTypeInformation
		
	Write-Host "inventory completed check $OutputCsvPath" -ForegroundColor Green
}
elseif ($OutGridView)
{
	$sysCollection `
	| select-object "Hostname", "FQDN", "CPU Core Thread", "CPU Count", "CPU Manufacturer", "CPU Name", "CPU Speed (MHz)", "CPU Type", `
					"RAM (MB)", "IPAddress", "Default Gateway", "DNS Domain", "Manufacturer", "Model ID", "Model Number", "Chassis Type", `
					"Serial Number", "Operating System", "OS Address Width (bits)", "OS Service Pack", "OS Version", "Location", `
					"DriveC", "DriveD", "DriveE", "DriveF", "Classification", "Firewall Status", "Fault Count", "Monitor Name", `
					"Operational Status", "Support Group", "Supported By", "Backed Up By", "Monitored By", "Patched By", `
					"Used For", "Is Virtual", ErrorMessage, ErrorItem `
	| sort -Property "Hostname" `
	| Out-GridView
	
	Write-Host "inventory completed check $OutputCsvPath" -ForegroundColor Green
}
else
{
	Write-Host "inventory completed check $OutputCsvPath" -ForegroundColor Green
	
	$sysCollection `
	| select-object "Hostname", "FQDN", "CPU Core Thread", "CPU Count", "CPU Manufacturer", "CPU Name", "CPU Speed (MHz)", "CPU Type", `
					"RAM (MB)", "IPAddress", "Default Gateway", "DNS Domain", "Manufacturer", "Model ID", "Model Number", "Chassis Type", `
					"Serial Number", "Operating System", "OS Address Width (bits)", "OS Service Pack", "OS Version", "Location", `
					"DriveC", "DriveD", "DriveE", "DriveF", "Classification", "Firewall Status", "Fault Count", "Monitor Name", `
					"Operational Status", "Support Group", "Supported By", "Backed Up By", "Monitored By", "Patched By", `
					"Used For", "Is Virtual", ErrorMessage, ErrorItem `
	| sort -Property "Hostname" `
	| FT -AutoSize
}