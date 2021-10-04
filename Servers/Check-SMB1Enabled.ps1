<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.153
	 Created on:   	4/13/2020 12:46 PM
	 Created by:   	jrs
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

param(
	[array]$Servers = (Get-Content -Path "$PSScriptRoot\ServerList.txt"),
	$OutputCsvPath = $null
)
$Date = Get-Date -Format yyyyMMdd

#$OutputCsvPath = "$PSScriptRoot\ServerSMB1Check_$Date.csv"

[System.Collections.ArrayList]$sysCollection = New-Object System.Collections.ArrayList($null)

foreach ($s in $Servers)
{
	Write-Host "Checking server: $s" -ForegroundColor Cyan
	
	try
	{
		$SMB1 = Invoke-Command -ComputerName $s -ScriptBlock {
			if ((Get-SmbServerConfiguration | Select EnableSMB1Protocol).EnableSMB1Protocol -eq $true)
			{
				Set-SmbServerConfiguration -EnableSMB1Protocol $false -Force
			}
			
			Get-SmbServerConfiguration | select EnableSMB1Protocol
		} -ErrorAction Stop
		
		$Info = @{
			"Hostname" = "$s"
			SMB1Enabled = $SMB1.EnableSMB1Protocol
		}
	}
	catch [System.Exception]
	{
		$Info = @{
			"Hostname" = "$s"
			ErrorMessage = $_.Exception.Message
		}	
	}
	finally
	{
		[void]$sysCollection.Add((New-Object PSObject -Property $Info))
	}
}

if ($OutputCsvPath)
{
	$sysCollection | select "Hostname", SMB1Enabled, ErrorMessage | Export-Csv -Path $OutputCsvPath -NoTypeInformation
}
else
{
	$sysCollection | select "Hostname", SMB1Enabled, ErrorMessage
}