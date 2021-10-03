<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.153
	 Created on:   	8/20/2019 9:27 AM
	 Created by:   	jrs
	 Organization: 	
	 Filename:     	Logging.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

[CmdletBinding()]
param (
	[Parameter()]
	[ValidateNotNullOrEmpty()]
	[string]$Message,
	[Parameter()]
	[ValidateNotNullOrEmpty()]
	[ValidateSet('Information', 'Warning', 'Error')]
	[string]$Severity = 'Information'
)

[pscustomobject]@{
	Time = (Get-Date -f g)
	Message = $Message
	Severity = $Severity
} | Export-Csv -Path "$PSScriptRoot\Logs\LogFile.csv" -Append -NoTypeInformation
