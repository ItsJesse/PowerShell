<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.142
	 Created on:   	8/9/2017 8:01 AM
	 Created by:   	jrs
	 Organization: Savage
	 Filename: ADTools.psm1     	
	===========================================================================
	.DESCRIPTION
		Get-UserInfo is a quick and easy search for an active directory users by using either first name,
		last name or username
	.EXAMPLE
		Search user by first name:
		Get-UserInfo -FirstName Jesse
	.EXAMPLE
		Search user by last name:
		Get-UserInfo -LastName Savage
	.EXAMPLE
		Search by username:
		Get-UserInfo -Username jsav
	.EXAMPLE
		Search user by first and last name:
		Get-UserInfo -FirstName Jesse -LastName Savage
	.EXAMPLE
		Search for user with additional properties:
		Get-UserInfo -FirstName Jesse -Properties Description
#>

function Get-UserInfo
{
	param (
		[Parameter()]
		[string]$FirstName,
		[string]$LastName,
		[Alias("Identity")]
		[string]$Username,
		[Parameter(HelpMessage="Enter one or more properties seperated by commas.")]
		[string[]]$Properties = $null
	)
	if ($Properties)
	{
		if ($FirstName)
		{
			Get-ADUser -Filter { GivenName -like $FirstName } -Properties $Properties
		}
		elseif ($LastName)
		{
			Get-ADUser -Filter { Surname -like $LastName } -Properties $Properties
		}
		elseif ($FirstName -and $LastName)
		{
			Get-ADUser -Filter { GivenName -like $FirstName -and Surname -like $LastName } -Properties $Properties
		}
		elseif ($Username)
		{
			Get-ADUser -Identity $Username -Properties $Properties
		}
	}
	else
	{
		if ($FirstName)
		{
			Get-ADUser -Filter { GivenName -like $FirstName }
		}
		elseif ($LastName)
		{
			Get-ADUser -Filter { Surname -like $LastName }
		}
		elseif ($FirstName -and $LastName)
		{
			Get-ADUser -Filter { GivenName -like $FirstName -and Surname -like $LastName }
		}
		elseif ($Username)
		{
			Get-ADUser -Identity $Username
		}
	}
}

<#	
	.DESCRIPTION
		Reset-Password allows you to reset a users password and you can choose whether to 
		force password change at logon or not. This command will also unlock the account automatically
		if the users account is locked out.
	.EXAMPLE
		Reset users password:
		Reset-Password -Identity test -NewPassword hello
	.EXAMPLE
		Reset users password and force password change at logon:
		Reset-Password -Identity test -NewPassword hello -ChangePasswordAtLogon
#>

function Reset-Password
{
	param (
		[Parameter(Mandatory=$true)]
		[string]$Identity,
		[Parameter(Mandatory=$true)]
		[string]$NewPassword,
		[switch]$ChangePasswordAtLogon
	)
	
	#Creates password as secure string
	$newpwd = ConvertTo-SecureString -String $NewPassword -AsPlainText -Force
	
	#Checks to see is change password at logon script switch is selected
	if ($ChangePasswordAtLogon)
	{
		#Sets users password and forces reset upon logging on
		Set-ADAccountPassword -Identity $Identity -NewPassword $newpwd -Reset -PassThru |
		Set-ADUser -ChangePasswordAtLogon $true
	}
	else
	{
		#Resets users password only without forcing reset upon logging in
		Set-ADAccountPassword -Identity $Identity -NewPassword $newpwd -Reset
	}
	
	#Check if account is locked
	$isLock = (Get-ADUser -Identity $Identity -Properties LockedOut).LockedOut
	
	#Unlock account if it is locked out
	if ($isLock)
	{
		Unlock-ADAccount -Identity $Identity
	}
}

<#	
	.DESCRIPTION
		Check-Niagara will display the results of the service Niagara on the HVAC service and display whether it
		is running or not.
	.EXAMPLE
		Check Niagara Service
		Check-Niagara
#>

function Check-Niagara
{
	Get-Service -Name Niagara -ComputerName HVAC
}

<#	
	.DESCRIPTION
		Start-Niagara will start the Niagara service on the HVAC server.
	.EXAMPLE
		Start the Niagara Service
		Start-Niagara
#>

function Start-Niagara
{
	Get-Service -Name Niagara -ComputerName HVAC | Set-Service -Status Running
}

<#	
	.DESCRIPTION
		Get-IT will retrieve the list of commands available within the ADTools and ExchangeTools Module.
	.EXAMPLE
		Get a list of commands from ADTools Module:
		Get-IT
#>

function Get-IT
{
	Get-Command -Module ADTools,ExchangeTools
}

<#	
	.DESCRIPTION
		Get-Jesse will retrieve the list of commands available within the ADTools and ExchangeTools Module.
	.EXAMPLE
		Get a list of commands from ADTools Module:
		Get-Jesse
#>

function Get-Jesse
{
	Get-Command -Module ADTools, ExchangeTools
}

<#
.DESCRIPTION
		Get-FolderReports will pull detailed information of all subfolders in the given path. It will retireve folder sizes,
		last write time, last access time and folder full path.
	.EXAMPLE
		This will get the size of subfolders of given directory and display the results in the powershell console:
		Get-FolderReports -Path C:\Users\username\documents\ -DisplayInConsole
	.EXAMPLE
		This will get the size of subfolders of given directory and export the data to a csv:
		Get-FolderReports -Path C:\Users\username\documents\ -ExportData
		cmdlet Get-FolderReports at command pipeline position 1
		Supply values for the following parameters:
		ExportPath: C:\Temp\test.csv
	.EXAMPLE
		This will get the size of subfolders of given directory and display the data in the console and export it to a csv:
		Get-FolderReports -Path C:\Users\username\documents\ -ExportData -DisplayInConsole
		cmdlet Get-FolderReports at command pipeline position 1
		Supply values for the following parameters:
		ExportPath: C:\Temp\test.csv
#>

function Get-FolderReports
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Path,
		<#[Parameter(Mandatory = $true)]
		[string]$ExportPath,#>
		[switch]$ExportData,
		[switch]$DisplayInConsole
	)
	
	DynamicParam
	{
		$paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
		$attributes = New-Object System.Management.Automation.ParameterAttribute
		$attributes.ParameterSetName = "__AllParameterSets"
		$attributes.Mandatory = $true
		$attributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
		$attributeCollection.Add($attributes)
		
		#If "-ExportData" is used, then add the "ExportPath" parameter
		if ($ExportData)
		{
			$dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ExportPath", [String], $attributeCollection)
			$paramDictionary.Add("ExportPath", $dynParam1)
		}
		
		return $paramDictionary
	}
	Process
	{
		$info = @()
		
		$colItems = Get-ChildItem $Path | Where-Object { $_.PSIsContainer -eq $true } | Sort-Object
		foreach ($i in $colItems)
		{
			$subFolderItems = Get-ChildItem $i.FullName -recurse -force | Where-Object { $_.PSIsContainer -eq $false } | Measure-Object -property Length -sum | Select-Object Sum
			#$i.FullName + " -- " + "{0:N2}" -f ($subFolderItems.sum / 1MB) + " MB"
			$size = "{0:N2}" -f ($subFolderItems.sum / 1MB)
			<#
			$i.FullName
			$size
			$i.LastWriteTime
			$i.LastAccessTime
			#>
			Write-Host "Scanning Folder: $($i.Fullname)"
			
			$props = @{
				'FolderPath'	 = $i.FullName;
				'Size(MB)'	     = $size;
				'LastWriteTime'  = $i.LastWriteTime;
				'LastAccessTime' = $i.LastAccessTime
			}
			
			$obj = New-Object -TypeName PSObject -Property $props
			
			if ($ExportData)
			{
				$obj | Export-Csv -Path "$($PSBoundParameters.ExportPath)" -Append -NoTypeInformation
			}
			if($DisplayInConsole)
			{
				$info += $obj
			}
		}
		if ($DisplayInConsole)
		{
			$info	
		}
	}
}

function Get-LoggedOnUser
{
	[CmdletBinding()]
	param
	(
		[Parameter()]
		[ValidateScript({ Test-Connection -ComputerName $_ -Quiet -Count 1 })]
		[ValidateNotNullOrEmpty()]
		[string[]]$ComputerName = $env:COMPUTERNAME,
		$ExportPath
	)
	foreach ($comp in $ComputerName)
	{
		$output = @{ 'ComputerName' = $comp }
		$output.UserName = (Get-WmiObject -Class win32_computersystem -ComputerName $comp).UserName
		[PSCustomObject]$output
		if ($ExportPath)
		{
			$output | Out-File -FilePath $ExportPath -Append
		}
	}
}