<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.142
	 Created on:   	8/7/2017 10:02 AM
	 Created by:   	jrs
	 Organization: 	
	 Filename: Get-AllMailboxInfo    	
	===========================================================================
	.DESCRIPTION
		This script will pull all mailboxes from exchange server and export selected information to a csv file.
		The CSV file will contain the users name, email address, mailbox size in MB and the users item count.
	.EXAMPLE
		Get-AllMailboxInfo -Path c:\scripts\test.csv
#>

function Get-AllMailboxInfo
{
	Param (
		[Parameter(Mandatory = $true,
			 HelpMessage = "Enter the location you want your CSV file to be saved.")]
		[Alias("OutputPath","ExportPath")]
		[string]$Path
	)
	
	$output = Get-Mailbox -ResultSize Unlimited | Select-Object DisplayName, PrimarySmtpAddress, @{ label = "TotalItemSize(MB)"; expression = { (Get-MailboxStatistics $_).TotalItemSize.Value.ToMB() } }, @{ label = "ItemCount"; expression = { (Get-MailboxStatistics $_).ItemCount } }, Database
	
	$output | Export-Csv $Path -NoTypeInformation
}

<#	
	.DESCRIPTION
		Get-EmailSendersIP will retrieve the senders original IP. The use of this cmdlet is to gather an IP
		of a SPAM email that made it through Kaspersky and so we can add it to the blacklist. You can narrow your search time 
		time frame by using the start and end parameters. Time parameters use 24 hour clock. If no date is specified
		but the time frame is, it will use the current day.
		Format is as followed to use start and end time:
		Date and time - "01/01/2017 09:00:00"
		Time - "09:00:00" or "9:00"
	.EXAMPLE
		This example will show the email senders original IP address:
		Get-EmailSendersIP -Sender spam@spammer.com
	.EXAMPLE
		This example will get the email senders original IP address within a specified time range.
		Get-EmailSendersIP -Sender spam@spammer.com -Start "01/01/2017 09:00:00" -End 01/01/2017 09:15:00"
#>

function Get-EmailSendersIP
{
	param (
		[Parameter(Mandatory = $true)]
		$Sender,
		$Start = $null,
		$End = $null
	)
	
	if ($Start -eq $null -and $End -eq $null)
	{
		#(Get-MessageTrackingLog -Sender $Sender).OriginalClientIp
		Get-MessageTrackingLog -Sender $Sender | Select-Object Sender, OriginalClientIp, Timestamp
	}
	elseif ($Start -ne $null -and $End -eq $null)
	{
		Get-MessageTrackingLog -Sender $Sender -Start $Start | Select-Object Sender, OriginalClientIp, Timestamp
	}
	elseif ($Start -eq $null -and $End -ne $null)
	{
		Get-MessageTrackingLog -Sender $Sender -End $End| Select-Object Sender, OriginalClientIp, Timestamp
	}
	elseif ($Start -ne $null -and $End -ne $null)
	{
		Get-MessageTrackingLog -Sender $Sender -Start $Start -End $End | Select-Object Sender, OriginalClientIp, TimeStamp
	}
	
}

function Export-Mailbox
{
	param (
		[Parameter(Mandatory = $true)]
		$Mailbox,
		[Parameter(Mandatory = $true)]
		$ExportPath,
		$Name
	)
	
	if($Name)
	{
		New-MailboxExportRequest -Mailbox $Mailbox -FilePath $ExportPath -Name $Name
	}
	else
	{
		New-MailboxExportRequest -Mailbox $Mailbox -FilePath $ExportPath
	}
	
}

function Get-ExchangeDatabaseSize
{
	# Get database size and whitespace within database
	Get-MailboxDatabase -Status | sort DatabaseSize -Descending | ft Name, DatabaseSize, AvailableNewMailboxSpace
}