<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.152
	 Created on:   	9/26/2018 10:44 AM
	 Created by:   	Jesse Savage
	 Organization: 	BBD
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		This Script will check the ABO Wheels logs of each wheels server to ensure the download completed correctly.
		If any of the downloads fail, we will receive an email with this information.
#>

function Send-Email
{
	param (
		$Subject,
		$Body
	)
	
	$To = @("test@fake.org")
	
	Send-MailMessage -SmtpServer "SmtpServer" -From "email@address.org" -To $To -Subject $Subject -Body $Body -UseSsl
	
}

#wheels servers in an array
$servers = @("chrwheels", "dionysus", "conwheels", "dovwheels", "salwheels", "m1server", "m1super", "m2server", "m2super", "m3server", "m3super", "m4server", "m4super")

$servers | foreach{
	Write-Host "Testing: $_" -ForegroundColor Yellow
		
	if (Test-Connection -ComputerName $_ -Count 1 -Quiet)
	{
		if ((Get-Service -ComputerName $_ -Name RemoteRegistry).Status -ne "Running")
		{
			Write-Host "Starting RemoteRegistry Service on computer: $_"
			Set-Service -ComputerName $_ -Name RemoteRegistry -Status Running -StartMode Automatic
		}
		
		if (Get-Process -ComputerName $_ -Name javaw)
		{
			#Line below will test to see if servers have overnight download error
			$logFile = Get-Content "\\$_\c$\Program Files (x86)\ABO SUITE\ABO WHEELS\logs\ABOWheels_0.txt" | Select-Object -last 10
			# $logFile = Get-Content "E:\Powershell\Wheels Scripts\ABOWheels_1.txt" | Select-Object -Last 10
			$logFileError = $logFile | Select-String "Process Command Failure", "severe", "BBCS-9302", "BBCS-4351", "Refresh Cancelled", "warning", "FAIL"
			$logFileDate = $logFile[$logFile.Length - 1].Substring(0, 19)
			
			# Future Date and Time use
			$LineTime = [datetime]::ParseExact($logFileDate, "MM/dd/yyyy HH:mm:ss", $null)
			# Current Date and Time
			$CurrentDate = Get-Date
			
			# If statement to see if last logged date is accurate to when download should of ran
			# Also see if logfile error is thrown
			if ($logFileError -and $LineTime -gt $CurrentDate.AddHours(-9))
			{
				# Body of the email
				$Body = $logFileError.ToString()
				
				# Subject of the email
				$Subject = $_ + " - Error with the overnight download/sync"
				
				# Send email with following parameters
				Send-Email -Subject $Subject -Body $Body
			}
			elseif ($LineTime -lt $CurrentDate.AddHours(-9))
			{
				# Body of email
				$Body = "Log file last entry date is older than expected.`n" + $logFile[$logFile.Length - 1]
				
				$Subject = $_ + " - Log file last entry date is older than expected"
				
				# Send email with following parameters
				Send-Email -Subject $Subject -Body $Body
			}
			
			<#
			# If statement to test different possible outcomes with overnight download
			if ($logFileError)
			{
				#Body of the email
				$Body = $logFileError.ToString()
				
				#Subject of the email
				$Subject = $_ + " - Error with the overnight download/sync"
				
				#Send email with following parameters
				Send-Email -Subject $Subject -Body $Body
			}
			#>
					
			<# Only send email if there is a log file error
			else
			{
				#Body of the email
				$date = (Get-Date).AddDays(-1).ToShortDateString()
				$Body = "No errors were found in ABO log files for server $_ on date $date.`n" + `
				$logFile[3] + "`n" + $logFile[4] + "`n" + $logFile[5] + "`n" + $logFile[6] + "`n" + $logFile[7] + "`n" + $logFile[8]
				
				#Subject of the email
				$Subject = $_ + " - Download ran without issues"
			}
			#>
			
		}
		else
		{
			$Subject = $_ + " - Wheels Application Not Open"
			$Body = "$_ does not have wheels open. Download did not run."
			
			Send-Email -Subject $Subject -Body $Body
		}
		
		#Check for .LCK file
		$lockFile = Get-ChildItem "\\$_\c$\Program Files (x86)\ABO SUITE\ABO WHEELS\" | where { $_.Extension -eq ".lck" }
		
		#If lockFile = true delete it
		if ($lockFile)
		{
			#Removes LCK file
			Remove-Item -Path $lockFile.Fullname
			
			#Append to body of email
			$Body = $Body + "`nRemoved LockFile - $($lockFile.Name)"
		}
	}
	else
	{
		$Subject = $_ + " - Computer Offline"
		$Body = "$_ is not online. Download did not run."
		
		Send-Email -Subject $Subject -Body $Body
	}
	
	
	#Send-Email -Subject $Subject -Body $Body
	
}


