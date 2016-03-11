<#
.SYNOPSIS
RemoveExchangeDomain.ps1 = Remove Domain from Exchange Organisation - Removes domain reference from Address Policies, checks for addressing in any users without Address Policy Applied, and removes accepted domain from Exchange Org.

.NOTES

Version 1.0, 11th March 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release

Author/Copyright:    Mike Parker - All rights reserved.
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
	
.PARAMETER DomainName
The domain name you want to remove from Microsoft Exchange

.PARAMETER Commit
Confirms that you want to commit changes live

.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
RemoveExchangeDomain.ps1 -DomainName TestDomain.com -Commit
This will remove any traces of TestDomain.com from your exchange environment and commit the changes.

.EXAMPLE
RemoveExchangeDomain.ps1 -DomainName LiveDomain.co.uk
The script will run and log any changes that would be made by the script, but no live changes will be completed.

#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$true )]
	[string]$DomainName,

	[Parameter( Mandatory=$false )]
	[switch]$Commit

)

############################################################################
# Functions Start 
############################################################################

#Retrieves the path the script has been run from
function Get-ScriptPath
{ Split-Path $myInvocation.ScriptName
}

#This function is used to write the log file
Function Write-Logfile()
{
 param( $logentry )
$timestamp = Get-Date -DisplayHint Time
"$timestamp $logentry" | Out-File $logfile -Append
Write-Host $logentry
}

#This function enables you to locate files using the file explorer
function Get-FileName($initialDirectory) { 
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
} #end function Get-FileName

############################################################################
# Functions end 
############################################################################


############################################################################
# Variables Start 
############################################################################

$scriptVersion = "1.0"

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logfile = "$myDir\RemoveExchangeDomain.log"

$start = Get-Date

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

# Confirm that the user has entered the correct domain name.

Write-Logfile "You have selected to remove the domain $DomainName from the Exchange Organisation."
$answer = Read-Host "Is this correct? (Y/N)"

If($answer.ToLower() -ne "y"){

	Write-Logfile "User has elected not to continue."

} 

Else{

	Write-Logfile "Removing the domain $domainName from the Exchange Organisation..."

	#First remove the domain from all Email Address Policies

	Get-EmailAddressPolicy | ForEach-Object { # Looking for domain in each email address policy

		$template = (Get-EmailAddressPolicy $_ ).EmailAddressPolicyTemplates;
		$newTemplate = @()

		ForEach($address in $template){

			if ($address -notlike '*$domainName*'){

				$newtemplate += $address

			}

		} # End of ForEach Address in Template

		if($Commit){

			Try{ #Update the Policy with the domain removed from the templates.
				$error.Clear()

				Set-EmailAddressPolicy $_ -EmailAddressPolicyTemplates $newTemplate
			}
			Catch{
				Write-Logfile "There was an error updating the email address policy..."
				Write-Logfile "$error"
			}
			Finally{
				if(!$error){
					Write-Logfile "Successfully updated Email Address Policy $_"
				}
				else{
					Write-Logfile "There was an error updating $_ "
				}
			} # End of Try, Catch, Finally
		} # End of If Commit
	} # End of Foreach Address Policy

	# Next, update all mailboxes

	Write-Logfile "Updating individual mailboxes..."

	$Mailboxes = @(Get-Mailbox -Filter {(EmailAddressPolicyEnabled -eq $False) -and (EmailAddresses -like '*$DomainName')} -ResultSize Unlimited)

	$itemCount = $mailboxes
	$itemCount = $itemCount.count
	$processedCount = 1
	$success = 0
	$failure = 0

	Foreach ($Mailbox in $Mailboxes)
	{
		$error.clear()

		Write-Progress -Activity "Processing.." -Status "User $processedCount of $itemCount" -PercentComplete ($processedCount / $itemCount * 100)

		try{

			Write-Logfile "******* Processing: $mailbox"
			$addresses = @($mailbox | Select -Expand EmailAddresses)
			$newAddresses = $addresses

			foreach ($address in $addresses)
			{
				Write-Logfile $address
				If($address -like "*$DomainName"){

					Write-Logfile "Removing Address $address"

					$newAddresses -= $address

				} # End of matching 

			} # End of address loop

			if($commit)
			{
				Set-Mailbox -Identity $Mailbox.Alias -EmailAddresses $newAddresses

			} # End of commit

		}
		catch{
			Write-Logfile "There was an error processing $Mailbox.Alias. Please review the log."

		}
		finally{
			if(!$error){
				$success++
			}
			else{
				$failure++
			}
		}

	} # End of Forech Mailbox
	Write-Logfile "$ItemCount records processed"
	Write-Logfile "$success records processed successfully."
	Write-Logfile "$failure records errored during processing." 

	# Finally remove the accepted domain from Exchange

	$AcceptedDomain = Get-AcceptedDomain | Where-Object {$_.DomainName -eq $DomainName }

	If($AcceptedDomain){
		Write-logfile "Accepted Domain found in Exchange."
		Write-Logfile "Removing Domain..."
		If($Commit){
			Remove-AcceptedDomain $AcceptedDomain.Identity
		}
	}
	Else{
		Write-logfile "Accepted domain $domainName not found in Exchange"
	}
} # End of Else

Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################