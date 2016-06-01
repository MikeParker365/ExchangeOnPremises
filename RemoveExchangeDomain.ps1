<#
.SYNOPSIS
RemoveExchangeDomain.ps1 = Remove Domain from Exchange Organisation - Removes domain reference from Address Policies, checks for addressing in any users without Address Policy Applied, and removes accepted domain from Exchange Org.

.NOTES

Version 1.0, 11th March 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release.
1.1     - Added mail users into the scope of the script.
1.2		- Added support for multiple domain removal at once.
		- Bug Fixes
		- Improved logging

Author/Copyright:    Mike Parker - All rights reserved.
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Remove Domain from Exchange Organisation - Removes domain reference from Address Policies, checks for addressing in any users without Address Policy Applied, and removes accepted domain from Exchange Org.

If a domain is manually added as a user's Primary SMTP Address (not via EAP) it will not be removed.

.PARAMETER DomainName
The domain name you want to remove from Microsoft Exchange

.PARAMETER DomainNames
The multiple domain names you want to remove from Microsoft Exchange

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

.EXAMPLE
RemoveExchangeDomain.ps1 -DomainNames "LiveDomain.co.uk","LiveDomain2.co.uk","LiveDomain.net" -Commit
The script will remove all of the specified domains from your Exchange environment one after the other.


#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$false )]
	[string]$DomainName,

	[Parameter( Mandatory=$false )]
	[array]$DomainNames,

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

############################################################################
# Functions end 
############################################################################

############################################################################
# Variables Start 
############################################################################

$scriptVersion = "1.2"

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$start = Get-Date

$dateforlog = $start.ToString("MM-dd-yyyy_hh-mm-ss")

$logfile = "$myDir\RemoveExchangeDomain-" + $dateforlog + ".log"


############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

# Confirm that the user has entered the correct domain name.
If($DomainName){
	Write-Logfile "You have selected to remove the domain $DomainName from the Exchange Organisation."
	$answer = Read-Host "Is this correct? (Y/N)"

	If($answer.ToLower() -ne "y"){

		Write-Logfile "User has elected not to continue."

	} 

	Else{
		If(!$Commit){
			Write-Warning "No changes will be made due to -Commit switch not being specified."
		}

		Write-Logfile "Removing the domain $domainName from the Exchange Organisation..."

		# First remove the domain from all Email Address Policies
    
		Write-Logfile "Checking email address policies for domain $DomainName"

		$Policies = Get-EmailAddressPolicy  
    
		ForEach($Policy in $Policies) { # Looking for domain in each email address policy

			Write-Logfile "Checking Policy $Policy"
			$PolicyName = $Policy.Identity
			#Write-Host $PolicyName
			$template = (Get-EmailAddressPolicy -Identity $PolicyName).EnabledEmailAddressTemplates
			#Write-Host $template
			$oldList = $template -join ","

			$newTemplate = $template

			ForEach($address in $template){

				if ($address -like "*$domainName"){
					try{
						$error.clear()
						Write-Logfile "Removing Address: $address"
						$newtemplate -= $address
						}
					catch{
						Write-Logfile $error
						}
					finally{
						if(!$error){
							Write-Logfile "$address removed from list"
						}
						else{
							Write-Logfile "There was an error"
							Write-Logfile $error
							}
						}
				}

			} # End of ForEach Address in Template
        
			$newList = $newTemplate -join ","

			If($oldList -eq $newList){

				# If templates match there are no changes to be made

				Write-LogFile "No changes required to Address Policy $policy"

				}

			Else{
				if($Commit){

					Try{ #Update the Policy with the domain removed from the templates.
						$error.Clear()
						Write-Logfile "Removing all entries for $domainName in Address Policy $Policy"
					

						Write-Logfile "Old Domains in Policy $Policy :"
						Write-Logfile $oldList
						Write-Logfile ""
						Write-Logfile "New Domains in Policy $Policy :"
						Write-Logfile $newList

						Write-Logfile "Updating Policy $Policy"
						Set-EmailAddressPolicy $Policy -EnabledEmailAddressTemplates $newTemplate
                    
						# Update the recipients
						Write-Logfile "Updating policy recipients..."
						Update-EmailAddressPolicy $Policy
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
			}# End of Else (Changes to be made)
		} # End of Foreach Address Policy

		# Next, update all mailboxes

		Write-Logfile "Updating individual mailboxes..."

		$Mailboxes = @(Get-Mailbox -ResultSize Unlimited | Where-Object {($_.EmailAddressPolicyEnabled -eq $False) -and ($_.EmailAddresses -like "*$DomainName")} )

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

				Write-Logfile "******* Processing: $mailbox *******"
				$addresses = $mailbox.EmailAddresses
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
				Write-Logfile $error

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
		Write-Logfile "$ItemCount Mailboxes processed"
		Write-Logfile "$success Mailboxes processed successfully."
		Write-Logfile "$failure Mailboxes errored during processing." 

	# Next, update all mail users

		Write-Logfile "Updating individual mailboxes..."

		$Mailboxes = @(Get-MailUser -ResultSize Unlimited | Where-Object {($_.EmailAddressPolicyEnabled -eq $False) -and ($_.EmailAddresses -like "*$DomainName")} )

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

				Write-Logfile "******* Processing: $mailbox *******"
				$addresses = $mailbox.EmailAddresses
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
					Set-MailUser -Identity $Mailbox.Alias -EmailAddresses $newAddresses

				} # End of commit

			}
			catch{
				Write-Logfile "There was an error processing $Mailbox.Alias. Please review the log."
				$error

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
		Write-Logfile "$ItemCount Mail Users processed"
		Write-Logfile "$success Mail Users processed successfully."
		Write-Logfile "$failure Mail Users errored during processing." 

		# Finally remove the accepted domain from Exchange

		$AcceptedDomain = Get-AcceptedDomain | Where-Object {$_.DomainName -eq $DomainName }

		If($AcceptedDomain){
			Write-logfile "Accepted Domain found in Exchange."
			Write-Logfile "Removing Domain..."
			If($Commit){
				try{

				$error.Clear()
				Remove-AcceptedDomain $AcceptedDomain.Identity -Confirm:$False
					}
				catch{
					Write-Logfile "There was an error removing the domain $AcceptedDomain.Identity"
					Write-Logfile $error
					}
				finally{
					if(!$error){
						Write-Logfile "Successfully removed $AcceptedDomain.Identity"
						}
					else{
						Write-Logfile "There was an error removing the domain. Please review the log."
						}
					} # End of try catch finally
			} # End of Commit
		} # End of handling the Accepted domain if found in Exchange
		Else{
			Write-logfile "Accepted domain $domainName not found in Exchange"
		}
		if(!$Commit)
		{
			Write-Warning "No changes made due to -Commit switch not being specified."

		}
	} # End of Else
}

ElseIf ($DomainNames){

	Write-Logfile "You have selected to remove multiple domains. ($DomainNames) "
	$answer = Read-Host "Is this correct Y/N"

	If($answer.ToLower() -eq "y"){

		Foreach($DomainName in $DomainNames){

			Write-Logfile "You have selected to remove the domain $DomainName from the Exchange Organisation."
			$answer = "y"

			If($answer.ToLower() -ne "y"){

				Write-Logfile "User has elected not to continue."

			} 

			Else{
				If(!$Commit){
					Write-Warning "No changes will be made due to -Commit switch not being specified."
				}

				Write-Logfile "Removing the domain $domainName from the Exchange Organisation..."

				# First remove the domain from all Email Address Policies
    
				Write-Logfile "Checking email address policies for domain $DomainName"

				$Policies = Get-EmailAddressPolicy  
    
				ForEach($Policy in $Policies) { # Looking for domain in each email address policy

					Write-Logfile "Checking Policy $Policy"
					$PolicyName = $Policy.Identity
					#Write-Host $PolicyName
					$template = (Get-EmailAddressPolicy -Identity $PolicyName).EnabledEmailAddressTemplates
					#Write-Host $template
					$oldList = $template -join ","

					$newTemplate = $template

					ForEach($address in $template){

						if ($address -like "*$domainName"){
							try{
								$error.clear()
								Write-Logfile "Removing Address: $address"
								$newtemplate -= $address
								}
							catch{
								Write-Logfile $error
								}
							finally{
								if(!$error){
									Write-Logfile "$address removed from list"
								}
								else{
									Write-Logfile "There was an error"
									Write-Logfile $error
									}
								}
						}

					} # End of ForEach Address in Template
        
					$newList = $newTemplate -join ","

					If($oldList -eq $newList){

						# If templates match there are no changes to be made

						Write-LogFile "No changes required to Address Policy $policy"

						}

					Else{
						if($Commit){

							Try{ #Update the Policy with the domain removed from the templates.
								$error.Clear()
								Write-Logfile "Removing all entries for $domainName in Address Policy $Policy"
					

								Write-Logfile "Old Domains in Policy $Policy :"
								Write-Logfile $oldList
								Write-Logfile ""
								Write-Logfile "New Domains in Policy $Policy :"
								Write-Logfile $newList

								Write-Logfile "Updating Policy $Policy"
								Set-EmailAddressPolicy $Policy -EnabledEmailAddressTemplates $newTemplate
                    
								# Update the recipients
								Write-Logfile "Updating policy recipients..."
								Update-EmailAddressPolicy $Policy
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
					}# End of Else (Changes to be made)
				} # End of Foreach Address Policy

				# Next, update all mailboxes

				Write-Logfile "Updating individual mailboxes..."

				$Mailboxes = @(Get-Mailbox -ResultSize Unlimited | Where-Object {($_.EmailAddressPolicyEnabled -eq $False) -and ($_.EmailAddresses -like "*$DomainName")} )

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

						Write-Logfile "******* Processing: $mailbox *******"
						$addresses = $mailbox.EmailAddresses
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
						Write-Logfile $error

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
				Write-Logfile "$ItemCount Mailboxes processed"
				Write-Logfile "$success Mailboxes processed successfully."
				Write-Logfile "$failure Mailboxes errored during processing." 

			# Next, update all mail users

				Write-Logfile "Updating individual mailboxes..."

				$Mailboxes = @(Get-MailUser -ResultSize Unlimited | Where-Object {($_.EmailAddressPolicyEnabled -eq $False) -and ($_.EmailAddresses -like "*$DomainName")} )

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

						Write-Logfile "******* Processing: $mailbox *******"
						$addresses = $mailbox.EmailAddresses
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
							Set-MailUser -Identity $Mailbox.Alias -EmailAddresses $newAddresses

						} # End of commit

					}
					catch{
						Write-Logfile "There was an error processing $Mailbox.Alias. Please review the log."
						$error

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
				Write-Logfile "$ItemCount Mail Users processed"
				Write-Logfile "$success Mail Users processed successfully."
				Write-Logfile "$failure Mail Users errored during processing." 

				# Finally remove the accepted domain from Exchange

				$AcceptedDomain = Get-AcceptedDomain | Where-Object {$_.DomainName -eq $DomainName }

				If($AcceptedDomain){
					Write-logfile "Accepted Domain found in Exchange."
					Write-Logfile "Removing Domain..."
					If($Commit){
						try{

						$error.Clear()
						Remove-AcceptedDomain $AcceptedDomain.Identity -Confirm:$False
							}
						catch{
							Write-Logfile "There was an error removing the domain $AcceptedDomain.Identity"
							Write-Logfile $error
							}
						finally{
							if(!$error){
								Write-Logfile "Successfully removed $AcceptedDomain.Identity"
								}
							else{
								Write-Logfile "There was an error removing the domain. Please review the log."
								}
							} # End of try catch finally
					} # End of Commit
				} # End of handling the Accepted domain if found in Exchange
				Else{
					Write-logfile "Accepted domain $domainName not found in Exchange"
				}
				if(!$Commit)
				{
					Write-Warning "No changes made due to -Commit switch not being specified."

				}
			} # End of Else
		}

	}


}
Write-Logfile "------------Processing Ended---------------------"
$end = Get-Date;
Write-Logfile "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
Write-Logfile "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################