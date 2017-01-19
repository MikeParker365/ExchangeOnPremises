<#
.SYNOPSIS
Script to automatically migrate Distribution Groups which are managed by users migrated to O365

.NOTES

Version 1.0, 15th December 2016
Revision History
---------------------------------------------------------------------
1.0 	- Initial release


Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.DESCRIPTION
Script to automatically migrate Distribution Groups which are managed by users to O365

.LINK
http://www.mikeparker365.co.uk

.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
#>

[CmdletBinding()]
param (

	[Parameter( Mandatory=$true )]
	[string]$MailboxesCSVPath,
	
	[Parameter( Mandatory=$true )]
	[boolean]$ReportOnly

)

$scriptVersion = "0.0"

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
function ShowError  ($msg){Write-Host "`n";Write-Host -ForegroundColor Red $msg;   Write-Logfile  $msg }
function ShowSuccess($msg){Write-Host "`n";Write-Host -ForegroundColor Green  $msg; Write-Logfile   ($msg)}
function ShowProgress($msg){Write-Host "`n";Write-Host -ForegroundColor Cyan  $msg; Write-Logfile   ($msg)}
function ShowInfo($msg){Write-Host "`n";Write-Host -ForegroundColor Yellow  $msg; Write-Logfile   ($msg)}
<#
function LogToFile   ($msg){$msg |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogSuccessToFile   ($msg){"Success: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
function LogErrorToFile   ($msg){"Error: $msg" |Out-File -Append -FilePath $logFile -ErrorAction:SilentlyContinue;}
#>

function Collect-GroupInfoForMigration($script:GroupsToMigrate){

	Foreach($g in $script:GroupsToMigrate){
		Get-DistributionGroup $groupName 

	}
}

function Collect-GroupMembersForMigration {

	$result = foreach($DG in $script:GroupsToMigrate){
		
		Get-DistributionGroupMember -Identity $DG | Select @{Label="Group Name";Expression={$DG}}, SamAccountName
		
	}	

	$result | Export-Csv $mydir\GroupMembershipExport.csv -NoTypeInformation
}

function Check-GroupsFromMigrationBatch($MailboxesCSV){

	$Mailboxes = Import-Csv $MailboxesCSV

	ForEach($Mailbox in $Mailboxes){

		$DGs = Get-DistributionGroup -ManagedBy $Mailbox.EmailAddress
		If($DGs){
			ForEach($DG in $DGs){

				$script:GroupsToMigrate += $DG

			}
		}
	}
}

function Remove-DuplicateGroupsFromList($GroupsToCheck){

	$script:GroupsToMigrate = $script:GroupsToMigrate | Select -Unique Name

}

function Create-DistributionGroups{

	$dgs = Import-Csv $mydir\GroupMembershipExport.csv

	Foreach($dg in $dgs){

	New-DistributionGroup -Name $dg.Name -Alias $dg.Alias -BypassNestedModerationEnabled $dg.BypassNestedModerationEnabled -CopyOwnerToMember $false -DisplayName $dg.DisplayName -IgnoreNamingPolicy $true -ManagedBy <MultiValuedProperty>] [-MemberDepartRestriction <Closed |
    Open | ApprovalRequired>] [-MemberJoinRestriction <Closed | Open | ApprovalRequired>] [-Members
    <MultiValuedProperty>] [-ModeratedBy <MultiValuedProperty>] [-ModerationEnabled <$true | $false>] [-Notes
    <String>] [-Organization <OrganizationIdParameter>] [-OrganizationalUnit <OrganizationalUnitIdParameter>]
    [-OverrideRecipientQuotas <SwitchParameter>] [-PrimarySmtpAddress <SmtpAddress>]
    [-RequireSenderAuthenticationEnabled <$true | $false>] [-RoomList <SwitchParameter>] [-SamAccountName <String>]
    [-SendModerationNotifications <Never | Internal | Always>] [-Type <Distribution | Security>] [-WhatIf
    <SwitchParameter>] [<CommonParameters>]

	}
}


############################################################################
# Functions end 
############################################################################

############################################################################
# Variables Start 
############################################################################

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$logfile = "$myDir\Add-SMTPAddresses.log"

$start = Get-Date

$script:GroupsToMigrate = @()
#$script:UniqueGroupsToMigrate = @()

############################################################################
# Variables End
############################################################################

############################################################################
# Script start   
############################################################################

Write-Logfile "Script started at $start";
Write-Logfile "Running script version $scriptVersion"

## WRITE YOUR SCRIPT HERE

# Collect the Distribution Groups that need migrating in the current mailbox migration batch

Check-GroupsFromMigrationBatch($MailboxesCSVPath)

# Remove duplicate entries in the list collected from the migration batch 

Remove-DuplicateGroupsFromList($script:GroupsToMigrate)


If($ReportOnly){

	ShowInfo "The following groups will be recreated in Exchange Online and removed from Exchange On-Premises."
	ShowProgress $script:GroupsToMigrate
}

Else{

	Collect-GroupInfoForMigration
	Collect-GroupMembersForMigration
}

##

ShowInfo "------------Processing Ended---------------------"
$end = Get-Date;
ShowInfo "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
ShowInfo "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################