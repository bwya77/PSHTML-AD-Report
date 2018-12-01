
#########################################
#                                       #
#            VARIABLES                  #
#                                       #
#########################################

#Company logo that will be displayed on the left, can be URL or UNC
$CompanyLogo = ""

#Logo that will be on the right side, UNC or URL
$RightLogo = "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png"

$ReportTitle = "Active Directory Report"

#Location the report will be saved to
$ReportSavePath = "C:\Automation\"



########################################
#Array of default Security Groups
$DefaultSGs = @(
		"Access Control Assistance Operators"
		"Account Operators"
		"Administrators"
		"Allowed RODC Password Replication Group"
		"Backup Operators"
		"Certificate Service DCOM Access"
		"Cert Publishers"
		"Cloneable Domain Controllers"
		"Cryptographic Operators"
		"Denied RODC Password Replication Group"
		"Distributed COM Users"
		"DnsUpdateProxy"
		"DnsAdmins"
		"Domain Admins"
		"Domain Computers"
		"Domain Controllers"
		"Domain Guests"
		"Domain Users"
		"Enterprise Admins"
		"Enterprise Key Admins"
		"Enterprise Read-only Domain Controllers"
		"Event Log Readers"
		"Group Policy Creator Owners"
		"Guests"
		"Hyper-V Administrators"
		"IIS_IUSRS"
		"Incoming Forest Trust Builders"
		"Key Admins"
		"Network Configuration Operators"
		"Performance Log Users"
		"Performance Monitor Users"
		"Pre–Windows 2000 Compatible Access"
		"Print Operators"
		"Protected Users"
		"RAS and IAS Servers"
		"RDS Endpoint Servers"
		"RDS Management Servers"
		"RDS Remote Access Servers"
		"Read-only Domain Controllers"
		"Remote Desktop Users"
		"Remote Management Users"
		"Replicator"
		"Schema Admins"
		"Server Operators"
		"Storage Replica Administrators"
		"System Managed Accounts Group"
		"Terminal Server License Servers"
		"Users"
		"Windows Authorization Access Group"
		"WinRMRemoteWMIUsers"
	)
	
	
	
	
	
$Table 					= New-Object 'System.Collections.Generic.List[System.Object]'
$LicenseTable 			= New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable 				= New-Object 'System.Collections.Generic.List[System.Object]'
$SharedMailboxTable		= New-Object 'System.Collections.Generic.List[System.Object]'
$GroupTypetable		    = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultGrouptable 		= New-Object 'System.Collections.Generic.List[System.Object]'
$IsLicensedUsersTable 	= New-Object 'System.Collections.Generic.List[System.Object]'
$ContactTable 			= New-Object 'System.Collections.Generic.List[System.Object]'
$MailUser 				= New-Object 'System.Collections.Generic.List[System.Object]'
$ContactMailUserTable 	= New-Object 'System.Collections.Generic.List[System.Object]'
$RoomTable 				= New-Object 'System.Collections.Generic.List[System.Object]'
$EquipTable 			= New-Object 'System.Collections.Generic.List[System.Object]'
$DomainAdminTable 		= New-Object 'System.Collections.Generic.List[System.Object]'
$ExpiringAccountsTable 	= New-Object 'System.Collections.Generic.List[System.Object]'
$CompanyInfoTable 		= New-Object 'System.Collections.Generic.List[System.Object]'
$securityeventtable 	= New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable 			= New-Object 'System.Collections.Generic.List[System.Object]'


# Get all users right away. Instead of doing several lookups, we will use this object to look up all the information needed.
$AllUsers = Get-ADUser -Filter *


#Company Information
$ADInfo 				= Get-ADDomain
$ForestObj 				= Get-ADForest
$DomainControllerobj    = Get-ADDomain

$Forest 				= $ADInfo.Forest
$InfrastructureMaster   = $DomainControllerobj.InfrastructureMaster
$RIDMaster 				= $DomainControllerobj.RIDMaster
$PDCEmulator 			= $DomainControllerobj.PDCEmulator
$DomainNamingMaster 	= $ForestObj.DomainNamingMaster
$SchemaMaster 			= $ForestObj.SchemaMaster


$obj = [PSCustomObject]@{
	'Domain'			    = $Forest
	'Infrastructure Master' = $InfrastructureMaster
	'RID Master'	  		= $RIDMaster
	'PDC Emulator'		    = $PDCEmulator
	'Domain Naming Master'  = $DomainNamingMaster
	'Schema Master'			= $SchemaMaster
	
}

$CompanyInfoTable.add($obj)

#Get Domain Admins
$DomainAdminMembers = Get-ADGroupMember "Domain Admins"

Foreach ($DomainAdminMember in $DomainAdminMembers)
{
	$Name = $DomainAdminMember.Name
	$Type = $DomainAdminMember.ObjectClass
	$Enabled = (Get-ADUser -filter * | Where-Object { $_.Name -eq $Name }).Enabled
	
	$obj = [PSCustomObject]@{
		'Name'		     = $Name
		'Enabled'    	 = $Enabled
		'Type' 		     = $Type
	}
	
	$DomainAdminTable.add($obj)
	
}



#Expiring Accounts
$LooseUsers = Search-ADAccount -AccountExpiring -UsersOnly 
Foreach ($LooseUser in $LooseUsers)
{
	$NameLoose 		= $LooseUser.Name
	$UPNLoose 		= $LooseUser.UserPrincipalName
	$ExpirationDate = $LooseUser.AccountExpirationDate
	$enabled 		= $LooseUser.Enabled

	
	$obj = [PSCustomObject]@{
		'Name'					   = $NameLoose
		'UserPrincipalName'	       = $UPNLoose
		'Expiration Date'		   = $LicensedLoose
		'Enabled' 				   = $StrongPasswordLoose
	}
	
	
	$ExpiringAccountsTable.add($obj)
}
If (($ExpiringAccountsTable).count -eq 0)
{
	$ExpiringAccountsTable = [PSCustomObject]@{
		'Information' = 'Information: No Users were found to expire soon'
	}
}


#Security Logs
$SecurityLogs = Get-EventLog -Newest 7 -LogName "Security" | Where-Object { $_.Message -like "*An account*" }
Foreach ($SecurityLog in $SecurityLogs)
{
	$TimeGenerated 	= $SecurityLog.TimeGenerated
	$EntryType 		= $SecurityLog.EntryType
	$Recipient 		= $SecurityLog.Message

	
	$obj = [PSCustomObject]@{
		'Time'  	= $TimeGenerated
		'Type' 		= $EntryType
		'Message'	= $Recipient

	}
	
	$SecurityEventTable.add($obj)
}
If (($securityeventtable).count -eq 0)
{
	$securityeventtable = [PSCustomObject]@{
		'Information' = 'Information: No recent security logs'
	}
}


#Tenant Domain
$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes

#Tenant Domain
$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes | ForEach-Object{
	$obj = [PSCustomObject]@{
		'UPN Suffixes' = $_
		'Valid'		   = "True"
	}
	$DomainTable.add($obj)
}



#Get groups and sort in alphabetical order
$Groups = Get-ADGroup -Filter * | Sort-Object DisplayName

$SecurityCount 		= 0
$MailSecurityCount 	= 0
$CustomGroup 		= 0
$DefaultGroup 		= 0


Foreach ($Group in $Groups)
{
	$DefaultADGroup = 'False'
	$Type = New-Object 'System.Collections.Generic.List[System.Object]'
	$Gemail = (Get-ADGroup $Group -properties mail).mail
	
	If (($group.GroupCategory -eq "Security") -and ($Gemail -ne $Null))
	{
		$MailSecurityCount++

	}
	If (($group.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
	{
		$SecurityCount++
	}
	
	If ($DefaultSGs -contains $Group.Name)
	{
		Write-Host "$($Group.name) is a default group"
		$DefaultADGroup = "True"
		$DefaultGroup++
	}
	Else
	{
		$CustomGroup++
	}
	
	
	
	if ($group.GroupCategory -eq "Distribution")
	{
		$Type = "Distribution Group"
	}
	if (($group.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
	{
		$Type = "Security Group"
	}
	If (($group.GroupCategory -eq "Security") -and (($Gemail) -ne $Null))
	{
		$Type = "Mail-Enabled Security Group"
	}
	
	If ($Group.Name -ne "Domain Users")
	{
		$Users = (Get-ADGroupMember -Identity $Group | Sort-Object DisplayName | Select-Object -ExpandProperty Name) -join ", "
	}
	Else
	{
		
		$Users = "Skipped Domain Users Membership"
		
	}
	
	$OwnerDN = Get-ADGroup -Filter { name -eq $Group.Name }  -Properties managedBy | Select-Object -ExpandProperty ManagedBy
	
	
	$Manager = Get-ADUser -Filter * | Where-Object { $_.distinguishedname -eq $OwnerDN } | Select-Object -ExpandProperty Name
	
	
	#$hash = New-Object PSObject -property @{ Name = "$GName"; Type = "$Type"; Members = "$Users" }
	
	
	$obj = [PSCustomObject]@{
		'Name'		     	= $Group.name
		'Type'		     	= $Type
		'Members' 		 	= $users
		'Managed By'	 	= $Manager
		'E-mail Address' 	= $GEmail
		'Default AD Group' 	= $DefaultADGroup
	}
	
	$table.add($obj)
}
If (($table).count -eq 0)
{
	$table = [PSCustomObject]@{
		'Information' = 'Information: No Groups were found'
	}
}

$obj1 = [PSCustomObject]@{
	'Name'  = 'Mail-Enabled Security Group'
	'Count' = $MailSecurityCount
}

$GroupTypetable.add($obj1)

$obj1 = [PSCustomObject]@{
	'Name'  = 'Security Group'
	'Count' = $SecurityCount
}
$GroupTypetable.add($obj1)

$DistroCount = ($Groups | Where-Object { $_.GroupCategory -eq "Distribution" }).Count
$obj1 = [PSCustomObject]@{
	'Name'  = 'Distribution Group'
	'Count' = $DistroCount
}

$GroupTypetable.add($obj1)


#Default Group Pie Chart
$obj1 = [PSCustomObject]@{
	'Name'  = 'Default Groups'
	'Count' = $DefaultGroup
}

$DefaultGrouptable.add($obj1)

$obj1 = [PSCustomObject]@{
	'Name'  = 'Custom Groups'
	'Count' = $CustomGroup
}

$DefaultGrouptable.add($obj1)



#Get all licenses
$Licenses = Get-AzureADSubscribedSku
#Split licenses at colon
Foreach ($License in $Licenses)
{
	$TextLic = $null
	
	$ASku = ($License).SkuPartNumber
	$TextLic = $Sku.Item("$ASku")
	If (!($TextLic))
	{
		$OLicense = $License.SkuPartNumber
	}
	Else
	{
		$OLicense = $TextLic
	}
	
	$TotalAmount = $License.PrepaidUnits.enabled
	$Assigned = $License.ConsumedUnits
	$Unassigned = ($TotalAmount - $Assigned)
	
	If ($TotalAmount -lt $LicenseFilter)
	{
		$obj = [PSCustomObject]@{
			'Name'			      = $Olicense
			'Total Amount'	      = $TotalAmount
			'Assigned Licenses'   = $Assigned
			'Unassigned Licenses' = $Unassigned
		}
		
		$licensetable.add($obj)
	}
}
If (($licensetable).count -eq 0)
{
	$licensetable = [PSCustomObject]@{
		'Information' = 'Information: No Licenses were found in the tenant'
	}
}


$IsLicensed = ($AllUsers | Where-Object { $_.assignedlicenses.count -gt 0 }).Count
$objULic = [PSCustomObject]@{
	'Name'  = 'Users Licensed'
	'Count' = $IsLicensed
}

$IsLicensedUsersTable.add($objULic)

$ISNotLicensed = ($AllUsers | Where-Object { $_.assignedlicenses.count -eq 0 }).Count
$objULic = [PSCustomObject]@{
	'Name'  = 'Users Not Licensed'
	'Count' = $IsNotLicensed
}

$IsLicensedUsersTable.add($objULic)
If (($IsLicensedUsersTable).count -eq 0)
{
	$IsLicensedUsersTable = [PSCustomObject]@{
		'Information' = 'Information: No Licenses were found in the tenant'
	}
}

Foreach ($User in $AllUsers)
{
	$ProxyA = New-Object 'System.Collections.Generic.List[System.Object]'
	$NewObject02 = New-Object 'System.Collections.Generic.List[System.Object]'
	$NewObject01 = New-Object 'System.Collections.Generic.List[System.Object]'
	$UserLicenses = ($user | Select-Object -ExpandProperty AssignedLicenses).SkuID
	If (($UserLicenses).count -gt 1)
	{
		Foreach ($UserLicense in $UserLicenses)
		{
			$UserLicense = ($licenses | Where-Object { $_.skuid -match $UserLicense }).SkuPartNumber
			$TextLic = $Sku.Item("$UserLicense")
			If (!($TextLic))
			{
				$NewObject01 = [PSCustomObject]@{
					'Licenses' = $UserLicense
				}
				$NewObject02.add($NewObject01)
			}
			Else
			{
				$NewObject01 = [PSCustomObject]@{
					'Licenses' = $textlic
				}
				
				$NewObject02.add($NewObject01)
			}
		}
	}
	Elseif (($UserLicenses).count -eq 1)
	{
		$lic = ($licenses | Where-Object { $_.skuid -match $UserLicenses }).SkuPartNumber
		$TextLic = $Sku.Item("$lic")
		If (!($TextLic))
		{
			$NewObject01 = [PSCustomObject]@{
				'Licenses' = $lic
			}
			$NewObject02.add($NewObject01)
		}
		Else
		{
			$NewObject01 = [PSCustomObject]@{
				'Licenses' = $textlic
			}
			$NewObject02.add($NewObject01)
		}
	}
	Else
	{
		$NewObject01 = [PSCustomObject]@{
			'Licenses' = $Null
		}
		$NewObject02.add($NewObject01)
	}
	
	$ProxyAddresses = ($User | Select-Object -ExpandProperty ProxyAddresses)
	If ($ProxyAddresses -ne $Null)
	{
		Foreach ($Proxy in $ProxyAddresses)
		{
			$ProxyB = $Proxy -split ":" | Select-Object -Last 1
			$ProxyA.add($ProxyB)
			
		}
		$ProxyC = $ProxyA -join ", "
	}
	Else
	{
		$ProxyC = $Null
	}
	
	$Name = $User.DisplayName
	$UPN = $User.UserPrincipalName
	$UserLicenses = ($NewObject02 | Select-Object -ExpandProperty Licenses) -join ", "
	$Enabled = $User.AccountEnabled
	$ResetPW = Get-User $User.DisplayName | Select-Object -ExpandProperty ResetPasswordOnNextLogon
	
	If ($IncludeLastLogonTimestamp -eq $True)
	{
		$LastLogon = Get-Mailbox $User.DisplayName | Get-MailboxStatistics -ErrorAction SilentlyContinue | Select-Object -ExpandProperty LastLogonTime -ErrorAction SilentlyContinue
		$obj = [PSCustomObject]@{
			'Name'						   = $Name
			'UserPrincipalName'		       = $UPN
			'Licenses'					   = $UserLicenses
			'Last Mailbox Logon'		   = $LastLogon
			'Reset Password at Next Logon' = $ResetPW
			'Enabled'					   = $Enabled
			'E-mail Addresses'			   = $ProxyC
		}
	}
	Else
	{
		$obj = [PSCustomObject]@{
			'Name'						   = $Name
			'UserPrincipalName'		       = $UPN
			'Licenses'					   = $UserLicenses
			'Reset Password at Next Logon' = $ResetPW
			'Enabled'					   = $Enabled
			'E-mail Addresses'			   = $ProxyC
		}
	}
	
	$usertable.add($obj)
}
If (($usertable).count -eq 0)
{
	$usertable = [PSCustomObject]@{
		'Information' = 'Information: No Users were found in the tenant'
	}
}


#Get all Shared Mailboxes
$SharedMailboxes = Get-Recipient -Resultsize unlimited | Where-Object { $_.RecipientTypeDetails -eq "SharedMailbox" }
Foreach ($SharedMailbox in $SharedMailboxes)
{
	$ProxyA = New-Object 'System.Collections.Generic.List[System.Object]'
	$Name = $SharedMailbox.Name
	$PrimEmail = $SharedMailbox.PrimarySmtpAddress
	$ProxyAddresses = ($SharedMailbox | Where-Object { $_.EmailAddresses -notlike "*$PrimEmail*" } | Select-Object -ExpandProperty EmailAddresses)
	If ($ProxyAddresses -ne $Null)
	{
		Foreach ($ProxyAddress in $ProxyAddresses)
		{
			$ProxyB = $ProxyAddress -split ":" | Select-Object -Last 1
			If ($ProxyB -eq $PrimEmail)
			{
				$ProxyB = $Null
			}
			$ProxyA.add($ProxyB)
			$ProxyC = $ProxyA
		}
	}
	Else
	{
		$ProxyC = $Null
	}
	
	$ProxyF = ($ProxyC -join ", ").TrimEnd(", ")
	
	$obj = [PSCustomObject]@{
		'Name'			   = $Name
		'Primary E-Mail'   = $PrimEmail
		'E-mail Addresses' = $ProxyF
	}
	
	
	
	$SharedMailboxTable.add($obj)
	
}
If (($SharedMailboxTable).count -eq 0)
{
	$SharedMailboxTable = [PSCustomObject]@{
		'Information' = 'Information: No Shared Mailboxes were found in the tenant'
	}
}


#Get all Contacts
$Contacts = Get-MailContact
#Split licenses at colon
Foreach ($Contact in $Contacts)
{
	
	$ContactName = $Contact.DisplayName
	$ContactPrimEmail = $Contact.PrimarySmtpAddress
	
	$objContact = [PSCustomObject]@{
		'Name'		     = $ContactName
		'E-mail Address' = $ContactPrimEmail
	}
	
	$ContactTable.add($objContact)
	
}
If (($ContactTable).count -eq 0)
{
	$ContactTable = [PSCustomObject]@{
		'Information' = 'Information: No Contacts were found in the tenant'
	}
}


#Get all Mail Users
$MailUsers = Get-MailUser
foreach ($MailUser in $mailUsers)
{
	$MailArray = New-Object 'System.Collections.Generic.List[System.Object]'
	$MailPrimEmail = $MailUser.PrimarySmtpAddress
	$MailName = $MailUser.DisplayName
	$MailEmailAddresses = ($MailUser.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($MailEmailAddress in $MailEmailAddresses)
	{
		$MailEmailAddressSplit = $MailEmailAddress -split ":" | Select-Object -Last 1
		$MailArray.add($MailEmailAddressSplit)
		
		
	}
	
	$UserEmails = $MailArray -join ", "
	
	$obj = [PSCustomObject]@{
		'Name'			   = $MailName
		'Primary E-Mail'   = $MailPrimEmail
		'E-mail Addresses' = $UserEmails
	}
	
	$ContactMailUserTable.add($obj)
}
If (($ContactMailUserTable).count -eq 0)
{
	$ContactMailUserTable = [PSCustomObject]@{
		'Information' = 'Information: No Mail Users were found in the tenant'
	}
}

$Rooms = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "RoomMailBox")'
Foreach ($Room in $Rooms)
{
	$RoomArray = New-Object 'System.Collections.Generic.List[System.Object]'
	
	$RoomName = $Room.DisplayName
	$RoomPrimEmail = $Room.PrimarySmtpAddress
	$RoomEmails = ($Room.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($RoomEmail in $RoomEmails)
	{
		$RoomEmailSplit = $RoomEmail -split ":" | Select-Object -Last 1
		$RoomArray.add($RoomEmailSplit)
	}
	$RoomEMailsF = $RoomArray -join ", "
	
	
	$obj = [PSCustomObject]@{
		'Name'			   = $RoomName
		'Primary E-Mail'   = $RoomPrimEmail
		'E-mail Addresses' = $RoomEmailsF
	}
	
	$RoomTable.add($obj)
}
If (($RoomTable).count -eq 0)
{
	$RoomTable = [PSCustomObject]@{
		'Information' = 'Information: No Room Mailboxes were found in the tenant'
	}
}


$EquipMailboxes = Get-Mailbox -ResultSize Unlimited -Filter '(RecipientTypeDetails -eq "EquipmentMailBox")'
Foreach ($EquipMailbox in $EquipMailboxes)
{
	$EquipArray = New-Object 'System.Collections.Generic.List[System.Object]'
	
	$EquipName = $EquipMailbox.DisplayName
	$EquipPrimEmail = $EquipMailbox.PrimarySmtpAddress
	$EquipEmails = ($EquipMailbox.EmailAddresses | Where-Object { $_ -cnotmatch '^SMTP' })
	foreach ($EquipEmail in $EquipEmails)
	{
		$EquipEmailSplit = $EquipEmail -split ":" | Select-Object -Last 1
		$EquipArray.add($EquipEmailSplit)
	}
	$EquipEMailsF = $EquipArray -join ", "
	
	$obj = [PSCustomObject]@{
		'Name'			   = $EquipName
		'Primary E-Mail'   = $EquipPrimEmail
		'E-mail Addresses' = $EquipEmailsF
	}
	
	
	$EquipTable.add($obj)
}
If (($EquipTable).count -eq 0)
{
	$EquipTable = [PSCustomObject]@{
		'Information' = 'Information: No Equipment Mailboxes were found in the tenant'
	}
}



$tabarray = @('Dashboard', 'Groups', 'Licenses', 'Users', 'Shared Mailboxes', 'Contacts', 'Resources')

#basic Properties 
$PieObject2 = Get-HTMLPieChartObject
$PieObject2.Title = "Office 365 Total Licenses"
$PieObject2.Size.Height = 250
$PieObject2.Size.width = 250
$PieObject2.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject2.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject2.ChartStyle.ColorSchemeName = "Generated7"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject2.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject2.DataDefinition.DataNameColumnName = 'Name'
$PieObject2.DataDefinition.DataValueColumnName = 'Total Amount'

#basic Properties 
$PieObject3 = Get-HTMLPieChartObject
$PieObject3.Title = "Office 365 Assigned Licenses"
$PieObject3.Size.Height = 250
$PieObject3.Size.width = 250
$PieObject3.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject3.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject3.ChartStyle.ColorSchemeName = "Generated5"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject3.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject3.DataDefinition.DataNameColumnName = 'Name'
$PieObject3.DataDefinition.DataValueColumnName = 'Assigned Licenses'

#basic Properties 
$PieObject4 = Get-HTMLPieChartObject
$PieObject4.Title = "Office 365 Unassigned Licenses"
$PieObject4.Size.Height = 250
$PieObject4.Size.width = 250
$PieObject4.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObject4.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObject4.ChartStyle.ColorSchemeName = "Generated4"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject4.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObject4.DataDefinition.DataNameColumnName = 'Name'
$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

#basic Properties 
$PieObjectGroupType = Get-HTMLPieChartObject
$PieObjectGroupType.Title = "Group Types"
$PieObjectGroupType.Size.Height = 250
$PieObjectGroupType.Size.width = 250
$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'


#basic Properties 
$PieObjectGroupType2 = Get-HTMLPieChartObject
$PieObjectGroupType2.Title = "Custom vs Default"
$PieObjectGroupType2.Size.Height = 250
$PieObjectGroupType2.Size.width = 250
$PieObjectGroupType2.ChartStyle.ChartType = 'doughnut'


#These file exist in the module directoy, There are 4 schemes by default
$PieObjectGroupType.ChartStyle.ColorSchemeName = "ColorScheme4"
#There are 8 generated schemes, randomly generated at runtime 
$PieObjectGroupType.ChartStyle.ColorSchemeName = "Generated8"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObjectGroupType.DataDefinition.DataNameColumnName = 'Name'
$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

##--LICENSED AND UNLICENSED USERS PIE CHART--##
#basic Properties 
$PieObjectULicense = Get-HTMLPieChartObject
$PieObjectULicense.Title = "License Status"
$PieObjectULicense.Size.Height = 250
$PieObjectULicense.Size.width = 250
$PieObjectULicense.ChartStyle.ChartType = 'doughnut'

#These file exist in the module directoy, There are 4 schemes by default
$PieObjectULicense.ChartStyle.ColorSchemeName = "ColorScheme3"
#There are 8 generated schemes, randomly generated at runtime 
$PieObjectULicense.ChartStyle.ColorSchemeName = "Generated3"
#you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectULicense.ChartStyle.ColorSchemeName = 'Random'

#Data defintion you can reference any column from name and value from the  dataset.  
#Name and Count are the default to work with the Group function.
$PieObjectULicense.DataDefinition.DataNameColumnName = 'Name'
$PieObjectULicense.DataDefinition.DataValueColumnName = 'Count'

$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += get-htmlopenpage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo

$rpt += Get-HTMLTabHeader -TabNames $tabarray
	$rpt += get-htmltabcontentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
		$rpt += Get-HtmlContentOpen -HeaderText "Dashboard"
			$rpt += Get-HTMLContentOpen -HeaderText "Company Information"
				$rpt += Get-HtmlContentTable $CompanyInfoTable
			$rpt += Get-HTMLContentClose

				$rpt += get-HtmlColumn1of2
					$rpt += Get-HtmlContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'
						$rpt += get-htmlcontentdatatable $DomainAdminTable -HideFooter
					$rpt += Get-HtmlContentClose
				$rpt += get-htmlColumnClose
				$rpt += get-htmlColumn2of2
					$rpt += Get-HtmlContentOpen -HeaderText 'Accounts Expiring Soon'
						$rpt += get-htmlcontentdatatable $ExpiringAccountsTable -HideFooter
					$rpt += Get-HtmlContentClose
				$rpt += get-htmlColumnClose

			$rpt += Get-HTMLContentOpen -HeaderText "Security Logs"
				$rpt += Get-HTMLContentDataTable $securityeventtable -HideFooter
			$rpt += Get-HTMLContentClose

		$rpt += Get-HTMLContentOpen -HeaderText "UPN Suffixes"
			$rpt += Get-HtmlContentTable $DomainTable
		$rpt += Get-HTMLContentClose
	$rpt += Get-HtmlContentClose
$rpt += get-htmltabcontentclose

$rpt += get-htmltabcontentopen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
	$rpt += Get-HTMLContentOpen -HeaderText "Active Directory Groups"
		$rpt += get-htmlcontentdatatable $Table -HideFooter
	$rpt += Get-HTMLContentClose
	$rpt += Get-HTMLContentOpen -HeaderText "Active Directory Groups Chart"
		$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
			$rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable
		$rpt += Get-HTMLColumnClose
		$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
			$rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType2 -DataSet $DefaultGrouptable
		$rpt += Get-HTMLColumnClose
	$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose
$rpt += get-htmltabcontentopen -TabName $tabarray[2] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
	$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licenses"
$rpt += get-htmlcontentdatatable $LicenseTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Licensing Charts"
$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObject2 -DataSet $GroupTypetable #$licensetable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObject3 -DataSet $DefaultGrouptable #$licensetable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentclose
$rpt += get-htmltabcontentclose
$rpt += get-htmltabcontentopen -TabName $tabarray[3] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Users"
$rpt += get-htmlcontentdatatable $UserTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Licensed & Unlicensed Users Chart"
$rpt += Get-HTMLPieChart -ChartObject $PieObjectULicense -DataSet $IsLicensedUsersTable
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose
$rpt += get-htmltabcontentopen -TabName $tabarray[4] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Shared Mailboxes"
$rpt += get-htmlcontentdatatable $SharedMailboxTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose
$rpt += get-htmltabcontentopen -TabName $tabarray[5] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Contacts"
$rpt += get-htmlcontentdatatable $ContactTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Mail Users"
$rpt += get-htmlcontentdatatable $ContactMailUserTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose
$rpt += get-htmltabcontentopen -TabName $tabarray[6] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Room Mailboxes"
$rpt += get-htmlcontentdatatable $RoomTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLContentOpen -HeaderText "Office 365 Equipment Mailboxes"
$rpt += get-htmlcontentdatatable $EquipTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += get-htmltabcontentclose

$rpt += Get-HTMLClosePage

$Day = (Get-Date).Day
$Month = (Get-Date).Month
$Year = (Get-Date).Year
$ReportName = ("$Day" + "-" + "$Month" + "-" + "$Year" + "-" + "AD Report")
Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath
