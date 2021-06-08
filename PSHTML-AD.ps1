 <#
.SYNOPSIS
    Generate graphed report for all Active Directory objects.

.DESCRIPTION
    Generate graphed report for all Active Directory objects.

.PARAMETER CompanyLogo
    Enter URL or UNC path to your desired Company Logo for generated report.

    -CompanyLogo "\\Server01\Admin\Files\CompanyLogo.png"

.PARAMETER RightLogo
    Enter URL or UNC path to your desired right-side logo for generated report.

    -RightLogo "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png"

.PARAMETER ReportTitle
    Enter desired title for generated report.

    -ReportTitle "Active Directory Report"

.PARAMETER Days
    Users that have not logged in [X] amount of days or more.

    -Days "30"

.PARAMETER UserCreatedDays
    Users that have been created within [X] amount of days.

    -UserCreatedDays "7"

.PARAMETER DaysUntilPWExpireINT
    Users password expires within [X] amount of days

    -DaysUntilPWExpireINT "7"

.PARAMETER ADModNumber
    Active Directory Objects that have been modified within [X] amount of days.

    -ADModNumber "3"

.NOTES
    Version: 1.0.3
    Author: Bradley Wyatt
    Date: 12/4/2018
    Modified: JBear 12/5/2018
    Bradley Wyatt 12/8/2018
    jporgand 12/6/2018
    michaelawilliams28 5/22/2020
#>


param (
	[Parameter(ValueFromPipeline = $true)]
	$script:loggingDate = (get-date -Format MM-dd-yyyy-hh:mm:ss),
	
	[Parameter(ValueFromPipeline = $true)]
	$script:logDate = (Get-Date -Format MM-dd-yyyy),
	
	[Parameter(ValueFromPipeline = $true)]
	$script:currentdir = (Get-Location),
	
	[Parameter(ValueFromPipeline = $true)]
	$script:steps = 25,
	#Company logo that will be displayed on the left, can be URL or UNC
	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
	[String]$script:CompanyLogo = "",
	#Logo that will be on the right side, UNC or URL

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
	[String]$script:RightLogo = "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png",
	#Title of generated report

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
	[String]$script:ReportTitle = "Active Directory Report",
	#Location the report will be saved to

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
	[String]$script:ReportSavePath = "$script:currentdir\report\",
	#Find users that have not logged in X Amount of days, this sets the days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have not logged on in more than [X] days. amount of days; Default: 30")]
	$script:Days = 30,
	#Get users who have been created in X amount of days and less

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have been created within [X] amount of days; Default: 7")]
	$script:UserCreatedDays = 7,
	#Get users whos passwords expire in less than X amount of days

	[Parameter(ValueFromPipeline = $true, HelpMessage = "Users password expires within [X] amount of days; Default: 7")]
	$script:DaysUntilPWExpireINT = 7,
	#Get AD Objects that have been modified in X days and newer

	[Parameter(ValueFromPipeline = $true, HelpMessage = "AD Objects that have been modified within [X] amount of days; Default: 3")]
	$script:ADModNumber = 3
	
	#CSS template located C:\Program Files\WindowsPowerShell\Modules\ReportHTML\1.4.1.1\
	#Default template is orange and named "Sample"
)


#Check for ReportHTML Module
Function Test-htmlModule {
	$script:Mod = Get-Module -ListAvailable -Name "ReportHTML"
	If (!$Mod) {
		# Write-Host "ReportHTML Module is not present, attempting to install it"
		Install-Module -Name ReportHTML -Force
		Import-Module ReportHTML -ErrorAction SilentlyContinue
	}
}
Function get-reportSettings {
	
	#                               Log
	$script:log = "$currentdir\$logdate.log"
	#                               report dir
	$script:reportdir = "$currentdir\report"

}
Function Write-ProgressHelper {
	param(
		[int]$StepNumber,
		[string]$Message
	)
	Write-Progress -Activity 'Building Report...' -Status $Message -PercentComplete (($StepNumber / $steps) * 100)
}
# Write progess to log file
Function New-LogWrite {
	Param ([string]$logstring)
	Add-content $log -value $logstring
}
# Wait Timer
Function New-WriteTime {
	Param ([Int]$time)
	Start-Sleep -Milliseconds $time
}
# Check and create for report folder
Function New-reportFolder {
	Write-ProgressHelper 1 "New-reportFolder"
	New-LogWrite "[$loggingDate]  Function New-reportFolder"
	If ($reportdir ) {
		If (Test-Path $reportdir) { 
			New-LogWrite  "[$loggingDate] report folder exists! "
		} 
		ElseIf (!(Test-Path $reportdir)) {  
			New-LogWrite  "[$loggingDate] Creating report folder "
			New-Item  $currentdir -Name "report" -ItemType "directory" | Out-File -Append -Encoding Default  $log
			If (!$?) {
				New-LogWrite  "Failed Creating report folder "
                
			}
		}
	}
	ElseIf (!$reportdir) {
                    
		New-LogWrite "[$loggingDate]  Failed New-reportFolder"
        
	}
	New-WriteTime 300
}

# Log starting Header
Function Set-Header {

	New-LogWrite  "[$loggingdate] Starting report....."
	New-LogWrite  "[$loggingdate] Report Settings....."
	New-LogWrite  "[$loggingdate] `$CompanyLogo = $CompanyLogo"
	New-LogWrite  "[$loggingdate] `$RightLogo = $RightLogo"
	New-LogWrite  "[$loggingdate] `$ReportTitle $ReportTitle"
	New-LogWrite  "[$loggingdate] `$ReportSavePath = $ReportSavePath"
	New-LogWrite  "[$loggingdate] `$Days = $Days"
	New-LogWrite  "[$loggingdate]  `$UserCreatedDays = $UserCreatedDays"
	New-LogWrite  "[$loggingdate] `$DaysUntilPWExpireINT = $DaysUntilPWExpireINT"
	New-LogWrite  "[$loggingdate] `$ADModNumber = $ADModNumber"
    
	New-WriteTime 300
}
# Convert object time
Function Set-FileTime ($FileTime) {
	New-LogWrite  "[$loggingdate] Function Set-FileTime"
	New-LogWrite  "[$loggingdate] `$FileTime $FileTime"
	$Date = [DateTime]::FromFileTime($FileTime)
	if ((!$date ) -or $Date -lt (Get-Date '1/1/1900') -or $date -eq 0) {
		'Never'
	}
	else {
		$Date
	}
}
# Default Security Groups
Function Get-DefaultSecurityGroups {
	Write-ProgressHelper 2 "Get-DefaultSecurityGroups"
	New-LogWrite  "[$loggingdate] Function Get-DefaultSecurityGroups " 
	$script:DefaultSGs = @(
		'Access Control Assistance Operators'
		'Account Operators'
		'Administrators'
		'Allowed RODC Password Replication Group'
		'Backup Operators'
		'Certificate Service DCOM Access'
		'Cert Publishers'
		'Cloneable Domain Controllers'
		'Cryptographic Operators'
		'Denied RODC Password Replication Group'
		'Distributed COM Users'
		'DnsUpdateProxy'
		'DnsAdmins'
		'Domain Admins'
		'Domain Computers'
		'Domain Controllers'
		'Domain Guests'
		'Domain Users'
		'Enterprise Admins'
		'Enterprise Key Admins'
		'Enterprise Read-only Domain Controllers'
		'Event Log Readers'
		'Group Policy Creator Owners'
		'Guests'
		'Hyper-V Administrators'
		'IIS_IUSRS'
		'Incoming Forest Trust Builders'
		'Key Admins'
		'Network Configuration Operators'
		'Performance Log Users'
		'Performance Monitor Users'
		'Print Operators'
		'Pre-Windows 2000 Compatible Access'
		'Protected Users'
		'RAS and IAS Servers'
		'RDS Endpoint Servers'
		'RDS Management Servers'
		'RDS Remote Access Servers'
		'Read-only Domain Controllers'
		'Remote Desktop Users'
		'Remote Management Users'
		'Replicator'
		'Schema Admins'
		'Server Operators'
		'Storage Replica Administrators'
		'System Managed Accounts Group'
		'Terminal Server License Servers'
		'Users'
		'Windows Authorization Access Group'
		'WinRMRemoteWMIUsers'
	)
	New-WriteTime 300
}
# Create Empty Tables
Function Get-CreateObjects {
	Write-ProgressHelper 3 "Get-CreateObjects"
	New-LogWrite  "[$loggingdate] Function Get-CreateObjects " 
	$script:Table = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:UserTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupTypetable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultGrouptable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:EnabledDisabledUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DomainAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ExpiringAccountsTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:securityeventtable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DomainTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUGPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupMembershipTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:PasswordExpirationTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:PasswordExpireSoonTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:EnterpriseAdminTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:NewCreatedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GroupProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:OUProtectionTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ADObjectTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ProtectedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputerProtectedTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:ComputersEnabledTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultComputersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:DefaultUsersinDefaultOUTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPUserTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPGroupsTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:TOPComputersTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GraphComputerOS = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:Type = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:LinkedGPOs = New-Object 'System.Collections.Generic.List[System.Object]'
	# $script:userphaventloggedonrecentlytable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$script:recentgpostable = New-Object 'System.Collections.Generic.List[System.Object]'

	New-WriteTime 300
}
# Get All AD Users
Function Get-AllUsers {
	Write-ProgressHelper 4 "Get-AllUsers"
	New-LogWrite  "[$loggingdate] Function Get-AllUsers"
	$script:AllUsers = Get-ADUser -Filter * -Properties *
	If (!$AllUsers) {
		New-LogWrite  "[$loggingdate] (!`$AllUsers) $true "
	}
	New-WriteTime 300
}
# Get all GPOs
Function Get-AllGPOs {
	Write-ProgressHelper 5 "Get-AllGPOs"
	New-LogWrite  "[$loggingdate] Function Get-AllGPOs"
	$script:GPOs = Get-GPO -All | Select-Object DisplayName, GPOStatus, ModificationTime, @{ Label = 'ComputerVersion'; Expression = { $_.computer.dsversion } }, @{ Label = 'UserVersion'; Expression = { $_.user.dsversion } }
	If (!$GPOs) {
		New-LogWrite  "[$loggingdate] (!`$GPOs) $true"
	}
	New-WriteTime 300
}
# Get AD Objects
Function Get-ADOBJ {
	Write-ProgressHelper 6 "Get-ADOBJ"
	New-LogWrite  "[$loggingdate] Function Get-ADOBJ "
	$setdate = (Get-Date).AddDays(- $ADModNumber)
	If ($setdate) {
		$script:ADObjs = Get-ADObject -Filter { whenchanged -gt $setdate -and ObjectClass -ne 'domainDNS' -and ObjectClass -ne 'rIDManager' -and ObjectClass -ne 'rIDSet' } -Properties *
		If (!$ADObjs) {
			New-LogWrite  "[$loggingdate] (!`$ADObjs) $true"
		}
	}
	New-WriteTime 300
}
# AD Objects to tables
Function Set-ADOBJ {
	Write-ProgressHelper 7 "Set-ADOBJ"
	New-LogWrite  "[$loggingdate] Function Set-ADOBJ "
	If ($ADObjs) {
		foreach ($ADObj in $ADObjs) {
			if ($ADObj.ObjectClass -eq 'GroupPolicyContainer') {
				$Name = $ADObj.DisplayName
			}
			else {
				$Name = $ADObj.Name
			}
			$adobjecttype = [PSCustomObject]@{
				'Name'         = $Name
				'Object Type'  = $ADObj.ObjectClass
				'When Changed' = $ADObj.WhenChanged
			}
			$script:ADObjectTable.Add($adobjecttype)
		}
	}
	ElseIf (!$ADObjs) {
		New-LogWrite  "[$loggingdate] (!`$ADObjs)  $true "
	}
	New-WriteTime 300
}
#  AD Trash Bin
Function Get-ADRecBin {
	Write-ProgressHelper 8 "Get-ADRecBin"
	New-LogWrite  "[$loggingdate] Function Get-ADRecBin"
	$ADRecycleBinStatus = (Get-ADOptionalFeature -Identity "Recycle Bin Feature").EnabledScopes
	If ($ADRecycleBinStatus) {
		if ($ADRecycleBinStatus.Count -lt 1) {
			$script:ADRecycleBin = 'Disabled'
		}
		else {	
			$script:ADRecycleBin = 'Enabled'
		}
	}
	ElseIf (!$ADRecycleBinStatus) {
		New-LogWrite  "[$loggingdate] (!`$ADRecycleBinStatus)  $true "
	}
	New-WriteTime 300
}
#  Get AD Domain, Forest info
Function Get-ADInfo {
	Write-ProgressHelper 9 "Get-ADInfo"
	New-LogWrite  "[$loggingdate] Function Get-ADInfo"
	$script:ADInfo = Get-ADDomain
	$script:ForestObj = Get-ADForest
	$script:DomainControllerobj = Get-ADDomain
    
	$script:Forest = $ADInfo.Forest
	$script:InfrastructureMaster = $DomainControllerobj.InfrastructureMaster
	$script:RIDMaster = $DomainControllerobj.RIDMaster
	$script:PDCEmulator = $DomainControllerobj.PDCEmulator
	$script:DomainNamingMaster = $ForestObj.DomainNamingMaster
	$script:SchemaMaster = $ForestObj.SchemaMaster


	$addomains = [PSCustomObject]@{
		'Domain'                = $Forest
		'AD Recycle Bin'        = $ADRecycleBin
		'Infrastructure Master' = $InfrastructureMaster
		'RID Master'            = $RIDMaster
		'PDC Emulator'          = $PDCEmulator
		'Domain Naming Master'  = $DomainNamingMaster
		'Schema Master'         = $SchemaMaster
	}
	$script:CompanyInfoTable.Add($addomains)
	New-WriteTime 300
}
# Get Created Users
Function Get-CreatedUsers {
	Write-ProgressHelper 10 "Get-CreatedUsers"
	New-LogWrite  "[$loggingdate] Function Get-CreatedUsers"
	# Get newly created users
	$When = ((Get-Date).AddDays(- $UserCreatedDays)).Date
	If ($When) {
		$NewUsers = $AllUsers | Where-Object { $_.whenCreated -ge $When }
		If ($NewUsers) {
			foreach ($Newuser in $Newusers) {
				$createduserss = [PSCustomObject]@{
					'Name'          = $Newuser.Name
					'Enabled'       = $Newuser.Enabled
					'Creation Date' = $Newuser.whenCreated
				}
				$script:NewCreatedUsersTable.Add($createduserss)
			}
		}
	}
	New-WriteTime 300
}
Function Get-DomainAdmins {
	Write-ProgressHelper 11 "Get-DomainAdmins"
	New-LogWrite  "[$loggingdate] Function Get-DomainAdmins"
	$script:DomainAdminMembers = Get-ADGroupMember 'Domain Admins'
	New-WriteTime 300
}
# Get Domain Admins
Function Set-DomainAdminsTable {
	Write-ProgressHelper 12 "Set-DomainAdminsTable"
	New-LogWrite  "[$loggingdate] Function Set-DomainAdminsTable"
	# Get Domain Admins
    
	If ($DomainAdminMembers) { 
		foreach ($DomainAdminMember in $DomainAdminMembers) {
			$Name = $DomainAdminMember.Name
			$Type = $DomainAdminMember.ObjectClass
			$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
			$adgroupmemebers = [PSCustomObject]@{
				'Name'    = $Name
				'Enabled' = $Enabled
				'Type'    = $Type
			}
			$script:DomainAdminTable.Add($adgroupmemebers)
			New-LogWrite  "[$loggingdate] `$adgroupmemebers $adgroupmemebers"
		}
	}
	New-WriteTime 300
}
# Get Enterprise Admins
Function Get-EntAdmins {
	Write-ProgressHelper 13 "Get-EntAdmins"
	New-LogWrite  "[$loggingdate] Function Get-EntAdmins "
	# Get Enterprise Admins
	$EnterpriseAdminsMembers = Get-ADGroupMember 'Enterprise Admins' -Server $SchemaMaster
	If ($EnterpriseAdminsMembers) {
		foreach ($EnterpriseAdminsMember in $EnterpriseAdminsMembers) {
			$Name = $EnterpriseAdminsMember.Name
			$Type = $EnterpriseAdminsMember.ObjectClass
			$Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled
			$entadmin = [PSCustomObject]@{
				'Name'    = $Name
				'Enabled' = $Enabled
				'Type'    = $Type
			}
			$script:EnterpriseAdminTable.Add($entadmin)
		}
	}
	New-WriteTime 300
}
# Get Ad Computers
Function Get-Computers {
	Write-ProgressHelper 14 "Get-Computers"
	New-LogWrite  "[$loggingdate] Function Get-Computers"
	$DefaultComputersOU = (Get-ADDomain).computerscontainer
	If ($DefaultComputersOU) {
		$DefaultComputers = Get-ADComputer -Filter * -Properties * -SearchBase "$DefaultComputersOU"
		If ($DefaultComputers) {
			foreach ($DefaultComputer in $DefaultComputers) {
				$computersenabled = [PSCustomObject]@{
					'Name'                  = $DefaultComputer.Name
					'Enabled'               = $DefaultComputer.Enabled
					'Operating System'      = $DefaultComputer.OperatingSystem
					'Modified Date'         = $DefaultComputer.Modified
					'Password Last Set'     = $DefaultComputer.PasswordLastSet
					'Protect from Deletion' = $DefaultComputer.ProtectedFromAccidentalDeletion
				}
				$script:DefaultComputersinDefaultOUTable.Add($computersenabled)
			}
		}
	}
	New-WriteTime 300
}
# AD Users
Function Get-ADUsers {
	Write-ProgressHelper 15 "Get-ADUsers"
	New-LogWrite "[$loggingDate]  Function Get-ADUsers "
	$DefaultUsersOU = (Get-ADDomain).UsersContainer
	If ($DefaultUsersOU ) {
		# $DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, @{ Name = 'lastlogon'; Expression = { $_.lastlogon } }, DistinguishedName
		$DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, DistinguishedName, LastLogon # @{ Name = 'lastlogon'; Expression = { Set-FileTime $_.lastlogon } }, 
		#$DefaultUsers = $Allusers | Where-Object { $_.DistinguishedName -like "*$($DefaultUsersOU)" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, DistinguishedName, lastlogon
		If ($DefaultUsers) {
			foreach ($DefaultUser in $DefaultUsers) {
				# $DefaultUserlogon = $DefaultUser | Set-FileTime $_.lastlogon
				New-LogWrite "[$loggingDate]  $DefaultUser "
				$DefaultUserlogon = $DefaultUser.Lastlogon 
				$DefaultUserlogon = Set-FileTime $DefaultUserlogon
				New-LogWrite "[$loggingDate]  $DefaultUserlogon "
				$aduserobject = [PSCustomObject]@{
					'Name'                    = $DefaultUser.Name
					'UserPrincipalName'       = $DefaultUser.UserPrincipalName
					'Enabled'                 = $DefaultUser.Enabled
					'Protected from Deletion' = $DefaultUser.ProtectedFromAccidentalDeletion
					'Last Logon'              = $DefaultUserlogon
					'Email Address'           = $DefaultUser.EmailAddress
				}
				$script:DefaultUsersinDefaultOUTable.Add($aduserobject)
			}
		}
	}
	ElseIf (!$DefaultUsersOU) {
		New-LogWrite "[$loggingDate] Failed: `$DefaultUsersOU $DefaultUsersOU"
	}
	New-WriteTime 300
}
# Expired Accounts
Function Get-ExpiredAccounts {
	Write-ProgressHelper 16 "Get-ExpiredAccounts"
	New-LogWrite "[$loggingDate]  Function Get-ExpiredAccounts "
	# Expiring Accounts
	$ExpiringAccounts = Search-ADAccount -AccountExpiring -UsersOnly
	If ($ExpiringAccounts) {
		foreach ($ExpiringAccount in $ExpiringAccounts) {
			$NameExpiringAccounts = $ExpiringAccount.Name
			$UPNExpiringAccounts = $ExpiringAccount.UserPrincipalName
			$ExpirationDate = $ExpiringAccount.AccountExpirationDate
			$enabled = $ExpiringAccount.Enabled
			$expaccount = [PSCustomObject]@{
		
				'Name'              = $NameExpiringAccounts
				'UserPrincipalName' = $UPNExpiringAccounts
				'Expiration Date'   = $ExpirationDate
				'Enabled'           = $enabled
			}
			$script:ExpiringAccountsTable.Add($expaccount)
		}
	}
	New-WriteTime 300
}
# Security Logs
Function Get-Seclogs {
	Write-ProgressHelper 17 "Get-Seclogs"
	New-LogWrite "[$loggingDate]  Function Get-Seclogs "
	# Security Logs
	$SecurityLogs = Get-EventLog -Newest 7 -LogName 'Security' | Where-Object { $_.Message -like '*An account*' }
	If ($SecurityLogs) {
		foreach ($SecurityLog in $SecurityLogs) {
			$TimeGenerated = $SecurityLog.TimeGenerated
			$EntryType = $SecurityLog.EntryType
			$Recipient = $SecurityLog.Message
			$securitylogss = [PSCustomObject]@{
				'Time'    = $TimeGenerated
				'Type'    = $EntryType
				'Message' = $Recipient
			}
			$script:SecurityEventTable.Add($securitylogss)
		}
	}
	New-WriteTime 300
}
# Domains
Function Get-Domains {
	Write-ProgressHelper 18 "Get-Domains"
	New-LogWrite "[$loggingDate]  Function Get-Domains "
	# Tenant Domain
	$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes
	If ($Domains) {
		ForEach ($Domain  in $Domains) {
			$domainforest = [PSCustomObject]@{
				'UPN Suffixes' = $Domain
				'True'    = $true  
			}
			$script:DomainTable.Add($domainforest)
			New-LogWrite "[$loggingDate]  `DomainTable $DomainTable"
		}
	}
	New-WriteTime 300
}
# Buld Group tables
Function Get-Groups {
	Write-ProgressHelper 19 "Get-Groups"
	New-LogWrite "[$loggingDate]  Function Get-Groups "

	$SecurityCount = 0
	$MailSecurityCount = 0
	$CustomGroup = 0
	$DefaultGroup = 0
	$Groupswithmemebrship = 0
	$Groupswithnomembership = 0
	$GroupsProtected = 0
	$GroupsNotProtected = 0

	# Get groups and get counts on Distribution, Non-Mail enabled Security, and Mail-Enabled groups 

	$Groups = Get-ADGroup -Filter * -Properties * #Initial gathering of AD groups.
	$DistroCount = ($Groups | Where-Object { $_.GroupCategory -eq 'Distribution' }).Count # Grab all Distribution Groups and count.
	$MailSecurityCount = ($Groups | Where-Object { $_.GroupCategory -eq 'Security' -and $null -ne $_.mail}).Count
	$SecurityCount = ($Groups | Where-Object { $_.GroupCategory -eq 'Security' -and $null -eq $_.mail}).Count

	foreach ($Group in $Groups) {
		$DefaultADGroup = 'False'						# Reset DefaultADGroup Variable
		$Manager = "" 								# Reset Manager Variable.	
		$Gemail = (Get-ADGroup $Group -Properties mail).mail			# $Gemail - Grab the E-mail address of a group.

		If ($Gemail -and $group.GroupCategory -eq 'Security') {			# If an E-mail address exists and if the group is a Security group.
			$Type = 'Mail-Enabled Security Group'				# 	Assign type to Mail-enabled security group
		}
		elseIf ($Gemail -and $group.GroupCategory -eq 'Distribution') {		# If an E-mail address exists and if the group is a Distribution group.
				$Type = 'Distribution Group'				#	Assign type to distibution group
		}
		elseIf (!$Gemail -and $group.GroupCategory -eq 'Security') {		# If an E-mail address doesn't exist and if the group is a security group.
				$Type = 'Security Group'				#	Assign type to (Non-mail enabled) Security group.
		}


		if ($Group.ProtectedFromAccidentalDeletion) {				# If the group is protected from accidental deletion
			$GroupsProtected++						# 	increase the protected count by one.
		}
		else {
			$GroupsNotProtected++						# Else increase the "not protected" count by one.
		}


		if ($DefaultSGs -contains $Group.Name) {				# Check the current group being checked against the list of default groups
			$DefaultADGroup = 'True'					#	If there's a match, group is a default group
			$DefaultGroup++							#	And increase the default group counter by one.
		}
		else {									# 	Else it's not a default group (aka custom)
			$CustomGroup++
		}


		if ($Group.Name -ne 'Domain Users') {					# Exclude the "Domain Users" group
			$Users = (Get-ADGroupMember -Identity $Group | Sort-Object DisplayName | Select-Object -ExpandProperty Name) -join ', ' # Concatinate and format a list of group members

			if (!$Users) {							# If there are not users
				$Groupswithnomembership++				#	Increase the count of groups with no users by one.
			}
			else {								# Else
				$Groupswithmemebrship++					#	Increase the count of group with users by one.
			}
		}


		$OwnerDN = Get-ADGroup -Filter { name -eq $Group.Name } -Properties managedBy | Select-Object -ExpandProperty ManagedBy #Grabs Managedby property of group.
			Try {
				$Manager = Get-ADUser -Filter { distinguishedname -like $OwnerDN } | Select-Object -ExpandProperty Name #Converts Managedby property to a name.
			}
			Catch {
				$groupname = $group.Name
				New-LogWrite "[$loggingDate]  Manager attribute:$Manager  on the group  $groupname  missing"
			}
		# $Manager = $AllUsers | Where-Object { $_.distinguishedname -eq $OwnerDN } | Select-Object -ExpandProperty Name


		$adgroupobject = [PSCustomObject]@{
		
			'Name'                    = $Group.name
			'Type'                    = $Type
			'Members'                 = $users
			'Managed By'              = $Manager
			'E-mail Address'          = $GEmail
			'Protected from Deletion' = $Group.ProtectedFromAccidentalDeletion
			'Default AD Group'        = $DefaultADGroup
		}
		$script:table.Add($adgroupobject)
	}
	# TOP groups table
	$objectmailgroups = [PSCustomObject]@{
		'Total Groups'                 = $Groups.Count
		'Mail-Enabled Security Groups' = $MailSecurityCount
		'Security Groups'              = $SecurityCount
		'Distribution Groups'          = $DistroCount
	}
	$script:TOPGroupsTable.Add($objectmailgroups)
    
	# Default Group Type Pie Chart
	$objectmailgroupssec = [PSCustomObject]@{
		'Name'  = 'Mail-Enabled Security Groups'
		'Count' = $MailSecurityCount
	}
	$script:GroupTypetable.Add($objectmailgroupssec)

	$secgroups = [PSCustomObject]@{
		'Name'  = 'Security Groups'
		'Count' = $SecurityCount
	}
	$script:GroupTypetable.Add($secgroups)

	$distgroups = [PSCustomObject]@{
		'Name'  = 'Distribution Groups'
		'Count' = $DistroCount
	}
	$script:GroupTypetable.Add($distgroups)

	# Default Group Pie Chart
	$defaultgroups = [PSCustomObject]@{
		'Name'  = 'Default Groups'
		'Count' = $DefaultGroup
	}
	$script:DefaultGrouptable.Add($defaultgroups)

	$customgroups = [PSCustomObject]@{
		'Name'  = 'Custom Groups'
		'Count' = $CustomGroup
	}
	$script:DefaultGrouptable.Add($customgroups)
	# Group Protection Pie Chart
	$protectedgroups = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $GroupsProtected
	}
	$script:GroupProtectionTable.Add($protectedgroups)

	$notprotectedgroups = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $GroupsNotProtected
	}
	$script:GroupProtectionTable.Add($notprotectedgroups)

	# Groups with membership vs no membership pie chart
	$groupwithmembers = [PSCustomObject]@{
		'Name'  = 'With Members'
		'Count' = $Groupswithmemebrship
	}
	$script:GroupMembershipTable.Add($groupwithmembers)

	$nogroupwithmembers = [PSCustomObject]@{
		'Name'  = 'No Members'
		'Count' = $Groupswithnomembership
	}
	$script:GroupMembershipTable.Add($nogroupwithmembers)
	New-WriteTime 300
}
# Build OU tables
Function Get-OU {
	Write-ProgressHelper 20 "Get-OU"
	New-LogWrite "[$loggingDate]  Function Get-OU  "

	$OUwithLinked = 0
	$OUwithnoLink = 0
	$OUProtected = 0
	$OUNotProtected = 0
	$OUs = Get-ADOrganizationalUnit -Filter * -Properties *
	foreach ($OU in $OUs) {
		if (($OU.linkedgrouppolicyobjects).length -lt 1) {
			$LinkedGPOs = 'None'
			$OUwithnoLink++
		}
		else {
			$OUwithLinked++
			$GPOslinks = $OU.linkedgrouppolicyobjects
			foreach ($GPOlink in $GPOslinks) {
				$Split1 = $GPOlink -split '{' | Select-Object -Last 1
				$Split2 = $Split1 -split '}' | Select-Object -First 1
				# $LinkedGPOs.Add((Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName)
				$LinkedGPOs += (Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName
				# $LinkedGPOs.Add($OUs.linkedgrouppolicyobjects.count)

			}
		}
		if ($OU.ProtectedFromAccidentalDeletion -eq $True) {
			$OUProtected++
		}
		else {
			$OUNotProtected++
		}
		$LinkedGPOs = $LinkedGPOs -join ', '
		$linkedgpoobjects = [PSCustomObject]@{
			'Name'                    = $OU.Name
			'Linked GPOs'             = $LinkedGPOs
			'Modified Date'           = $OU.WhenChanged
			'Protected from Deletion' = $OU.ProtectedFromAccidentalDeletion
		}
		$script:OUTable.Add($linkedgpoobjects)
	}
	# OUs with no GPO Linked
	$nolinkedous = [PSCustomObject]@{
		'Name'  = "OUs with no GPO's linked"
		'Count' = $OUwithnoLink
	}
	$script:OUGPOTable.Add($nolinkedous)
	$linkedous = [PSCustomObject]@{
		'Name'  = "OUs with GPO's linked"
		'Count' = $OUwithLinked
	}
	$script:OUGPOTable.Add($linkedous)
	# OUs Protected Pie Chart
	$protectedous = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $OUProtected
	}
	$script:OUProtectionTable.Add($protectedous)
	$notprotectedous = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $OUNotProtected
	}
	$script:OUProtectionTable.Add($notprotectedous)
	New-WriteTime 300

}
# Build User Tables - Protected and pwd expiring
Function Get-Users {
	Write-ProgressHelper 21 "Get-Users"
	New-LogWrite "[$loggingDate]  Function Get-Users "

	$UserEnabled = 0
	$UserDisabled = 0
	$UserPasswordExpires = 0
	$UserPasswordNeverExpires = 0
	$ProtectedUsers = 0
	$NonProtectedUsers = 0
	## $UsersWIthPasswordsExpiringInUnderAWeek = 0
	## $UsersNotLoggedInOver30Days = 0
	## $AccountsExpiringSoon = 0
	## Get users that haven't logged on in X amount of days, var is set at start of script
	foreach ($User in $AllUsers) {
		# $AttVar = $User | Select-Object Enabled, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, Name, SamAccountName, EmailAddress, AccountExpirationDate, @{ Name = 'lastlogon'; Expression = { $_.lastlogon } }, DistinguishedName
		$AttVar = $User | Select-Object Enabled, PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, Name, SamAccountName, EmailAddress, AccountExpirationDate, DistinguishedName, LastLogon
		$defaultuserlastlogon = $User.lastlogon
		$defaultuserlastlogon = Set-FileTime $defaultuserlastlogon 
		New-LogWrite "[$loggingDate] `$defaultuserlastlogon $defaultuserlastlogon "
		$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
		if ((($AttVar.PasswordNeverExpires) -eq $False) -and (($AttVar.Enabled) -ne $false)) {
			# Get Password last set date
			$passwordSetDate = ($User | ForEach-Object { $_.PasswordLastSet })
			if (!$passwordSetDate) {
				$daystoexpire = 'User has never logged on'
			}
			else {
				# Check for Fine Grained Passwords
				$PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
				if ($PasswordPol) {
					$maxPasswordAge = ($PasswordPol).MaxPasswordAge
				}
				$expireson = $passwordsetdate.AddDays($maxPasswordAge)
				$today = (Get-Date)
				# Gets the count on how many days until the password expires and stores it in the $daystoexpire var
				$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
			}
		}
		else {
			$daystoexpire = 'N/A'
		}
		New-LogWrite "((`"$defaultuserlastlogon`" -like `"Never`") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))"
		If ($defaultuserlastlogon) {
			New-LogWrite "((`"$defaultuserlastlogon`" -like `"Never`") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))"
			If (("$defaultuserlastlogon" -like "Never") -or ($($User.Enabled) -eq $True -and $defaultuserlastlogon -lt $((Get-Date).AddDays(- $Days)))) {
				New-LogWrite "[$loggingDate]  If ($User -and $AttVar -and $Days )"
				New-LogWrite "[$loggingDate] psobject lastlogon =  $defaultuserlastlogon"
				if (($User.Enabled -eq $True)) {
					$objlastlogonobject = [PSCustomObject]@{
						'Name'                        = $User.Name
						'UserPrincipalName'           = $User.UserPrincipalName
						'Enabled'                     = $AttVar.Enabled
						'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
						'Last Logon'                  = $defaultuserlastlogon
						'Password Never Expires'      = $AttVar.PasswordNeverExpires
						'Days Until Password Expires' = $daystoexpire
					}
					$script:userphaventloggedonrecentlytable.Add($objlastlogonobject)
					New-LogWrite "`$userphaventloggedonrecentlytable $userphaventloggedonrecentlytable"
				}
			}
		}
		# Items for protected vs non protected users
		if ($User.ProtectedFromAccidentalDeletion -eq $False) {
			$NonProtectedUsers++
		}
		else {
			$ProtectedUsers++
		}
		# Items for the enabled vs disabled users pie chart
		if (($AttVar.PasswordNeverExpires) -ne $false) {
			$UserPasswordNeverExpires++
		}
		else {
			$UserPasswordExpires++
		}
		# Items for password expiration pie chart
		if (($AttVar.Enabled) -ne $false) {
			$UserEnabled++
		}
		else {
			$UserDisabled++
		}
		$Name = $User.Name
		$UPN = $User.UserPrincipalName
		$Enabled = $AttVar.Enabled
		$EmailAddress = $AttVar.EmailAddress
		$AccountExpiration = $AttVar.AccountExpirationDate
		$PasswordExpired = $AttVar.PasswordExpired
		$PasswordLastSet = $AttVar.PasswordLastSet
		$PasswordNeverExpires = $AttVar.PasswordNeverExpires
		# $daysUntilPWExpire = $daystoexpire
		$adobjectouser = [PSCustomObject]@{
			'Name'                        = $Name
			'UserPrincipalName'           = $UPN
			'Enabled'                     = $Enabled
			'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
			'Last Logon'                  = $defaultuserlastlogon
			# 'Last Logon'                  = $LastLogon
			'Email Address'               = $EmailAddress
			'Account Expiration'          = $AccountExpiration
			'Change Password Next Logon'  = $PasswordExpired
			'Password Last Set'           = $PasswordLastSet
			'Password Never Expires'      = $PasswordNeverExpires
			'Days Until Password Expires' = $daystoexpire
		}
		$script:usertable.Add($adobjectouser)
		if ($daystoexpire -lt $DaysUntilPWExpireINT) {
			$noadobjectouser = [PSCustomObject]@{
				'Name'                        = $Name
				'Days Until Password Expires' = $daystoexpire
			}
			$script:PasswordExpireSoonTable.Add($noadobjectouser)

		}
	}

	# Data for users enabled vs disabled pie graph
	$userenabled = [PSCustomObject]@{
		'Name'  = 'Enabled'
		'Count' = $UserEnabled
	}
	$script:EnabledDisabledUsersTable.Add($userenabled)

	$userdisabled = [PSCustomObject]@{
		'Name'  = 'Disabled'
		'Count' = $UserDisabled
	}
	$script:EnabledDisabledUsersTable.Add($userdisabled)

	# Data for users password expires pie graph
	$userpasswordexp = [PSCustomObject]@{
		'Name'  = 'Password Expires'
		'Count' = $UserPasswordExpires
	}
	$script:PasswordExpirationTable.Add($userpasswordexp)

	$userpasswordneverexp = [PSCustomObject]@{
		'Name'  = 'Password Never Expires'
		'Count' = $UserPasswordNeverExpires
	}
	$script:PasswordExpirationTable.Add($userpasswordneverexp)

	#Data for protected users pie graph
	$protecteduser = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $ProtectedUsers
	}
	$script:ProtectedUsersTable.Add($protecteduser)

	$notprotecteduser = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $NonProtectedUsers
	}
	$script:ProtectedUsersTable.Add($notprotecteduser)

	if ($userphaventloggedonrecentlytable.count -eq 0) {
		# if (($userphaventloggedonrecentlytable).Information) {
		$script:UHLONXD = '0'
	}
	Else {
		$script:UHLONXD = $userphaventloggedonrecentlytable.Count
	}
	# TOP User table
	If ($ExpiringAccounts) {
		# If (($ExpiringAccounts).Information) {
		$expaccounts = [PSCustomObject]@{
			'Total Users'                                                           = $AllUsers.Count
			"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
			'Expiring Accounts'                                                     = $ExpiringAccounts.Count
			"Users Haven't Logged on in $Days Days or more"                         = $UHLONXD
		}
		$script:TOPUserTable.Add($expaccounts)
	}
	Else {
		$expaccounts = [PSCustomObject]@{
			'Total Users'                                                           = $AllUsers.Count
			"Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.Count
			'Expiring Accounts'                                                     = '0'
			"Users Haven't Logged on in $Days Days or more"                         = $UHLONXD
		}
		$script:TOPUserTable.Add($expaccounts)
	}
	New-WriteTime 300
}
# GPO Tables
Function Get-GPOs {
	Write-ProgressHelper 22 "Get-GPOs"
	New-LogWrite "[$loggingDate]  Get-GPOs "
	foreach ($GPO in $GPOs) {
		$gpoobject = [PSCustomObject]@{
			'Name'             = $GPO.DisplayName
			'Status'           = $GPO.GpoStatus
			'Modified Date'    = $GPO.ModificationTime
			'User Version'     = $GPO.UserVersion
			'Computer Version' = $GPO.ComputerVersion
		}
		$script:GPOTable.Add($gpoobject)
	}
	New-WriteTime 300
}
# Recent GPOs Table
Function Get-RecentGPOs {
	Write-ProgressHelper 22 "Get-RecentGPOs"
	New-LogWrite "[$loggingDate]  Function Get-RecentGPOs"
	$createdinthelast = ((Get-Date).AddDays( - 30)).Date
	$grouppolicyrecent = Get-GPO -all | Select-Object DisplayName, GPOStatus, CreationTime , @{ Label = 'ComputerVersion'; Expression = { $_.computer.dsversion } }, @{ Label = 'UserVersion'; Expression = { $_.user.dsversion } } 
	foreach ($grouppolicyrecent in $grouppolicyrecent) {
		$grouppolicyrecenttest5 = $grouppolicyrecent.CreationTime
		If (($grouppolicyrecenttest5 -gt $createdinthelast)) {
			$gpocreatedobject = [PSCustomObject]@{
				'Name'             = $grouppolicyrecent.DisplayName
				'Status'           = $grouppolicyrecent.GpoStatus
				'Created Date'     = $grouppolicyrecenttest5
				'User Version'     = $grouppolicyrecent.UserVersion
				'Computer Version' = $grouppolicyrecent.ComputerVersion
			}
			$script:recentgpostable.Add($gpocreatedobject)
		}
	}
	New-WriteTime 300
}
# AD Computers tables - Protected
Function Get-Adcomputers {
	Write-ProgressHelper 23 "Get-Adcomputers"
	New-LogWrite "[$loggingDate]  Function Get-Adcomputers "
	$Computers = Get-ADComputer -Filter * -Properties *

	$ComputersProtected = 0
	$ComputersNotProtected = 0
	$ComputerEnabled = 0
	$ComputerDisabled = 0
	# Only search for versions of windows that exist in the Environment
	$WindowsRegex = '(Windows (Server )?(\d+|XP)?( R2)?).*'
	$OsVersions = $Computers | Select-Object OperatingSystem -unique | ForEach-Object {
		if ($_.OperatingSystem -match $WindowsRegex ) { 
			return $matches[1]
		}
		elseif ($_.OperatingSystem) {
			return $_.OperatingSystem
		}
	} | Select-Object -unique | Sort-Object
	$OsObj = [PSCustomObject]@{ }
	$OsVersions | ForEach-Object {
		$OsObj | Add-Member -Name $_ -Value 0 -Type NoteProperty
	}
	foreach ($Computer in $Computers) {
		if ($Computer.ProtectedFromAccidentalDeletion -eq $True) {
			$ComputersProtected++
		}
		else {
			$ComputersNotProtected++
		}
		if ($Computer.Enabled -eq $True) {
			$ComputerEnabled++
		}
		else {
			$ComputerDisabled++
		}
		$computerlastlogin = $Computer.lastLogonTimestamp
		$computerlastlogin = Set-FileTime $computerlastlogin
		$computerobjects = [PSCustomObject]@{
			'Name'                  = $Computer.Name
			'Enabled'               = $Computer.Enabled
			'Operating System'      = $Computer.OperatingSystem
			'Modified Date'         = $Computer.Modified
			'Last Login'            = $computerlastlogin
			'Password Last Set'     = $Computer.PasswordLastSet
			'Protect from Deletion' = $Computer.ProtectedFromAccidentalDeletion
		}
		$script:ComputersTable.Add($computerobjects)
		if ($Computer.OperatingSystem -match $WindowsRegex) {
			$OsObj."$($matches[1])"++
		}
	}
	#Pie chart breaking down OS for computer obj
	$OsObj.PSObject.Properties | ForEach-Object {
		$script:GraphComputerOS.Add([PSCustomObject]@{'Name' = $_.Name; 'Count' = $_.Value })
	}
	#Data for TOP Computers data table
	$OsObj | Add-Member -Name 'Total Computers' -Value $Computers.Count -Type NoteProperty
	$script:TOPComputersTable.Add($OsObj)
	#Data for protected Computers pie graph
	$protectedcomputers = [PSCustomObject]@{
		'Name'  = 'Protected'
		'Count' = $ComputersProtected
	}
	$script:ComputerProtectedTable.Add($protectedcomputers)

	$notprotectedcomputers = [PSCustomObject]@{
		'Name'  = 'Not Protected'
		'Count' = $ComputersNotProtected
	}
	$script:ComputerProtectedTable.Add($notprotectedcomputers)

	#Data for enabled/vs Computers pie graph
	$computersenabled = [PSCustomObject]@{
		'Name'  = 'Enabled'
		'Count' = $ComputerEnabled
	}
	$script:ComputersEnabledTable.Add($computersenabled)

	$computersdisabled = [PSCustomObject]@{
		'Name'  = 'Disabled'
		'Count' = $ComputerDisabled
	}
	$script:ComputersEnabledTable.Add($computersdisabled)
	New-WriteTime 300
}
# Checking for Empty Tables
Function compare-tables {
	Write-ProgressHelper 24 "compare-tables"
	New-LogWrite "[$loggingDate]  Function compare-tables"
	if (!$PasswordExpireSoonTable) {
		$pwdexp = [PSCustomObject]@{
			Information = ' No users were found to have passwords expiring soon'
		}
		$script:PasswordExpireSoonTable.Add($pwdexp)
	}
	if (!$userphaventloggedonrecentlytable) {
		$noobjlastlogonobject = [PSCustomObject]@{
			Information = " No Users were found to have not logged on in $Days days or more"
		}
		$script:userphaventloggedonrecentlytable.Add($noobjlastlogonobject)
	}
	if (!$ADObjectTable ) {
		$noadobjecttype = [PSCustomObject]@{
			Information = ' No AD Objects have been modified recently'
		}
		$script:ADObjectTable.Add($noadobjecttype)
	}
	if (!$CompanyInfoTable) {
		$noaddomains = [PSCustomObject]@{
			Information = ' Could not get items for table'
		}
		$script:CompanyInfoTable.Add($noaddomains)
	}
	if (!$NewCreatedUsersTable ) {
		$nocreateduserss = [PSCustomObject]@{
			Information = ' No new users have been recently created'
		}
		$script:NewCreatedUsersTable.Add($nocreateduserss)
	}
	if (!$DomainAdminTable) {
		$noadgroupmemebers = [PSCustomObject]@{
			Information = ' No Domain Admin Members were found'
		}
		$script:DomainAdminTable.Add($noadgroupmemebers)
	}
	if (!$EnterpriseAdminTable) {
		$noentadmin = [PSCustomObject]@{
			Information = ' Enterprise Admin members were found'
		}
		$script:EnterpriseAdminTable.Add($noentadmin)
	}
	if (!$DefaultComputersinDefaultOUTable ) {
		$computersdisabled = [PSCustomObject]@{
			Information = ' No computers were found in the Default OU'
		}
		$script:DefaultComputersinDefaultOUTable.Add($computersdisabled)
	}
	if (!$DefaultUsersinDefaultOUTable ) {
		$noaduserobject = [PSCustomObject]@{
			Information = ' No Users were found in the default OU'
		}
		$script:DefaultUsersinDefaultOUTable.Add($noaduserobject)
	}
	if (!$SecurityEventTable) {
		$nosecuritylogss = [PSCustomObject]@{
			Information = ' No logon security events were found'
		}
		$script:SecurityEventTable.Add($nosecuritylogss)
	}
	if (!$DomainTable ) {
		$nodomainforest = [PSCustomObject]@{
			Information = ' No UPN Suffixes were found'
		}
		$script:DomainTable.Add($nodomainforest)
	}
	if (!$table ) {
		$noadgroupobject = [PSCustomObject]@{
			Information = ' No Groups were found'
		}
		$script:table.Add($noadgroupobject)
	}
	if (!$usertable ) {
		$adobjectousers = [PSCustomObject]@{
			Information = ' No users were found'
		}
		$script:usertable.Add($adobjectousers)
	}
	if (!$OUTable ) {
		$nolinkedgpoobjects = [PSCustomObject]@{
			Information = ' No OUs were found'
		}
		$script:OUTable.Add($nolinkedgpoobjects)
	}
	if (!$GPOTable ) {
		$gpoobject = [PSCustomObject]@{
			Information = ' No Group Policy Obejects were found'
		}
		$script:GPOTable.Add($gpoobject)
	}
	if (!$ComputersTable ) {
		$compuerobjects = [PSCustomObject]@{
			Information = ' No computers were found'
		}
		$script:ComputersTable.Add($compuerobjects)
	}
	If (!$ExpiringAccountsTable ) {
		$noexpaccount = [PSCustomObject]@{
			Information = ' No Users were found to expire soon' 
		}
		$script:ExpiringAccountsTable.Add($noexpaccount)
	}
	If (!$recentgpostable ) {
		$norecentgpostable = [PSCustomObject]@{
			Information = ' No GPOs were created recently' 
		}
		$script:recentgpostable.Add($norecentgpostable)
	}
	New-WriteTime 300
}
# Build Report
Function Get-Report {
	Write-ProgressHelper 25 "Get-Report"
	New-LogWrite "[$loggingDate]  Function Get-Report"

	$tabarray = @('Dashboard', 'Groups', 'Organizational Units', 'Users', 'Group Policy', 'Computers')


	##--OU Protection PIE CHART--##
	#  Basic Properties 
	$PO12 = Get-HTMLPieChartObject
	$PO12.Title = 'Organizational Units Protected from Deletion'
	$PO12.Size.Height = 250
	$PO12.Size.width = 250
	$PO12.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PO12.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PO12.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PO12.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PO12.DataDefinition.DataNameColumnName = 'Name'
	$PO12.DataDefinition.DataValueColumnName = 'Count'

	##--Computer OS Breakdown PIE CHART--##
	$PieObjectComputerObjOS = Get-HTMLPieChartObject
	$PieObjectComputerObjOS.Title = 'Computer Operating Systems'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'Random'

	##--Computers Protection PIE CHART--##
	#  Basic Properties 
	$PieObjectComputersProtected = Get-HTMLPieChartObject
	$PieObjectComputersProtected.Title = 'Computers Protected from Deletion'
	$PieObjectComputersProtected.Size.Height = 250
	$PieObjectComputersProtected.Size.width = 250
	$PieObjectComputersProtected.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectComputersProtected.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectComputersProtected.DataDefinition.DataValueColumnName = 'Count'

	#  #--Computers Enabled PIE CHART--##
	#  Basic Properties 
	$PieObjectComputersEnabled = Get-HTMLPieChartObject
	$PieObjectComputersEnabled.Title = 'Computers Enabled vs Disabled'
	$PieObjectComputersEnabled.Size.Height = 250
	$PieObjectComputersEnabled.Size.width = 250
	$PieObjectComputersEnabled.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectComputersEnabled.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectComputersEnabled.DataDefinition.DataValueColumnName = 'Count'

	##--USERS Protection PIE CHART--##
	#  Basic Properties 
	$PieObjectProtectedUsers = Get-HTMLPieChartObject
	$PieObjectProtectedUsers.Title = 'Users Protected from Deletion'
	$PieObjectProtectedUsers.Size.Height = 250
	$PieObjectProtectedUsers.Size.width = 250
	$PieObjectProtectedUsers.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectProtectedUsers.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectProtectedUsers.DataDefinition.DataValueColumnName = 'Count'

	#  Basic Properties 
	$PieObjectOUGPOLinks = Get-HTMLPieChartObject
	$PieObjectOUGPOLinks.Title = 'OU GPO Links'
	$PieObjectOUGPOLinks.Size.Height = 250
	$PieObjectOUGPOLinks.Size.width = 250
	$PieObjectOUGPOLinks.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectOUGPOLinks.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectOUGPOLinks.DataDefinition.DataValueColumnName = 'Count'

	#  Basic Properties 
	$PieObject4 = Get-HTMLPieChartObject
	$PieObject4.Title = 'Office 365 Unassigned Licenses'
	$PieObject4.Size.Height = 250
	$PieObject4.Size.width = 250
	$PieObject4.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObject4.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObject4.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObject4.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObject4.DataDefinition.DataNameColumnName = 'Name'
	$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

	#  Basic Properties 
	$PieObjectGroupType = Get-HTMLPieChartObject
	$PieObjectGroupType.Title = 'Group Types'
	$PieObjectGroupType.Size.Height = 250
	$PieObjectGroupType.Size.width = 250
	$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'

	#  Pie Chart Groups with members vs no members
	$PieObjectGroupMembersType = Get-HTMLPieChartObject
	$PieObjectGroupMembersType.Title = 'Group Membership'
	$PieObjectGroupMembersType.Size.Height = 250
	$PieObjectGroupMembersType.Size.width = 250
	$PieObjectGroupMembersType.ChartStyle.ChartType = 'doughnut'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'ColorScheme1'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'Generated1'
	$PieObjectGroupMembersType.ChartStyle.ColorSchemeName = 'Random'
	$PieObjectGroupMembersType.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupMembersType.DataDefinition.DataValueColumnName = 'Count'

	#  Basic Properties 
	$PieObjectGroupType2 = Get-HTMLPieChartObject
	$PieObjectGroupType2.Title = 'Custom vs Default Groups'
	$PieObjectGroupType2.Size.Height = 250
	$PieObjectGroupType2.Size.width = 250
	$PieObjectGroupType2.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'ColorScheme1'

	# There are 8 generated schemes, randomly generated at runtime 
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Generated1'

	# you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectGroupType.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

	##--Enabled users vs Disabled Users PIE CHART--##
	#  Basic Properties 
	$EnabledDisabledUsersPieObject = Get-HTMLPieChartObject
	$EnabledDisabledUsersPieObject.Title = 'Enabled vs Disabled Users'
	$EnabledDisabledUsersPieObject.Size.Height = 250
	$EnabledDisabledUsersPieObject.Size.width = 250
	$EnabledDisabledUsersPieObject.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$EnabledDisabledUsersPieObject.DataDefinition.DataNameColumnName = 'Name'
	$EnabledDisabledUsersPieObject.DataDefinition.DataValueColumnName = 'Count'

	##--PasswordNeverExpires PIE CHART--##
	#  Basic Properties 
	$PWExpiresUsersTable = Get-HTMLPieChartObject
	$PWExpiresUsersTable.Title = 'Password Expiration'
	$PWExpiresUsersTable.Size.Height = 250
	$PWExpiresUsersTable.Size.Width = 250
	$PWExpiresUsersTable.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PWExpiresUsersTable.DataDefinition.DataNameColumnName = 'Name'
	$PWExpiresUsersTable.DataDefinition.DataValueColumnName = 'Count'

	##--Group Protection PIE CHART--##
	#   Basic Properties 
	$PieObjectGroupProtection = Get-HTMLPieChartObject
	$PieObjectGroupProtection.Title = 'Groups Protected from Deletion'
	$PieObjectGroupProtection.Size.Height = 250
	$PieObjectGroupProtection.Size.width = 250
	$PieObjectGroupProtection.ChartStyle.ChartType = 'doughnut'

	#  These file exist in the module directoy, There are 4 schemes by default
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'ColorScheme1'

	#  There are 8 generated schemes, randomly generated at runtime 
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'Generated1'

	#  you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'Random'

	#  Data defintion you can reference any column from name and value from the  dataset.  
	#  Name and Count are the default to work with the Group function.
	$PieObjectGroupProtection.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectGroupProtection.DataDefinition.DataValueColumnName = 'Count'

	#  Dashboard Report
	$FinalReport = New-Object 'System.Collections.Generic.List[System.Object]'
	$FinalReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo))
	$FinalReport.Add($(Get-HTMLTabHeader -TabNames $tabarray))
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[0] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Company Information'))
	$FinalReport.Add($(Get-HTMLContentTable $CompanyInfoTable))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Groups'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DomainAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Objects in Default OUs'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Computers'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DefaultComputersinDefaultOUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DefaultUsersinDefaultOUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "AD Objects Modified in Last $ADModNumber Days "))
	$FinalReport.Add($(Get-HTMLContentDataTable $ADObjectTable))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Expiring Items'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText " Users with Passwords Expiring in less than $DaysUntilPWExpireINT days "))
	$FinalReport.Add($(Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days or more"))
	$FinalReport.Add($(Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Security Logs'))
	$FinalReport.Add($(Get-HTMLContentDataTable $securityeventtable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'UPN Suffixes'))
	$FinalReport.Add($(Get-HTMLContentTable $DomainTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	# Groups Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[1] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Groups Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPGroupsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Groups'))
	$FinalReport.Add($(Get-HTMLContentDataTable $Table -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumn1of2))

	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $DomainAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'))
	$FinalReport.Add($(Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Groups Chart'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupType2 -DataSet $DefaultGrouptable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupMembersType -DataSet $GroupMembershipTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 4 -ColumnCount 4))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectGroupProtection -DataSet $GroupProtectionTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#  Organizational Unit Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[2] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Organizational Units'))
	$FinalReport.Add($(Get-HTMLContentDataTable $OUTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Organizational Units Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectOUGPOLinks -DataSet $OUGPOTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PO12 -DataSet $OUProtectionTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentclose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#  Users Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[3] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPUserTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Active Directory Users'))
	$FinalReport.Add($(Get-HTMLContentDataTable $UserTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Expiring Items'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days"))
	$FinalReport.Add($(Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Accounts'))
	$FinalReport.Add($(Get-HTMLColumn1of2))
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days or more"))
	$FinalReport.Add($(Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumn2of2))

	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Users Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $EnabledDisabledUsersPieObject -DataSet $EnabledDisabledUsersTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PWExpiresUsersTable -DataSet $PasswordExpirationTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectProtectedUsers -DataSet $ProtectedUsersTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#   GPO Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[4] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Group Policies'))
	$FinalReport.Add($(Get-HTMLContentDataTable $GPOTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	#$FinalReport.Add($(Get-HTMLTabContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Recently Created GPOs'))
	$FinalReport.Add($(Get-HTMLContentDataTable $recentgpostable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	$FinalReport.Add($(Get-HTMLTabContentClose))

	#  Computers Report
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[5] -TabHeading ('Report: ' + (Get-Date -Format MM-dd-yyyy))))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Overivew'))
	$FinalReport.Add($(Get-HTMLContentTable $TOPComputersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers'))
	$FinalReport.Add($(Get-HTMLContentDataTable $ComputersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Charts'))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputersProtected -DataSet $ComputerProtectedTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputersEnabled -DataSet $ComputersEnabledTable))
	$FinalReport.Add($(Get-HTMLColumnClose))
	$FinalReport.Add($(Get-HTMLContentclose))

	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText 'Computers Operating System Breakdown'))
	$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectComputerObjOS -DataSet $GraphComputerOS))
	$FinalReport.Add($(Get-HTMLContentclose))

	$FinalReport.Add($(Get-HTMLTabContentClose))
	$FinalReport.Add($(Get-HTMLClosePage))

	$Day = (Get-Date).Day
	$Month = (Get-Date).Month
	$Year = (Get-Date).Year
	$ReportName = ("$Day - $Month - $Year - AD Report")
	Write-ProgressHelper 25 "Save-HTMLReport"
	Save-HTMLReport -ReportContent $FinalReport -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath
}
Test-htmlModule
get-reportSettings
New-reportFolder
Set-Header
Get-DefaultSecurityGroups
Get-CreateObjects
Get-AllUsers
Get-AllGPOs
Get-ADOBJ
Set-ADOBJ
Get-ADRecBin
Get-ADInfo
Get-CreatedUsers
Get-DomainAdmins
Set-DomainAdminsTable
Get-EntAdmins
Get-Computers
Get-ADUsers
Get-ExpiredAccounts
Get-Seclogs
Get-Domains
Get-Groups
Get-OU
Get-Users
Get-GPOs
Get-RecentGPOs
Get-Adcomputers
compare-tables 
Get-Report