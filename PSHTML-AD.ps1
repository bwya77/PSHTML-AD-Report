function LastLogonConvert ($ftDate)
{
    $date = [DateTime]::FromFileTime($ftDate)
    If ($date -lt (Get-Date '1/1/1900') -or $date -eq 0 -or $date -eq $Null) { "Never" }
    Else { $date }
} # End function LastLogonConvert

#########################################
#                                       #
#            VARIABLES                  #
#                                       #
#########################################

# Company logo that will be displayed on the left, can be URL or UNC
$CompanyLogo = ""

# Logo that will be on the right side, UNC or URL
$RightLogo = "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png"

$ReportTitle = "Active Directory Report"

# Location the report will be saved to
$ReportSavePath = "C:\Automation\"

# Find users that have not logged in X Amount of days, this sets the days
$Days = 1

# Get users who have been created in X amount of days and less
$UserCreatedDays = 7

# Get users whos passwords expire in less than X amount of days
$DaysUntilPWExpireINT = 7

# Get AD Objects that have been modified in X days and newer
$ADNumber = 3

# CSS templatee located C:\Program Files\WindowsPowerShell\Modules\ReportHTML\1.4.1.1\
# Default template is orange and named "Sample"


########################################
# Check for ReportHTML Module
$Mod = Get-Module -ListAvailable -Name "ReportHTML"
If ($Null -eq $Mod)
{
    Write-Host "ReportHTML Module is not present, attempting to install it"
    Install-Module -Name ReportHTML -Force
    Import-Module ReportHTML -ErrorAction SilentlyContinue
}

# Array of default Security Groups
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





$Table                              = New-Object 'System.Collections.Generic.List[System.Object]'
$OUTable                            = New-Object 'System.Collections.Generic.List[System.Object]'
$UserTable                          = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupTypetable                     = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultGrouptable                  = New-Object 'System.Collections.Generic.List[System.Object]'
$EnabledDisabledUsersTable          = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainAdminTable                   = New-Object 'System.Collections.Generic.List[System.Object]'
$ExpiringAccountsTable              = New-Object 'System.Collections.Generic.List[System.Object]'
$CompanyInfoTable                   = New-Object 'System.Collections.Generic.List[System.Object]'
$securityeventtable                 = New-Object 'System.Collections.Generic.List[System.Object]'
$DomainTable                        = New-Object 'System.Collections.Generic.List[System.Object]'
$OUGPOTable                         = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupMembershipTable               = New-Object 'System.Collections.Generic.List[System.Object]'
$PasswordExpirationTable            = New-Object 'System.Collections.Generic.List[System.Object]'
$PasswordExpireSoonTable            = New-Object 'System.Collections.Generic.List[System.Object]'
$userphaventloggedonrecentlytable   = New-Object 'System.Collections.Generic.List[System.Object]'
$EnterpriseAdminTable               = New-Object 'System.Collections.Generic.List[System.Object]'
$NewCreatedUsersTable               = New-Object 'System.Collections.Generic.List[System.Object]'
$GroupProtectionTable               = New-Object 'System.Collections.Generic.List[System.Object]'
$OUProtectionTable                  = New-Object 'System.Collections.Generic.List[System.Object]'
$GPOTable                           = New-Object 'System.Collections.Generic.List[System.Object]'
$ADObjectTable                      = New-Object 'System.Collections.Generic.List[System.Object]'
$ProtectedUsersTable                = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputersTable                     = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputerProtectedTable             = New-Object 'System.Collections.Generic.List[System.Object]'
$ComputersEnabledTable              = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultComputersinDefaultOUTable   = New-Object 'System.Collections.Generic.List[System.Object]'
$DefaultUsersinDefaultOUTable       = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPUserTable                       = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPGroupsTable                     = New-Object 'System.Collections.Generic.List[System.Object]'
$TOPComputersTable                  = New-Object 'System.Collections.Generic.List[System.Object]'
$GraphComputerOS                    = New-Object 'System.Collections.Generic.List[System.Object]'

# Get all users right away. Instead of doing several lookups, we will use this object to look up all the information needed.
$AllUsers   = Get-ADUser -Filter * -Properties *
$GPOs       = Get-GPO -all | Select-Object DisplayName, GPOStatus, ModificationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }

<###########################
         Dashboard
############################>


$dte    = (Get-Date).AddDays(- $ADNumber)
$ADObjs = Get-ADObject -Filter { whenchanged -gt $dte -and ObjectClass -ne "domainDNS" -and ObjectClass -ne "rIDManager" -and ObjectClass -ne "rIDSet" } -Properties *
ForEach ($ADObj in $ADObjs)
{
    If ($ADObj.ObjectClass -eq "GroupPolicyContainer")
    {
        $Name = $ADObj.DisplayName
    }
    Else
    {
        $Name = $ADObj.Name
    }
    $obj = [PSCustomObject]@{
        'Name'  = $Name
        'Object Type' = $ADObj.ObjectClass
        'When Changed' = $ADObj.WhenChanged
    }

    $ADObjectTable.add($obj)

}



$ADRecycleBinStatus = (Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes
If ($ADRecycleBinStatus.Count -lt 1)
{
    $ADRecycleBin = "Disabled"
}
Else
{
    $ADRecycleBin = "Enabled"
}

# Company Information
$ADInfo                 = Get-ADDomain
$ForestObj              = Get-ADForest
$DomainControllerobj    = Get-ADDomain

$Forest                 = $ADInfo.Forest
$InfrastructureMaster   = $DomainControllerobj.InfrastructureMaster
$RIDMaster              = $DomainControllerobj.RIDMaster
$PDCEmulator            = $DomainControllerobj.PDCEmulator
$DomainNamingMaster     = $ForestObj.DomainNamingMaster
$SchemaMaster           = $ForestObj.SchemaMaster


$obj = [PSCustomObject]@{
    'Domain'                = $Forest
    'AD Recycle Bin'        = $ADRecycleBin
    'Infrastructure Master' = $InfrastructureMaster
    'RID Master'            = $RIDMaster
    'PDC Emulator'          = $PDCEmulator
    'Domain Naming Master'  = $DomainNamingMaster
    'Schema Master'         = $SchemaMaster

}

$CompanyInfoTable.add($obj)

# Get newly created users
$When = ((Get-Date).AddDays(-$UserCreatedDays)).Date
$NewUsers = Get-ADUser -Filter { whenCreated -ge $When } -Properties whenCreated
ForEach ($Newuser in $Newusers)
{
    $obj = [PSCustomObject]@{
        'Name'          = $Newuser.Name
        'Enabled'       = $Newuser.Enabled
        'Creation Date' = $Newuser.whenCreated
    }

    $NewCreatedUsersTable.add($obj)

}

# Get Domain Admins
$DomainAdminMembers = Get-ADGroupMember "Domain Admins"

ForEach ($DomainAdminMember in $DomainAdminMembers)
{
    $Name       = $DomainAdminMember.Name
    $Type       = $DomainAdminMember.ObjectClass
    $Enabled    = (Get-ADUser -Filter * | Where-Object { $_.Name -eq $Name }).Enabled

    $obj = [PSCustomObject]@{
        'Name'      = $Name
        'Enabled'   = $Enabled
        'Type'      = $Type
    }

    $DomainAdminTable.add($obj)

}


# Get Enterprise Admins
$EnterpriseAdminsMembers = Get-ADGroupMember "Enterprise Admins"

ForEach ($EnterpriseAdminsMember in $EnterpriseAdminsMembers)
{
    $Name       = $EnterpriseAdminsMember.Name
    $Type       = $EnterpriseAdminsMember.ObjectClass
    $Enabled    = (Get-ADUser -Filter * | Where-Object { $_.Name -eq $Name }).Enabled

    $obj = [PSCustomObject]@{
        'Name'    = $Name
        'Enabled' = $Enabled
        'Type'    = $Type
    }

    $EnterpriseAdminTable.add($obj)

}

$DefaultComputersOU = (Get-ADDomain).computerscontainer
$DefaultComputers   = Get-ADComputer -Filter * -Properties * -SearchBase "$DefaultComputersOU"
ForEach ($DefaultComputer in $DefaultComputers)
{
    $obj = [PSCustomObject]@{
        'Name'                  = $DefaultComputer.Name
        'Enabled'               = $DefaultComputer.Enabled
        'Operating System'      = $DefaultComputer.OperatingSystem
        'Modified Date'         = $DefaultComputer.Modified
        'Password Last Set'     = $DefaultComputer.PasswordLastSet
        'Protect from Deletion' = $DefaultComputer.ProtectedFromAccidentalDeletion
    }


    $DefaultComputersinDefaultOUTable.add($obj)

}

$DefaultUsersOU = (Get-ADDomain).userscontainer
$DefaultUsers   = Get-ADUser -Filter * -Properties * -SearchBase "$DefaultUsersOU" | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, @{ Name = 'lastlogon'; Expression = { LastLogonConvert $_.lastlogon } }, DistinguishedName
ForEach ($DefaultUser in $DefaultUsers)
{
    $obj = [PSCustomObject]@{
        'Name'                    = $DefaultUser.Name
        'UserPrincipalName'       = $DefaultUser.UserPrincipalName
        'Enabled'                 = $DefaultUser.Enabled
        'Protected from Deletion' = $DefaultUser.ProtectedFromAccidentalDeletion
        'Last Logon'              = $DefaultUser.LastLogon
        'Email Address'           = $DefaultUser.EmailAddress

    }


    $DefaultUsersinDefaultOUTable.add($obj)
}

# Expiring Accounts
$LooseUsers = Search-ADAccount -AccountExpiring -UsersOnly
ForEach ($LooseUser in $LooseUsers)
{
    $NameLoose      = $LooseUser.Name
    $UPNLoose       = $LooseUser.UserPrincipalName
    $ExpirationDate = $LooseUser.AccountExpirationDate
    $enabled        = $LooseUser.Enabled


    $obj = [PSCustomObject]@{
        'Name'              = $NameLoose
        'UserPrincipalName' = $UPNLoose
        'Expiration Date'   = $ExpirationDate
        'Enabled'           = $enabled
    }


    $ExpiringAccountsTable.add($obj)
}
If (($ExpiringAccountsTable).count -eq 0)
{
    $ExpiringAccountsTable = [PSCustomObject]@{
        'Information' = 'Information: No Users were found to expire soon'
    }
}


# Security Logs
$SecurityLogs = Get-EventLog -Newest 7 -LogName "Security" | Where-Object { $_.Message -like "*An account*" }
ForEach ($SecurityLog in $SecurityLogs)
{
    $TimeGenerated = $SecurityLog.TimeGenerated
    $EntryType     = $SecurityLog.EntryType
    $Recipient     = $SecurityLog.Message


    $obj = [PSCustomObject]@{
        'Time'    = $TimeGenerated
        'Type'    = $EntryType
        'Message' = $Recipient

    }

    $SecurityEventTable.add($obj)
}
If (($securityeventtable).count -eq 0)
{
    $securityeventtable = [PSCustomObject]@{
        'Information' = 'Information: No recent security logs'
    }
}


# Tenant Domain
$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes

# Tenant Domain
$Domains = Get-ADForest | Select-Object -ExpandProperty upnsuffixes | ForEach-Object{
    $obj = [PSCustomObject]@{
        'UPN Suffixes' = $_
        'Valid'        = "True"
    }
    $DomainTable.add($obj)
}

<###########################
        Groups
############################>


# Get groups and sort in alphabetical order
$Groups = Get-ADGroup -Filter * -Properties *

$SecurityCount          = 0
$MailSecurityCount      = 0
$CustomGroup            = 0
$DefaultGroup           = 0
$Groupswithmemebrship   = 0
$Groupswithnomembership = 0
$GroupsProtected        = 0
$GroupsNotProtected     = 0


ForEach ($Group in $Groups)
{
    $DefaultADGroup = 'False'
    $Type           = New-Object 'System.Collections.Generic.List[System.Object]'
    $Gemail         = (Get-ADGroup $Group -Properties mail).mail

    If (($group.GroupCategory -eq "Security") -and ($Gemail -ne $Null))
    {
        $MailSecurityCount++

    }
    If (($group.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
    {
        $SecurityCount++
    }

    If ($Group.ProtectedFromAccidentalDeletion -eq $True)
    {
        $GroupsProtected++
    }
    Else
    {
        $GroupsNotProtected++
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



    If ($group.GroupCategory -eq "Distribution")
    {
        $Type = "Distribution Group"
    }
    If (($group.GroupCategory -eq "Security") -and (($Gemail) -eq $Null))
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
        If (!($Users))
            {
                $Groupswithnomembership++
            }
            Else
            {
                $Groupswithmemebrship++
            }
        }
        Else
        {

            $Users = "Skipped Domain Users Membership"

        }

        $OwnerDN = Get-ADGroup -Filter { name -eq $Group.Name }  -Properties managedBy | Select-Object -ExpandProperty ManagedBy


    $Manager = Get-ADUser -Filter * | Where-Object { $_.distinguishedname -eq $OwnerDN } | Select-Object -ExpandProperty Name


    $obj = [PSCustomObject]@{
        'Name'                    = $Group.name
        'Type'                    = $Type
        'Members'                 = $users
        'Managed By'              = $Manager
        'E-mail Address'          = $GEmail
        'Protected from Deletion' = $Group.ProtectedFromAccidentalDeletion
        'Default AD Group'        = $DefaultADGroup
    }

    $table.add($obj)
}
If (($table).count -eq 0)
{
    $table = [PSCustomObject]@{
        'Information' = 'Information: No Groups were found'
    }
}

# TOP groups table
$obj1 = [PSCustomObject]@{
    'Total Groups'                 = $Groups.count
    'Mail-Enabled Security Groups' = $MailSecurityCount
    'Security Groups'              = $SecurityCount
    'Distribution Groups'          = $DistroCount
}
$TOPGroupsTable.add($obj1)

$obj1 = [PSCustomObject]@{
    'Name'  = 'Mail-Enabled Security Groups'
    'Count' = $MailSecurityCount
}

$GroupTypetable.add($obj1)

$obj1 = [PSCustomObject]@{
    'Name'  = 'Security Groups'
    'Count' = $SecurityCount
}
$GroupTypetable.add($obj1)

$DistroCount = ($Groups | Where-Object { $_.GroupCategory -eq "Distribution" }).Count
$obj1 = [PSCustomObject]@{
    'Name'  = 'Distribution Groups'
    'Count' = $DistroCount
}

$GroupTypetable.add($obj1)


# Default Group Pie Chart
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


# Group Protection Pie Chart
$obj1 = [PSCustomObject]@{
    'Name'  = 'Protected'
    'Count' = $GroupsProtected
}

$GroupProtectionTable.add($obj1)

$obj1 = [PSCustomObject]@{
    'Name'  = 'Not Protected'
    'Count' = $GroupsNotProtected
}

$GroupProtectionTable.add($obj1)


# Groups with membership vs no membership pie chart

$objmem = [PSCustomObject]@{
    'Name'  = 'With Members'
    'Count' = $Groupswithmemebrship
}

$GroupMembershipTable.add($objmem)

$objmem = [PSCustomObject]@{
    'Name'  = 'No Members'
    'Count' = $Groupswithnomembership
}

$GroupMembershipTable.add($objmem)


<###########################
    Organizational Units
############################>

# Get all OU's'
$OUs = Get-ADOrganizationalUnit -Filter * -Properties *

$OUwithLinked   = 0
$OUwithnoLink   = 0
$OUProtected    = 0
$OUNotProtected = 0

ForEach ($OU in $OUs)
{
    $LinkedGPOs = New-Object 'System.Collections.Generic.List[System.Object]'
    If (($OU.linkedgrouppolicyobjects).length -lt 1)
    {
        $LinkedGPOs = "None"
        $OUwithnoLink++
    }
    Else
    {
        $OUwithLinked++
        $GPOs = $OU.linkedgrouppolicyobjects
        ForEach ($GPO in $GPOs)
        {
            $Split1 = $GPO -split "{" | Select-Object -Last 1
            $Split2 = $Split1 -split "}" | Select-Object -First 1
            $LinkedGPOs.add((Get-GPO -Guid $Split2 -ErrorAction SilentlyContinue).DisplayName )
        }


    }

    If ($OU.ProtectedFromAccidentalDeletion -eq $True)
    {
        $OUProtected++
    }
    Else
    {
        $OUNotProtected++
    }


    $LinkedGPOs = $LinkedGPOs -join ", "
    $obj = [PSCustomObject]@{
        'Name'                    = $OU.Name
        'Linked GPOs'             = $LinkedGPOs
        'Modified Date'           = $OU.WhenChanged
        'Protected from Deletion' = $OU.ProtectedFromAccidentalDeletion
    }

    $OUTable.add($obj)

}
If (($OUTable).count -eq 0)
{
    $OUTable = [PSCustomObject]@{
        'Information' = 'Information: No Organizational Units were found'
    }
}


# OUs with no GPO Linked
$obj1 = [PSCustomObject]@{
    'Name'  = "OU's with no GPO's linked"
    'Count' = $OUwithnoLink
}

$OUGPOTable.add($obj1)

$obj2 = [PSCustomObject]@{
    'Name'  = "OU's with GPO's linked"
    'Count' = $OUwithLinked
}

$OUGPOTable.add($obj2)


# OUs Protected Pie Chart
$obj1 = [PSCustomObject]@{
    'Name'  = "Protected"
    'Count' = $OUProtected
}

$OUProtectionTable.add($obj1)

$obj2 = [PSCustomObject]@{
    'Name'  = "Not Protected"
    'Count' = $OUNotProtected
}

$OUProtectionTable.add($obj2)



<###########################
           USERS
############################>
$UserEnabled                            = 0
$UserDisabled                           = 0
$UserPasswordExpires                    = 0
$UserPasswordNeverExpires               = 0
$ProtectedUsers                         = 0
$NonProtectedUsers                      = 0
$UsersWIthPasswordsExpiringInUnderAWeek = 0
$UsersNotLoggedInOver30Days             = 0
$AccountsExpiringSoon                   = 0

ForEach ($User in $AllUsers)
{

    $AttVar = Get-ADUser -Filter {Name -eq $User.Name} -Properties * | Select-Object Enabled,PasswordExpired, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, Name, SamAccountName, EmailAddress, AccountExpirationDate, @{ Name = 'lastlogon'; Expression = { LastLogonConvert $_.lastlogon } }, DistinguishedName

    $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    If ((($AttVar.PasswordNeverExpires) -eq $False) -and (($AttVar.Enabled) -ne $False))
    {
        # Get Password last set date
        $passwordSetDate = (Get-ADUser $user -Properties * | ForEach-Object { $_.PasswordLastSet })
        If ($Null -eq $passwordSetDate)
        {
            $daystoexpire = "User has never logged on"
        }
        Else
        {
            # Check for Fine Grained Passwords
            $PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
            If (($PasswordPol) -ne $Null)
            {
                $maxPasswordAge = ($PasswordPol).MaxPasswordAge
            }

            $expireson = $passwordsetdate + $maxPasswordAge
            $today     = (Get-Date)
            # Gets the count on how many days until the password expires and stores it in the $daystoexpire var
            $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days

        }
    }
    Else
    {

        $daystoexpire = "N/A"

    }

    # Get users that haven't logged on in X amount of days, var is set at start of script
    If (($User.Enabled -eq $True) -and ($User.LastLogonDate -lt (Get-Date).AddDays(-$Days)) -and ($User.LastLogonDate -ne $NULL))
    {
        $obj = [PSCustomObject]@{
            'Name'                        = $User.Name
            'UserPrincipalName'           = $User.UserPrincipalName
            'Enabled'                     = $AttVar.Enabled
            'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
            'Last Logon'                  = $AttVar.lastlogon
            'Password Never Expires'      = $AttVar.PasswordNeverExpires
            'Days Until Password Expires' = $daystoexpire
        }
        $userphaventloggedonrecentlytable.add($obj)
    }
    If (($userphaventloggedonrecentlytable).count -eq 0)
    {
        $userphaventloggedonrecentlytable = [PSCustomObject]@{
            'Information' = "Information: No Users were found to have not logged on in $Days days"
        }
    }

    # Items for protected vs non protected users
    If ($User.ProtectedFromAccidentalDeletion -eq $False)
    {
        $NonProtectedUsers++
    }
    Else
    {
        $ProtectedUsers++
    }

    # Items for the enabled vs disabled users pie chart
    If (($AttVar.PasswordNeverExpires) -ne $False)
    {
        $UserPasswordNeverExpires++
    }
    Else
    {
        $UserPasswordExpires++
    }

    # Items for password expiration pie chart
    If (($AttVar.Enabled) -ne $False)
    {
        $UserEnabled++
    }
    Else
    {
        $UserDisabled++
    }





    $Name                   = $User.Name
    $UPN                    = $User.UserPrincipalName
    $Enabled                = $AttVar.Enabled
    $LastLogon              = $AttVar.lastlogon
    $EmailAddress           = $AttVar.EmailAddress
    $AccountExpiration      = $AttVar.AccountExpirationDate
    $PasswordExpired        = $AttVar.PasswordExpired
    $PasswordLastSet        = $AttVar.PasswordLastSet
    $PasswordNeverExpires   = $AttVar.PasswordNeverExpires
    $daysUntilPWExpire      = $daystoexpire




    $obj = [PSCustomObject]@{
        'Name'                        = $Name
        'UserPrincipalName'           = $UPN
        'Enabled'                     = $Enabled
        'Protected from Deletion'     = $User.ProtectedFromAccidentalDeletion
        'Last Logon'                  = $LastLogon
        'Email Address'               = $EmailAddress
        'Account Expiration'          = $AccountExpiration
        'Change Password Next Logon'  = $PasswordExpired
        'Password Last Set'           = $PasswordLastSet
        'Password Never Expires'      = $PasswordNeverExpires
        'Days Until Password Expires' = $daystoexpire
    }
    $usertable.add($obj)

    If ($daystoexpire -lt $DaysUntilPWExpireINT)
    {
        $obj = [PSCustomObject]@{
            'Name'                  = $Name
            'Days Until Password Expires' = $daystoexpire
        }
        $PasswordExpireSoonTable.add($obj)
    }
}
If (($usertable).count -eq 0)
{
    $usertable = [PSCustomObject]@{
        'Information' = 'Information: No Users were found'
    }
}


# Data for users enabled vs disabled pie graph
$objULic = [PSCustomObject]@{
    'Name'  = 'Enabled'
    'Count' = $UserEnabled
}

$EnabledDisabledUsersTable.add($objULic)


$objULic = [PSCustomObject]@{
    'Name'  = 'Disabled'
    'Count' = $UserDisabled
}

$EnabledDisabledUsersTable.add($objULic)


# Data for users password expires pie graph
$objULic = [PSCustomObject]@{
    'Name'  = 'Password Expires'
    'Count' = $UserPasswordExpires
}

$PasswordExpirationTable.add($objULic)


$objULic = [PSCustomObject]@{
    'Name'  = 'Password Never Expires'
    'Count' = $UserPasswordNeverExpires
}

$PasswordExpirationTable.add($objULic)


# Data for protected users pie graph
$objULic = [PSCustomObject]@{
    'Name'  = 'Protected'
    'Count' = $ProtectedUsers
}

$ProtectedUsersTable.add($objULic)


$objULic = [PSCustomObject]@{
    'Name'  = 'Not Protected'
    'Count' = $NonProtectedUsers
}

$ProtectedUsersTable.add($objULic)


# TOP User table
$objULic = [PSCustomObject]@{
    'Total Users'                                                           = $AllUsers.count
    "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" = $PasswordExpireSoonTable.count
    'Expiring Accounts'                                                     = $ExpiringAccountsTable.count
    "Users Haven't Logged on in $Days Days"                                 = $userphaventloggedonrecentlytable.count
}

$TOPUserTable.add($objULic)


<###########################
    Group Policy
############################>
$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'
$GPOs = Get-GPO -all | Select-Object DisplayName, GPOStatus, ModificationTime, @{ Label = "ComputerVersion"; Expression = { $_.computer.dsversion } }, @{ Label = "UserVersion"; Expression = { $_.user.dsversion } }
ForEach ($GPO in $GPOs)
{

    $obj = [PSCustomObject]@{
        'Name'             = $GPO.DisplayName
        'Status'           = $GPO.GpoStatus
        'Modified Date'    = $GPO.ModificationTime
        'User Version'     = $GPO.UserVersion
        'Computer Version' = $GPO.ComputerVersion
    }

    $GPOTable.add($obj)
}


<###########################
         Computers
############################>
$Computers             = Get-ADComputer -Filter * -Properties *
$ComputersProtected    = 0
$ComputersNotProtected = 0
$ComputerEnabled       = 0
$ComputerDisabled      = 0

$Server2016   = 0
$Server2012   = 0
$Server2012R2 = 0
$Server2008R2 = 0
$Windows10    = 0
$Windows8     = 0
$Windows7     = 0
$Server2012R2 = 0


ForEach ($Computer in $Computers)
{

    If ($Computer.ProtectedFromAccidentalDeletion -eq $True)
    {
        $ComputersProtected++
    }
    Else
    {
        $ComputersNotProtected++
    }

    If ($Computer.Enabled -eq $True)
    {
        $ComputerEnabled++
    }
    Else
    {
        $ComputerDisabled++
    }

    $obj = [PSCustomObject]@{
        'Name'                  = $Computer.Name
        'Enabled'               = $Computer.Enabled
        'Operating System'      = $Computer.OperatingSystem
        'Modified Date'         = $Computer.Modified
        'Password Last Set'     = $Computer.PasswordLastSet
        'Protect from Deletion' = $Computer.ProtectedFromAccidentalDeletion

    }

    $ComputersTable.add($obj)


    If ($Computer.OperatingSystem -like "*Server 2016*")
    {
        $Server2016++
    }
    ElseIf ($Computer.OperatingSystem -like "*Server 2012 R2*")
    {
        $Server2012R2++
    }
    ElseIf ($Computer.OperatingSystem -like "*Server 2012*")
    {
        $Server2012++
    }
    ElseIf ($Computer.OperatingSystem -like "*Server 2008 R2*")
    {
        $Server2008R2++
    }
    ElseIf ($Computer.OperatingSystem -like "*Windows 10*")
    {
        $Windows10++
    }
    ElseIf ($Computer.OperatingSystem -like "*Windows 8*")
    {
        $Windows8++
    }
    Elseif ($Computer.OperatingSystem -like "*Windows 7*")
    {
        $Windows7++
    }

}
If (($ComputersTable).count -eq 0)
{
    $ComputersTable = [PSCustomObject]@{
        'Information' = 'Information: No Computers were found'
    }
}

# Data for TOP Computers data table
$objULic = [PSCustomObject]@{
    'Total Computers' = $Computers.count
    "Server 2016"     = $Server2016
    "Server 2012 R2"  = $Server2012R2
    "Server 2012"     = $Server2012
    "Server 2008 R2"  = $Server2008R2
    "Windows 10"      = $Windows10
    "Windows 8"       = $Windows8
    "Windows 7"       = $Windows7
}
$TOPComputersTable.add($objULic)

# Pie chart breaking down OS for computer obj
$objULic = [PSCustomObject]@{
    'Name' = "Server 2016"
    "Count"= $Server2016
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Server 2012 R2"
    "Count" = $Server2012R2
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Server 2012"
    "Count" = $Server2012
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Server 2008 R2"
    "Count" = $Server2008R2
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Windows 10"
    "Count" = $Windows10
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Windows 8"
    "Count" = $Windows8
}
$GraphComputerOS.add($objULic)

$objULic = [PSCustomObject]@{
    'Name'  = "Windows 7"
    "Count" = $Windows7
}
$GraphComputerOS.add($objULic)
################################

# Data for protected Computers pie graph
$objULic = [PSCustomObject]@{
    'Name'  = 'Protected'
    'Count' = $ComputerProtected
}

$ComputerProtectedTable.add($objULic)


$objULic = [PSCustomObject]@{
    'Name'  = 'Not Protected'
    'Count' = $ComputersNotProtected
}

$ComputerProtectedTable.add($objULic)


# Data for enabled/vs Computers pie graph
$objULic = [PSCustomObject]@{
    'Name'  = 'Enabled'
    'Count' = $ComputerEnabled
}

$ComputersEnabledTable.add($objULic)


$objULic = [PSCustomObject]@{
    'Name'  = 'Disabled'
    'Count' = $ComputerDisabled
}

$ComputersEnabledTable.add($objULic)



$tabarray = @('Dashboard', 'Groups', 'Organizational Units', 'Users', 'Group Policy', 'Computers')

##--OU Protection PIE CHART--##
# Basic Properties
$PO12                      = Get-HTMLPieChartObject
$PO12.Title                = "Organizational Units Protected from Deletion"
$PO12.Size.Height          = 250
$PO12.Size.width           = 250
$PO12.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PO12.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PO12.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PO12.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PO12.DataDefinition.DataNameColumnName  = 'Name'
$PO12.DataDefinition.DataValueColumnName = 'Count'

##--Computer OS Breakdown PIE CHART--##
$PieObjectComputerObjOS = Get-HTMLPieChartObject
$PieObjectComputerObjOS.Title = "Computer Operating Systems"
# These file exist in the module directoy, There are 4 schemes by default
$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectComputerObjOS.ChartStyle.ColorSchemeName = 'Random'


##--Computers Protection PIE CHART--##
# Basic Properties
$PieObjectComputersProtected                      = Get-HTMLPieChartObject
$PieObjectComputersProtected.Title                = "Computers Protected from Deletion"
$PieObjectComputersProtected.Size.Height          = 250
$PieObjectComputersProtected.Size.width           = 250
$PieObjectComputersProtected.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PieObjectComputersProtected.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectComputersProtected.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectComputersProtected.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectComputersProtected.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectComputersProtected.DataDefinition.DataValueColumnName = 'Count'


##--Computers Enabled PIE CHART--##
# Basic Properties
$PieObjectComputersEnabled                      = Get-HTMLPieChartObject
$PieObjectComputersEnabled.Title                = "Computers Enabled vs Disabled"
$PieObjectComputersEnabled.Size.Height          = 250
$PieObjectComputersEnabled.Size.width           = 250
$PieObjectComputersEnabled.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectComputersEnabled.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectComputersEnabled.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectComputersEnabled.DataDefinition.DataValueColumnName = 'Count'


##--USERS Protection PIE CHART--##
# Basic Properties
$PieObjectProtectedUsers                      = Get-HTMLPieChartObject
$PieObjectProtectedUsers.Title                = "Users Protected from Deletion"
$PieObjectProtectedUsers.Size.Height          = 250
$PieObjectProtectedUsers.Size.width           = 250
$PieObjectProtectedUsers.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectProtectedUsers.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectProtectedUsers.DataDefinition.DataValueColumnName = 'Count'

# Basic Properties
$PieObjectOUGPOLinks                      = Get-HTMLPieChartObject
$PieObjectOUGPOLinks.Title                = "OU GPO Links"
$PieObjectOUGPOLinks.Size.Height          = 250
$PieObjectOUGPOLinks.Size.width           = 250
$PieObjectOUGPOLinks.ChartStyle.ChartType = 'doughnut'

# These file exist in the module directoy, There are 4 schemes by default
$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = "ColorScheme4"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = "Generated5"
# You can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectOUGPOLinks.ChartStyle.ColorSchemeName = 'Random'

# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectOUGPOLinks.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectOUGPOLinks.DataDefinition.DataValueColumnName = 'Count'

# Basic Properties
$PieObject4                      = Get-HTMLPieChartObject
$PieObject4.Title                = "Office 365 Unassigned Licenses"
$PieObject4.Size.Height          = 250
$PieObject4.Size.width           = 250
$PieObject4.ChartStyle.ChartType = 'doughnut'

# These file exist in the module directoy, There are 4 schemes by default
$PieObject4.ChartStyle.ColorSchemeName = "ColorScheme4"
# There are 8 generated schemes, randomly generated at runtime
$PieObject4.ChartStyle.ColorSchemeName = "Generated4"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObject4.ChartStyle.ColorSchemeName = 'Random'

# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObject4.DataDefinition.DataNameColumnName  = 'Name'
$PieObject4.DataDefinition.DataValueColumnName = 'Unassigned Licenses'

# Basic Properties
$PieObjectGroupType                      = Get-HTMLPieChartObject
$PieObjectGroupType.Title                = "Group Types"
$PieObjectGroupType.Size.Height          = 250
$PieObjectGroupType.Size.width           = 250
$PieObjectGroupType.ChartStyle.ChartType = 'doughnut'

# Pie Chart Groups with members vs no members
$PieObjectGroupMembersType                                    = Get-HTMLPieChartObject
$PieObjectGroupMembersType.Title                              = "Group Membership"
$PieObjectGroupMembersType.Size.Height                        = 250
$PieObjectGroupMembersType.Size.width                         = 250
$PieObjectGroupMembersType.ChartStyle.ChartType               = 'doughnut'
$PieObjectGroupMembersType.ChartStyle.ColorSchemeName         = "ColorScheme4"
$PieObjectGroupMembersType.ChartStyle.ColorSchemeName         = "Generated8"
$PieObjectGroupMembersType.ChartStyle.ColorSchemeName         = 'Random'
$PieObjectGroupMembersType.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectGroupMembersType.DataDefinition.DataValueColumnName = 'Count'

# Basic Properties
$PieObjectGroupType2                      = Get-HTMLPieChartObject
$PieObjectGroupType2.Title                = "Custom vs Default Groups"
$PieObjectGroupType2.Size.Height          = 250
$PieObjectGroupType2.Size.width           = 250
$PieObjectGroupType2.ChartStyle.ChartType = 'doughnut'


# These file exist in the module directoy, There are 4 schemes by default
$PieObjectGroupType.ChartStyle.ColorSchemeName = "ColorScheme4"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectGroupType.ChartStyle.ColorSchemeName = "Generated8"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectGroupType.ChartStyle.ColorSchemeName = 'Random'

# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectGroupType.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectGroupType.DataDefinition.DataValueColumnName = 'Count'

##--Enabled users vs Disabled Users PIE CHART--##
# Basic Properties
$EnabledDisabledUsersPieObject                      = Get-HTMLPieChartObject
$EnabledDisabledUsersPieObject.Title                = "Enabled vs Disabled Users"
$EnabledDisabledUsersPieObject.Size.Height          = 250
$EnabledDisabledUsersPieObject.Size.width           = 250
$EnabledDisabledUsersPieObject.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$EnabledDisabledUsersPieObject.DataDefinition.DataNameColumnName = 'Name'
$EnabledDisabledUsersPieObject.DataDefinition.DataValueColumnName = 'Count'


##--PasswordNeverExpires PIE CHART--##
# Basic Properties
$PWExpiresUsersTable                      = Get-HTMLPieChartObject
$PWExpiresUsersTable.Title                = "Password Expiration"
$PWExpiresUsersTable.Size.Height          = 250
$PWExpiresUsersTable.Size.width           = 250
$PWExpiresUsersTable.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PWExpiresUsersTable.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PWExpiresUsersTable.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PWExpiresUsersTable.DataDefinition.DataNameColumnName  = 'Name'
$PWExpiresUsersTable.DataDefinition.DataValueColumnName = 'Count'


##--Group Protection PIE CHART--##
# Basic Properties
$PieObjectGroupProtection                      = Get-HTMLPieChartObject
$PieObjectGroupProtection.Title                = "Groups Protected from Deletion"
$PieObjectGroupProtection.Size.Height          = 250
$PieObjectGroupProtection.Size.width           = 250
$PieObjectGroupProtection.ChartStyle.ChartType = 'doughnut'
# These file exist in the module directoy, There are 4 schemes by default
$PieObjectGroupProtection.ChartStyle.ColorSchemeName = "ColorScheme3"
# There are 8 generated schemes, randomly generated at runtime
$PieObjectGroupProtection.ChartStyle.ColorSchemeName = "Generated3"
# you can also ask for a random scheme.  Which also happens if you have too many records for the scheme
$PieObjectGroupProtection.ChartStyle.ColorSchemeName = 'Random'
# Data defintion you can reference any column from name and value from the  dataset.
# Name and Count are the default to work with the Group function.
$PieObjectGroupProtection.DataDefinition.DataNameColumnName  = 'Name'
$PieObjectGroupProtection.DataDefinition.DataValueColumnName = 'Count'




####################
# Dashboard Report #
####################


$rpt = New-Object 'System.Collections.Generic.List[System.Object]'
$rpt += Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo
$rpt += Get-HTMLTabHeader -TabNames $tabarray

$rpt += Get-HTMLTabContentOpen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Company Information"
$rpt += Get-HTMLContentTable $CompanyInfoTable
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Groups"
$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'
$rpt += Get-HTMLContentDataTable $DomainAdminTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'
$rpt += Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Objects in Default OUs"
$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Computers'
$rpt += Get-HTMLContentDataTable $DefaultComputersinDefaultOUTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -HeaderText 'Users'
$rpt += Get-HTMLContentDataTable $DefaultUsersinDefaultOUTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "AD Objects Modified in Last $ADNumber Days"
$rpt += Get-HTMLContentDataTable $ADObjectTable
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Expiring Items"

$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days"
$rpt += Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose

$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'
$rpt += Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose

$rpt += Get-HTMLContentClose


$rpt += Get-HTMLContentOpen -HeaderText "Accounts"

$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days"
$rpt += Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"
$rpt += Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Security Logs"
$rpt += Get-HTMLContentDataTable $securityeventtable -HideFooter
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "UPN Suffixes"
$rpt += Get-HTMLContentTable $DomainTable
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLTabContentClose

# Groups Report
$rpt += Get-HTMLTabContentOpen -TabName $tabarray[1] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))

$rpt += Get-HTMLContentOpen -HeaderText "Groups Overivew"
$rpt += Get-HTMLContentTable $TOPGroupsTable -HideFooter
$rpt += Get-HTMLContentClose

    $rpt += Get-HTMLContentOpen -HeaderText "Active Directory Groups"
        $rpt += Get-HTMLContentDataTable $Table -HideFooter
    $rpt += Get-HTMLContentClose

    $rpt += Get-HTMLColumn1of2
        $rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText 'Domain Administrators'
            $rpt += Get-HTMLContentDataTable $DomainAdminTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLColumnClose
    $rpt += Get-HTMLColumn2of2
        $rpt += Get-HTMLContentOpen -HeaderText 'Enterprise Administrators'
            $rpt += Get-HTMLContentDataTable $EnterpriseAdminTable -HideFooter
        $rpt += Get-HTMLContentClose
    $rpt += Get-HTMLColumnClose


    $rpt += Get-HTMLContentOpen -HeaderText "Active Directory Groups Chart"
        $rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 4
            $rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType -DataSet $GroupTypetable
        $rpt += Get-HTMLColumnClose
        $rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 4
            $rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupType2 -DataSet $DefaultGrouptable
        $rpt += Get-HTMLColumnClose
        $rpt += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 4
            $rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupMembersType -DataSet $GroupMembershipTable
        $rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 4 -ColumnCount 4
$rpt += Get-HTMLPieChart -ChartObject $PieObjectGroupProtection -DataSet $GroupProtectionTable
$rpt += Get-HTMLColumnClose
    $rpt += Get-HTMLContentClose
$rpt += Get-HTMLTabContentClose

# Organizational Unit Report
$rpt += Get-HTMLTabContentOpen -TabName $tabarray[2] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
    $rpt += Get-HTMLContentOpen -HeaderText "Organizational Units"
$rpt += Get-HTMLContentDataTable $OUTable -HideFooter
$rpt += Get-HTMLContentClose


$rpt += Get-HTMLContentOpen -HeaderText "Organizational Units Charts"
$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObjectOUGPOLinks -DataSet $OUGPOTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PO12 -DataSet $OUProtectionTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentclose

$rpt += Get-HTMLTabContentClose

# Users Report
$rpt += Get-HTMLTabContentOpen -TabName $tabarray[3] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))

$rpt += Get-HTMLContentOpen -HeaderText "Users Overivew"
$rpt += Get-HTMLContentTable $TOPUserTable -HideFooter
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Active Directory Users"
$rpt += Get-HTMLContentDataTable $UserTable -HideFooter
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Expiring Items"
$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days"
$rpt += Get-HTMLContentDataTable $PasswordExpireSoonTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -HeaderText 'Accounts Expiring Soon'
$rpt += Get-HTMLContentDataTable $ExpiringAccountsTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose

$rpt += Get-HTMLContentClose



$rpt += Get-HTMLContentOpen -HeaderText "Accounts"
$rpt += Get-HTMLColumn1of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Users Haven't Logged on in $Days Days"
$rpt += Get-HTMLContentDataTable $userphaventloggedonrecentlytable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumn2of2
$rpt += Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"
$rpt += Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLColumnClose

$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Users Charts"
$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3
$rpt += Get-HTMLPieChart -ChartObject $EnabledDisabledUsersPieObject -DataSet $EnabledDisabledUsersTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3
$rpt += Get-HTMLPieChart -ChartObject $PWExpiresUsersTable -DataSet $PasswordExpirationTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3
$rpt += Get-HTMLPieChart -ChartObject $PieObjectProtectedUsers -DataSet $ProtectedUsersTable
$rpt += Get-HTMLColumnClose

$rpt += Get-HTMLContentClose

$rpt += Get-HTMLTabContentClose

# GPO Report
$rpt += Get-HTMLTabContentOpen -TabName $tabarray[4] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))
$rpt += Get-HTMLContentOpen -HeaderText "Group Policies"
$rpt += Get-HTMLContentDataTable $GPOTable -HideFooter
$rpt += Get-HTMLContentClose
$rpt += Get-HTMLTabContentClose

# Computers Report
$rpt += Get-HTMLTabContentOpen -TabName $tabarray[5] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))

$rpt += Get-HTMLContentOpen -HeaderText "Computers Overivew"
$rpt += Get-HTMLContentTable $TOPComputersTable -HideFooter
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Computers"
$rpt += Get-HTMLContentDataTable $ComputersTable -HideFooter
$rpt += Get-HTMLContentClose

$rpt += Get-HTMLContentOpen -HeaderText "Computers Charts"
$rpt += Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObjectComputersProtected -DataSet $ComputerProtectedTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 2
$rpt += Get-HTMLPieChart -ChartObject $PieObjectComputersEnabled -DataSet $ComputersEnabledTable
$rpt += Get-HTMLColumnClose
$rpt += Get-HTMLContentclose

$rpt += Get-HTMLContentOpen -HeaderText "Computers Operating System Breakdown"
$rpt += Get-HTMLPieChart -ChartObject $PieObjectComputerObjOS -DataSet $GraphComputerOS
$rpt += Get-HTMLContentclose

$rpt += Get-HTMLTabContentClose

$rpt += Get-HTMLClosePage

$Day        = (Get-Date).Day
$Month      = (Get-Date).Month
$Year       = (Get-Date).Year
$ReportName = ("$Day" + "-" + "$Month" + "-" + "$Year" + "-" + "AD Report")

Save-HTMLReport -ReportContent $rpt -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath
