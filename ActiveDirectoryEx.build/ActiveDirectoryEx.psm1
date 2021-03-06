function Clear-ADObjectAdminCount {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [System.DirectoryServices.DirectoryEntry]
        $InputObject,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [switch]
        $Force
    )    

    begin {
    }
    process {
        $InputObject.Put("adminCount",@())
        if ($Force) {
            $p = 'y'
        }
        else {
            $p = Read-Host "Confirm (y/n)"
        }
        if ($p.ToLower().Trim() -eq 'y') {
            $InputObject.SetInfo()
        }
        else {
            $InputObject = $null
            Write-Error "Operation 'SetInfo for adminCount' aborted by user."
        }
    }
    end {
    }
}

function Convert-ADdwAdminSDExMask {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        $InputObject
    )

    begin { 
        $Domain_Alias_RID_Account_Ops = [int]"0x1"
        $Domain_Alias_RID_System_Ops  = [int]"0x2"
        $Domain_Alias_RID_Print_Ops   = [int]"0x4"
        $Domain_Alias_RID_Backup_Ops  = [int]"0x8"
        [array]$return = @()
    }
    process {
        if ($InputObject -band $Domain_Alias_RID_Account_Ops) {
            $return += "DOMAIN_ALIAS_RID_ACCOUNT_OPS"
        }
        if ($InputObject -band $Domain_Alias_RID_System_Ops) {
            $return += "DOMAIN_ALIAS_RID_SYSTEM_OPS"
        }
        if ($InputObject -band $Domain_Alias_RID_Print_Ops) {
            $return += "DOMAIN_ALIAS_RID_PRINT_OPS"
        }
        if ($InputObject -band $Domain_Alias_RID_Backup_Ops) {
            $return += "DOMAIN_ALIAS_RID_BACKUP_OPS"
        }
        return $return
    }
    end {
    }
}

function ConvertTo-ByteRid {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [int]$Rid
    )

    begin {
    }
    process {
        [byte[]]$return = @()
        $return += "0x{0}" -f ("{0:x8}" -f $Rid).Substring(6,2)
        $return += "0x{0}" -f ("{0:x8}" -f $Rid).Substring(4,2)
        $return += "0x{0}" -f ("{0:x8}" -f $Rid).Substring(2,2)
        $return += "0x{0}" -f ("{0:x8}" -f $Rid).Substring(0,2)
        return $return
    }
    end {
    }
}

function ConvertTo-StringSid {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [System.Byte]$ByteSid
    )

    begin {
        $return = ""
    }
    process {
        $ByteSid | 
            ForEach-Object {
                $return = "{0}\{1:x2}" -f $return, $_
            }
    }
    end {
        return $return
    }
}

function Get-ADChildGroup {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [string]$DistinguishedName,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [switch]$Recurse,       
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=2)]
        [string[]]$RecurseDistinguishedNames=@()
    )    

    begin {
    }
    process {
        $DistinguishedName.Substring($DistinguishedName.ToLower().IndexOf("dc=")) |
            Get-ADDirectoryEntry | 
            Get-ADObject -LDAPFilter "(&(objectCategory=Group)(memberOf=$DistinguishedName))" |
            foreach {
                $_
                if ($Recurse) {
                    if (!$RecurseDistinguishedNames.Contains($_.distinguishedName.ToLower())) {
                        $RecurseDistinguishedNames += $_.distinguishedName.ToLower()
                        $_ | 
                            Select-Object -ExpandProperty distinguishedName |
                            Get-ADChildGroup -Recurse -RecurseDistinguishedNames $RecurseDistinguishedNames
                    }
                }
            }
    }
    end {
    }
}

function Get-ADChildUser {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [string]$DistinguishedName
    )    

    begin {
    }
    process {
        $DistinguishedName.Substring($DistinguishedName.ToLower().IndexOf("dc=")) |
            Get-ADDirectoryEntry | 
            Get-ADObject -LDAPFilter "(&(objectCategory=User)(memberOf=$DistinguishedName))"
    }
    end {
    }
}

function Get-ADDirectoryEntry {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DistinguishedName
    )

    begin {
    }
    process {
        return [adsi]"LDAP://$DistinguishedName"
    }
    end {
    }
}

function Get-ADDomain {
    [CmdletBinding()]

    param (
        [switch]$All
    )

    begin { 
    }
    process {
        if ($All) {
            Get-ADForest
            Get-ADForest | 
                Select-Object -ExpandProperty subRefs | 
                Where-Object {
                    $_.StartsWith("DC=") -and 
                    !$_.StartsWith("DC=ForestDnsZones,") -and 
                    !$_.StartsWith("DC=DomainDnsZones,")
                } |
                Get-ADDirectoryEntry
        } 
        else {
            Get-ADDirectoryEntry "rootDSE" | 
                Select-Object -ExpandProperty defaultNamingContext | 
                Get-ADDirectoryEntry
        }
    }
    end {
    }
}

function Get-ADDomainControllerList {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]
        [ValidateNotNullOrEmpty()]
        [system.directoryservices.directoryentry]$Domain=(Get-ADDomain)
    )

    begin {
    }
    process {
        $Domain | 
            Select-Object -ExpandProperty masteredby | 
            foreach { 
                $_.split(",")[1].split("=")[1] 
            }
    }
    end {
    }
}

function Get-ADdSHeuristics {
    [CmdletBinding()]

    param (
    )    

    begin {
    }
    process {
        if (Test-ADdSHeuristics) {
            [string]$dSHeuristics = "CN=Configuration,{0}" -f (Get-ADForest | Select-Object -ExpandProperty distinguishedName) | 
                Get-ADDirectoryEntry |
                Get-ADObject -LDAPFilter "(name=Directory Service)" -Properties dSHeuristics | 
                Select-Object -ExpandProperty dSHeuristics

            while ($dSHeuristics.Length -lt 25) {
                $dSHeuristics += 0
            }

            return @{
                fSupFirstLastANR                                = [int]("0x{0}" -f $dSHeuristics.Substring(0,1))
                fSupLastFirstANR                                = [int]("0x{0}" -f $dSHeuristics.Substring(1,1))
                fDoListObject                                   = [int]("0x{0}" -f $dSHeuristics.Substring(2,1))
                fDoNickRes                                      = [int]("0x{0}" -f $dSHeuristics.Substring(3,1))
                fLDAPUsePermMod                                 = [int]("0x{0}" -f $dSHeuristics.Substring(4,1))
                ulHideDSID                                      = [int]("0x{0}" -f $dSHeuristics.Substring(5,1))
                fLDAPBlockAnonOps                               = [int]("0x{0}" -f $dSHeuristics.Substring(6,1))
                fAllowAnonNSPI                                  = [int]("0x{0}" -f $dSHeuristics.Substring(7,1))
                fUserPwdSupport                                 = [int]("0x{0}" -f $dSHeuristics.Substring(8,1))
                tenthChar                                       = [int]("0x{0}" -f $dSHeuristics.Substring(9,1))
                fSpecifyGUIDOnAdd                               = [int]("0x{0}" -f $dSHeuristics.Substring(10,1))
                fDontStandardizeSDs                             = [int]("0x{0}" -f $dSHeuristics.Substring(11,1))
                fAllowPasswordOperationsOverNonSecureConnection = [int]("0x{0}" -f $dSHeuristics.Substring(12,1))
                fDontPropagateOnNoChangeUpdate                  = [int]("0x{0}" -f $dSHeuristics.Substring(13,1))
                fComputeANRStats                                = [int]("0x{0}" -f $dSHeuristics.Substring(14,1))
                dwAdminSDExMask                                 = [int]("0x{0}" -f $dSHeuristics.Substring(15,1))
                fKVNOEmuW2K                                     = [int]("0x{0}" -f $dSHeuristics.Substring(16,1))
                fLDAPBypassUpperBoundsOnLimits                  = [int]("0x{0}" -f $dSHeuristics.Substring(17,1))
                fDisableAutoIndexingOnSchemaUpdate              = [int]("0x{0}" -f $dSHeuristics.Substring(18,1))
                twentiethChar                                   = [int]("0x{0}" -f $dSHeuristics.Substring(19,1))
                DoNotVerifyUPNAndOrSPNUniqueness                = [int]("0x{0}" -f $dSHeuristics.Substring(20,1))
                MinimumGetChangesRequestVersion                 = [int]("0x{0}" -f $dSHeuristics.Substring(21,2))
                MinimumGetChangesReplyVersion                   = [int]("0x{0}" -f $dSHeuristics.Substring(23,2))
            }
        }
        else {
            return $null
        }
    }
    end {
    }
}

function Get-ADdwAdminSDExMask {
    [CmdletBinding()]

    param (
    )
    
    begin {
    }
    process {
        $dSHeuristics = Get-ADdSHeuristics
        if ($dSHeuristics) {
            $dSHeuristics.dwAdminSDExMask | 
                Convert-ADdwAdminSDExMask
        }
    }
    end {
    }
}

function Get-ADForest {
    [CmdletBinding()]

    param (
    )

    begin {  
    }
    process {
        (Get-ADDirectoryEntry "rootDSE" | Select-Object -ExpandProperty configurationNamingContext).replace("CN=Configuration,","") | 
            Get-ADDirectoryEntry
    }
    end {
    }
}

function Get-ADObject {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $LDAPFilter,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [System.DirectoryServices.DirectoryEntry]
        $SearchRoot=(New-Object System.DirectoryServices.DirectoryEntry),
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=2)]
        [string[]]
        $Properties = @(
            "name",
            "sAMAccountName",
            "distinguishedName",
            "userAccountControl"
        ),
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=3)]
        [string]
        $SearchScope='subtree',
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=4)]
        [int]
        $PageSize=116        
    )

    begin {
    }
    process {
        $search = New-Object System.DirectoryServices.DirectorySearcher
        $search.SearchRoot = $SearchRoot
        $search.PageSize = $PageSize    
        $search.SearchScope = $SearchScope
        
        $Properties | 
            ForEach-Object {
                $search.PropertiesToLoad.Add($_) | Out-Null
            }

        $search.filter = $LDAPFilter

        $rs = $search.FindAll()
        $rs | 
            ForEach-Object {
                New-Object -TypeName psobject -Property $_.Properties
            }
    }
    end  {
    }
}

function Get-ADObjectWellKnownRid {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0,ParameterSetName="ParameterSet1")]
        [int]$Rid,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=1,ParameterSetName="ParameterSet2")]
        [string]$Name,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2,ParameterSetName="ParameterSet3")]
        [switch]$All,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=3,ParameterSetName="ParameterSet4")]
        [switch]$Protected
    )

    begin { 
        $hash = @{  

        # https://msdn.microsoft.com/en-us/library/cc223144.aspx
            DOMAIN_USER_RID_ADMIN                  = 0x000001F4
            DOMAIN_USER_RID_GUEST                  = 0x000001F5
            DOMAIN_USER_RID_KRBTGT                 = 0x000001F6
            DOMAIN_GROUP_RID_ADMINS                = 0x00000200
            DOMAIN_GROUP_RID_USERS                 = 0x00000201
            DOMAIN_GROUP_RID_COMPUTERS             = 0x00000203
            DOMAIN_GROUP_RID_CONTROLLERS           = 0x00000204
            DOMAIN_GROUP_RID_CERT_PUBLISHERS       = 0x00000205
            DOMAIN_GROUP_RID_SCHEMA_ADMINS         = 0x00000206
            DOMAIN_GROUP_RID_ENTERPRISE_ADMINS     = 0x00000207
            DOMAIN_GROUP_RID_POLICY_CREATOR_OWNERS = 0x00000208
            DOMAIN_GROUP_RID_READONLY_CONTROLLERS  = 0x00000209

        # https://msdn.microsoft.com/en-us/library/windows/desktop/aa379649(v=vs.85).aspx
            DOMAIN_ALIAS_RID_ADMINS                         = 0x00000220 # A local group used for administration of the domain.    
            DOMAIN_ALIAS_RID_USERS                          = 0x00000221 # A local group that represents all users in the domain.
            DOMAIN_ALIAS_RID_GUESTS                         = 0x00000222 # A local group that represents guests of the domain.
            DOMAIN_ALIAS_RID_POWER_USERS                    = 0x00000223 # A local group used to represent a user or set of users who expect to treat a system as if it were their personal computer rather than as a workstation for multiple users.
            DOMAIN_ALIAS_RID_ACCOUNT_OPS                    = 0x00000224 # A local group that exists only on systems running server operating systems. This local group permits control over nonadministrator accounts.
            DOMAIN_ALIAS_RID_SYSTEM_OPS                     = 0x00000225 # A local group that exists only on systems running server operating systems. This local group performs system administrative functions, not including security functions. It establishes network shares, controls printers, unlocks workstations, and performs other operations.
            DOMAIN_ALIAS_RID_PRINT_OPS                      = 0x00000226 # A local group that exists only on systems running server operating systems. This local group controls printers and print queues.
            DOMAIN_ALIAS_RID_BACKUP_OPS                     = 0x00000227 # A local group used for controlling assignment of file backup-and-restore privileges.
            DOMAIN_ALIAS_RID_REPLICATOR                     = 0x00000228 # A local group responsible for copying security databases from the primary domain controller to the backup domain controllers. These accounts are used only by the system.
            DOMAIN_ALIAS_RID_RAS_SERVERS                    = 0x00000229 # A local group that represents RAS and IAS servers. This group permits access to various attributes of user objects.
            DOMAIN_ALIAS_RID_PREW2KCOMPACCESS               = 0x0000022A # A local group that exists only on systems running Windows 2000 Server. For more information, see Allowing Anonymous Access.
            DOMAIN_ALIAS_RID_REMOTE_DESKTOP_USERS           = 0x0000022B # A local group that represents all remote desktop users.
            DOMAIN_ALIAS_RID_NETWORK_CONFIGURATION_OPS      = 0x0000022C # A local group that represents the network configuration. 
            DOMAIN_ALIAS_RID_INCOMING_FOREST_TRUST_BUILDERS = 0x0000022D # A local group that represents any forest trust users.
            DOMAIN_ALIAS_RID_MONITORING_USERS               = 0x0000022E # A local group that represents all users being monitored.
            DOMAIN_ALIAS_RID_LOGGING_USERS                  = 0x0000022F # A local group responsible for logging users.
            DOMAIN_ALIAS_RID_AUTHORIZATIONACCESS            = 0x00000230 # A local group that represents all authorized access.
            DOMAIN_ALIAS_RID_TS_LICENSE_SERVERS             = 0x00000231 # A local group that exists only on systems running server operating systems that allow for terminal services and remote access.
            DOMAIN_ALIAS_RID_DCOM_USERS                     = 0x00000232 # A local group that represents users who can use Distributed Component Object Model (DCOM).
            DOMAIN_ALIAS_RID_IUSERS                         = 0X00000238 # A local group that represents Internet users.
            DOMAIN_ALIAS_RID_CRYPTO_OPERATORS               = 0x00000239 # A local group that represents access to cryptography operators.
            DOMAIN_ALIAS_RID_CACHEABLE_PRINCIPALS_GROUP     = 0x0000023B # A local group that represents principals that can be cached.
            DOMAIN_ALIAS_RID_NON_CACHEABLE_PRINCIPALS_GROUP = 0x0000023C # A local group that represents principals that cannot be cached.
            DOMAIN_ALIAS_RID_EVENT_LOG_READERS_GROUP        = 0x0000023D # A local group that represents event log readers.
            DOMAIN_ALIAS_RID_CERTSVC_DCOM_ACCESS_GROUP      = 0x0000023E # The local group of users who can connect to certification authorities using Distributed Component Object Model (DCOM).
        }     
    }
    process {
        $hash.GetEnumerator() | 
            Where-Object {
                $_.Value -eq $Rid
            }

        $hash.GetEnumerator() | 
            Where-Object {
                $_.Name -eq $Name
            }

        if ($All) {
            $hash
        }

        if ($Protected) {
            # https://technet.microsoft.com/en-us/library/2009.09.sdadminholder.aspx
            "DOMAIN_ALIAS_RID_ACCOUNT_OPS",
            "DOMAIN_USER_RID_ADMIN",
            "DOMAIN_ALIAS_RID_ADMINS",
            "DOMAIN_ALIAS_RID_BACKUP_OPS",
            "DOMAIN_GROUP_RID_ADMINS",
            "DOMAIN_GROUP_RID_CONTROLLERS",
            "DOMAIN_GROUP_RID_ENTERPRISE_ADMINS",
            "DOMAIN_USER_RID_KRBTGT",
            "DOMAIN_ALIAS_RID_PRINT_OPS",
            "DOMAIN_GROUP_RID_READONLY_CONTROLLERS",
            "DOMAIN_ALIAS_RID_REPLICATOR",
            "DOMAIN_GROUP_RID_SCHEMA_ADMINS",
            "DOMAIN_ALIAS_RID_SYSTEM_OPS" |
                Get-ADObjectWellKnownRid
        }
    }
    end {
    }
}

function Get-ADOrphanProtectedGroups {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)]
        [system.directoryservices.directoryentry]$Domain=(Get-ADDomain)
    )

    begin {
    }
    process {
        [string[]]$pGroups = $Domain | 
            Get-ADProtectedGroups -Recurse |
            Select-Object -ExpandProperty distinguishedname

        $Domain | 
            Get-ADObject -LDAPFilter "(&(objectCategory=Group)(adminCount=1))" |
            ForEach-Object {
                if ($_.distinguishedname -notin $pGroups) {
                    $_
                }
            }
    }
    end {
    }
}

function Get-ADOrphanProtectedUsers {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)]
        [system.directoryservices.directoryentry]$Domain=(Get-ADDomain)
    )

    begin {
    }
    process {
        [string[]]$pUsers = $Domain | 
            Get-ADProtectedUsers |
            Select-Object -ExpandProperty distinguishedname

        $Domain | 
            Get-ADObject -LDAPFilter "(&(objectCategory=User)(adminCount=1))" |
            ForEach-Object {
                if ($_.distinguishedname -notin $pUsers) {
                    $_
                }
            }
    }
    end {
    }
}

function Get-ADProtectedGroups {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=0)]
        [switch]$Default,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [switch]$Recurse,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=2)]
        [system.directoryservices.directoryentry]$Domain=(Get-ADDomain)
    )

    begin {
    }
    process {
        $domainSID = $Domain | 
            Select-Object -ExpandProperty objectSid
        $domainSID[1] = 5

        if ($Default) {
            $pObjects =  Get-ADObjectWellKnownRid -Protected 
        }
        else {
            $pObjects = Get-ADObjectWellKnownRid -Protected | 
                Where-Object {
                    $_.Name -notin (Get-ADdwAdminSDExMask)
                }
        }

        $pObjects | 
            ForEach-Object {
                if ($_.Name.Contains("_ALIAS_")) {
                    $searchSid = [byte[]]@(1,2,0,0,0,0,0,5,32,0,0,0) + ($_.Value | ConvertTo-ByteRid) | 
                            ConvertTo-StringSid
                }
                else {
                    $searchSid = $domainSID + ($_.Value | ConvertTo-ByteRid) |
                        ConvertTo-StringSid 
                } 

                $pObject = $Domain |
                    Get-ADObject -Properties objectCategory, distinguishedName -LDAPFilter (
                        "objectSid={0}" -f $searchSid
                    )

                if ($pObject.objectCategory.ToLower().Contains('group')) {
                    $pObject
                    if ($Recurse) {
                        $pObject | 
                            Select-Object -ExpandProperty distinguishedName |
                            Get-ADChildGroup -Recurse
                    }   
                }
            } | 
            Select-Object -Unique distinguishedName
    }
    end {
    }
}

function Get-ADProtectedUsers {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=0)]
        [switch]$Default,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=1)]
        [system.directoryservices.directoryentry]$Domain=(Get-ADDomain)
    )

    begin {
    }
    process {
        $domainSID = $Domain | 
            Select-Object -ExpandProperty objectSid
        $domainSID[1] = 5

        if ($Default) {
            $pObjects =  Get-ADObjectWellKnownRid -Protected 
        }
        else {
            $pObjects = Get-ADObjectWellKnownRid -Protected | 
                Where-Object {
                    $_.Name -notin (Get-ADdwAdminSDExMask)
                }
        }

        $pObjects | 
            ForEach-Object {
                if ($_.Name.Contains("_ALIAS_")) {
                    $searchSid = [byte[]]@(1,2,0,0,0,0,0,5,32,0,0,0) + ($_.Value | ConvertTo-ByteRid) | 
                            ConvertTo-StringSid
                }
                else {
                    $searchSid = $domainSID + ($_.Value | ConvertTo-ByteRid) |
                        ConvertTo-StringSid 
                } 

                $pObject = $Domain |
                    Get-ADObject -Properties objectCategory, distinguishedName -LDAPFilter (
                        "objectSid={0}" -f $searchSid
                    )

                if ($pObject.objectCategory.ToLower().Contains('person')) {
                    $pObject
                }
            } | 
            Select-Object -Unique distinguishedName
    
        if (!$Default) {
            $Domain | 
                Get-ADProtectedGroups -Recurse | 
                Select-Object -ExpandProperty distinguishedName |
                Get-ADChildUser | 
                select -Unique distinguishedname
        }
    }
    end {
    }
}

function Set-ADObjectAccessRuleProtection {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [System.DirectoryServices.DirectoryEntry]$InputObject,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=1)]
        [bool]$IsProtected,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=2)]
        [bool]$PreserveInheritance,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=3)]
        [switch]$Force
    )    

    begin {
    }
    process {
        #$InputObject.psbase.Options.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl
        $InputObject.psbase.ObjectSecurity.SetAccessRuleProtection($IsProtected, $PreserveInheritance)
        if ($Force) {
            $p = 'y'
        }
        else {
            $p = Read-Host "Confirm (y/n)"
        }
        if ($p.ToLower().Trim() -eq 'y') {
            $InputObject.psbase.CommitChanges()
        }
        else {
            $InputObject = $null
            Write-Error "Operation 'SetAccessRuleProtection' aborted by user."
        }
    }
    end {
    }
}

function Test-ADdSHeuristics {
    [CmdletBinding()]

    param (
    )    

    begin {
    }  
    process {
        ("CN=Configuration,{0}" -f (Get-ADForest | Select-Object -ExpandProperty distinguishedName) | 
            Get-ADDirectoryEntry |
            Get-ADObject -LDAPFilter "(name=Directory Service)" -Properties dSHeuristics | 
            Get-Member |
            Select-Object -ExpandProperty Name) -contains 'dSHeuristics'
    }
    end {
    }    
}

