function Get-GpoPermission {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        [Microsoft.GroupPolicy.Gpo]   
        $InputObject
    )

    begin {
    }
    process {
        [adsi]('LDAP://CN={' + $InputObject.Id.Guid + '},CN=Policies,CN=System,' + ([adsi]'LDAP://RootDSE').defaultNamingContext) |
            Get-ADObjectAccessRight |
            ForEach-Object {
                $_.Parent_canonicalName = $_.Parent_canonicalName.ToLower().Replace("$($InputObject.Id.Guid.ToLower())","$($InputObject.DisplayName)")
                $_.Parent_distinguishedName = $_.Parent_distinguishedName.ToLower().Replace("$($InputObject.Id.Guid.ToLower())","$($InputObject.DisplayName)")
                $_
            }
    }
    end {
    } 
}

function Set-GpoPermission {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        $InputObject,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [switch]
        $Force = $false
    )

    begin {
    }
    process {
        $gpo = Get-GPO -DisplayName $InputObject.Parent_canonicalName.split('/')[-1].Trim('{}')
        Write-Verbose -Message ("[    ] Set-GpoPermission : {0}" -f $gpo.DisplayName)
        $_.Parent_canonicalName = $_.Parent_canonicalName.ToLower().Replace("$($gpo.DisplayName.ToLower())","$($gpo.Id.Guid)")
        $_.Parent_distinguishedName = $_.Parent_distinguishedName.ToLower().Replace("$($gpo.DisplayName.ToLower())","$($gpo.Id.Guid)")
        $_.Parent_distinguishedName |
            Get-ADDirectoryEntry |             
            Remove-ADObjectAccessRight `
                -IdentityReference 'NT AUTHORITY\Authenticated Users' `
                -ActiveDirectoryRights 'ExtendedRight' `
                -AccessControlType 'Allow' `
                -ObjectType 'edacfd8f-ffb3-11d1-b41d-00a0c968f939' `
                -InheritanceType 'All' `
                -InheritedObjectType '00000000-0000-0000-0000-000000000000' `
                -Confirm:(-not $Force)       
        $_ | Set-ADObjectAccessRight -Force:($Force)
    }
    end {
    } 
}

function Get-GpoLink {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        [Microsoft.GroupPolicy.Gpo]   
        $InputObject
    )

    begin {
        $gpoLinks = [string[]](Get-ADOrganizationalUnit -Filter * | Select-Object -ExpandProperty distinguishedName) + `
            [string[]]([adsi]'LDAP://RootDSE').defaultNamingContext |
                Get-GPInheritance |
                Select-Object -ExpandProperty GpoLinks
    }
    process {
        $gpoLinks |
            Where-Object {
                $_.DisplayName -eq $InputObject.DisplayName
            } |
            ForEach-Object {
                if ($_.Enabled) {$linkEnabled = 'Yes'} else {$linkEnabled = 'No'}
                if ($_.Enforced) {$linkEnforced = 'Yes'} else {$linkEnforced = 'No'}
                New-Object -TypeName psobject -Property @{
                   'Id' = $_.GpoId
                   'DisplayName' = $_.DisplayName
                   'LinkOrder' = $_.Order
                   'LinkEnabled' = $linkEnabled
                   'LinkEnforced' = $linkEnforced
                   'LinkTarget' = $_.Target
                }
            }
    }
    end {
    } 
}

function Set-GpoLink {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        $InputObject
    )

    begin {
    }
    process {
        Write-Verbose -Message ("[    ] Set-GpoLink : {0} -> {1}" -f $InputObject.DisplayName, $InputObject.LinkTarget)
            if (
                (Get-GPInheritance -Target $InputObject.LinkTarget |
                Select-Object -ExpandProperty GpoLinks |
                Select-Object -ExpandProperty DisplayName) -contains $InputObject.DisplayName
            ) {
                Set-GPLink -Name $_.DisplayName -Target $_.LinkTarget -Order $_.LinkOrder -LinkEnabled $_.LinkEnabled -Enforced $_.LinkEnforced 
            }
            else {
                New-GPLink -Name $_.DisplayName -Target $_.LinkTarget -Order $_.LinkOrder -LinkEnabled $_.LinkEnabled -Enforced $_.LinkEnforced 
            }
    }
    end {
    }
}

function Get-GpoWmiFilter {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        [Microsoft.GroupPolicy.Gpo]
        $InputObject
    )

    begin {
    }
    process {
        if ($InputObject.WmiFilter) {
            $wmiFilterGuid = $InputObject.WmiFilter.Path.Split("{")[1].Split("}")[0]

            $objWmiFilter = ActiveDirectory\Get-ADObject -LDAPFilter "(&(objectClass=msWMI-Som)(Name={$wmiFilterGuid}))" `
                -Properties "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2"

            $wmiFilterName        = $objWmiFilter | Select-Object -ExpandProperty "msWMI-Name"
            $wmiFilterDescription = $objWmiFilter | Select-Object -ExpandProperty "msWMI-Parm1"
            $wmiFilterQueryList   = $objWmiFilter | Select-Object -ExpandProperty "msWMI-Parm2"            

            New-Object -TypeName psobject -Property @{
                'Id' = $InputObject.GpoId
                'DisplayName' = $InputObject.DisplayName
                'wmiFilterName' = $wmiFilterName
                'WmiFilterDescription' = $wmiFilterDescription
                'wmiFilterQueryList' = $wmiFilterQueryList
             }
        }
    }
    end {
    }
}

function Set-GpoWmiFilter {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        $InputObject
    )

    begin {
    }
    process {
        Write-Verbose -Message ("[    ] Set-GpoWmiFilter : {0} -> {1}" -f $InputObject.WmiFilterName, $InputObject.DisplayName)

        $objWmiFilter = ActiveDirectory\Get-ADObject -LDAPFilter "(&(objectClass=msWMI-Som)(msWMI-Name=$($InputObject.WmiFilterName)))" `
                -Properties "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2"

        if (($objWmiFilter | Measure-Object).Count -gt 1) {
            $wmiFilterName =  $objWmiFilter | Select-Object -ExpandProperty "msWMI-Name"
            Write-Error -Message "There are multiple WMI Filters named '$wmiFilterName'."
            break
        }
        elseif (($objWmiFilter | Measure-Object).Count -eq 1) {
            $wmiFilterName =  $objWmiFilter | Select-Object -ExpandProperty "msWMI-Name"
            $wmiFilterDescription =  $objWmiFilter | Select-Object -ExpandProperty "msWMI-Parm1"
            $wmiFilterQueryList =  $objWmiFilter | Select-Object -ExpandProperty "msWMI-Parm2"
            
            if (($wmiFilterQueryList -ne $InputObject.WmiFilterQueryList) -and ($wmiFilterDescription -ne $InputObject.WmiFilterDescription)) {
                Write-Error -Message "A WMI Filter named '$wmiFilterName' already exists with a different Query List and Description."
                break
            }
            elseif ($wmiFilterQueryList -ne $InputObject.WmiFilterQueryList) {
                Write-Error -Message "A WMI Filter named '$wmiFilterName' already exists with a different Query List."
                break
            }
            elseif ($wmiFilterDescription -ne $InputObject.WmiFilterDescription) {
                Write-Error -Message "A WMI Filter named '$wmiFilterName' already exists with a different Description."
                break
            }
        } 
        else {
            $defaultNamingContext = (Get-ADRootDSE).DefaultNamingContext

            $guid = [System.Guid]::NewGuid()
            $msWMICreationDate = (Get-Date).ToUniversalTime().ToString("yyyyMMddhhmmss.ffffff-000")
            
            $otherAttributes = @{
                "msWMI-Name" = $InputObject.wmiFilterName;
                "msWMI-Parm1" = $InputObject.WmiFilterDescription;
                "msWMI-Parm2" = $InputObject.wmiFilterQueryList;
                "msWMI-ID"= "{$guid}";
                "instanceType" = 4;
                "showInAdvancedViewOnly" = "TRUE";
                "distinguishedname" = "CN={$guid},CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext";
                "msWMI-ChangeDate" = $msWMICreationDate; 
                "msWMI-CreationDate" = $msWMICreationDate
            }

            if ($InputObject.WmiFilterDescription -eq $null) { 
                $otherAttributes.Remove("msWMI-Parm1") 
            }
            if ($InputObject.wmiFilterQueryList -eq $null) {
                $otherAttributes.Remove("msWMI-Parm2") 
            }
                
            New-ADObject -Name "{$guid}" -Type "msWMI-Som" -Path ("CN=SOM,CN=WMIPolicy,CN=System,$defaultNamingContext") -OtherAttributes $otherAttributes -PassThru
        }

        $objWmiFilter = ActiveDirectory\Get-ADObject -LDAPFilter "(&(objectClass=msWMI-Som)(msWMI-Name=$($InputObject.wmiFilterName)))" `
            -Properties "msWMI-Name", "msWMI-Parm1", "msWMI-Parm2"

        $gpDomain = New-Object -Type Microsoft.GroupPolicy.GPDomain

        $gpo = Get-GPO -DisplayName $InputObject.DisplayName
        $gpo.WmiFilter = $gpDomain.GetWmiFilter('MSFT_SomFilter.ID="' + $objWmiFilter.Name + '",Domain="' + $gpDomain.DomainName +'"')
    }
    end {
    }
}

function Backup-Gpo {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)] 
        [Microsoft.GroupPolicy.Gpo]   
        $InputObject,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=1)]
        $Path
    )

    begin {
        if (Test-Path -Path $Path) {
            Write-Error -Message ("An item with the specified name '{0}' already exists." -f $Path)
            break
        }
        $Path = New-Item -Force -ItemType Directory -Path $Path
    }
    process {
        if ($InputObject.DisplayName.EndsWith(' ')) {
            Write-Warning ("Not backing up policy '{0}' because it's DisplayName ends with a space (' ')." -f $InputObject.DisplayName)
        }
        elseif ($InputObject.DisplayName.StartsWith(' ')) {
            Write-Warning ("Not backing up policy '{0}' because it's DisplayName starts with a space (' ')." -f $InputObject.DisplayName)
        }
        else {
            $InputObject | 
                GroupPolicy\Backup-GPO -Path $Path |
                Export-Csv -Append -NoTypeInformation -Path (Join-Path -Path $Path -ChildPath 'policy.config.csv')
            $InputObject |
                Get-GpoPermission |
                Export-Csv -Append -NoTypeInformation -Path (Join-Path -Path $Path -ChildPath 'ace.config.csv')
            $InputObject | 
                Get-GpoLink |
                Export-Csv -Append -NoTypeInformation -Path (Join-Path -Path $Path -ChildPath 'link.config.csv')  
            $InputObject | 
                Get-GpoWmiFilter |
                Export-Csv -Append -NoTypeInformation -Path (Join-Path -Path $Path -ChildPath 'wmifilter.config.csv')                                 
        }
    }
    end {
        ActiveDirectory\Get-ADObject -SearchBase ('CN=Partitions,CN=Configuration,' + ([adsi]'LDAP://RootDSE').defaultNamingContext) -Filter * -Properties netbiosname | 
            Select-Object -ExpandProperty netbiosname |
            Out-File -FilePath (Join-Path -Path $Path -ChildPath 'netbiosname')

        Import-Csv -Path (Join-Path -Path $Path -ChildPath 'policy.config.csv') | 
            Select-Object -ExpandProperty DisplayName |
            New-GpoMigrationTable -Path $Path
    }
}

function Restore-Gpo {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=0)]
        $Path,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=1)]
        [switch]$IncludeLinks,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=2)]
        [switch]$LinksOnly,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=3)]
        [switch]$IncludeWmiFilters, 
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=4)]
        [switch]$WmiFiltersOnly,        
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=5)]
        [switch]$IncludePermissions,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=6)]
        [switch]$PermissionsOnly,
        [Parameter(Mandatory=$false,ValueFromPipeline=$false,Position=7,ParameterSetName='Parameter Set 1')]
        [switch]$DoNotMigrateSAMAccountName,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=8,ParameterSetName='Parameter Set 2')]
        [switch]$MigrateSAMAccountNameSameAsSource,          
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=9,ParameterSetName='Parameter Set 3')]
        [switch]$MigrateSAMAccountNameByRelativeName,         
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=10,ParameterSetName='Parameter Set 4')]
        [switch]$MigrateSAMAccountNameUsingMigrationTable,
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=11,ParameterSetName='Parameter Set 4')]
        [string]$MigrationTable
    )

    begin {
        $netBiosName_source = Get-Content (Join-Path -Path $Path -ChildPath 'netbiosname').trim()
        $netBiosName_target = ActiveDirectory\Get-ADObject -SearchBase ('CN=Partitions,CN=Configuration,' + ([adsi]'LDAP://RootDSE').defaultNamingContext) -Filter * -Properties netbiosname | 
            Select-Object -ExpandProperty netbiosname

        if ($DoNotMigrateSAMAccountName) {
           $MigrationTable = (Join-Path -Path $Path -ChildPath 'destination.none.migtable') 
        }

        if ($MigrateSAMAccountNameSameAsSource) {
            $MigrationTable = (Join-Path -Path $Path -ChildPath 'destination.sameassource.migtable')
        }

        if ($MigrateSAMAccountNameByRelativeName) {
            $MigrationTable = (Join-Path -Path $Path -ChildPath 'destination.byrelativename.migtable')
        }

        if ($MigrateSAMAccountNameUsingMigrationTable) {
            if (-not (Test-Path $MigrationTable)) {
                Write-Error -Message "Migration table file not found '{$MigrationTable}'."
                break
            }
        }
    }
    process {
        if ((-not $LinksOnly) -and (-not $WmiFiltersOnly) -and (-not $PermissionsOnly)) {
            Import-Csv -Path (Join-Path -Path $Path -ChildPath 'policy.config.csv') | 
                ForEach-Object {
                    Write-Verbose -Message ("[    ] Restore-Gpo : {0}" -f $_.DisplayName)
                    if ($MigrationTable.Length -gt 0) {
                        GroupPolicy\Import-GPO -CreateIfNeeded -BackupGpoName $_.DisplayName -TargetName $_.DisplayName -Path $Path -MigrationTable $MigrationTable
                    } 
                    else {
                        GroupPolicy\Import-GPO -CreateIfNeeded -BackupGpoName $_.DisplayName -TargetName $_.DisplayName -Path $Path
                    }
                }
        }

        if ($IncludePermissions -or $PermissionsOnly) {
            Import-Csv -Path (Join-Path -Path $Path -ChildPath 'ace.config.csv') | 
                ForEach-Object {
                    if ($_.IsInherited -eq 'False') {
                        $_.__AddRemoveIndicator = 1
                    }
                    $_.IdentityReference = $_.IdentityReference.ToUpper().Replace("$($netBiosName_source.ToUpper())\","$($netBiosName_target.ToUpper())\")
                    if ($_.Parent_canonicalName -like "*/*") {
                        $_.Parent_canonicalName = (([adsi]'LDAP://RootDSE').defaultNamingContext.ToString() -replace('dc=','') -replace(',','.')) + $_.Parent_canonicalName.SubString($_.Parent_canonicalName.IndexOf('/'))
                    }
                    else {
                        $_.Parent_canonicalName = ([adsi]'LDAP://RootDSE').defaultNamingContext -replace('dc=','') -replace(',','.')
                    }    
                    $_.Parent_distinguishedName = $_.Parent_distinguishedName.SubString(0,$_.Parent_distinguishedName.ToUpper().IndexOf('DC=')) + ([adsi]"LDAP://RootDSE").defaultNamingContext
                    $_ | Set-GpoPermission -Force -Verbose:$VerbosePreference
                }
        }

        if ($IncludeWmiFilters -or $WmiFiltersOnly) {
            Import-Csv -Path (Join-Path -Path $Path -ChildPath 'wmifilter.config.csv') | 
                ForEach-Object {
                    if ($_.WmiFilterDescription.Length -lt 1) {
                        $_.WmiFilterDescription = $null
                    }
                    if ($_.WmiFilterQueryList.Length -lt 1) {
                        $_.WmiFilterQueryList = $null
                    }        
                    $_ | Set-GpoWmiFilter -Verbose:$VerbosePreference
                }   
        }

        if ($IncludeLinks -or $LinksOnly) {
            Import-Csv -Path (Join-Path -Path $Path -ChildPath 'link.config.csv') | 
                ForEach-Object {        
                    $_.LinkTarget = $_.LinkTarget.SubString(0,$_.LinkTarget.ToUpper().IndexOf('DC=')) + ([adsi]"LDAP://RootDSE").defaultNamingContext
                    $_ | Set-GpoLink -Verbose:$VerbosePreference
                }
        }
    }
    end {
    }
}

function New-GpoMigrationTable {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$false,Position=0)]
        $Path,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)]
        [string]
        $GpoDisplayName
    )

    begin {
        $netbiosname = ActiveDirectory\Get-ADObject -SearchBase ('CN=Partitions,CN=Configuration,' + ([adsi]'LDAP://RootDSE').defaultNamingContext) -Filter * -Properties netbiosname | 
            Select-Object -ExpandProperty netbiosname
        
        [string[]]$samAccountNames = @()
        
        [string[]]$migTable_mapN = '<?xml version="1.0" encoding="utf-16"?>'
        [string[]]$migTable_mapS = '<?xml version="1.0" encoding="utf-16"?>'
        [string[]]$migTable_mapR = '<?xml version="1.0" encoding="utf-16"?>'

        $migTable_mapN += '<MigrationTable xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable">'
        $migTable_mapS += '<MigrationTable xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable">'
        $migTable_mapR += '<MigrationTable xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable">'
    }
    process {
        $r = Get-GPOReport -DisplayName $GpoDisplayName -ReportType Html
        [regex]::Matches($r,'(?xi) ' + $netbiosname + ' \\ [\w\s]+') |
            Select-Object -ExpandProperty value | 
            ForEach-Object {
                if (-not ($samAccountNames -contains $_)) {
                    $samAccountNames += $_
                    $o = ActiveDirectory\Get-ADObject -LDAPFilter ("(samAccountName={0})" -f ($_ | Split-Path -Leaf))

                    switch ($o | Select-Object -ExpandProperty ObjectClass) {
                        'user'          {$type = 'User'}
                        'inetOrgPerson' {$type = 'User'}
                        'computer'      {$type = 'Computer'}
                        'group'         {
                                            $type = "{0}Group" -f ($o | Get-ADGroup | Select-Object -ExpandProperty GroupScope) -replace 'domain',''
                                        }
                        default         {$type = 'Unknown' }
                    }

                    $migTable_mapN += '<Mapping>'
                    $migTable_mapS += '<Mapping>'
                    $migTable_mapR += '<Mapping>'

                    $migTable_mapN += "<Type>{0}</Type>" -f $type
                    $migTable_mapS += "<Type>{0}</Type>" -f $type
                    $migTable_mapR += "<Type>{0}</Type>" -f $type

                    $migTable_mapN += "<Source>{0}</Source>" -f $_
                    $migTable_mapS += "<Source>{0}</Source>" -f $_
                    $migTable_mapR += "<Source>{0}</Source>" -f $_
                    
                    $migTable_mapN += '<DestinationNone />'
                    $migTable_mapS += '<DestinationSameAsSource />'
                    $migTable_mapR += '<DestinationByRelativeName />'

                    $migTable_mapN += '</Mapping>'
                    $migTable_mapS += '</Mapping>'
                    $migTable_mapR += '</Mapping>'
                }
            }
    }
    end {
        $migTable_mapN += '</MigrationTable>'
        $migTable_mapS += '</MigrationTable>'
        $migTable_mapR += '</MigrationTable>'
        
        $migTable_mapN | Out-File -FilePath (Join-Path -Path $Path -ChildPath 'destination.none.migtable')
        $migTable_mapS | Out-File -FilePath (Join-Path -Path $Path -ChildPath 'destination.sameassource.migtable')
        $migTable_mapR | Out-File -FilePath (Join-Path -Path $Path -ChildPath 'destination.byrelativename.migtable')
    }
}


# This check is not ideal, need to replace with a propper .psd1 manifest
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}
if (-not (Get-Module -Name GroupPolicy)) {
    Import-Module GroupPolicy
}
if (-not (Get-Module -Name ADObjectAccessRight)) {
    Write-Error -Message "This module requires the 'ADObjectAccessRight' module. Import the 'ADObjectAccessRight' module and then try to import this module again."
    Export-ModuleMember
    break
}

Export-ModuleMember `
    -Function @(
        'Get-GpoPermission',
        'Set-GpoPermission',
        'Get-GpoWmiFilter',
        'Set-GpoWmiFilter',
        'Get-GpoLink',
        'Set-GpoLink',
        'Backup-Gpo',
        'Restore-Gpo',
        'New-GpoMigrationTable'
    )
