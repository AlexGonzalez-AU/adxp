function Get-GpoReport_AllowLogonRights {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    $sb = {
        foreach ($gpo in (Get-GPO -All -Domain $Domain)) {
            if (!$gpo.displayname) {
                Write-Warning -Message ("GPO is missing from SYSVOL '{0}'" -f $gpo.id.guid)
                continue
            }
            Write-Host ("Searching: '{0}'..." -f $gpo.displayname)

            ($gpo | Get-GPOReport -ReportType Html -Domain $Domain).split("`n") |
                Select-String -CaseSensitive -Pattern @(
                    "^<tr><td>Access this computer from the network", 
                    "^<tr><td>Allow log on locally",
                    "^<tr><td>Allow log on through Terminal Services",
                    "^<tr><td>Log on as a batch job",
                    "^<tr><td>Log on as a service"
                ) |
                ForEach-Object {
                    foreach ($identity in ($_.line.replace("<tr><td>","").replace("</td><td>","`t").replace("</td></tr>","").split("`t")[1].replace(", ",",").split(","))) {
                        New-Object -TypeName psobject -Property @{ 
                            GpoDisplayName = $gpo.DisplayName
                            GpoGuid = $gpo.id.guid
                            UserRightsAssignment = $_.line.replace("<tr><td>","").replace("</td><td>","`t").replace("</td></tr>","").split("`t")[0]
                            Identity = $identity
                        }
                    }
                }
        }
    }

    if ($AsCsvFile) {
        & $sb |
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoReport_AllowLogonRights.csv" -f $Domain.Split('.').ToUpper())    
    }
    else {
        & $sb
    }
}

function Get-GpoReport_DenyLogonRights {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    $sb = {
        foreach ($gpo in (Get-GPO -All -Domain $Domain)) {
            if (!$gpo.displayname) {
                Write-Warning -Message ("GPO is missing from SYSVOL '{0}'" -f $gpo.id.guid)
                continue
            }
            Write-Host ("Searching: '{0}'..." -f $gpo.displayname)

            ($gpo | Get-GPOReport -ReportType Html -Domain $Domain).split("`n") |
                Select-String -CaseSensitive -Pattern @(
                    "^<tr><td>Deny access to this computer from the network", 
                    "^<tr><td>Deny log on as a batch job",
                    "^<tr><td>Deny log on as a service",
                    "^<tr><td>Deny log on locally",
                    "^<tr><td>Deny log on through Terminal Services"
                ) |
                ForEach-Object {
                    foreach ($identity in ($_.line.replace("<tr><td>","").replace("</td><td>","`t").replace("</td></tr>","").split("`t")[1].replace(", ",",").split(","))) {
                        New-Object -TypeName psobject -Property @{ 
                            GpoDisplayName = $gpo.DisplayName
                            GpoGuid = $gpo.id.guid
                            UserRightsAssignment = $_.line.replace("<tr><td>","").replace("</td><td>","`t").replace("</td></tr>","").split("`t")[0]
                            Identity = $identity
                        }
                    }
                }
        }
    }

    if ($AsCsvFile) {
        & $sb |
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoReport_DenyLogonRights.csv" -f $Domain.Split('.').ToUpper())    
    }
    else {
        & $sb
    }
}

function Get-GpoReport_GpoAcls {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    $sb = {
        Get-GPO -All -Domain $Domain | 
            ForEach-Object {
                $gpo = $_
                $gpo | Get-GPPermission -All -Domain $Domain | foreach {
                    $ace = $_
                    New-Object -TypeName psobject -Property @{
                        Gpo_DisplayName = $gpo.DisplayName
                        Gpo_Id = $gpo.Id
                        Trustee_Domain = $ace.Trustee.Domain
                        Trustee_Name = $ace.Trustee.Name
                        Trustee_Sid = $ace.Trustee.Sid
                        Trustee_SidType = $ace.Trustee.SidType
                        Permission = $ace.Permission
                        Inherited = $ace.Inherited
                    }
                }
            }
    }

    if ($AsCsvFile) {
        & $sb | 
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoReport_GpoAcls.csv" -f $Domain.Split('.').ToUpper())   
    }
    else {
        & $sb
    }
}

function Get-GpoReport_GpoInheritanceBlockedOus {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    $sb = {
        Get-ADOrganizationalUnit -Filter * -Server $Domain -Properties canonicalName | 
            ForEach-Object {
                Write-Host ("Checking: '{0}'..." -f $_.canonicalName)
                $_ | Get-GPInheritance -Domain $Domain
            } | 
            Select-Object -Property Name, ContainerType, Path, GpoInheritanceBlocked, @{
                n = "GpoLinks"; e = {($_ | Select-Object -ExpandProperty GpoLinks | Sort-Object -Property Order | Select-Object -ExpandProperty DisplayName) -join "`n"}
            }, @{
                n = "InheritedGpoLinks"; e = {($_ | Select-Object -ExpandProperty InheritedGpoLinks | Sort-Object -Property Order | Select-Object -ExpandProperty DisplayName) -join "`n"}
            }
    }

    if ($AsCsvFile) {
        & $sb | 
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoReport_GpoInheritanceBlockedOus.csv" -f $Domain.Split('.').ToUpper())   
    }
    else {
        & $sb
    }
}

function Get-GpoReport_GpoLinks {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    Set-StrictMode -Version 2
    $ErrorActionPreference = 'Stop'

    $sb = {
        function Test-XmlProperty {
            [CmdletBinding()]

            param (
                [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
                $XmlPath,
                [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
                [string]$Property
            )

            [string[]]$properties = $XmlPath | 
                Get-Member -MemberType Properties | 
                Select-Object -ExpandProperty Name

            $properties -Contains($Property)
        }

        foreach ($gpo in (Get-GPO -All -Domain $Domain)) {
            if (!$gpo.displayname) {
                Write-Warning -Message ("GPO is missing from SYSVOL '{0}'" -f $gpo.id.guid)
                continue
            }
            Write-Host ("Checking: '{0}'..." -f $gpo.displayname)

            [xml]$xmlReport = $gpo | Get-GPOReport -ReportType Xml -Domain $Domain

            if (!($xmlReport | Test-XmlProperty -Property 'GPO')) {
                continue
            }
            if (!($xmlReport.GPO | Test-XmlProperty -Property 'LinksTo')) {
                New-Object -TypeName psobject -Property @{
                    GpoDisplayName = $gpo.DisplayName
                    GpoGuid = $gpo.id.guid
                    LinkEnabled = $null
                    LinkEnforced = $null
                    LinkName = $null
                    LinkPath = $null
                }  
                continue
            }
            foreach ($link in $xmlReport.GPO.LinksTo) {
                New-Object -TypeName psobject -Property @{
                    GpoDisplayName = $gpo.DisplayName
                    GpoGuid = $gpo.id.guid
                    LinkEnabled = $link.Enabled
                    LinkEnforced = $link.NoOverride
                    LinkName = $link.SOMName
                    LinkPath = $link.SOMPath
                }       
            }    
        }
    }

    if ($AsCsvFile) {
        & $sb |
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoLinks.csv" -f $Domain.Split('.').ToUpper())    
    }
    else {
        & $sb
    }
}

function Get-GpoReport_RestrictedGroups {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [switch]$AsCsvFile
    )

    Set-StrictMode -Version 2
    $ErrorActionPreference = 'Stop'

    $sb = {
        function Test-XmlProperty {
            [CmdletBinding()]
    
            param (
                [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
                $XmlPath,
                [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
                [string]$Property
            )
    
            [string[]]$properties = $XmlPath | 
                Get-Member -MemberType Properties | 
                Select-Object -ExpandProperty Name
    
            $properties -Contains($Property)
        }

        foreach ($gpo in (Get-GPO -All -Domain $Domain)) { 
            if (!$gpo.displayname) {
                Write-Warning -Message ("GPO is missing from SYSVOL '{0}'" -f $gpo.id.guid)
                continue
            }
            Write-Host ("Searching: '{0}'..." -f $gpo.displayname)

                [xml]$xmlReport = $gpo | Get-GPOReport -ReportType Xml -Domain $Domain

                if (!($xmlReport | Test-XmlProperty -Property 'GPO')) {
                    continue
                }
                if (!($xmlReport.GPO | Test-XmlProperty -Property 'Computer')) {
                    continue
                }
                if (!($xmlReport.GPO.Computer | Test-XmlProperty -Property 'ExtensionData')) {
                    continue
                }

                foreach ($ExtensionData in $xmlReport.GPO.Computer.ExtensionData) {
                    if (!($ExtensionData | Test-XmlProperty -Property 'Extension')) {
                        continue
                    }
                    if (($ExtensionData.Extension | Test-XmlProperty -Property 'RestrictedGroups')) {
                        foreach ($restrictedGroup in $ExtensionData.Extension.RestrictedGroups) {
                            $groupName = $null
                            if (($restrictedGroup.GroupName | Test-XmlProperty -Property 'Name')) {
                                $groupName = $restrictedGroup.GroupName.Name.'#text'
                            }
                            $groupSid = $null            
                            if (($restrictedGroup.GroupName | Test-XmlProperty -Property 'Sid')) {
                                $groupSid = $restrictedGroup.GroupName.Sid.'#text'
                            }        
            
                            if (!($restrictedGroup | Test-XmlProperty -Property 'Member')) {
                                New-Object -TypeName psobject -Property @{
                                    GpoDisplayName = $gpo.DisplayName
                                    GpoGuid = $gpo.id.guid
                                    Action = 'R'
                                    RestrictedGroupName = $groupName
                                    RestrictedGroupSid = $groupSid
                                    MemberName = '<empty>'
                                    MemberSid = '<empty>'
                                    LocalGroupName = $null
                                    LocalGroupSid = $null
                                    SettingType = "Policy:RestrictedGroups"
                                }
                            }
                            else {
                                foreach ($member in $restrictedGroup.Member) {
                                    $memberName = $null
                                    if (($member | Test-XmlProperty -Property 'Name')) {
                                        $memberName = $member.Name.'#text'
                                    }
                                    $memberSid = $null            
                                    if (($member | Test-XmlProperty -Property 'Sid')) {
                                        $memberSid = $member.Sid.'#text'
                                    }        
                                    New-Object -TypeName psobject -Property @{
                                        GpoDisplayName = $gpo.DisplayName
                                        GpoGuid = $gpo.id.guid
                                        Action = 'R'
                                        RestrictedGroupName = $groupName
                                        RestrictedGroupSid = $groupSid
                                        MemberName = $memberName
                                        MemberSid = $memberSid
                                        LocalGroupName = $null
                                        LocalGroupSid = $null
                                        SettingType = "Policy:RestrictedGroups"
                                    }
                                }
                            }
                        }
                    }
            
                    if (($ExtensionData.Extension | Test-XmlProperty -Property 'LocalUsersAndGroups')) {
                        if (($ExtensionData.Extension.LocalUsersAndGroups | Test-XmlProperty -Property 'Group')) {
                            foreach ($group in $ExtensionData.Extension.LocalUsersAndGroups.Group) {
                                $groupName = $null
                                if (($group.Properties | Test-XmlProperty -Property 'groupName')) {
                                    $groupName = $group.Properties.groupName
                                }
                                $groupSid = $null            
                                if (($group.Properties | Test-XmlProperty -Property 'groupSid')) {
                                    $groupSid = $group.Properties.groupSid
                                }  
                                if (!($group.Properties.Members | Test-XmlProperty -Property 'Member')) {
                                    New-Object -TypeName psobject -Property @{
                                        GpoDisplayName = $gpo.DisplayName
                                        GpoGuid = $gpo.id.guid
                                        Action = $group.Properties.action
                                        RestrictedGroupName = $null
                                        RestrictedGroupSid = $null
                                        MemberName = '<empty>'
                                        MemberSid = '<empty>'
                                        LocalGroupName = $groupName
                                        LocalGroupSid = $groupSid
                                        SettingType = "Preferences:LocalUsersAndGroups"
                                    }
                                }
                                else {            
                                    foreach ($member in $group.Properties.Members.Member) {
                                        $memberName = $null
                                        if (($member | Test-XmlProperty -Property 'name')) {
                                            $memberName = $member.name
                                        }
                                        $memberSid = $null            
                                        if (($member | Test-XmlProperty -Property 'sid')) {
                                            $memberSid = $member.sid
                                        }                                        
                                        New-Object -TypeName psobject -Property @{
                                            GpoDisplayName = $gpo.DisplayName
                                            GpoGuid = $gpo.id.guid
                                            Action = $group.Properties.action
                                            RestrictedGroupName = $null
                                            RestrictedGroupSid = $null
                                            MemberName = "[{0}] {1}" -f $member.action, $memberName
                                            MemberSid = "[{0}] {1}" -f $member.action, $memberSid
                                            LocalGroupName = $groupName
                                            LocalGroupSid = $groupSid
                                            SettingType = "Preferences:LocalUsersAndGroups"
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
        }
    }

    if ($AsCsvFile) {
        & $sb | 
            Export-Csv -NoTypeInformation -Path (".\{0} Get-GpoReport_RestrictedGroups.csv" -f $Domain.Split('.').ToUpper())   
    }
    else {
        & $sb
    }
}

function Search-GpoReport {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,Position=0)]    
        [string]$Domain = (New-Object system.directoryservices.directoryentry).distinguishedname.tolower().replace('dc=','').replace(',','.'),
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=1)]    
        [string]$SearchString
    )

    foreach ($gpo in (Get-GPO -All -Domain $Domain)) {
        Write-Host -ForegroundColor Green ("Searching: '{0}'..." -f $gpo.displayname)

        ($gpo | Get-GPOReport -ReportType Html -Domain $Domain).split("`n") |
            Select-String -CaseSensitive -Pattern $SearchString
    }
}


# This check is not ideal, need to replace with a propper .psd1 manifest
if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory
}
if (-not (Get-Module -Name GroupPolicy)) {
    Import-Module GroupPolicy
}
