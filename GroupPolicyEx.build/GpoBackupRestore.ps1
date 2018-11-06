[string[]]$module = @()

Get-Content -Path ..\GroupPolicyEx\Modules\GpoBackupRestore.psd1 |
    Select-String  -SimpleMatch '..\CmdLets\' |
    Select-Object -ExpandProperty Line |
    ForEach-Object {
        $module += Get-Content -Path $_.Trim().Trim(",").Trim("'").ToLower().Replace('..\cmdlets\','..\GroupPolicyEx\CmdLets\')
        $module += ""
    }

$module += @'

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
'@    

$module |
    Out-File .\GroupPolicyEx.build\GpoBackupRestore.psm1