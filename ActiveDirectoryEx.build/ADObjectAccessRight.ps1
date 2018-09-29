[string[]]$module = @()

Get-Content -Path ..\ActiveDirectoryEx\Modules\ADObjectAccessRight.psd1 |
    Select-String  -SimpleMatch '..\CmdLets\' |
    Select-Object -ExpandProperty Line |
    ForEach-Object {
        $module += Get-Content -Path $_.Trim().Trim(",").Trim("'").ToLower().Replace('..\cmdlets\','..\ActiveDirectoryEx\CmdLets\')
        $module += ""
    }

$module += @'

Get-ADRightsObjectGuids
Get-ADSchemaObjectGuids

Export-ModuleMember `
    -Function @(
        'Get-ADDirectoryEntry',
        'Add-ADObjectAccessRight',
        'Remove-ADObjectAccessRight',
        'Remove-ADObjectAccessRightHelper',
        'Get-ADObjectAccessRight',
        'Set-ADObjectAccessRight'
    ) `
    -Variable @(
        '_ADRightsObjectGuids',
        '_ADRightsObjectNames',
        '_ADSchemaObjectGuids',
        '_ADSchemaObjectNames'   
    )
'@

$module |
    Out-File .\ActiveDirectoryEx.build\ADObjectAccessRight.psm1