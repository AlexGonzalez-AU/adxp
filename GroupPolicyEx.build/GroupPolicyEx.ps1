[string[]]$module = @()

Get-Content -Path ..\GroupPolicyEx\Modules\GroupPolicyEx.psd1 |
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
'@    

$module |
    Out-File .\GroupPolicyEx.build\GroupPolicyEx.psm1