[string[]]$module = @()

Get-Content -Path ..\ActiveDirectoryEx\Modules\ActiveDirectoryEx.psd1 |
    Select-String  -SimpleMatch '..\CmdLets\' |
    Select-Object -ExpandProperty Line |
    ForEach-Object {
        $module += Get-Content -Path $_.Trim().Trim(",").Trim("'").ToLower().Replace('..\cmdlets\','..\ActiveDirectoryEx\CmdLets\')
        $module += ""
    }

$module |
    Out-File .\ActiveDirectoryEx.build\ActiveDirectoryEx.psm1