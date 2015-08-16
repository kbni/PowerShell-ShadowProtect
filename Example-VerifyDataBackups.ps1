$env:PSModulePath += ";."
Import-Module -Force ShadowProtect

$DebugPreference = "Continue"

$servers = @('KBNI-BNE-DC1', 'KBNI-BNE-TS1')
$lasts = @()

ForEach($serverName in $servers) {
    Get-ShadowProtectBackupHistory -ComputerName $serverName | Where-Object { $_.ImageFile -match 'C_VOL' } | Select-Object -Last 1 | ForEach-Object { $lasts += $_ }
}

$lasts | Mount-ShadowProtectImage -Password password | Use-VerifyShadowProtectImage | Use-VerifyCopyShadowProtectImage | Dismount-ShadowProtectImage
$lasts | Select ImageFile,ComputerName,IsVerified,IsCopyVerified
