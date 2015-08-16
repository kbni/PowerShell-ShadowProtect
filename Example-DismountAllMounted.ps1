$env:PSModulePath += ";."
Import-Module -Force ShadowProtect

$DebugPreference = "Continue"

Get-ShadowProtectImageMounts | Dismount-ShadowProtectImage | Select ImageFile,IsMounted