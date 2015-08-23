<#
    .SYNOPSIS
        Installs PowerShell-ShadowProtectModule
	
    .DESCRIPTION
        Uses an extremely primative method to install this module on remote servers

    .PARAMETER Credentials
        Credential to use to install module on remote server

    .PARAMETER ServerList
        Array of servers to run script against. This can be an array of one.
        See also: -ComputerName
	
    .PARAMETER InstallFrom
        Path to install from

    .EXAMPLE
        .\RemoteInstall.ps1 -ServerList (Get-Content .\ServerList.txt) -InstallFrom .
        Installs module to each server in ServerList.txt

    .LINK
        Wiki: http://co.kbni.net/display/PROJ/PowerShell-ShadowProtect
        Github repo: https://github.com/kbni/PowerShell-ShadowProtect
#>
Param(
    [Parameter(Mandatory=$true)]
    [String[]]
    $ServerList,

    [Parameter(Mandatory=$false)]
    $InstallFrom,

    [Parameter(Mandatory=$false)]
    $Credentials
)

$DebugPreference = "Continue"
$ErrorPreference = "Stop"

If(!$UseCredential -and $Credential) {
    $UseCredential = $Credential
}

ForEach($server in $serverList) {
    $s = New-PSSession -ComputerName $server -Credential $UseCredential -ErrorAction Stop
    Invoke-Command -InputObject $InstallFrom -Session $s -ScriptBlock {
        $modulePath = [Environment]::GetEnvironmentVariable("PSModulePath").split(';') | Where-Object { (Test-Path $_) -And ($_ -match 'system32') } | Select-Object -First 1
        If($modulePath) {
            New-Item -Type Directory -Path "$modulePath\ShadowProtect" -ErrorAction SilentlyContinue | Out-Null
            Get-ChildItem -Path $input | Copy-Item -Destination "$modulePath\ShadowProtect" -Verbose

            $p = Get-ExecutionPolicy
            Set-ExecutionPolicy ByPass # Of course, our module is not signed. You might want to sign it in you environment.
            Import-Module ShadowProtect -Verbose 
            Set-ExecutionPolicy $p
        }
        Else {
            Write-Error "Unable to install to $env:COMPUTERNAME - can't find PS module path!"
        }
    }
    Remove-PSSession -Session $s | Out-Null
}
