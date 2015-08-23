Param(
    [Parameter(Mandatory=$false)]
    [String]
    $Operation,

    [Parameter(Mandatory=$false)]
    $ScriptBlock,

    [Parameter(Mandatory=$false)]
    [String[]]
    $ServerList,

    [Parameter(Mandatory=$false)]
    [String]
    $ComputerName,

    [Parameter(Mandatory=$false)]
    [Int]
    $Throttle = 8,

    [Parameter(Mandatory=$false)]
    $Credential,
    
    [Parameter(Mandatory=$false)]
    [String[]]
    $Select,

    [Parameter(Mandatory=$false)]
    [Hashtable]
    $Options,

    [Switch]
    $SimpleLog
)

$Operations = @{
    LastBackups = {
        $backups = @()
        $b = Get-ShadowProtectBackupHistory -Range Month -ImageNotMatches 'System Reserved'
        ForEach ($Letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray()) {
            $b | Where-Object { $_.ImageFile -match "${Letter}_VOL" } | Select -Last 1 | ForEach-Object { $backups += $_ }
        }
        $backups
    }
    VerifyLastBackups = {
        Param($Options)

        Import-Module ShadowProtect -Force
        . "$((Get-Module -ListAvailable ShadowProtect).ModuleBase)\Invoke-TokenManipulation.ps1"

        $VerbosePreference = "Continue"

        net use * /del /y | out-null

        $backups = @()
        $b = Get-ShadowProtectBackupHistory -Range Month -ImageNotMatches 'System Reserved'
        ForEach ($Letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.ToCharArray()) {
            $b | Where-Object { $_.ImageFile -match "${Letter}_VOL" } | Select -Last 1 | ForEach-Object { $backups += $_ }
        }

        If($backups) {
            Use-ShadowProtectMountDriver

            $networkShareMount = $Null
            $networkShareMountRes = $False

            Try {
                if (($backups | Select-Object -Last 1).ImageFile -match '^(\\\\[^\\]+\\[^\\]+)') {
                    $networkShareMount = Get-AvailableDriveLetter
                    $uncPath = $matches[0]
                    #$net = new-object -ComObject WScript.Network
                    Write-Verbose "Mounting ${uncPath} to ${networkShareMount}"
                    net use ${networkShareMount}: $uncPath /user:$($Options['CIFSUsername']) $($Options['CifsPassword']) | Out-Null
                    If($LastExitCode -gt 0) {
                        $backups | ForEach-Object { $_.SimpleLog += "- - Unable to access ${uncPath}" }
                        $networkShareMountRes = $False
                    }
                    Else {
                        New-PSDrive -Name $networkShareMount -PSProvider FileSystem -Root "${networkShareMount}:" | Out-Null
                        $networkShareMountRes = $True
                    }
                }
            }
            Catch {
                $error[0]
                Write-Error "Unable to map ${uncPath} to ${networkShareMount}:" -ErrorAction Continue
            }

            If( -Not $networkShareMount -Or $networkShareMountRes ) {
                try {
                    ForEach($bu in $backups) {
                        Write-Verbose -Verbose "Mounting image.."
                        $bu | Mount-ShadowProtectImage -Password password | Out-Null
                        Write-Verbose -Verbose "Mounted $($bu.ImageFile) to $($bu.DriveLetter)"
                        Invoke-TokenManipulation -ImpersonateUser -User "NT AUTHORITY\System"
                        $bu | Use-VerifyCopyShadowProtectImage | Out-Null
                        Invoke-TokenManipulation -RevToSelf
                        $bu | Dismount-ShadowProtectImage | Out-Null
                    }
                }
                catch {
                    $error | Select -First 2 | fl
                }

                If($networkShareMount) {
                    Write-Verbose "Dismounting ${networkShareMount}"
                    net use "${networkShareMount}:" /del /y | Out-Null
                }
            }
        }

        $backups
    }
    MountedImages = {
        Get-ShadowProtectImageMounts
    }
    UnmountImages = {
        Get-ShadowProtectImageMounts | Dismount-ShadowProtectImage
    }
}

$DebugPreference = "Continue"
$ErrorActionPreference = "Stop"
$VerbosePreference = "Continue"

$UseBlock = $Null
$sessions = @{}
$jobs = @{}
$results = @{}
$all_results = @()

If ($ComputerName) {
    $ServerList = @($ComputerName)
}

If ($Credential) {
    $UseCredential = $Credential
}

If ($Operation) {
    $UseBlock = $Operations[$Operation]
}

If ($ScriptBlock) {
    $UseBlock = $ScriptBlock
}

If (!$UseBlock) {
    Write-Error "You must specify either an -Operation or -ScriptBlock"
}

ForEach ($serverName in $serverList) {
    Write-Verbose "Connecting to ${serverName} as $($UseCredential.UserName)"
    $sessions[$serverName] = New-PSSession -ComputerName $serverName -Credential $UseCredential -Verbose
}

ForEach ($serverName in $serverList) {
    $RunningJobs = @(Get-Job | Where-Object { $_.State -eq 'Running' })
    If ($RunningJobs.Count -eq $Throttle) {
        Write-Verbose "Hit throttle: $Throttle"
        $RunningJobs | Wait-Job -Any | Out-Null
    }
    $jobs[$serverName] = Invoke-Command -AsJob -Session $sessions[$serverName] -ArgumentList $Options -ScriptBlock $UseBlock
}

If ( @(Get-Job | Where-Object { $_.State -eq 'Running' }) ) {
    Write-Verbose "Waiting for jobs to complete.."
    Get-Job | Wait-Job | Out-Null
}

ForEach($serverName in $serverList) {
    $results[$serverName] = Receive-Job $jobs[$serverName]
    $results[$serverName] | ForEach-Object { $all_results += $_ }
}

$sessions.Values | Remove-PSSession | Out-Null

If($Select) {
    $all_results | Format-Table $Select
}
ElseIf($SimpleLog) {
    $lines = @()
    ForEach($serverName in $serverList) {
        If($serverList.Count -gt 1) {
            $lines += "- ${serverName}"
        }
        $results[$serverName] | ForEach-Object {
            ForEach($line in $_.SimpleLog) {
                $lines += $line
                If($serverList.Count -gt 1) {
                    $lines[-1] = "- " + $lines[-1]
                }
            }
        }
    }

    $lines
}
Else {
    $all_results
}