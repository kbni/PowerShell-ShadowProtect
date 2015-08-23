<#
    .SYNOPSIS
        Perform multiple operations on multiple servers (and their ShadowProtect backups).
	
    .DESCRIPTION
        Intended to be a one-stop tool for quickly verifying backups are running in the fleet.
        I'd consider this work-in-progress quality, but it's not likely more work will happen. :)

    .PARAMETER Operation
        The type of operation you wish to perform, these are limited to:
            LastBackups -> Returns the last volume backups for each system
            VerifyLastBackups -> Copy-verifies the last volume backups for each system
            MountedImages -> Returns the currently mounted backups for each system
            UnmountImages -> Unmounts all currently mounted backups for each system
        See also: -ScriptBlock

    .PARAMETER ScriptBlock
        Pass a PowerShell ScriptBlock to each server. This will let you do pretty much anything
        you want, even non-ShadowProtect related tasks. Although, if simply running commands
        against multiple servers is your goal, there are far better scripts out there.

    .PARAMETER Options
        Each of the Operations above and the ScriptBlock are passed in $Options, which is
        expected to be a Hashtable. Really only necessary for VerifyLastBackups, where Options
        should contain the following key/values:
            ImagePassword = 'ShadowProtectImagePassword'
            CIFSUsername = 'Additional username to use to connect to the CIFS Share'
            CIFSPassword = 'Additional password to use to connect to the CIFS Share'

    .PARAMETER ServerList
        Array of servers to run script against. This can be an array of one.
        See also: -ComputerName

    .PARAMETER ComputerName
        Specify a single server to run a script against. Instead of -Serverlist

    .PARAMETER Throttle
        How many servers should perform their operation at the same time.
        The default is set to 8. I would restrict this further if all servers were
        hitting the same backup device (and you intended on reading from that device).


    .PARAMETER SimpleLog
        Some of the functions will append a really basic log entry to the SimpleLog
        property of each ShadowProtectBackup-Object. If you just want to dump this and
        don't require further inspection, this is the way to go.

        Currently only recommended for VerifyLastBackups, however.

    .PARAMETER Select
        Select properties of each ShadowProtectBackup-Object.
        Pretty much just calls Format-Table $Select
	
    .EXAMPLE
        .\RemoteOp.ps1 -ServerList (Get-Content .\ServerList.txt) -Operation VerifyLastBackups `
            -Options @{CIFSUsername="KBNI-BNE-NAS2\BackupUser";CIFSPassword="BackupPass";ImagePassword="hunter2"} `
            -Select ImageFile,ComputerName,IsCopyVerified
        Verifies the last backups for each server, and outputs whether or not they have been verified by copying files

    .LINK
        Wiki: http://co.kbni.net/display/PROJ/PowerShell-ShadowProtect
        Github repo: https://github.com/kbni/PowerShell-ShadowProtect
#>
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