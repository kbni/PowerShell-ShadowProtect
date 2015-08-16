$ErrorActionPreference = "Stop"

Function Get-AvailableDriveLetter {
    $driveList = [System.IO.DriveInfo]::getdrives().Name -replace ":.",''
    "HIKLMNOPQRSTUVWXYZ".ToCharArray() | Where-Object { $driveList -notcontains $_ } | Select-Object -Last 1
}

Function Get-PathOfSPIMount {
	$try_executables = @(
		"C:\Program Files (x86)\StorageCraft\ShadowProtect\mount.exe"
		"C:\Program Files\StorageCraft\ShadowProtect\mount.exe"
		"C:\Program Files (x86)\ShadowProtect\mount.exe"
		"C:\Program File\ShadowProtect\mount.exe"
	)
    $exe = $try_executables | Where-Object { Test-Path $_ } | Select-Object -Last 1
    If($exe) {
        Return $exe
    }
    Write-Error "Unable to locate ShadowProtect's mount.exe"
}

Function Get-PathOfSPImage {
    $try_executables = @(
		"C:\Program Files (x86)\StorageCraft\ShadowProtect\image.exe"
		"C:\Program Files\StorageCraft\ShadowProtect\image.exe"
		"C:\Program Files (x86)\ShadowProtect\image.exe"
		"C:\Program File\ShadowProtect\image.exe"
	)
    $exe = $try_executables | Where-Object { Test-Path $_ } | Select-Object -Last 1
    If($exe) {
        Return $exe
    }
    Write-Error "Unable to locate ShadowProtect's image.exe"
}

Function New-ShadowProtectBackupObject {
    Param(
        [Parameter(Mandatory=$false)]$EventLogEntry,
        [Parameter(Mandatory=$false)]$MountData,
        [Parameter(Mandatory=$false)]$ImageFile,
        [Parameter(Mandatory=$false)]$UseDate
    )

    $bu = New-Object PSObject -Property @{
        ImageFile = $ImageFile
        Date = $Null
        Completed = $False
        ComputerName = $Null

        IsVerified = $False
        IsCopyVerified = $False
        IsIncremental = $Null
        StartTime = $Null
        EndTime = $_.TimeGenerated
        LogFile = $Null
        Md5File = $Null
        WriteBufferFile = $Null
        WriteBufferLength = $Null
        IncrementalFile = $Null
        MountPointDir = $Null
        GenerateIncrementalOnDismount = $Null
        KeepWriteBufferOnDismount = $Null
        VolumeDevice = $Null
        VolumeNumber = $Null
        ReadOnly = $Null
        MountTime = $Null
        EventLogEntry = $EventLogEntry
        PartitionLength = $Null
        IsMounted = $False
        VerifyCopyCount = $Null
        VerifyCopyErrors = $Null
        VerifyCopyItems = $Null
        DriveLetter = $Null

        SimpleLog = @()
    }

    $bu.PSObject.TypeNames.Insert(0,'ShadowProtect.BackupInformation')
    Update-TypeData -TypeName ShadowProtect.BackupInformation -DefaultDisplayPropertySet ImageFile,Date,Completed -ErrorAction SilentlyContinue
    
    $bu | Add-Member -MemberType ScriptMethod -Name GetLog -Value { Get-Item $this.LogFile }
    $bu | Add-Member -MemberType ScriptMethod -Name LogExists -Value { Test-Path $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name GetImage -Value { Get-Item $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name ImageExists -Value { Test-Path $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name VerifyImage -Value { $this | Verify-ShadowProtectImage }
    $bu | Add-Member -MemberType ScriptMethod -Name MountImage -Value { $this | Mount-ShadowProtectImage }
    $bu | Add-Member -MemberType ScriptMethod -Name DismountImage -Value { $this | Dismount-ShadowProtectImage }  
    $bu | Add-Member -MemberType ScriptMethod -Name ParseEventLog -Value {
        $this.ComputerName = $this.EventLogEntry.MachineName
        $from_log = $this.EventLogEntry.Message.Split("`n") | Parse-Multiline -DefaultProperties @('Code','StartTime','LogFile','Message','ImageFile','BackupStatus')
        $this.Completed = $from_log.BackupStatus -match 'complete'
        $this.LogFile = $from_log.LogFile
        $this.StartTime = $from_log.StartTime
        $this.ImageFile = $from_log.ImageFile
        $this.Date = $this.StartTime
    }

    If($bu.EventLogEntry) {
        $bu.ParseEventLog()
    }

    If($MountData) {
        $bu.DriveLetter = $MountData.DriveLetter
    }

    If($UseDate) {
        $bu.Date = $UseDate
    }

    If($bu.ImageFile) {
        If($bu.Complete) {
            $bu.Md5File = $bu.ImageFile -replace '.sp[if]$','.md5'
            $bu.IsIncremental = $bu -match '.spf'
        }
        $bu.SimpleLog += "- $($bu.ImageFile.split('\')[-1])"
    }

    Else {
        Write-Error "Unable to gleen ImageFile from input"
    }

    Write-Output $bu
}

Function Get-ShadowProtectBackupHistory {
    Param(
        [Parameter(Mandatory=$false)][DateTime]$EndDate,
        [Parameter(Mandatory=$false)][DateTime]$StartDate,
        [Parameter(Mandatory=$false)][String]$ComputerName,
        [Parameter(Mandatory=$false)][Switch]$CheckMounted
    )

    $EventLogParams = @{
        Source = 'ShadowProtectSvc'
        LogName = 'Application'
    }
    
    If ($EndDate) { $EventLogParams['Before'] = $EndDate }
    If ($StartDate) { $EventLogParams['After'] = $StartDate }
    If ($ComputerName) { $EventLogParams['ComputerName'] = $ComputerName }
	
    Get-EventLog @EventLogParams| Where-Object { $_.EventID -in 1120..1122 } | Sort-Object -Property TimeGenerated | ForEach-Object {
        $bu = New-ShadowProtectBackupObject -EventLogEntry $_
        Write-Output $bu
    }
}

Function Mount-ShadowProtectImage {
    Param(
        [string]$Password
    )

    If($Password) { $Password = "p=$Password" }

    $output = @()
    $exe = Get-PathOfSPIMount
    $DriveTimeout = 300

    ForEach($bu in $input) {
        $output += $bu

        $next_drive = Get-AvailableDriveLetter

        if($bu.IsIncremental) {
            $full_image = $bu.ImageFile -replace '-i\d+.spi$','.spf'
            &$exe s $full_image i=$full_image $bu.ImageFile d=$next_drive $Password
        }
        else {
            &$exe s $bu.ImageFile d=$next_drive $Password
        }
        $mount_res = $LastExitCode

        If ($mount_res -eq 0) {
            Write-Debug "Mounting $($bu.GetImage().Name) to $next_drive"
        }
        Else {
            Write-Error "Unable to mount $($bu.GetImage().Name) to $next_drive"
        }

        $sw = [diagnostics.stopwatch]::StartNew()
        $foundDrive = $Null
        While (($sw.elapsed.seconds -lt $DriveTimeout) -And (-Not $foundDrive)) {
            $foundDrive = [System.IO.DriveInfo]::getdrives() | Where-Object { $_.Name -eq "$next_drive`:\" } | Select-Object -First 1
            Start-Sleep 1
        }

        If (($mount_res -eq 0) -And $foundDrive) {
            Write-Debug "Successfully mounted $($bu.GetImage().Name) to $next_drive"
            $bu.DriveLetter = "$($next_drive):"
            $bu.IsMounted = $True
        }
    }

    Return $output
}

Function Dismount-ShadowProtectImage {
	$exe = Get-PathOfSPIMount
    ForEach($bu in $input) {
        Write-Debug "Calling $exe d $($bu.DriveLetter)"
        &$exe d $bu.DriveLetter
        If ($LASTEXITCODE -eq 0) {
            $bu.IsMounted = $False
            $bu.SimpleLog += "- - $($bu.GetImage().BaseName) dismounted"
        }
        Else {
            Write-Error "Command ($exe d $($bu.DriveLetter)) returned $LASTEXITCODE"
        }

        Write-Output $bu
    }
}

Function Get-ShadowProtectImageFiles {
    Param(
        [ValidateScript({Test-Path $_})][String]$Path
    )
    $imageFiles = Get-ChildItem -Path $Path -Recurse | Where-Object { $_.Extension -eq '.spf' -Or $_.Extension -eq '.spi' } | Sort-Object -Property LastWriteTime
    ForEach($ImageFile in $imageFiles) {
        New-ShadowProtectBackupObject -ImageFile $ImageFile.FullName -UseDate $ImageFile.LastWriteTime
    }
}

Function Get-ShadowProtectImageMounts {
    Param([switch]$Raw)

    $temp_file = [System.IO.Path]::GetTempFileName()
	$exe = Get-PathOfSPIMount
    &$exe e | Out-File -Encoding Ascii $temp_file
    $mounts = Get-Content $temp_file
    Remove-Item $temp_file

    If($Raw) {
        $mounts
        Return
    }

    $properties = @{
        "VolumeNumber" = "Volume Number";
        "DriveLetter" = "Drive Letter";
        "PartitionLength" = "Partition Length";
        "ImageFile" = "Image File";
        "WriteBufferFile" = "Write Buffer File";
        "MountPointDir" = "Mount Point Directory";
        "WriteBufferLength" = "Data Bytes in Write Buffer File";
        "VolumeDevice" = "Volume Device";
        "ReadOnly" = "Read Only";
        "GenerateIncrementalOnDismount" = "Generate Incremental On Dismount";
        "KeepWriteBufferOnDismount" = "Keep Write Buffer On Dismount";
    }
    
    $current = $Null
    $save = $False

    $mounts | Parse-Multiline | ForEach-Object {
        New-ShadowProtectBackupObject -ImageFile $_.ImageFile.split('|')[-1] -MountData $_
    }
}

Function Use-VerifyShadowProtectImage {
	$exe = Get-PathOfSPImage

    ForEach($bu in $input) {
        Write-Debug ("Verifying against {0}" -f $bu.Md5File)
        $bu.IsVerified = $False
        &cmd.exe /U /C $exe v $bu.Md5File | out-file -Encoding ascii temp.txt
        $contents = (Get-Content -Encoding unicode temp.txt) -replace "[^\u0020-\u007E]",""
        remove-item -force temp.txt
        foreach($line in $contents) {
            If($line -match "SUCCESS") {
                $bu.IsVerified = $True
                $bu.SimpleLog += "- - Verified checksum matches ShadowProtect"
            }
        }
        Write-Output $bu
    }
}

Function Use-VerifyCopyShadowProtectImage {
    Param(
        $MinLength = 1024*512, # 512KiB
        $MaxLength = 1024*1024*50, # 50MiB
        $MaxDirs = 5,
        $MaxFiles = 5,
        $TempPath = $env:TEMP
    )

    Write-Debug "Looking for files between $MinLength and $MaxLength, will copy to $TempPath"

    foreach($bu in $input) {
        $ErrorCount = 0
        $CopyCount = 0

        $DismountAfter = $False
        If($bu.IsMounted -ne $True) {
            $bu | Mount-ShadowProtectImage | Out-Null
            $DismountAfter = $True
        }

        If($bu.IsMounted -ne $False) {
            $prev_loc = Get-Location
            New-PSDrive -Name $bu.DriveLetter[0] -PSProvider FileSystem -Root "$($bu.DriveLetter)\"
            Set-Location -Path $bu.DriveLetter -ErrorAction Stop
            $parents = Get-ChildItem $bu.DriveLetter -Filter *.* | Where-Object { $_.Name -notmatch '(Windows|PerfLogs)' } | Get-Random -Count $MaxDirs | Select -First $MaxDirs

            $items = @()
            Foreach($parent in $parents) {
                $test_conds = { ( ! $_.PSIsContainer ) -And ($_.Length -gt $MinLength) -And ($_.Length -lt $MaxLength) }
                Get-ChildItem $parent -Recurse | Where-Object $test_conds | Get-Random -Count $MaxFiles | Select -First $MaxFiles | %{
                    $items += $_
                }
            }

            $bu.SimpleLog += "- - Located $($items.Count) files to copy"

            Foreach($TestFile in $items) {
                $CopyToFile = "$TempPath\$($TestFile.Name)"
                $CopyCount += 1
                try {
                    Write-Debug "Copying $TestFile to $CopyToFile"
                    Copy-Item $TestFile.FullName -Destination $CopyToFile -ErrorAction Stop
                    Remove-Item $CopyToFile -ErrorAction SilentlyContinue | Out-Null
                    $bu.SimpleLog += "- - - Successful Copy: $($TestFile.Name)"
                }
                catch {
                    Remove-Item $CopyToFile -ErrorAction SilentlyContinue | Out-Null
                    $ErrorCount += 1
                    $bu.SimpleLog += "- - - Unsuccessful Copy: $($TestFile.Name)"
                }
            }

            Set-Location $prev_loc 
        
            $bu.VerifyCopyErrors = $ErrorCount
            $bu.VerifyCopyCount = $CopyCount
            $bu.VerifyCopyItems = $items
            $bu.IsCopyVerified = ( ($ErrorCount -eq 0) -And ($CopyCount -gt 0) )
            If($bu.IsCopyVerified) {
                $bu.SimpleLog += "- - Successful test-copy"
            }
            Else {
                $bu.SimpleLog += "- - UNSUCCESSFUL Test-copy"
            }

            If($DismountAfter -eq $True) {
                $bu | Dismount-ShadowProtectImage | Out-Null
            }
        }

        Write-Output $bu
    }
}

Function Parse-Multiline {
    Param(
        $DefaultProperties = @()
    )

    $rows = @()
    $row = $Null
    $properties = $DefaultProperties

    ForEach($line in $input) {
        $idx = $line.IndexOf(':')
        If($idx -gt 0) {
            $left = (Get-Culture).TextInfo.ToTitleCase($line.SubString(0, $idx)) -replace ' ',''
            $right = $line.SubString($idx+1).Trim()
            If($row -eq $Null) { $row = @{} }
            If($left -notin $properties) { $properties += $left }
            $row[$left] = $right
        }
        ElseIf($row.Keys.Count -gt 0) {
            $rows += $row
            $row = $Null
        }
    }

    $rows | %{
        $NewObject = New-Object PSObject
        ForEach($p in $properties) {
            $NewObject | Add-Member -NotePropertyName $p -NotePropertyValue $_[$p]
        }

        Write-Output $NewObject
    }
}

Export-ModuleMember -Function Get-ShadowProtectBackupHistory
Export-ModuleMember -Function Get-ShadowProtectImageFiles
Export-ModuleMember -Function Get-ShadowProtectImageMounts
Export-ModuleMember -Function Mount-ShadowProtectImage
Export-ModuleMember -Function Dismount-ShadowProtectImage
Export-ModuleMember -Function Use-VerifyShadowProtectImage
Export-ModuleMember -Function Use-VerifyCopyShadowProtectImage