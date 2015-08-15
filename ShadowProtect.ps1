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
        [string]$ImageFile,
        [string]$SourcedFrom
    )

    $bu = New-Object PSObject
    $bu | Add-Member -NotePropertyName ImageFile -NotePropertyValue $ImageFile
    $bu | Add-Member -NotePropertyName SourcedFrom -NotePropertyValue $SourcedFrom

    $bu | Add-Member -NotePropertyName IsVerified -NotePropertyValue $False
    $bu | Add-Member -NotePropertyName IsCopyVerified -NotePropertyValue $False
    $bu | Add-Member -NotePropertyName Completed -NotePropertyValue $False
    $bu | Add-Member -NotePropertyName Checksum -NotePropertyValue $False
    $bu | Add-Member -NotePropertyName IsIncremental -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName StartTime -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName EndTime -NotePropertyValue $_.TimeGenerated
    $bu | Add-Member -NotePropertyName LogFile -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName Md5File -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName WriteBufferFile -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName WriteBufferLength -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName IncrementalFile -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName MountPointDir -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName GenerateIncrementalOnDismount -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName KeepWriteBufferOnDismount -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName VolumeDevice -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName VolumeNumber -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName ReadOnly -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName SimpleLog -NotePropertyValue @()
    $bu | Add-Member -NotePropertyName MountTime -NotePropertyValue $Null
    $bu | Add-Member -MemberType NoteProperty -Name EventLogEntry -Value $_
    $bu | Add-Member -MemberType NoteProperty -Name PartitionLength -Value "Unknown"

    
    $bu | Add-Member -NotePropertyName IsMounted -NotePropertyValue $False
    $bu | Add-Member -NotePropertyName VerifyCopyCount -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName VerifyCopyErrors -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName VerifyCopyItems -NotePropertyValue $Null
    $bu | Add-Member -NotePropertyName DriveLetter -NotePropertyValue $Null
    
    $bu | Add-Member -MemberType ScriptMethod -Name GetLog -Value { Get-Item $this.LogFile }
    $bu | Add-Member -MemberType ScriptMethod -Name LogExists -Value { Test-Path $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name GetImage -Value { Get-Item $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name ImageExists -Value { Test-Path $this.ImageFile }
    $bu | Add-Member -MemberType ScriptMethod -Name VerifyImage -Value { $this | Verify-ShadowProtectImage }
    $bu | Add-Member -MemberType ScriptMethod -Name MountImage -Value { $this | Mount-ShadowProtectImage }
    $bu | Add-Member -MemberType ScriptMethod -Name UnmountImage -Value { $this | Unmount-ShadowProtectImage }

    $defaultProperties = @(‘ImageFile', 'IsCopyVerified', 'IsVerified')
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $bu | Add-Member MemberSet PSStandardMembers $PSStandardMembers
    

    $InitFromEventLog = {
        If($this.EventLogEntry) {
            foreach($line in $this.EventLogEntry.Message.split("`n")) {
                if($line -match 'Backup status') { $this.Completed = ( $line -match 'completed') }
                if($line -match 'Log file') { $this.LogFile = $line.SubString($line.IndexOf(":")+2).Trim() }
                if($line -match 'Start time') { $this.StartTime = $line.SubString($line.IndexOf(":")+2).Trim() }
                if($line -match 'Image file') {
                    $this.ImageFile = $line.SubString($line.IndexOf(":")+2).Trim()
                    $this.Md5File = $this.ImageFile -replace '.sp[if]$','.md5'
                    $this.IsIncremental = $line -match '.spf'
                    If($this.Md5File -And (Test-Path $this.Md5File)) {
                        $this.Checksum = (Get-Content $this.Md5File ).ToString().Split(" ")[0]
                    }
                    else {
                        $this.Checksum = "Unknown"
                    }
                }
            }
            $this.SimpleLog += "- Image: $($this.ImageFile)"
            Return $this
        }
    }

    $InitFromMountRes = {
        $this.IsIncremental = $this.ImageFile -contains '.spi'
        If(!$this.EventLogEntry) {
	        $log = Get-EventLog -LogName "Application" -Source "ShadowProtectSvc" | Where-Object { $_.EventID -eq 1120 -or $_.EventID -eq 1121 -or $_.EventID -eq 1122 -And $_.Message.ToLower().IndexOf($this.ImageFile.ToLower()) -gt 0 }
            If($log) {
                $this.EventLogEntry = $log
            }
        }
        Return $this
    }

    $bu | Add-Member -MemberType ScriptMethod -Name InitFromEventLog -Value $InitFromEventLog
    $bu | Add-Member -MemberType ScriptMethod -Name InitFromMountRes -Value $InitFromMountRes
    
    Return $bu
}

Function Get-ShadowProtectBackupHistory {
	$EndDate = (Get-Date).AddDays(1)
	$StartDate = (Get-Date).AddMonths(-12)
	$results = @()
	
	$logs = Get-EventLog -LogName "Application" -Source "ShadowProtectSvc" -Before $EndDate -After $StartDate
    $logs | Where-Object { $_.EventID -eq 1120 -or $_.EventID -eq 1121 -or $_.EventID -eq 1122 } | % {
        $bu_job = New-ShadowProtectBackupObject
        $bu_job.EventLogEntry = $_
        $bu_job.InitFromEventLog() | Out-Null
        $bu_job.InitFromMountRes() | Out-Null
        $results += $bu_job
    }
	Return ( $results | Sort-Object -Property EndTime )
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
            $bu.SimpleLog += "- - $($bu.GetImage().BaseName) mounted" 
        }
        Else {
            $bu.SimpleLog += "- - $($bu.GetImage().BaseName) unable to mount" 
        }
    }

    Return $output
}

Function Unmount-ShadowProtectImage {
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
    ForEach($imageFile in $imageFiles) {
        $bu = New-ShadowProtectBackupObject
        $bu.ImageFile = $imageFile.FullName
        $bu.SourcedFrom = "FileSystem"
        Write-Output $bu
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
        "IncrementalFile" = "Incremental File";
        "MountPointDir" = "Mount Point Directory";
        "WriteBufferLength" = "Data Bytes in Write Buffer File";
        "VolumeDevice" = "Volume Device";
        "ReadOnly" = "Read Only";
        "GenerateIncrementalOnDismount" = "Generate Incremental On Dismount";
        "KeepWriteBufferOnDismount" = "Keep Write Buffer On Dismount";
    }
    
    $current = $Null
    $save = $False
    foreach($line in $mounts) {
        If(!$line -And $save) {
            $current.InitFromMountRes() | Out-Null
            $current.InitFromEventLog() | Out-Null
            Write-Output $current
            $current = $Null
            $save = $False
        }
        If($line) {
            ForEach($key in $properties.KEYS.GetEnumerator()) {
                If($current -eq $Null) {
                    $current = New-ShadowProtectBackupObject
                }
                $matchText = "$($properties[$key]): "
                If($line.IndexOf($matchText) -eq 0) {
                    $val = ($line -replace $matchText,'').Trim()
                    If($val -eq "TRUE") { $val = $True }
                    If($val -eq "FALSE") { $val = $True }
                    If($key -eq "ImageFile") { $val = $val.split('|')[-1] }
                    If($val -eq "") { $val = $Null }
                    If($key -eq "PartitionLength") {
                        $val -match '(\d+) bytes?' | Out-Null
                        $val = [long]($matches[1])
                    }
                    $current.($key) = $val
                    $save = $true
                }
            }
        }
    }
}

Function Verify-ShadowProtectImage {
	$exe = Get-PathOfSPImage

    $output = @()
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
        $output += $bu
    }

    Return $output
}

Function VerifyCopy-ShadowProtectImage {
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
                $bu | Unmount-ShadowProtectImage | Out-Null
            }
        }

        $bu
    }
}

#$res = Get-ShadowProtectBackupHistory | Where-Object { $_.ImageFile -match 'C_VOL' } | Select-Object -last 1 | Verify-ShadowProtectImage | Mount-ShadowProtectImage -Password password | VerifyCopy-ShadowProtectImage
#$res.SimpleLog -join "`r`n"
#$history = Get-ShadowProtectBackupHistory
#Unmount-All-ShadowProtectImages
#Start-Sleep 3
#Get-ShadowProtectBackupHistory | Where-Object { $_.ImageFile -match "C_VOL" } | Select-Object -Last 1 | Mount-ShadowProtectImage | VerifyCopy-ShadowProtectImage
#Unmount-All-ShadowProtectImages