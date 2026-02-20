# ============================================================
# SETUP.PS1 — Automated Disk Partition + Windows Update Script
# ============================================================
Start-Transcript -Path "C:\setup_log.txt" -Append

# SHRINK DISK
$disk = Get-Disk | Where-Object { $_.Number -eq 0 }
$totalSizeGB = [math]::Round($disk.Size / 1GB)
$targetSizeGB = 0

# Set target based on disk size
if ($totalSizeGB -ge 150 -and $totalSizeGB -le 260) { $targetSizeGB = 100 }
elseif ($totalSizeGB -ge 350 -and $totalSizeGB -le 550) { $targetSizeGB = 200 }

# Only run if target is valid AND 'Data' drive doesn't exist yet
if ($targetSizeGB -gt 0 -and !(Get-Volume -FileSystemLabel "Data" -ErrorAction SilentlyContinue)) {
    try {
        # Guard: only shrink if C: is currently larger than the target
        $currentSize = (Get-Partition -DriveLetter C).Size
        if ($currentSize -gt ($targetSizeGB * 1GB)) {
            Write-Host "Shrinking C: to $targetSizeGB GB..."
            Resize-Partition -DriveLetter C -Size ($targetSizeGB * 1GB) -ErrorAction Stop
        } else {
            Write-Host "C: is already at or below $targetSizeGB GB, skipping shrink."
        }

        Write-Host "Creating and Formatting Data partition..."
        $newPart = New-Partition -DiskNumber 0 -UseMaximumSize -AssignDriveLetter
        Format-Volume -DriveLetter $newPart.DriveLetter -FileSystem NTFS -NewFileSystemLabel "Data" -Full:$false -Force -Confirm:$false

        # Verify the new partition was created successfully
        $dataVol = Get-Volume -FileSystemLabel "Data" -ErrorAction SilentlyContinue
        if ($dataVol) {
            Write-Host "Disk partitioning successful! Data drive is $($dataVol.DriveLetter):" -ForegroundColor Green
        } else {
            Write-Host "Warning: Data partition may not have been created correctly." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Disk operation failed: $_" -ForegroundColor Yellow
    }
}

# INSTALL OFFICE 2019
# Find the USB drive by its volume label
Write-Host "Looking for Office installer on WINDOWS_10 drive..."
$usbDrive = Get-Volume -FileSystemLabel "WINDOWS_10" -ErrorAction SilentlyContinue
if ($usbDrive) {
    $officePath = "$($usbDrive.DriveLetter):\Office2019\install.exe"
    if (Test-Path $officePath) {
        Write-Host "Found Office installer at $officePath. Installing..." -ForegroundColor Cyan
        try {
            $officeProc = Start-Process -FilePath $officePath -ArgumentList "/quiet", "/norestart" -Wait -PassThru -ErrorAction Stop
            if ($officeProc.ExitCode -eq 0 -or $officeProc.ExitCode -eq 3010) {
                Write-Host "Office 2019 installed successfully! (Exit code: $($officeProc.ExitCode))" -ForegroundColor Green
            } else {
                Write-Host "Office installer finished with exit code $($officeProc.ExitCode) — may need review." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Office installation failed: $_" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Office installer not found at expected path: $officePath" -ForegroundColor Yellow
    }
} else {
    Write-Host "Could not find a drive with label 'WINDOWS_10'. Skipping Office install." -ForegroundColor Yellow
}

# IMMEDIATE PERSISTENCE (Runs every time the PC starts until deleted)
$LocalPath = "$env:SystemDrive\setup.ps1"
$RunOnceValue = "powershell.exe -ExecutionPolicy Bypass -File `"$LocalPath`""
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -Value $RunOnceValue

# SECURITY & MODULE LOAD
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
}
Get-ChildItem -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate" -Recurse | Unblock-File
Import-Module PSWindowsUpdate -Force

# THE UPDATE ENGINE (With Error Recovery)
Write-Host "Checking for Windows Updates..."
try {
    # AutoReboot will reboot the machine automatically if needed; script ends here on reboot
    Get-WindowsUpdate -AcceptAll -Install -AutoReboot -Confirm:$false -ErrorAction Stop
} catch {
    Write-Host "Encountered a Windows Update error. Attempting to repair and restart..." -ForegroundColor Yellow
    Stop-Service -Name wuauserv, bits -Force
    Remove-Item -Path "$env:WinDir\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv, bits
    Restart-Computer -Force
    exit
}

# FINAL STAGE
# If we reach here, Get-WindowsUpdate did NOT trigger a reboot, meaning no reboot is needed.
# Check reboot status to be sure.
$rebootPending = Get-WURebootStatus -ErrorAction SilentlyContinue
if (-not $rebootPending) {
    Write-Host "Windows is fully updated. No reboot pending." -ForegroundColor Green

    # FINAL DESTRUCTION
    Write-Host "Finalizing system... Disabling Auto-Login and removing RunOnce."
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -ErrorAction SilentlyContinue

    $Winlogon = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
    Set-ItemProperty -Path $Winlogon -Name "AutoAdminLogon" -Value "0"
    Remove-ItemProperty -Path $Winlogon -Name "AutoLogonCount" -ErrorAction SilentlyContinue

    Write-Host "Setup complete!" -ForegroundColor Green
} else {
    Write-Host "Reboot still pending after updates. Will resume on next boot." -ForegroundColor Yellow
}

Stop-Transcript
exit
