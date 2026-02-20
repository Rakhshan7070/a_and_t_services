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
        Write-Host "Shrinking C: to $targetSizeGB GB..."
        Resize-Partition -DriveLetter C -Size ($targetSizeGB * 1GB) -ErrorAction Stop
        
        Write-Host "Creating and Formatting Data partition..."
        $newPart = New-Partition -DiskNumber 0 -UseMaximumSize -AssignDriveLetter
        Format-Volume -DriveLetter $newPart.DriveLetter -FileSystem NTFS -NewFileSystemLabel "Data" -Full:$false -Force -Confirm:$false
        
        Write-Host "Disk partitioning successful!" -ForegroundColor Green
    } catch {
        Write-Host "Disk operation failed: $_" -ForegroundColor Yellow
    }
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
    # Attempt update with AutoReboot so it doesn't ask [Y/N]
    Get-WindowsUpdate -AcceptAll -Install -AutoReboot -Confirm:$false -ErrorAction Stop
} catch {
    Write-Host "Encountered a Windows Update error. Attempting to repair and restart..."
    Stop-Service -Name wuauserv, bits -Force
    Remove-Item -Path "$env:WinDir\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv, bits
    # Trigger a reboot to refresh the system state
    Restart-Computer -Force
    exit
}

# FINAL STAGE (Runs only if all updates are finished)
if (-not (Get-WURebootStatus -ErrorAction SilentlyContinue)) {
    Write-Host "Windows is fully updated. Removing auto-resume..."
    Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -ErrorAction SilentlyContinue
}

# WINDOWS UPDATE (With 5-Minute Timeout for End-of-Support systems)
Write-Host "Starting Update Check (Max 5 minute wait for end-of-life systems)..."
$timeout = New-TimeSpan -Minutes 5
$sw = [diagnostics.stopwatch]::StartNew()

while ($sw.Elapsed -lt $timeout) {
    $status = (Get-Service wuauserv).Status
    $installer = Get-Process "TrustedInstaller" -ErrorAction SilentlyContinue
    
    if ($status -ne 'Running' -and !$installer) { break }
    
    Write-Host "Updates/Services active... waiting ($($sw.Elapsed.Seconds)s elapsed)"
    Start-Sleep -Seconds 30
}

# FINAL DESTRUCTION (Runs after updates finish OR timeout hits)
Write-Host "Finalizing system... Disabling Auto-Login and RunOnce."
Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -ErrorAction SilentlyContinue

$Winlogon = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty -Path $Winlogon -Name "AutoAdminLogon" -Value "0"
Remove-ItemProperty -Path $Winlogon -Name "AutoLogonCount" -ErrorAction SilentlyContinue

Write-Host "Setup complete!"
exit
