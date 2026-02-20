# Shrink Disk size
$disk = Get-Disk | Where-Object { $_.Number -eq 0 }
$totalSizeGB = [math]::Round($disk.Size / 1GB)
$targetSizeGB = 0
if ($totalSizeGB -ge 150 -and $totalSizeGB -le 260) {
    $targetSizeGB = 100
}
elseif ($totalSizeGB -ge 350 -and $totalSizeGB -le 550) {
    $targetSizeGB = 200
}
if ($targetSizeGB -gt 0) {
    try {
        Write-Host "Shrinking C: to $targetSizeGB GB..."
        Resize-Partition -DriveLetter C -Size ($targetSizeGB * 1GB) -ErrorAction Stop
        
        Write-Host "Creating new Data partition..."
        # Create the partition
        $newPart = New-Partition -DiskNumber 0 -UseMaximumSize -AssignDriveLetter
        
        Write-Host "Formatting Data drive ($($newPart.DriveLetter):) silently..."
        # Fix: Added -Force to prevent the "Format" popup window from appearing
        Format-Volume -DriveLetter $newPart.DriveLetter -FileSystem NTFS -NewFileSystemLabel "Data" -Full:$false -Force -Confirm:$false
        
        Write-Host "Disk partitioning successful!" -ForegroundColor Green
    } catch {
        Write-Host "Disk operation failed: $_" -ForegroundColor Red
    }
}


# 1. IMMEDIATE PERSISTENCE (Runs every time the PC starts until deleted)
$LocalPath = "$env:SystemDrive\setup.ps1"
$RunOnceValue = "powershell.exe -ExecutionPolicy Bypass -File `"$LocalPath`""
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -Value $RunOnceValue

# 2. SECURITY & MODULE LOAD
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
}
Get-ChildItem -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate" -Recurse | Unblock-File
Import-Module PSWindowsUpdate -Force

# 3. THE UPDATE ENGINE (With Error Recovery)
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

# 4. FINAL STAGE (Runs only if all updates are finished)
if (-not (Get-WURebootStatus -ErrorAction SilentlyContinue)) {
    Write-Host "Windows is fully updated. Removing auto-resume..."
    Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -ErrorAction SilentlyContinue

    Write-Host "Starting Office 2019 Installation..."
    # (Insert your Office ODT download and install commands here)
    
    Write-Host "Activating Office..."
    # (Insert your activation script here)
}
