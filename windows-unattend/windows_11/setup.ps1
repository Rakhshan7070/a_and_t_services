# ------------------------------------------------------------------------------
# 0.5 SCRIPT CREDITS & HEADER
# ------------------------------------------------------------------------------
Clear-Host
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "         AUTOMATED SYSTEM DEPLOYMENT & CONFIGURATION           " -ForegroundColor White
Write-Host "         Developed by: Rakhshan Ali                            " -ForegroundColor Green
Write-Host "         Everything is open-source!                            " -ForegroundColor Yellow
Write-Host "         GitHub: https://github.com/Rakhshan7070/a_and_t_services" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

# ==============================================================================
# STAGE 1 - INSTALL SCRIPT
# Handles: Disk, Software, Activation
# Ends by: Writing Stage2_Update.ps1 to disk and registering it in RunOnce
# ==============================================================================

# ------------------------------------------------------------------------------
# 0. SELF-ELEVATION
# ------------------------------------------------------------------------------
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# ------------------------------------------------------------------------------
# 2. DISK PARTITIONING
# ------------------------------------------------------------------------------
$disk = Get-Disk | Where-Object { $_.Number -eq 0 }
$totalSizeGB = [math]::Round($disk.Size / 1GB)
$targetSizeGB = 0

if ($totalSizeGB -ge 150 -and $totalSizeGB -le 260) { 
    $targetSizeGB = 100 
} elseif ($totalSizeGB -ge 350 -and $totalSizeGB -le 550) { 
    $targetSizeGB = 200 
}

if ($targetSizeGB -gt 0 -and !(Get-Volume -FileSystemLabel "Data" -ErrorAction SilentlyContinue)) {
    try {
        Resize-Partition -DriveLetter C -Size ($targetSizeGB * 1GB) -ErrorAction Stop
        $newPart = New-Partition -DiskNumber 0 -UseMaximumSize -AssignDriveLetter
        Format-Volume -DriveLetter $newPart.DriveLetter -FileSystem NTFS -NewFileSystemLabel "Data" -Full:$false -Force -Confirm:$false
    } catch { 
        Write-Host "Disk operation skipped: $_" -ForegroundColor Yellow
    }
}

# ------------------------------------------------------------------------------
# 3. DYNAMIC USB DETECTION
# ------------------------------------------------------------------------------
$usbDrive = Get-WmiObject Win32_Volume | Where-Object { $_.Label -eq "WINDOWS_11" }
if ($usbDrive) {
    $usb = $usbDrive.Name
    Write-Host "[+] Found USB Drive at $usb" -ForegroundColor Green
} else {
    Write-Host "[!] ERROR: USB labeled 'WINDOWS_11' not found!" -ForegroundColor Red
    Pause; exit
}

# ------------------------------------------------------------------------------
# 3.5 COPY OS-SPECIFIC ASSETS TO DESKTOP
# ------------------------------------------------------------------------------
$desktopPath = "$env:PUBLIC\Desktop"
$assetName   = "assets"
$assetSource = Join-Path $usb $assetName
$assetDest   = Join-Path $desktopPath $assetName

if (Test-Path $assetSource) {
    Write-Host "[*] Copying $assetName to Desktop..." -ForegroundColor Cyan
    
    # Perform the copy
    Copy-Item -Path $assetSource -Destination $desktopPath -Recurse -Force
    
    # Confirmation Check
    if (Test-Path $assetDest) {
        Write-Host "[+] VERIFIED: $assetName was successfully copied to the Desktop." -ForegroundColor Green
    } else {
        Write-Host "[!] ERROR: Copy command ran, but $assetName is missing from the Desktop!" -ForegroundColor Red
    }
} else {
    Write-Host "[!] $assetName folder not found on the USB drive. Skipping." -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# 4. INSTALL CHROME
# ------------------------------------------------------------------------------
$chrome = Join-Path $usb "Chrome\install.exe"
if (Test-Path $chrome) {
    Write-Host "[*] Installing Google Chrome..." -ForegroundColor Cyan
    Start-Process $chrome -ArgumentList "/silent", "/install" -Wait
}

# ------------------------------------------------------------------------------
# 5. INSTALL ADOBE READER
# ------------------------------------------------------------------------------
if (Test-Path "$usb\Adobe\install.exe") {
    Write-Host "[*] Installing Adobe Reader silently..." -ForegroundColor Cyan
    Start-Process "$usb\Adobe\install.exe" `
        -ArgumentList "/sAll /rs /msi /norestart /quiet" `
        -Wait
}

# ------------------------------------------------------------------------------
# 6. INSTALL HD SENTINEL
# ------------------------------------------------------------------------------
$hdInstaller = Join-Path $usb "Hdsential\install.exe"
if (Test-Path $hdInstaller) {
    Write-Host "[*] Installing HD Sentinel..." -ForegroundColor Cyan
    $proc = Start-Process $hdInstaller -ArgumentList "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/IACCEPTTHEAGREEMENT" -PassThru
    $proc.WaitForExit(30000)
}

# ------------------------------------------------------------------------------
# 7. COPY & INSTALL OFFICE 2021
# ------------------------------------------------------------------------------
$officeUsbSource   = Join-Path $usb "Office2021"
$officeDesktopDest = Join-Path $desktopPath "Office2021"
$officeInstaller   = Join-Path $officeDesktopDest "install.exe"

$wordPath  = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

if (!(Test-Path $wordPath)) {
    if (Test-Path $officeUsbSource) {
        Write-Host "[*] Copying Office 2021 to Desktop..." -ForegroundColor Cyan
        Copy-Item -Path $officeUsbSource -Destination $desktopPath -Recurse -Force
        
        if (Test-Path $officeInstaller) {
            Write-Host "[*] Installing Microsoft Office 2021 from Desktop..." -ForegroundColor Cyan
            # Launch the installer, but DO NOT forcefully kill the background processes afterwards!
            Start-Process $officeInstaller -ArgumentList "/quiet", "/norestart"
        } else {
            Write-Host "[!] ERROR: Office installer not found on Desktop after copy!" -ForegroundColor Red
        }
    } else {
        Write-Host "[!] ERROR: Office2021 folder not found on USB!" -ForegroundColor Red
    }
} else {
    Write-Host "[+] Office is already installed. Skipping copy and install." -ForegroundColor Green
}

# Office Verification Watchdog (10 min)
Write-Host "[*] Searching for Office files (10 min watchdog)..." -ForegroundColor Yellow
$timer = [System.Diagnostics.Stopwatch]::StartNew()
$found = $false

while ($timer.Elapsed.TotalMinutes -lt 10) {
    if ((Test-Path $wordPath) -or (Test-Path $excelPath)) {
        Write-Host "`n[+] Office files found! Proceeding to activation..." -ForegroundColor Green
        $found = $true
        break
    }
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 10
}
$timer.Stop()

# ------------------------------------------------------------------------------
# 8. OFFICE ACTIVATION (waits 3 minutes after Office files are found)
# ------------------------------------------------------------------------------
if ($found) {
    Write-Host "[+] Office found. Waiting 3 minutes before activation..." -ForegroundColor Yellow
    for ($i = 180; $i -gt 0; $i--) {
        Write-Host "`r    Time remaining: $([math]::Floor($i/60))m $($i % 60)s  " -NoNewline
        Start-Sleep -Seconds 1
    }
    Write-Host "`n[*] Downloading MAS Office Activator..." -ForegroundColor Cyan

    # Check internet before downloading
    $netReady = $false
    for ($i = 1; $i -le 10; $i++) {
        if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $netReady = $true; break }
        Write-Host "    Waiting for internet... ($i/10)" -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }

    if (-not $netReady) {
        Write-Host "[!] No internet. Skipping Office activation." -ForegroundColor Red
    } else {
        try {
            $officeUrl  = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/Separate-Files-Version/Activators/Ohook_Activation_AIO.cmd"
            $officeTmp  = "$env:TEMP\Office_Ohook_Activation.cmd"
            Invoke-WebRequest -Uri $officeUrl -OutFile $officeTmp -UseBasicParsing -ErrorAction Stop
            Write-Host "[+] Launching Office Activator (/Ohook)..." -ForegroundColor Green
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$officeTmp`" /Ohook" -Wait -WindowStyle Normal
            Remove-Item -Path $officeTmp -Force -ErrorAction SilentlyContinue
            Write-Host "[+] Office activation completed." -ForegroundColor Green
        } catch {
            Write-Host "[!] Office activation download failed: $_" -ForegroundColor Red
        }
    }
}

# ------------------------------------------------------------------------------
# 9. WINDOWS ACTIVATION (MAS HWID)
# ------------------------------------------------------------------------------
Write-Host "Verifying Windows Activation..." -ForegroundColor Cyan

# Collect all matching license objects and check if any is activated (LicenseStatus = 1)
$LicenseObjects = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" |
    Where-Object { $_.ApplicationID -eq '55c28233-d71d-4411-8190-76544a4cbd42' }

$IsActivated = ($LicenseObjects | Where-Object { $_.LicenseStatus -eq 1 }).Count -gt 0

if ($IsActivated) {
    Write-Host "[+] Windows is already activated. Skipping." -ForegroundColor Green
} else {
    Write-Host "[*] Windows is not activated. Running MAS HWID..." -ForegroundColor Cyan

    # Make sure internet is available before attempting download
    $netReady = $false
    for ($i = 1; $i -le 10; $i++) {
        if (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet) { $netReady = $true; break }
        Write-Host "    Waiting for internet before activation... ($i/10)" -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }

    if (-not $netReady) {
        Write-Host "[!] No internet connection. Skipping Windows activation." -ForegroundColor Red
    } else {
        try {
            $url     = "https://raw.githubusercontent.com/massgravel/Microsoft-Activation-Scripts/refs/heads/master/MAS/Separate-Files-Version/Activators/HWID_Activation.cmd"
            $tempCmd = "$env:TEMP\HWID_Activation.cmd"
            Invoke-WebRequest -Uri $url -OutFile $tempCmd -UseBasicParsing -ErrorAction Stop
            $MASProcess = 'echo. | "' + $tempCmd + '" /HWID'
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $MASProcess" -Wait -Verb RunAs
            Remove-Item -Path $tempCmd -Force -ErrorAction SilentlyContinue
            Write-Host "[+] MAS activation completed." -ForegroundColor Green
        } catch {
            Write-Host "[!] Activation download failed: $_" -ForegroundColor Red
        }
    }
}

# ==============================================================================
# 10. CREATE STAGE 2 SCRIPT ON DISK
# ==============================================================================
$Stage2Dir  = "C:\Windows\Setup\Scripts"
$Stage2Path = "$Stage2Dir\Stage2_Update.ps1"

if (!(Test-Path $Stage2Dir)) {
    New-Item -ItemType Directory -Path $Stage2Dir -Force | Out-Null
}

# Written as a literal here-string so no variables expand prematurely
$Stage2Content = @'
# ==============================================================================
# STAGE 2 - WINDOWS UPDATE ENGINE
# Persists via RunOnce on every reboot until all updates are complete.
# Destroys itself when done.
# ==============================================================================

# SELF-ELEVATION
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$MyPath = $MyInvocation.MyCommand.Path

# Disable QuickEdit mode so clicking the window never freezes the script
try {
    $RegPath = "HKCU:\Console"
    Set-ItemProperty -Path $RegPath -Name "QuickEdit" -Value 0 -ErrorAction SilentlyContinue
    # Also disable it for the current session via Win32 API
    $code = "using System; using System.Runtime.InteropServices;" +
            "public class ConsoleHelper {" +
            "    [DllImport(`"kernel32.dll`")] public static extern System.IntPtr GetStdHandle(int n);" +
            "    [DllImport(`"kernel32.dll`")] public static extern bool GetConsoleMode(System.IntPtr h, out uint m);" +
            "    [DllImport(`"kernel32.dll`")] public static extern bool SetConsoleMode(System.IntPtr h, uint m);" +
            "}"
    Add-Type -TypeDefinition $code -ErrorAction SilentlyContinue
    $handle = [ConsoleHelper]::GetStdHandle(-10) # STD_INPUT_HANDLE
    $mode   = 0
    [ConsoleHelper]::GetConsoleMode($handle, [ref]$mode) | Out-Null
    $mode = $mode -band (-bnot 0x0040) # Remove ENABLE_QUICK_EDIT_MODE flag
    [ConsoleHelper]::SetConsoleMode($handle, $mode) | Out-Null
} catch { <# Non-critical, continue anyway #> }

# RE-REGISTER IN RUNONCE (so we survive the next reboot if needed)
$RunOnceValue = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$MyPath`""
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "Stage2_Update" -Value $RunOnceValue
Write-Host "[+] RunOnce re-registered. Will resume after reboot if needed." -ForegroundColor Cyan

# ENSURE PSWindowsUpdate IS AVAILABLE
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "[*] Installing PSWindowsUpdate module..." -ForegroundColor Cyan
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
}
Get-ChildItem -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate" -Recurse | Unblock-File
Import-Module PSWindowsUpdate -Force

# RUN WINDOWS UPDATES (with 5 minute timeout)
Write-Host "[*] Checking for Windows Updates (5 min timeout)..." -ForegroundColor Cyan
$updatesFound = $false
try {
    $updateJob = Start-Job -ScriptBlock {
        Import-Module PSWindowsUpdate -Force
        Get-WindowsUpdate -AcceptAll -Install -AutoReboot -Confirm:$false -ErrorAction Stop
    }

    # Wait up to 5 minutes for the job to return results
    $completed = Wait-Job -Job $updateJob -Timeout 300

    if ($completed) {
        $jobResult = Receive-Job -Job $updateJob -ErrorAction SilentlyContinue
        if ($jobResult) { $updatesFound = $true }
    } else {
        # Job timed out after 5 min - updates may be installing in background, let it run
        Write-Host "[*] Update check timed out after 5 min - updates may be running in background." -ForegroundColor Yellow
        $updatesFound = $true
    }
    Remove-Job -Job $updateJob -Force -ErrorAction SilentlyContinue

} catch {
    Write-Host "[!] Update error encountered. Repairing Windows Update service..." -ForegroundColor Yellow
    Stop-Service -Name wuauserv, bits -Force
    Remove-Item -Path "$env:WinDir\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service -Name wuauserv, bits
    Write-Host "[*] Rebooting to retry..." -ForegroundColor Yellow
    Restart-Computer -Force
    exit
}

# If updates were found/installed, let AutoReboot handle the reboot
if ($updatesFound) {
    Write-Host "[*] Updates were found/installed. Waiting for reboot..." -ForegroundColor Yellow
    Start-Sleep -Seconds 60
    Restart-Computer -Force
    exit
}

# ==============================================================================
# NO UPDATES FOUND - CONFIRM FOR 10 MINUTES BEFORE SELF-DESTRUCT
# ==============================================================================
Write-Host "[*] No updates found. Confirming for 10 minutes before finalizing..." -ForegroundColor Cyan

$confirmClean = $true
$checkInterval = 60   # check every 60 seconds
$totalChecks   = 10   # 10 checks = 10 minutes

for ($i = 1; $i -le $totalChecks; $i++) {
    Write-Host "    Confirmation check $i/$totalChecks - waiting 60 seconds..." -ForegroundColor Cyan
    Start-Sleep -Seconds $checkInterval

    # Check for any pending updates
    $pendingUpdates = Get-WindowsUpdate -ErrorAction SilentlyContinue
    $rebootStatus   = Get-WURebootStatus -ErrorAction SilentlyContinue
    $rebootPending  = ($rebootStatus.RebootRequired -eq $true)

    if ($pendingUpdates) {
        Write-Host "    [!] Updates appeared during confirmation. Restarting update cycle..." -ForegroundColor Yellow
        $confirmClean = $false
        Restart-Computer -Force
        exit
    }

    if ($rebootPending) {
        Write-Host "    [!] Reboot required detected. Rebooting..." -ForegroundColor Yellow
        Restart-Computer -Force
        exit
    }

    Write-Host "    [+] Check $i/$totalChecks passed - no updates, no reboot pending." -ForegroundColor Green
}

if (-not $confirmClean) { exit }

# ==============================================================================
# ALL CLEAR - SELF-DESTRUCT
# ==============================================================================
Write-Host "`n[+] 10 minute confirmation passed. System is fully updated." -ForegroundColor Green

# Remove RunOnce entry so we never run again
Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "Stage2_Update" -ErrorAction SilentlyContinue

# Disable Auto-Login
$Winlogon = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
Set-ItemProperty    -Path $Winlogon -Name "AutoAdminLogon" -Value "0"
Remove-ItemProperty -Path $Winlogon -Name "AutoLogonCount" -ErrorAction SilentlyContinue

Write-Host "[+] Auto-login disabled." -ForegroundColor Green

# Show completion message to the user
Add-Type -AssemblyName System.Windows.Forms
$msg = "System Setup Complete!`n`nAll Windows Updates have been installed successfully.`nThis machine is fully configured and ready for use.`n`nDeveloped by: Rakhshan Ali`nEverything is open-source!`nGitHub: https://github.com/Rakhshan7070/a_and_t_services"
[System.Windows.Forms.MessageBox]::Show(
    $msg,
    "IT Deployment - Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

# Delete this script from disk
Remove-Item -Path $MyPath -Force
'@

$Stage2Content | Out-File -FilePath $Stage2Path -Encoding UTF8 -Force
Write-Host "[+] Stage 2 script written to: $Stage2Path" -ForegroundColor Green

# ==============================================================================
# 11. REGISTER STAGE 2 IN RUNONCE & REBOOT
# ==============================================================================
$RegValue = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$Stage2Path`""
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "Stage2_Update" -Value $RegValue

Write-Host "`n[+] Stage 1 complete. Handing off to Stage 2 Update Engine on next boot." -ForegroundColor Green
Write-Host "[*] Rebooting in 5 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 5
Restart-Computer -Force