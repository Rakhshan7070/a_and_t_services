
# 1. IMMEDIATE PERSISTENCE
# We set this FIRST so if Windows force-reboots, we are already set to resume
$LocalPath = "$env:SystemDrive\setup.ps1"
$RunOnceValue = "powershell.exe -ExecutionPolicy Bypass -File "$LocalPath""
Set-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -Value $RunOnceValue

# 2. FIX SECURITY & LOAD MODULE
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
}
Get-ChildItem -Path "$env:ProgramFiles\WindowsPowerShell\Modules\PSWindowsUpdate" -Recurse | Unblock-File
Import-Module PSWindowsUpdate -Force

# WINDOWS UPDATE
# 3. THE "YES TO ALL" UPDATE COMMAND
# -AutoReboot tells the module to restart the PC automatically without asking [Y/N]
Write-Host "Starting fully automated updates. The system may reboot automatically..."
Get-WindowsUpdate -AcceptAll -Install -AutoReboot -Confirm:$false

# 4. FINAL CHECK & OFFICE INSTALL
# This part only runs if the command above finishes without needing a reboot
if (-not (Get-WURebootStatus -ErrorAction SilentlyContinue)) {
    Write-Host "Updates complete. Proceeding to Office 2019..."
    
    # REMOVE the RunOnce so it doesn't loop forever after finished
    Remove-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce" -Name "AutoSetup" -ErrorAction SilentlyContinue
