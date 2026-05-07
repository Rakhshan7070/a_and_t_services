# 🚀 A&T Services — Automated Windows Deployment

> Fully automated Windows 10/11 deployment solution for refurbished and fresh systems.

This project was originally developed for **A&T Services** to simplify and accelerate Windows deployments on refurbished laptops and desktops. Instead of manually installing Windows, partitioning drives, installing software, activating Office, and waiting through endless update cycles, this setup automates the entire workflow from boot to completion.

Everything in this repository is open-source and fully customizable for your own deployment environment.

---

# ✨ Features

This deployment uses a custom `autounattend.xml` configuration combined with PowerShell automation to create a truly zero-touch Windows installation experience.

---

## 🧠 Intelligent Disk Partitioning

Automatically detects total drive size and configures partitions accordingly:

| Total Disk Size | System (`C:`) | Data (`D:`) |
|-----------------|----------------|-------------|
| Small Drives    | 100GB          | Remaining Space |
| Large Drives    | 200GB          | Remaining Space |

- Creates and formats partitions automatically
- Labels the secondary partition as `Data`
- No manual disk setup required

---

## 📦 Dynamic Asset Migration

The deployment automatically detects the Windows installation USB based on its volume label:

- `WINDOWS_10`
- `WINDOWS_11`

Required deployment assets are copied silently to the system during setup.

---

## ⚙️ Silent Software Installation

Installs essential applications completely unattended:

- Google Chrome
- Adobe Acrobat Reader
- Hard Disk Sentinel
- Microsoft Office

No prompts, clicks, or user interaction required.

---

## 🔑 Automatic Activation

Once internet connectivity is detected, the script automatically activates:

### Windows Activation
- MAS HWID Activation

### Office Activation
- MAS Ohook Activation

Completely silent and automated.

---

# 🔄 Two-Stage Update Engine

The deployment process is split into two automated stages.

---

## 🟦 Stage 1 — Initial Setup

Handles:

- System configuration
- Disk partitioning
- Software installation
- Windows activation
- Office activation

After completion, the system automatically reboots into Stage 2.

---

## 🟩 Stage 2 — Persistent Windows Update Engine

Stage 2 creates a persistent update script using Windows `RunOnce` registry entries to survive reboots.

It will continuously:

- Scan for Windows Updates
- Download updates
- Install updates
- Reboot automatically when required

The engine keeps running until:

✅ No pending updates remain  
✅ System remains update-free for 10 minutes

After completion, it automatically:

- Cleans temporary deployment files
- Disables auto-login
- Removes update engine persistence
- Displays a completion notification

---

# 📁 Repository Structure

```text
├── windows_10/
│   ├── autounattend.xml   # Windows 10 unattended setup configuration
│   └── setup.ps1          # Stage 1 deployment script for Windows 10
│
├── windows_11/
│   ├── autounattend.xml   # Windows 11 unattended setup configuration
│   └── setup.ps1          # Stage 1 deployment script for Windows 11
│
└── README.md              # Project documentation
```

---

# 🛠️ Prerequisites & USB Setup

Your Windows installation USB must follow the required structure exactly for automation to work correctly.

---

## 1️⃣ Create Bootable Windows USB

Use either:

- Rufus
- Windows Media Creation Tool

to create a bootable Windows installation drive.

---

## 2️⃣ Set Correct USB Volume Label

### For Windows 10:
```text
WINDOWS_10
```

### For Windows 11:
```text
WINDOWS_11
```

The deployment script relies on these exact labels to locate the USB drive.

---

## 3️⃣ Copy `autounattend.xml`

Place the `autounattend.xml` file directly in the root of the USB drive.

Example:

```text
USB Root/
├── autounattend.xml
```

---

## 4️⃣ Create Required Assets Structure

Create an `assets` folder in the root of the USB and organize installers exactly like this:

```text
USB Root/
├── autounattend.xml
├── assets/
│   ├── Chrome/
│   │   └── install.exe
│   │
│   ├── Adobe/
│   │   └── install.exe
│   │
│   ├── Hdsential/
│   │   └── install.exe
│   │
│   └── Office2019/       # Use Office2021 for Windows 11
│       └── install.exe
```

> ⚠️ Important:  
> The PowerShell scripts expect these exact paths and filenames.  
> If you modify them, you must also update the paths inside `setup.ps1`.

---

# ⚠️ Critical Configuration — Wi-Fi Automation

By default, the `autounattend.xml` contains placeholder Wi-Fi credentials.

If you do not configure them, Windows Setup will pause at the network screen and break the zero-touch installation flow.

---

## 🔧 Configure Wi-Fi Credentials

Open `autounattend.xml` in:

- VS Code
- Notepad++
- Any text editor

Search for:

```text
WIFI_NAME
WIFI_PASSWORD
```

Replace them with your actual Wi-Fi:

```text
WIFI_NAME
WIFI_PASSWORD
```

Save the file back to the USB drive.

---

# 🚀 Usage

## Step 1
Prepare the USB drive using the instructions above.

---

## Step 2
Insert the USB into the target machine and boot from it.

---

## Step 3
Sit back and relax ☕

The deployment will automatically:

- Partition the disk
- Configure Windows
- Bypass OOBE
- Create a local admin account
- Connect to Wi-Fi
- Install software
- Activate Windows & Office
- Install all Windows Updates

---

# 🔁 Deployment Workflow

```text
Boot USB
   ↓
Windows Installation
   ↓
OOBE Bypass
   ↓
Auto Login
   ↓
Stage 1 Deployment
   ↓
Reboot
   ↓
Stage 2 Update Engine
   ↓
Automatic Reboots (if needed)
   ↓
Cleanup & Finalization
   ↓
✅ System Ready
```

---

# 💡 Notes

## 🕒 Office Installation Watchdog

Microsoft Office installs quietly in the background and can take time to appear.

The deployment uses a watchdog timer that:

- Waits up to 10 minutes
- Continuously checks for `WINWORD.EXE`
- Activates Office only after detection

This prevents failed activations.

---

## 🌐 Internet Connection Required

Internet access is required for:

- Windows activation
- Office activation
- Windows Updates

---

## 🧩 Unattend Generator

The `autounattend.xml` files were generated using:

- Schneegans Unattend Generator

---

## 🔓 Activation Credits

Activation functionality is powered by:

- Microsoft Activation Scripts (MAS)

Huge thanks to the MAS developers for their excellent open-source work.

---

# ❤️ Developed By

## Rakhshan Ali

Built for **A&T Services** to make IT deployment faster, cleaner, and fully automated.

---

# 📜 License

This project is open-source.

Feel free to:

- Use
- Modify
- Fork
- Improve

for personal or commercial deployment workflows.

---