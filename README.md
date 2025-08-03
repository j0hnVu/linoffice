# LinOffice - Microsoft Office launcher for Linux

**LinOffice enables you to run Microsoft Office applications in Linux.** In the background, it runs a virtual machine with Windows and Office installed which is then accessed via FreeRDP. LinOffice's aim is to be a "1-click" setup for Microsoft Office in Linux.

By default, **Microsoft Office 2024** (Home & Retail) is installed, containing Word, Excel, Powerpoint, Onenote and Outlook. If you have a **Microsoft Office 365** subscription you can log in with your Microsoft account and it should upgrade from 2024 to 365.

The project utilises [Winapps](https://github.com/winapps-org/winapps), [Dockur/Windows](https://github.com/dockur/windows), FreeRDP, Podman, and Qemu. You *can* run LinOffice alongside WinApps - all the names and RDP ports have been changed to avoid conflicts between the two.

# Screenshot

![ScreenShot](screenshot.png).
<img src="https://raw.githubusercontent.com/eylenburg/linoffice/refs/heads/main/screenshot2.png" width="394" height="398" />

# Features

- [x] **Automatic non-interactive setup**
- [x] **Run Microsoft Office apps as if they were native Linux apps**
- [x] **Office apps have access to /home folder and Linux clipboard**
- [x] Run Publisher or Access, if manually installed by user
- [x] Quickly access useful tools such as Powershell, Command Prompt, Windows Explorer, Registry Editor, Office Language Settings
- [x] Run a full Windows session via RDP or VNC
- [x] Automatic suspension of the Windows container when inactive, resume when starting an Office app
- [x] Automatic deletion of Office lock files (like `~$file.docx`)
- [x] Force time sync in Windows after Linux host wakes up from sleep, to avoid time drift (time zone is always set to UTC for simplicity)
- [x] Automatic detection of language, date format, thousand and decimal separator, currency symbol, keyboard layout etc
- [x] Toggle settings such as display scaling, auto-suspend on/off,  network on/off (to prevent Windows and Office from "phoning home") 
- [x] Circumvent geo-blocking for the Office installation
- [x] Install updates for Windows and Office
- [x] Tidy Quick Access pane in Windows File Explorer
- [x] Troubleshooting functions e.g. force cleanup of Office lock files, recreate .desktop files for the Office apps, reboot the Windows container, container health check

</details>

**<details><summary>Planned features</summary>**
    
### Planned features

- [ ] Logo to use for the GUI
- [ ] Deliver as Flatpak or AppImage, which would have these benefits:
    - Bundles dependencies such as FreeRDP and Podman-Compose; only Podman would need to be installed on the system already
    - Installation and uninstallation more straight-forward for Linux beginners 
- [ ] APPDATA folder should not be hardcoded (in `setup.sh`, `linoffice.sh`, `mainwindow.py`, `installer.py`, `linoffice.py`, `TimeSync.ps1`, `FirstRunRDP.ps1`, and `RegistryOverride.ps1`) or at least only hardcoded in one of them and then read by the others (like `uninstall.sh` is doing).
- [ ] Flatpak would not allow creatingfiles in the `config` folder inside the app's folder (`compose.yaml`, `linoffice.conf`, `oem/registry/regional-settings.reg`), so it might be better to to move this whole folder out of the app directory and change the file references (perhaps copy to the APPDATA folder)


### Nice to have but lower priority

- [ ] ARM support (should be possible as both Windows and Office have ARM versions)
- [ ] Support for non-core Office apps (e.g. Access, Publisher, Visio)
- [ ] Support for older Office versions (e.g. 2016, 2019, 2021)

</details>

### Comparison with other options to run Microsoft Office in Linux

| |Linoffice|Winapps|Windows VM|Crossover|Wine|
|:----|:----|:----|:----|:----|:----|
|Latest supported Office versions|**2024** (installed by default) and **365** (upgrade after login)|All up to **2024** and **365**|All up to **2024** and **365**|All up to **2016**, **365** may or may not work ([List](https://www.codeweavers.com/compatibility?browse=&app_desc=&company=&rating=&platform=&date_start=&date_end=&name=microsoft+excel&search=app#cxlinux))|All up to **2016** ([List](https://appdb.winehq.org/objectManager.php?sClass=application&iId=11))|
|Office components|All, but installed by default: Excel, Word, Powerpoint, OneNote, Outlook|All: Excel, Word, Powerpoint, OneNote, Outlook, Access, Publisher, Visio, Project|All: Excel, Word, Powerpoint, OneNote, Outlook, Access, Publisher, Visio, Project|Excel, Word, Powerpoint|Excel, Word, Powerpoint|
|Bugginess|🐞 Issues when moving or resizing Office windows or when working with multiple open Office windows|🐞 Issues when moving or resizing Office windows or when working with multiple open Office windows|✅ No bugs, working just as Microsoft intended|🐞🐞 UI bugs, crashes, failing to save files, certain features may not work|🐞🐞🐞 UI bugs, crashes, failing to save files, certain features may not work, installation or activation may fail|
|Cost|Free|Free|Free|$60/€60/£60 for Crossover|Free|
|Activation with MAS|✅Yes|✅Yes|✅Yes|❌No|❌No|
|RAM & CPU use|🔴 Significant (Windows VM)|🔴 Significant (Windows VM)|🔴 Significant (Windows VM)|🟢 Modest|🟢 Modest|
|Integration into Linux|✅ App launchers<br />✅ File associations<br />✅ Access to `/home`<br />✅ Shared clipboard|✅ App launchers<br />✅ File associations<br />✅ Access to `/home`<br />✅ Shared clipboard|❌ App launchers<br />❌ File associations<br />⚠️ Access to `/home`\*<br />⚠️ Shared clipboard\*|✅ App launchers<br />✅ File associations<br />✅ Access to `/home`<br />✅ Shared clipboard|✅ App launchers<br />✅ File associations<br />✅ Access to `/home`<br />✅ Shared clipboard|
|Installation & setup|🟢 Easy:<br />1. Install dependencies<br />2. Run LinOffice installer|🔴 Complicated:<br />1. Install dependencies<br />2. Set up Windows VM<br />3. Install Office in VM<br />4. Create config file<br />5. Run WinApps installer|🟡 Medium:<br />1. Install dependencies<br />2. Set up Windows VM<br />3. Install Office in VM|🟢 Easy:<br />1. Install Crossover<br />2. Download Office installer and open it in Crossover|🔴 Complicated; installation and activation will fail without arcane workarounds\**, no automatic creation of `.desktop` files|

\*requires manual setup, e.g. VirtIO-win and WinFSP or Samba when using Virt-Manager/Qemu, or Virtualbox Guest Additions for Virtualbox

\**PlayOnLinux 4 has install scripts for Office 2016 and below that may or may not work

# Installation

All Linux distributions are supported.

**Hardware requirements:**
- Sufficient resources (the Windows VM is allowed to take up to 4 GB RAM, 64 GB storage and 4 CPU cores if needed). The script will check that you have at least 8 GB RAM and 64 GB free storage before proceeding.
- Virtualization support (using kvm)
- x86_64 CPU (ARM is currently not supported, although it would be possible)
- Fast Internet connection as you will need to download several GB from Micrsoft

## Quickinstall

To download and install LinOffice and all the necessary dependencies, run this in a terminal:

```
curl -sSL https://github.com/eylenburg/linoffice/raw/refs/heads/main/quickstart.sh -o quickstart.sh && chmod +x quickstart.sh && ./quickstart.sh
```

This should automatically install all the required dependencies for most distros and _should_ work with most distros.

>[!NOTE]
>The script will install Podman, Podman-Compose, FreeRDP3, Python3 and PySide6 on your computer. If they are not available in your repo, they will be installed via Flatpak or Pip where possible.

>[!WARNING]
>The quickstart script will fail if you are using an immutable distro (e.g. Fedora Atomic), NixOS, Slackware, Gentoo, or Alpine. In this case you will need to do install the dependencies yourself (see below).

## Manual install

### Dependencies

Dependencies that need to be installed:
- **Podman**
- **Podman-Compose**
- **FreeRDP**  (v3)
- **Python3** (optional, for the GUI)
- **PySide6** (optional, for the GUI)

<details><summary>How to install the dependencies</summary>

- Ubuntu & Debian: `sudo apt install podman podman-compose freerdp3 python3 pyside6` 
  - make sure your repo has FreeRDP v3 is only available from Debian 13 and Ubuntu 24.04 onwards; in older versions you need to use backports or install the Flatpak version of FreeRDP
  - PySide6 is only available from Debian 13 and Ubuntu 24.10 (!) onwards; in older versions you need to use backports or install them via `pip`
- Arch: `sudo pacman -Syu podman podman-compose freerdp python pyside6`
- OpenSUSE: `sudo zypper install podman podman-compose freerdp python3 python3-pyside6`
- Fedora: `sudo dnf install podman podman-compose freerdp python3 python-pyside6`
- Fedora Atomic: `rpm-ostree install podman-compose freerdp python-pyside6`, then reboot (podman and python3 are already preinstalled)
- OpenMandriva: `sudo dnf install podman podman-compose freerdp python pyside6`

Podman-Compose and PySide6 can also be installed [via `pip`](https://pip.pypa.io/en/stable/installation/).
FreeRDP can also be installed as a [Flatpak](https://flathub.org/apps/com.freerdp.FreeRDP), but make sure to give it permission to access the /home folder.

</details>

### Run the installer

First, install the dependencies (see above).

Then:
1. Download this repo (e.g. [release version](https://github.com/eylenburg/linoffice/releases) or [latest git version](https://github.com/eylenburg/linoffice/archive/refs/heads/main.zip))
2. Unzip and save in a convenient folder (e.g. `~/.local/bin/linoffice`)
3. Either run the setup script (`chmod +x setup.sh && ./setup.sh`) or start the GUI which will automatically set things up (`python3 gui/linoffice.py`)

The installation should do everything automatically but will take quite a while. You need to download about 8 GB in total and wait until both Windows and Office are installed. In my experience, on a modern laptop (2023 mid-range AMD Ryzen CPU) and with fast Internet (250 Mbps download), it took about 15 minutes all in (breakdown: 3 minutes Windows download, 8 minutes Windows install, 4 minutes Office download and install).

If the setup succeeds without issues, please [share your system setup](https://github.com/eylenburg/linoffice/issues/15) which will be very helpful in order to know where LinOffice works out of the box. 

<details><summary>Notes on the Windows version</summary>

By default, Windows 11 Pro will be installed. If you want, you can also install Windows 10, which should be a bit snappier. To do that, replace `VERSION: "11"` with `VERSION: "10"` in the `config/compose.yaml.default` file _before_ running `setup.sh`. Microsoft will end mainstream support for Windows 10 in October 2025, but after the installation you can use Microsoft Activation Script (MAS) to extend Windows 10 updates until 2028.

Unfortunately it is not allowed to redistribute Microsoft software, otherwise I would have just prepared a pre-made VM with Office installed, which would cut down the installation time and make this whole project much simpler. At the moment, the script downloads Windows, installs it into a VM, then downloads Office and installs it in the VM, as well as various other tweaks to integrate Office.

</details>

## Updating

- LinOffice: There's a built-in updater in the GUI. Otherwise, you can manually download the newest version and replace the old files with it. Don't delete all of your old files (as there are some files that will be created during the initial setup), just overwrite any existing ones.
    - If you already have a working Windows VM from LinOffice v1.0.7 and below, and you want to update, all the files in `linoffice/config/oem` will need to be manually copied again into `C:\OEM` in the Windows VM, overwriting any existing ones. In the RDP session (`./linoffice.sh windows`) you can access your Linux `/home` folder at `\\tsclient\home`.
- Windows & Office: Should auto-update but there's also an update script in the GUI (or you can run `./linoffice.sh update` from the terminal.)

## Uninstall

You can run the `uninstall.sh` to remove everything or click on the Uninstall button in the "Troubleshooting" menu.

<details><summary>Where are the files saved?</summary>

If you want to manually remove the files:
- The self-contained folder where you have saved the `linoffice.sh` script
- The appdata folder for temporary files is in `~/.local/share/linoffice`
- The `.desktop files` (Excel, Onenote, Outlook, Powerpoint, Word) will be created in `~/.local/share/applications`
- The Podman containers, which include the Windows VM, can be removed with `podman rm -f LinOffice && podman volume rm linoffice_data`

</details>

# Usage

### Starting Office applications

After installation, you should find the launchers for the Office applications in your app menu.

### Opening Office files

You can open files from your file manager with Right-click -> Open with. 

### Using the GUI

You will find a launcher for "LinOffice" in your app menu.

### Using the terminal

If you are using the terminal commands often, you might want to create an alias by running `echo "alias linoffice='/home/user/.local/bin/linoffice/linoffice.sh'" >> ~/.bashrc && source ~/.bashrc` - but make sure to adjust the username and path to where your `linoffice.sh` script is saved. This will enable you to run the LinOffice script from anywhere with the `linoffice` command.

- `linoffice [excel|word|powerpoint|onenote|outlook]`: runs one of the predefined Office applications
- `linoffice manual [msaccess.exe|mspub.exe]`: run Microsoft Access or Microsoft Publisher, if installed (they are not part of the default Office version installed by LinOffice)
- `linoffice manual [explorer.exe|regedit.exe|powershell.exe|cmd.exe]`: runs a specific Windows app in the Windows PATH
- `linoffice manual "C:\Program Files\Microsoft Office\root\Office16\SETLANG.EXE"`: like above, but for any application (here: Microsoft Office Language Preferences tool)
- `linoffice windows`:  shows the whole Windows desktop in an RDP session
- `linoffice update`: runs an update script for Windows in Powershell
- `linoffice reset`: kills all FreeRDP processes, cleans up Office lock files, and reboots the Windows VM
- `linoffice cleanup [--full|--reset]`: cleans up Office lock files (such as ~$file.xlsx) in the home folder and removable media; `--full` cleans all files regardless of creation date, `--reset` resets the last cleanup timestamp

### Office activation 

You will need an Office 2024 license key or Office 365 subscription to use Office. During the first 5 days after installation, you can use Office without activation by clicking on "I have a product key" and then on the "X" of the window where you are supposed to enter your product key.

The **Microsoft Activation Scripts (MAS)** will also work if you have, let's say, trouble with activation - just run `./linoffice.sh manual powershell.exe` from the script's directory to open a Powershell window where you can then paste the command to run MAS. 


# Troubleshooting

### Problems with the setup

If the setup fails, try running it a second time. Sometimes that does the trick.
If your Windows VM installs successfully but Office doesn't seem to be installed, you can trigger a re-installation of Office using `/.setup.sh --installoffice`. You can also manually install Office by accessing your virtual machine through `127.0.0.1:8006` in the browser.
If you have installed Office but it wasn't picked up by the setup script, you can let it check for Office again by running `./setup.sh --firstrun`.
If you need to re-create the .desktop files (app launchers) you can do it by running `./setup.sh --desktop` (or the equivalent button in the GUI)
If something breaks later on you can re-run the requirements check with `./setup.sh --healthcheck` (or the equivalent button in the GUI).

If you can't get the setup to work, please [create a bug report ("setup didn't work")](https://github.com/eylenburg/linoffice/issues) with these information:
- The `windows_install.log` (in `~/.local/share/linoffice`)
- The `setup.log`, `setup_office.log`, and `setup_rdp.log` (if they exist) in `C:\OEM` in the Windows VM
  - You can access the VM through via RDP with the command `xfreerdp /cert:ignore /u:MyWindowsUser /p:MyWindowsPassword /v:127.0.0.1 /port:3388` or alternatively by accessing `127.0.0.1:8006` in the browser (password is `MyWindowsPassword`), although the latter method doesn't have clipboard sharing.
- Your system information (LinOffice version, Linux distribution, desktop environment, Wayland or X11, how did you install podman, podman-compose and freerdp?)

### Window management

In my experience, window management can be wonky, particularly if you're using Wayland instead of X11 or if you're using multiple monitors.

<details><summary>Examples</summary>
    
- Using multiple Office documents/windows can be tricky. For example, opening an Office window might not open until you start it a second time and you may or may not get two windows then. Or, opening a new Office window might close already open ones. Don't panic - your documents are not lost. Just launch the latest Office window again and you should now see both the new one and old one. Sometimes, opening a new windows might also have the quirk that the focus is on the older Office window which is sitting in the background. The solution is to minimize the old one so that the focus is gone..
- Moving and resizing windows does not always work well, particularly on setups with multiple monitors. Most desktop environments use the shortcut "Meta or Alt + Left-Click" for moving a window and and "Meta or Alt + Right-Click" for resizing; this is a very reliable to move around or resize Office windows.
- Dialog boxes (commonly encountered when working with charts in Excel for example) spawn as new, separate windows, but they can sometimes appear behind the main window and at the same time block the main window until you close the dialog box.
- Dialog windows may also have a bad size, e.g. the "edit chart data" window in Excel often cuts off the "OK" at the bottom. The solution, again, is to resize the window using a shortcut like "Meta or Alt + Right-Click" in order to access the "OK" button.

</details>

I believe that these are FreeRDP issues. If it becomes too bad, you can try `./linoffice.sh reset` to kill all FreeRDP processes and reboot the Windows VM - but be aware that you will lose any unsaved Office documents this way.

### Wrong keyboard layout

Theoretically, this should be done automatically by the setup script but it might fall back to the US layout if it doesn't detect your Linux keyboard layout or can't match it to a Microsoft keyboard layout. There are two ways to manually set the keyboard layout:

Option 1: 
- In the GUI, you can choose the keyboard layout in Settings. If you're not using the GUI you can achieve the same thing manually:
    - In the LinOffice folder, open the `config/linoffice.conf` and find the row saying `RDP_KBD=""`. 
    - Check [this list of keyboard layouts](https://github.com/eylenburg/linoffice/blob/main/config/languages.csv) to find the numeric code for your keyboard layout in the first code. 
    - Edit the line in the config file like these examples: `RDP_KBD="/kbd:layout:0x0809"` for the UK keyboard or `RDP_KBD="/kbd:layout:0x0407"` for the German keyboard

Option 2 (if the above doesn't work):
- Access the Windows VM, either via RDP (`./linoffice.sh windows`) or VNC (`127.0.0.1:8006` in the browser, password is `MyWindowsPassword`)
- Open the Windows Settings app and set your keyboard layout in there.
- Open the Command Prompt (cmd.exe) and enter: `REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Keyboard Layout" /v IgnoreRemoteKeyboardLayout /t REG_DWORD /d 1` (this tells Windows to use its own keyboard layout and not whatever the RDP client uses)

### Orphaned lock files

There is a strange issue that Microsoft Office will not clean up the lock files in the Linux /home folder. If you open, say, "Book1.xlsx" then Excel will create a file called "~$Book1.xlsx" which is just a few bytes in size and serves the purpose of "locking" this file so other users can't edit it at the same time. Normally these files should be deleted when you close the file, but this doesn't happen for whatever reason. The /home folder is mounted by FreeRDP and appears in Windows as a network drive accessed via RDP Drive Redirection (RDPDR).

LinOffice searches and deletes these lock files when the last Office process is closed. If this fails for any reason you can manually delete all lock files by running `./linoffice.sh cleanup --full` or by using the button in the "Troubleshooting" window.

<details><summary>How to hide these lock files in KDE's Dolphin file manager</summary>

1. Go to KDE `System Settings` -> `Default Applications` -> `File Associations`, then search for the mime type corresponding to .xlsx (in this case it's called `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`). Select the existing filename pattern (`*.xlsx`) and click `- Remove` and then click `+ Add` and enter `[!~][!$]*.xlsx`. Do the same for docx and pptx, and, if you use them, odt, ods, odp, docm, xlsm and pptm. 

(By default files starting with `~$` have the mime type `application/x-trash`. By making the above change, a file like "~$Book1.xlsx" will be seen as a trash file rather than a spreadsheet.)

2. Open Dolphin, go to `Configure Dolphin` -> `View` and check `[x] Also hide backup files when hiding files`. 

("Backup files" in this case actually refers to all files with the `application/x-trash` mime type.)

</details>

# Legal information

This project is licensed under the GNU AGPL 3. 

The main script (`linoffice.sh`) is forked from [Winapps](https://github.com/winapps-org/winapps), AGPL license.

The Windows VM is set up using the [Dockur/Windows](https://github.com/dockur/windows) OCI container, MIT license.

Windows and Office are directly downloaded from Microsoft. This project contains only open-source code and does not distribute any copyrighted material. Any product keys found in the code are just generic placeholders provided by Microsoft for trial purposes. You will need to provide your own product keys to activate Windows and Office.
