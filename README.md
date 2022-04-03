# LE Shortcut Creator
Create shortcuts for games/applications that will run them through Locale Emulator.

Written in PowerShell (don't let the \*.cmd extention fool you).  
Compatible with Windows 7, Windows 8.1, Windows 10, and Windows 11.

<p align="center"><img src="screenshot-win7.png?raw=true" alt="Windows 7" width="49%" /><img src="screenshot-win81.png?raw=true" alt="Windows 8.1" width="49%" /><img src="screenshot-win10.png?raw=true" alt="Windows 10" width="49%" /><img src="screenshot-win11.png?raw=true" alt="Windows 11" width="49%" /></p>

## Features
- File drag'n'drop: Drop files on the user interface.
- Batch mode: Add multiple files at once and create shortcut files for all of them.
- Non-GUI mode: Give the script a list of files as parameters and shortcut files will be created without any user interaction.
- Config file: Can remember last used settings between executions.
- Launcher: Can be used to directly launch applications through Locale Emulator without creating a shortcut file first.

## Install Instructions
Download `LEShortcutCreator.cmd` from the [releases page](https://github.com/Svintooo/LEShortcutCreator/releases).

If Locale Emulator is installed it will autodetect its location. Alternatively
it can autodetect an uninstalled Locale Emulator if `LEShortcutCreator.cmd` is
placed in the same directory.

**IMPORTANT:** Make sure to at least once run `LEInstaller.exe` from Locale Emulator (even if you do not want to have it installed). This creates some necessary dll-files that Locale Emulator needs to function.

## Usage Instructions
Either double click `LEShortcutCreator.cmd` to get a Graphical User Interface.  
The interface will let you create shortcuts that uses Locale Emulator.
Also supports file drag'n'drop.

OR drag'n'drop files onto the `LEShortcutCreator.cmd` file directly to create
shortcut files automatically without any graphical interface interaction. These
shortcut files will be created in the same directory as the script file.


## Locale Emulator
A software that can run applications (that has no Unicode support)
with a different locale than the systems default.  
It is usually used to run Japanese Games on non-Japanese versions of
Windows.

Excellent software, highly recommended.
URL: https://pooi.moe/Locale-Emulator/


## This Script
Creates shortcut files for applications that will run them through
Locale Emulator.


## Why?
Shortcut files created by Locale Emulator itself each require their own
config file. This config file is created simultaneously and stored
in the targeted applications install directory.

This script creates shortcut files that can use Locale Emulator
without these extra config files. They instead uses Locale Emulator's
global config file inside Locale Emulator's install directory (`LEConfig.xml`).


## PowerShell in a \*.cmd file?
Yes. The script has a header written in BATCH, but everything else in the file is PowerShell code.  
The only thing the BATCH code does is to execute the rest of the file as a PowerShell script.

I want the script to be executed by double clicking it.  
PowerShell files (\*.ps1) does not allow double clicking for security reasons.  
But BATCH files (\*.bat, \*.cmd) does not have this restriction.  
So by wrapping the PowerShell code in a BATCH script I circumvent this restriction.


## Common Problems
**Problem:** The "Locale Emulator location" text field is automatically filled on start. But the text field is pink, and the "Edit List" button is greyed out.  
**Solution:** Either reinstall Locale Emulator (run `LEInstaller.exe`), or put the `LEShortcutCreator.cmd` file in the same directory as Locale Emulator.  
**Reason:** This happens if Locale Emulator was once installed, and later the Locale Emulator files were moved/deleted without uninstalling it first.

**Problem:** No languages are listed in the dropdown menu.  
**Solution:** Click the "Edit List" button. This will start the Locale Emulator GUI.  
Just close this GUI (you do not need to use it).  
The default languages should now show up in the dropdown menu.  
**Reason:** The Locale Emulator file `LEConfig.xml` has been deleted. Running the GUI will recreate it with the default languages.

**Problem:** Clicking the "Edit List" button gives a weird error message "Could not load file or assembly..."  
**Solution:** Just run and immediately close `LEInstaller.exe` from Locale Emulator folder (you do not need to perform the actual install).  
**Reason:** Locale Emulator needs some dll-files to function that is not included in the download. Instead, these dll-files are created when running `LEInstaller.exe` for the first time.
