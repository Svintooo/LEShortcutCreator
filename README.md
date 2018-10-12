# LE Shortcut Creator
Create shortcuts for games/applications that will run them through Locale Emulator

Written in PowerShell. Should be compatible with Windows 7, Windows 8.1, and Windows 10.

![alt text](https://raw.githubusercontent.com/Svinto/LEShortcutCreator/master/screenshot.png)

## Usage Instructions
Either launch the script normally to get a Graphical User Interface.
The interface will let you create shortcuts that uses Locale
Emulator. Supports file drag'n'drop.

OR drag'n'drop files onto the script file to create shortcut files
automatically without any graphical interface interaction. These
shortcut files will be created in the same directory as the script
file.


## Locale Emulator
A software that can run applications (that has no Unicode support)
with a different locale than the systems default.
Usually used to run Japanese Games on non-Japanese versions of
Windows.

Exellent software, highly recommended.
URL: https://pooi.moe/Locale-Emulator/


## This Script
Creates shortcut files for applications that will run them through
Locale Emulator.


## Why?
Shortcut files created by Locale Emulator each requires their own
config file. This config file is created simultaneously and stored
in the targeted applications install directory.

This script creates shortcut files that can use Locale Emulator
without these extra config files. They instead uses Locale Emulator's
global config file inside Locale Emulator's install directory.

The script can also be used as a launcher for Locale Emulator.
Applications can be launched directly from the Graphical User
Interface.
