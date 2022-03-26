<# ::: Batch code starts here :::::::::::::::::::::::::::::::::::::::::::::::::
   :
   : Batch (*.BAT/*.CMD) wrapper for Powershell scripts
   :
   : Inspiration:
   : - https://stackoverflow.com/questions/29645/#8597794
   : - https://stackoverflow.com/questions/9074476/#9074483
   :
   : PowerShell arguments explained:
   : -Sta                     Enables support for file drag'n'drop in Windows forms.
   : -NoProfile               Makes sure no PowerShell settings disrupts script execution.
   : -ExecutionPolicy Bypass  Makes sure no PowerShell settings disrupts script execution.
   : -WindowStyle hidden      Hides the console window.

@echo off
setlocal

:: Run :: Execute this file directly as a PowerShell script
:: No temporary files needed, but errors messages is harder to understand.
set POWERSHELL_BAT_ARGS=%0 %*
if defined POWERSHELL_BAT_ARGS set POWERSHELL_BAT_ARGS=%POWERSHELL_BAT_ARGS:"=\"%
if defined POWERSHELL_BAT_ARGS set POWERSHELL_BAT_ARGS=%POWERSHELL_BAT_ARGS:$=`$%
PowerShell -Sta -NoProfile -ExecutionPolicy Bypass -WindowStyle hidden -Command Invoke-Expression $('$args=@(^&{$args} %POWERSHELL_BAT_ARGS%);'+[String]::Join([char]10,$(Get-Content '%~f0')))
EXIT /B

:: Run DEBUG :: Execute a temporary copy of this file with a *.ps1 extension
:: By running a real *.ps1 file errors messages gets easier to understand.
set TMPFILE=%~d0%~p0%~n0.debug.ps1
COPY "%~f0" "%TMPFILE%" >NUL
PowerShell -Sta -NoProfile -ExecutionPolicy Bypass -File "%TMPFILE%" %*
DEL "%TMPFILE%" >NUL
PAUSE
EXIT /B
#>



## ::: PowerShell Code Starts here ::::::::::::::::::::::::::::::::::::::::::::

### ### ### ### ### ### ### ### ### ### ### ### ### ### ### ### ### ### ### ###
###                                                                         ###
###  ==============  Locale  Emulator  Shortcut  Creator  ===============   ###
###                                                                         ###
###  Usage Instructions                                                     ###
###  ------------------                                                     ###
###  Either launch the script normally to get a Graphical User Interface.   ###
###  The interface will let you create shortcuts that uses Locale           ###
###  Emulator.                                                              ###
###                                                                         ###
###  OR drag'n'drop files onto the script file to create shortcut files     ###
###  automatically without any graphical interface interaction. These       ###
###  shortcut files will be created in the same directory as the script     ###
###  file.                                                                  ###
###                                                                         ###
###                                                                         ###
###  Locale Emulator                                                        ###
###  ---------------                                                        ###
###  A software that can run applications (that has no Unicode support)     ###
###  with a different locale than the systems default.                      ###
###  Usually used to run Japanese Games on non-Japanese versions of         ###
###  Windows.                                                               ###
###                                                                         ###
###  Exellent software, highly recommended.                                 ###
###  URL: https://pooi.moe/Locale-Emulator/                                 ###
###                                                                         ###
###                                                                         ###
###  This Script                                                            ###
###  -----------                                                            ###
###  Creates shortcut files for applications that will run them through     ###
###  Locale Emulator.                                                       ###
###                                                                         ###
###                                                                         ###
###  Why?                                                                   ###
###  ----                                                                   ###
###  Shortcut files created by Locale Emulator each requires their own      ###
###  config file. This config file is created simultaneously and stored     ###
###  in the targeted applications install directory.                        ###
###                                                                         ###
###  This script creates shortcut files that can use Locale Emulator        ###
###  without these extra config files. They instead uses Locale Emulator's  ###
###  global config file inside Locale Emulator's install directory.         ###
###                                                                         ###
###  The script can also be used as a launcher for Locale Emulator.         ###
###  Applications can be launched directly from the Graphical User          ###
###  Interface.                                                             ###
###                                                                         ###
###                                                                         ###
###  Author                                                                 ###
###  ------                                                                 ###
###  Svintoo, 2018-10-11                                                    ###
###                                                                         ###



# Stop script on any error
$ErrorActionPreference = "Stop"


#REGION BEGIN Script File Location {

# Figure out the absolute path to this script
if (Test-Path -LiteralPath $MyInvocation.MyCommand.Definition) {
  # Script file is run as a PowerShell (*.ps1) script.
  # This method of finding the script file path is compatible with PowerShell V2 and up.
  $script_path  = $MyInvocation.MyCommand.Definition
  $file_targets = $Args
} else {
  # Script file is run as a Batch (*.cmd/*.bat) script.
  # Script file path should be in the first argument.
  $script_path, $file_targets = $Args
}

$script_file = Split-Path $script_path -Leaf
$script_dir  = Split-Path $script_path -Parent
#ENDREGION Script File Location }



#REGION BEGIN Files {
$Files = New-Object PSObject -Property @{
  Script = New-Object PSObject -Property @{
    Config = Join-Path $script_dir ([IO.Path]::GetFileNameWithoutExtension($script_file) + ".config.xml")
  }
  LE     = New-Object PSObject -Property @{
    Runner = "LEProc.exe"
    Config = "LEConfig.xml"
    Editor = "LEGUI.exe"
  }
}
#ENDREGION Files }



#REGION BEGIN Create Shortcut {
function Create-LEShortcutFile {
  Param (
    <##  No  Alias  ##>[String]$Target,              # Mandatory
    [Alias("Args")    ][String]$TargetArgs,
    [Alias("LEPath")  ][String]$LocaleEmulatorPath,  # Mandatory
    [Alias("LangName")][String]$LanguageName,
    [Alias("LangID")  ][String]$LanguageID,          # Mandatory
    [Alias("Path")    ][String]$ShortcutFilePath,    # Mandatory
    [Alias("WorkDir") ][String]$WorkingDirectory
  )

  if (-Not $Target -Or -Not $LocaleEmulatorPath -Or -Not $LanguageID -Or -Not $ShortcutFilePath) {
    throw "Create-LEShortcutFile: Missing arguments "
  }

  if (-Not $WorkingDirectory) { $WorkingDirectory = Split-Path $Target -Parent }

  if ($LanguageName) {
    $Description = "\(^-^)/ $LanguageName with Locale Emulator \(^-^)/"
  } else {
    $Description = "\(^-^)/ Run with Locale Emulator \(^-^)/"
  }

  if (-Not (Test-Path -LiteralPath $ShortcutFilePath)) {
    New-Item -ItemType file $ShortcutFilePath | Out-Null
  }

  $Shell = New-Object -ComObject Shell.Application

  $ShortcutDirectory = Split-Path $ShortcutFilePath -Parent
  $ShortcutFilename  = Split-Path $ShortcutFilePath -Leaf

  $Shortcut                  = $Shell.NameSpace($ShortcutDirectory).ParseName($ShortcutFilename).GetLink
  $Shortcut.Description      = $Description
  $Shortcut.Path             = $LocaleEmulatorPath
  $Shortcut.Arguments        = "-runas `"$LanguageID`" `"$Target`" $TargetArgs".Trim()
  $Shortcut.WorkingDirectory = $WorkingDirectory
  $Shortcut.SetIconLocation($Target, 0)

  $Shortcut.Save() | Out-Null
}
#ENDREGION Create Shortcut }



#REGION BEGIN Get Config {
<#
<?xml version="1.0" encoding="UTF-8" ?>
<Config>
  <Language>$Language</Language>
  <DefaultSaveDirectory>$DefaultSaveDirectory</DefaultSaveDirectory>
  <LELocation>$LELocation</LELocation>
</Config>
#>
function Get-ConfigFile {
  $ConfigPath = $Files.Script.Config

  if (Test-Path -LiteralPath $ConfigPath -PathType Leaf -ErrorAction SilentlyContinue) {
    $config = try{ ([Xml](Get-Content $ConfigPath)).Config } catch {}
  }

  # Return $Null if config file is not found or can not be parsed
  return $config
}
#ENDREGION Get Config }



#REGION BEGIN FileSystem Paths {
<# # Explanation to why a Regular Expression solution is used for normalizing paths # #
 # (Join-Path $path1 '') -Eq (Join-Path $path2 '')
 #   o Handles paths that doesn't exists:        "C:\fake\" == "C:\fake\"
 #   o Handles both \ and /:                     "C:\asdf\" == "C:/asdf/"
 #   o Handles missing \ at end of line:         "C:\asdf\" == "C:\asdf"
 #   o Handles multiple \ at end of line:        "C:\asdf\" == "C:\asdf\\"
 #   x Doesn't handle multiple \ anywhere else:  "C:\\asdf" != "C:\asdf"
 #   x Doesn't handle spaces before \ and end:   "C: \asdf" != "C:\asdf"  #NOTE: Only in PowerShell.
 #                                               "C:\asdf " != "C:\asdf"         Not valid in Batch (cmd).
 # (Get-ItemProperty -LiteralPath $path1).FullName -Eq (Get-ItemProperty -LiteralPath $path2).FullName
 #   x Doesn't handle paths that doesn't exists: "C:\fake\" != "C:\fake\"
 #   o Handles both \ and /:                     "C:\asdf\" == "C:/asdf/"
 #   x Doesn't handle missing \ at end of line:  "C:\asdf\" != "C:\asdf"
 #   o Handles multiple \ at end of line:        "C:\asdf\" == "C:\asdf\\"
 #   o Handles multiple \ anywhere else:         "C:\\asdf" != "C:\asdf"
 #   o Handles spaces before \ and end:          "C: \asdf" == "C:\asdf"  #NOTE: Only in PowerShell.
 #                                               "C:\asdf " != "C:\asdf"         Not valid in Batch (cmd).
 # -Not (Compare-Object (Get-ItemProperty -LiteralPath $path1) (Get-ItemProperty -LiteralPath $path2))
 #   NO! Returns $Null if EQUAL. Code is hard to understand.
 #   AND doesn't handle paths that doesn't exists.
 #>

# All strings that resolve to the same file system path are here normalized to the same string
function Normalize-Path([String]$path) {
  # C:/path/somewhere  => C:\path\somewhere
  # C:\path\\somewhere => C:\path\somewhere
  # C:\path\somewhere\ => C:\path\somewhere
  # C:\path \somewhere => C:\path\somewhere  # PowerShell ignores whitespaces at end of file/folder names
  $path -Replace "/", "\" `
        -Replace "\\\\+", "\" `
        -Replace "\\$", "" `
        -Replace " +\\", "\"
}

# Compare two or more file system paths
function Equal-Paths {
  if ($Args.Count -LT 2) {
    throw "Equal-Paths: At least two paths are needed for comparison"
  }

  $reference_path = Normalize-Path $Args[0]

  for ($i = 1 ; $i -LT $Args.Count ; $i++) {
    $path = Normalize-Path $Args[$i]

    if ($path -NE $reference_path) {
      return $False
    }
  }

  return $True
}
#ENDREGION FileSystem Paths }



#REGION BEGIN Locale Emulator Functions {
function Verify-LEDirectory([String]$Directory) {
  # Test-Path arguments '-ErrorAction SilentlyContinue' is ignored when:
  # - using arguments '-PathType Container', OR
  # - receiving an empty string ("") as path.
  # Hence the try-catch.
  $dir_exists = try { Test-Path -LiteralPath $Directory -PathType Container } catch { $False }

  if (-Not $dir_exists) { return $False }
  if (-Not (Test-Path -LiteralPath (Join-Path -Path $Directory -ChildPath $Files.LE.Runner))) { return $False }

  return $True
}

function Get-LEDirectories {
  Param (
    [String]$ConfigLEDirectory = ""
  )

  $LEDirs = New-Object System.Collections.ArrayList

  # Config directory
  if ($ConfigLEDirectory) {
    $DirectoryPath = $ConfigLEDirectory
    $Valid = Verify-LEDirectory $DirectoryPath
    $LEDirs.Add((New-Object PSObject -Property @{Path = $DirectoryPath; Valid = $Valid})) | Out-Null
  }

  # Script directory
  $DirectoryPath = $script_dir
  $Valid = Verify-LEDirectory $DirectoryPath
  if ($Valid) {
    $LEDirs.Add((New-Object PSObject -Property @{Path = $DirectoryPath; Valid = $Valid})) | Out-Null
  }

  # Install directory
  if (-Not (Test-Path "HKCR:")) { New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null }
  $regKeyName = "HKCR:\CLSID\{C52B9871-E5E9-41FD-B84D-C5ACADBEC7AE}\InprocServer32"
  $regValueName = "CodeBase"
  if (Test-Path $regKeyName) {
    $regKey = Get-Item $regKeyName
    $LE_DLLPath = [String](Get-ItemProperty $regKey.PSPath).$regValueName
    if ($LE_DLLPath) {
      #Ex: $LE_DLLPath = "file:///C:/path/to/Locale Emulator/LEContextMenuHandler.DLL"
      $DirectoryPath = $LE_DLLPath -Replace '^file:///','' -Replace '/[^/]+$','' -Replace '/','\'
      $Valid = Verify-LEDirectory $DirectoryPath
      $LEDirs.Add((New-Object PSObject -Property @{Path = $DirectoryPath; Valid = $Valid})) | Out-Null
    }
  }

  return ,$LEDirs
}

# Returns a list of the profiles in Locale Emulator's config file.
function Get-LELanguages([String]$Directory) {
  # ArrayList is used instead of Powershell arrays because:
  # - Faster when changing the number of elements in the array.
  # - required when used with ComboBoxes in Windows Forms.
  $Languages = New-Object System.Collections.ArrayList

  [XML]$Config = try {
    Get-Content (Join-Path -Path $Directory -ChildPath $Files.LE.Config)
  } catch {
    # Return empty language list if the LE config file doesn't exist
    return ,$Languages
  }

  if ($Config.LEConfig.Profiles.Profile -Is [Xml.XmlLinkedNode]) {
    # Only one <Profile> element in XML document
    $Profiles = @($Config.LEConfig.Profiles.Profile)
  } elseif ($Config.LEConfig.Profiles.Profile -Is [Array]) {
    # Multiple <Profile> elements in XML document
    $Profiles = $Config.LEConfig.Profiles.Profile
  }

  $Profiles | % {
    if ($_.Name -And $_.Guid) {
      $Language = New-Object PSObject -Property @{Name = $_.Name; Guid = $_.Guid}
      $Languages.Add($Language) | Out-Null
    }
  }

  return ,$Languages
}
#ENDREGION Locale Emulator Functions }









### ### ### ### ###  The Automatic Part  ### ### ### ### ###
###                                                      ###
###  No user interaction here.                           ###
###                                                      ###
###  File paths are received through script arguments    ###
###  (by drag'n'drop) and shortcuts are created          ###
###  automatically.                                      ###
###                                                      ###
###  NOTE: Settings in the config file are used here     ###
###                                                      ###

#REGION BEGIN NoGUI {
Add-Type -AssemblyName PresentationFramework

if ($file_targets.Count -GT 0) {
  # Verify that all targets exists
  ForEach ($Target in $file_targets) {
    $TargetExists = Test-Path -LiteralPath $Target -PathType Leaf -ErrorAction SilentlyContinue
    if (-Not $TargetExists) {
      [System.Windows.MessageBox]::Show("File not found: $Target")
      Exit
    }
  }

  # Parse config file (if it exists)
  $Config = Get-ConfigFile

  # Find Locale Emulator
  if ($Config) {
    $LEDirectories = Get-LEDirectories -ConfigLEDirectory $Config.LELocation
    $LanguageName = $Config.Language
  } else {
    $LEDirectories = Get-LEDirectories
  }

  if ($LEDirectories.Count -Eq 0) {
    [System.Windows.MessageBox]::Show("Can't find Locale Emulator")
    Exit
  }

  $LEDirectory = $LEDirectories[0]

  if (-Not $LEDirectory.Valid) {
    [System.Windows.MessageBox]::Show("Configured path to Locale Emulator is invalid: ${$LEDirectory.Path}")
    Exit
  }

  $LocaleEmulatorPath = Join-Path $LEDirectory.Path $Files.LE.Runner

  # Fetch all available Languages (Profiles) configured in Locale Emulator
  $Languages = Get-LELanguages $LEDirectory.Path

  if ($Languages.Count -Eq 0) {
    [System.Windows.MessageBox]::Show("Locale Emulator has no configured profiles: ${Join-Path $LEDirectory.Path $Files.LE.Config}")
    Exit
  }

  # Select one language
  if ($LanguageName) {
    $Language = $Languages | Where { $_.Name -Eq $LanguageName } | Select -First 1
  }

  if (-Not $Language) { $Language = $Languages[0] }

  # Create all shortcut files and save them to the script directory
  ForEach ($Target in $file_targets) {
    $ShortcutFilename = [IO.Path]::GetFileNameWithoutExtension($Target)
    $ShortcutFilename = "(LE)" + $ShortcutFilename + ".lnk"
    $ShortcutFilePath = Join-Path $script_dir $ShortcutFilename

    Create-LEShortcutFile -Target             $Target `
                          -LocaleEmulatorPath $LocaleEmulatorPath `
                          -LanguageName       $Language.Name `
                          -LanguageID         $Language.Guid `
                          -ShortcutFilePath   $ShortcutFilePath
  }

  Exit
}
#ENDREGION NoGUI }








### ### ### ### ###  The Interactive Part  ### ### ### ### ###
###                                                        ###
###  Displays a Graphical User Interface (GUI)             ###
###                                                        ###
###  A user can here select files and configure settings   ###
###  before deciding to either create shortcuts OR launch  ###
###  a selected file directly.                             ###
###                                                        ###
###  Selected settings can be saved and remembered.        ### 
###                                                        ###

#REGION BEGIN GUI {
<# This form was created using POSHGUI.com  a free online gui designer for PowerShell #>
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()



#REGION BEGIN GUI Colors {
$TextBoxColors = New-Object PSObject -Property @{
  None    = [System.Drawing.Color]::Empty
  Invalid = "255,235,238"  # Light Pink (textbox bgcolor, used when text is invalid)
}
#ENDREGION GUI Colors }



#REGION BEGIN GUI Elements {
$Form                            = New-Object System.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Size(400,400)  #'400,400'
$Form.Text                       = "LE Shortcut Creator"
$Form.TopMost                    = $False
$Form.MaximizeBox                = $False
$Form.SizeGripStyle              = "Hide"
$Form.StartPosition              = "CenterScreen"  # Manual, CenterScreen, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
$Form.FormBorderStyle            = 'FixedDialog'  # None, FixedSingle, Fixed3D, FixedDialog, Sizable, FixedToolWindow, SizableToolWindow

$labels_x                        =   9
$browse_buttons_width            =  74
$browse_buttons_height           =  24
$browse_buttons_x                = 316
$browse_textbox_width            = 304
$browse_textbox_height           =  22
$lang_buttons_width              =  88
$lang_buttons_height             =  26
$bottom_buttons_y                = 360

$Target_Label                    = New-Object System.Windows.Forms.Label
$Target_Label.Text               = "Target(s)"
$Target_Label.AutoSize           = $True
$Target_Label.Width              = 25
$Target_Label.Height             = 10
$Target_Label.Location           = New-Object System.Drawing.Point($labels_x,5)

$Target_TextBox                  = New-Object System.Windows.Forms.TextBox
$Target_TextBox.Multiline        = $False
$Target_TextBox.Width            = $browse_textbox_width
$Target_TextBox.Height           = $browse_textbox_height
$Target_TextBox.AutoSize         = $False #Win11
$Target_TextBox.Location         = New-Object System.Drawing.Point(11,24)
Add-Member -InputObject $Target_TextBox -MemberType "NoteProperty" -Name "Targets" -Value @()
# Note that the initial directory fallback below is modified during execution
Add-Member -InputObject $Target_TextBox -MemberType "NoteProperty" -Name "FileBrowserInitialDirectoryFallback" `
                                                                   -Value $script_dir

$Target_Button                   = New-Object System.Windows.Forms.Button
$Target_Button.Text              = "Browse..."
$Target_Button.Width             = $browse_buttons_width
$Target_Button.Height            = $browse_buttons_height
$Target_Button.Location          = New-Object System.Drawing.Point($browse_buttons_x,23)

$Args_Label                      = New-Object System.Windows.Forms.Label
$Args_Label.Text                 = "Target Arguments (optional)"
$Args_Label.AutoSize             = $True
$Args_Label.Width                = 25
$Args_Label.Height               = 10
$Args_Label.Location             = New-Object System.Drawing.Point($labels_x,59)

$Args_TextBox                    = New-Object System.Windows.Forms.TextBox
$Args_TextBox.Multiline          = $False
$Args_TextBox.Width              = $browse_textbox_width + $browse_buttons_width
$Args_TextBox.Height             = $browse_textbox_height
$Args_TextBox.AutoSize           = $False #Win11
$Args_TextBox.Location           = New-Object System.Drawing.Point(11,78)

$Lang_Groupbox                   = New-Object System.Windows.Forms.Groupbox
$Lang_Groupbox.Height            =  56
$Lang_Groupbox.Width             = 378
$Lang_Groupbox.Text              = "Language"
$Lang_Groupbox.Location          = New-Object System.Drawing.Point(11,119)

$Lang_ComboBox                   = New-Object System.Windows.Forms.ComboBox
$Lang_ComboBox.Width             = 266
$Lang_ComboBox.Height            =  24
$Lang_ComboBox.Location          = New-Object System.Drawing.Point(9,19)
$Lang_ComboBox.DropDownStyle     = "DropDownList"
$Lang_ComboBox.DisplayMember     = "Name"
$Lang_ComboBox.ValueMember       = "Guid"
$Lang_ComboBox.DataSource        = New-Object System.Collections.ArrayList(@(,@(New-Object PSObject -Property @{Name="";Guid=""})))

$Lang_Button_Edit                = New-Object System.Windows.Forms.Button
$Lang_Button_Edit.Text           = "Edit List"
$Lang_Button_Edit.Width          = $lang_buttons_width
$Lang_Button_Edit.Height         = $lang_buttons_height
$Lang_Button_Edit.Location       = New-Object System.Drawing.Point(282,18)
$Lang_Button_Edit.Enabled        = $False

$SaveSettings_Checkbox           = New-Object System.Windows.Forms.Checkbox
$SaveSettings_Checkbox.Text      = "Remember Settings"
$SaveSettings_Checkbox.AutoSize  = $True
$SaveSettings_Checkbox.AutoCheck = $True
$SaveSettings_Checkbox.Location  = New-Object System.Drawing.Point(($labels_x+2),187)

$SaveDir_Label                   = New-Object System.Windows.Forms.Label
$SaveDir_Label.Text              = "Default Save directory"
$SaveDir_Label.AutoSize          = $True
$SaveDir_Label.Width             = 25
$SaveDir_Label.Height            = 10
$SaveDir_Label.Location          = New-Object System.Drawing.Point($labels_x,223)

$SaveDir_TextBox                 = New-Object System.Windows.Forms.TextBox
$SaveDir_TextBox.Multiline       = $False
$SaveDir_TextBox.Width           = $browse_textbox_width
$SaveDir_TextBox.Height          = $browse_textbox_height
$SaveDir_TextBox.AutoSize        = $False #Win11
$SaveDir_TextBox.Location        = New-Object System.Drawing.Point(11,242)
$SaveDir_TextBox.text            = [Environment]::GetFolderPath("Desktop")
Add-Member -InputObject $SaveDir_TextBox -MemberType "NoteProperty" -Name "TextWas" -Value $SaveDir_TextBox.text

$SaveDir_Button                  = New-Object System.Windows.Forms.Button
$SaveDir_Button.Text             = "Browse..."
$SaveDir_Button.Width            = $browse_buttons_width
$SaveDir_Button.Height           = $browse_buttons_height
$SaveDir_Button.Location         = New-Object System.Drawing.Point($browse_buttons_x,241)

$LELocation_Label                = New-Object System.Windows.Forms.Label
$LELocation_Label.Text           = "Locale Emulator location"
$LELocation_Label.AutoSize       = $True
$LELocation_Label.Width          = 25
$LELocation_Label.Height         = 10
$LELocation_Label.Location       = New-Object System.Drawing.Point($labels_x,277)

$LELocation_TextBox              = New-Object System.Windows.Forms.TextBox
$LELocation_TextBox.Multiline    = $False
$LELocation_TextBox.Width        = $browse_textbox_width
$LELocation_TextBox.Height       = $browse_textbox_height
$LELocation_TextBox.AutoSize     = $False #Win11
$LELocation_TextBox.BackColor    = $TextBoxColors.Invalid
$LELocation_TextBox.Location     = New-Object System.Drawing.Point(11,294)
Add-Member -InputObject $LELocation_TextBox -MemberType "NoteProperty" -Name "TextWas"  -Value ""
Add-Member -InputObject $LELocation_TextBox -MemberType "NoteProperty" -Name "Valid"    -Value $False
Add-Member -InputObject $LELocation_TextBox -MemberType "NoteProperty" -Name "NewValid" -Value $Null

$LELocation_Button               = New-Object System.Windows.Forms.Button
$LELocation_Button.Text          = "Browse..."
$LELocation_Button.Width         = $browse_buttons_width
$LELocation_Button.Height        = $browse_buttons_height
$LELocation_Button.Location      = New-Object System.Drawing.Point($browse_buttons_x,293)

$Run_Button                      = New-Object System.Windows.Forms.Button
$Run_Button.Text                 = "Run Target"
$Run_Button.Width                = 115
$Run_Button.Height               =  30
$Run_Button.Location             = New-Object System.Drawing.Point(10,$bottom_buttons_y)
$Run_Button.Enabled              = $False

$Save_Button                     = New-Object System.Windows.Forms.Button
Add-Member -InputObject $Save_Button -MemberType "NoteProperty" -Name "TextSingular" -Value "Save Shortcut"
Add-Member -InputObject $Save_Button -MemberType "NoteProperty" -Name "TextPlural"   -Value "Save Shortcuts"
$Save_Button.Text                = $Save_Button.textSingular
$Save_Button.Width               = 115
$Save_Button.Height              =  30
$Save_Button.Location            = New-Object System.Drawing.Point(142,$bottom_buttons_y)
$Save_Button.Enabled             = $False

$Quit_Button                     = New-Object System.Windows.Forms.Button
$Quit_Button.Text                = "Quit"
$Quit_Button.Width               = 115
$Quit_Button.Height              =  30
$Quit_Button.Location            = New-Object System.Drawing.Point(275,$bottom_buttons_y)

$form_controls = @($Target_TextBox,$Target_Button,$Target_Label,$Args_TextBox,$Args_Label,$Lang_Groupbox,$SaveSettings_Checkbox,$SaveDir_TextBox,$SaveDir_Button,$SaveDir_Label,$LELocation_TextBox,$LELocation_Button,$LELocation_Label,$Run_Button,$Save_Button,$Quit_Button)
$lang_controls = @($Lang_ComboBox,$Lang_Button_Edit)
$controls = $form_controls + $lang_controls

$Form.controls.AddRange($form_controls)
$Lang_Groupbox.controls.AddRange($lang_controls)
#ENDREGION GUI Elements }



#REGION BEGIN GUI LELocations Elements }
$Form_LELocations                  = New-Object System.Windows.Forms.Form
$Form_LELocations.ClientSize       = New-Object System.Drawing.Size(400,182)
$Form_LELocations.Text             = "LE Locations"
$Form_LELocations.TopMost          = $False
$Form_LELocations.MaximizeBox      = $False
$Form_LELocations.SizeGripStyle    = "Hide"
$Form_LELocations.StartPosition    = "CenterScreen"  # Manual, CenterScreen, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
$Form_LELocations.FormBorderStyle  = 'FixedDialog'   # None, FixedSingle, Fixed3D, FixedDialog, Sizable, FixedToolWindow, SizableToolWindow
Add-Member -InputObject $Form_LELocations -MemberType "NoteProperty" -Name "ButtonAction" -Value ""

$LELocations_Label                 = New-Object System.Windows.Forms.Label
$LELocations_Label.Text            = "Which Locale Emultor do you want to use?"
$LELocations_Label.AutoSize        = $True
$LELocations_Label.Height          = 10
$LELocations_Label.Location        = New-Object System.Drawing.Point(9,8)

$LELocations_ListBox               = New-Object System.Windows.Forms.ListBox
$LELocations_ListBox.SelectionMode = "one"
$LELocations_ListBox.Sorted        = $False
$LELocations_ListBox.HorizontalScrollbar = $True
$LELocations_ListBox.Width         = 378
$LELocations_ListBox.Height        = 100
$LELocations_ListBox.Location      = New-Object System.Drawing.Point(11,32)
$LELocations_ListBox.DisplayMember = "Name"
$LELocations_ListBox.ValueMember   = "Value"
$LELocations_ListBox.DataSource = New-Object System.Collections.ArrayList(@(,@(New-Object PSObject -Property @{Name=""; Value=""})))
Add-Member -InputObject $LELocations_ListBox -MemberType "NoteProperty" -Name "LastSelection" -Value ""

$LELocations_ButtonCancel          = New-Object System.Windows.Forms.Button
$LELocations_ButtonCancel.Text     = "Cancel"
$LELocations_ButtonCancel.width    = 115
$LELocations_ButtonCancel.Height   =  30
$LELocations_ButtonCancel.Location = New-Object System.Drawing.Point(10,142)

$LELocations_ButtonBrowse          = New-Object System.Windows.Forms.Button
$LELocations_ButtonBrowse.Text     = "Browse"
$LELocations_ButtonBrowse.Width    = 115
$LELocations_ButtonBrowse.Height   =  30
$LELocations_ButtonBrowse.Location = New-Object System.Drawing.Point(142,142)

$LELocations_ButtonSelect          = New-Object System.Windows.Forms.Button
$LELocations_ButtonSelect.Text     = "Select"
$LELocations_ButtonSelect.Width    = 115
$LELocations_ButtonSelect.Height   =  30
$LELocations_ButtonSelect.Location = New-Object System.Drawing.Point(275,142)

$lelocation_controls = @($LELocations_ListBox,$LELocations_Label,$LELocations_ButtonCancel,$LELocations_ButtonBrowse,$LELocations_ButtonSelect)
$Form_LELocations.controls.AddRange($lelocation_controls)
#ENDREGION GUI LELocations Elements }



#REGION BEGIN GUI Functions {

# Tries to find a valid absolute path in a string (best effort).
# Should return an empty string on failure.
function Find-ValidPath([String]$String) {
  $Path = $String.Trim(' "') -Replace '^[^a-zA-Z]+',''
  $Path = Normalize-Path $Path

  #if ($Path.Length -LT 2 -Or $Path -Match '^[a-zA-Z][^:]') {
  if ($Path.Length -LT 2 -Or -Not ($Path -Match '^[a-zA-Z]+:')) {
    return ""
  }

  if ($Path) {
    # Shorten the path until a valid one is found
    while ($Path -NE "") {
      # Test-Path arguments '-ErrorAction SilentlyContinue' is ignored when:
      # - using arguments '-PathType Container', OR
      # - receiving an empty string ("") as path.
      # Hence the try-catch.
      $valid_path = try { Test-Path -LiteralPath $Path -PathType Container } catch { $False }
      if ($valid_path) { break }
      $Path = Split-Path $Path -Parent
    }
  }

  return $Path
}

function Show-SaveFileDialog {
  Param (
    [String]$Title                    = "",
    [String]$InitialFilename          = "",
    [String]$InitialDirectory         = "",
    [String]$InitialDirectoryFallback = "",
    [String]$Filter                   = "",
    [Int]   $FilterIndex              = 0
  )

  $SaveFileDialog = New-Object -Typename System.Windows.Forms.SaveFileDialog
  $SaveFileDialog.Title            = $Title
  $SaveFileDialog.Filter           = $Filter
  $SaveFileDialog.FilterIndex      = $FilterIndex
  $SaveFileDialog.FileName         = $InitialFilename
  $SaveFileDialog.ShowHelp         = $True

  if ($InitialDirectory -And ($InitialDirectory = Find-ValidPath $InitialDirectory)) {
    $SaveFileDialog.InitialDirectory = $InitialDirectory
  } elseif ($InitialDirectoryFallback) {
    $SaveFileDialog.InitialDirectory = $InitialDirectoryFallback
  }

  $dialog_result = $SaveFileDialog.ShowDialog()

  if ($dialog_result -Eq [System.Windows.Forms.DialogResult]::OK) {
    return $SaveFileDialog.Filename
  } else {
    return $Null
  }
}

function Show-OpenFileDialog {
  Param (
    [String]$InitialDirectory         = "",
    [String]$InitialDirectoryFallback = "",
    [String]$Filter                   = "",
    [Int]   $FilterIndex              = 0,
    [Switch]$CheckFileExists,
    [Switch]$Multiselect
  )

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.Filter          = $Filter  #"All files (*.*)| *.*|Applications (*.exe)| *.exe"
  $OpenFileDialog.FilterIndex     = $FilterIndex
  $OpenFileDialog.CheckFileExists = $CheckFileExists
  $OpenFileDialog.Multiselect     = $Multiselect
  $OpenFileDialog.ShowHelp        = $True  # OpenFileDialog won't show unless the help button is displayed

  if ($InitialDirectory -And ($InitialDirectory = Find-ValidPath $InitialDirectory)) {
    $OpenFileDialog.InitialDirectory = $InitialDirectory
  } elseif ($InitialDirectoryFallback) {
    $OpenFileDialog.InitialDirectory = $InitialDirectoryFallback
  }

  $dialog_result = $OpenFileDialog.ShowDialog()

  if ($dialog_result -NE [System.Windows.Forms.DialogResult]::OK) {
    return $Null
  } elseif ($OpenFileDialog.FileNames.Count -Eq 1) {
    return $OpenFileDialog.FileName
  } else {
    return '"' + ($OpenFileDialog.FileNames -Join '" "') + '"'
  }
} #END Function Show-OpenFileDialog


function Show-FolderBrowserDialog([String]$Description, [bool]$ShowNewFolderButton = $False, [String]$InitialDirectory = "") {
  $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $FolderBrowserDialog.Description         = $Description
  $FolderBrowserDialog.RootFolder          = "Desktop"  # Desktop, MyComputer, MyDocuments, Favorites, Personal, DesktopDirectory
  $FolderBrowserDialog.ShowNewFolderButton = $ShowNewFolderButton

  if ($InitialDirectory -And ($InitialDirectory = Find-ValidPath $InitialDirectory)) {
    $FolderBrowserDialog.SelectedPath = $InitialDirectory
  } else {
    $FolderBrowserDialog.SelectedPath = [Environment]::GetFolderPath("Desktop")
  }

  $dialog_result = $FolderBrowserDialog.ShowDialog()

  if ($dialog_result -NE [System.Windows.Forms.DialogResult]::OK) {
    return $Null
  } else {
    return $FolderBrowserDialog.SelectedPath
  }
} #END Function Show-FolderBrowserDialog


function Show-LELocationDialog($LEPaths, $DefaultSelection) {
  $Options = New-Object System.Collections.ArrayList
  if (-Not $DefaultSelection) { $DefaultSelection = $LELocations_ListBox.LastSelection }

  ForEach ($LEPath in $LEPaths) {
    $Options.Add((New-Object PSObject -Property @{Name=$LEPath; Value=$LEPath})) | Out-Null
  }

  $LELocations_ListBox.DataSource = $Options

  $Option = $LELocations_ListBox.DataSource | Where { $_.Value -Eq $DefaultSelection } | Select -First 1
  if ($Option) { $LELocations_ListBox.SelectedItem = $Option }

  #$Form_LELocations.Location = Get-NewGUIPosition $Form_LELocations
  $Form_LELocations.ShowDialog() | Out-Null

  Switch -Exact ($Form_LELocations.ButtonAction) {
    "Select" { return $LELocations_ListBox.SelectedValue }
    "Browse" { return "browse" }
    "Cancel" { return $Null }
    default  { Throw "Fatal Error: This should never happen" }
  }
}

function GUI-Update-LE {
  GUI-Update-Languages
  GUI-Update-RunCreate
}

function GUI-Disable-LE {
  GUI-Disable-Languages
  GUI-Disable-RunCreate
}

function GUI-Update-RunCreate {
  if (-Not $LELocation_TextBox.Valid -Or -Not $Lang_ComboBox.Enabled) {
    $Run_Button.Enabled  = $False
    $Save_Button.Enabled = $False
    return
  }

  if ($Target_TextBox.Targets.Count -Eq 1) {
    $Run_Button.Enabled  = $True
    $Save_Button.Enabled = $True
    $Save_Button.Text    = $Save_Button.TextSingular
  } elseif ($Target_TextBox.Targets.Count -ge 2) {
    $Run_Button.Enabled  = $False
    $Save_Button.Enabled = $True
    $Save_Button.Text    = $Save_Button.TextPlural
  }
}

function GUI-Disable-RunCreate {
  $Run_Button.Enabled  = $False
  $Save_Button.Enabled = $False
}

function GUI-Update-Languages {
  $Languages = Get-LELanguages -Directory $LELocation_TextBox.Text

  if ($Languages.Count -Eq 0) {
    # ComboBox.DataSource must at all times be a 
    # System.Collections.ArrayList containing at least one PSObject.
    # Otherwise it will start displaying the elements incorrectly in the GUI.
    $empty_language = New-Object PSObject -Property @{Name=""; Guid=""}
    $Languages.Add($empty_language) | Out-Null
    $Lang_ComboBox.DataSource = $Languages
    $Lang_ComboBox.Enabled = $False
    return
  }

  $Lang_ComboBox.Enabled = $True

  $language_differences = Compare-Object -Property Name,Guid $Languages $Lang_ComboBox.DataSource
  $current_selection_name = $Lang_ComboBox.SelectedItem.Name

  if ($language_differences) {
    $Lang_ComboBox.DataSource = $Languages

    # If possible, change the selection to the same language as before
    $Language = $Lang_ComboBox.DataSource | Where { $_.Name -Eq $current_selection_name } | Select -First 1
    if ($Language) { $Lang_ComboBox.SelectedItem = $Language }
  }
}

function GUI-Disable-Languages {
  $Lang_ComboBox.Enabled = $False
}

function Save-ConfigFile {
  $Language             = $Lang_ComboBox.SelectedItem.Name
  $DefaultSaveDirectory = $SaveDir_TextBox.Text
  $LELocation           = $LELocation_TextBox.Text

  $Config = @"
<?xml version="1.0" encoding="UTF-8" ?>
<Config>
  <Language>$Language</Language>
  <DefaultSaveDirectory>$DefaultSaveDirectory</DefaultSaveDirectory>
  <LELocation>$LELocation</LELocation>
</Config>
"@

  $ConfigPath = $Files.Script.Config
  $Config | Out-File $ConfigPath
}

function Quit-Application {
  if ($SaveSettings_Checkbox.Checked) {
    Save-ConfigFile
  } elseif (Test-Path -LiteralPath $Files.Script.Config -PathType Leaf -ErrorAction SilentlyContinue) {
    Remove-Item -Path $Files.Script.Config
  }

  $Form.Close()
}
#ENDREGION GUI Functions }



#REGION BEGIN GUI Events {
# Here there be event code for the main window

# Mouse drop support
#NOTE: This requires that PowerShell was executed with: -Sta
$Form.AllowDrop = $True
$Form.Add_DragEnter({ $_.Effect = 'Copy' })
$Form.Add_DragDrop({
  if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::Text)) {
    # If the dropped data is: plain text
    $Target_TextBox.Text = $_.Data.GetData([Windows.Forms.DataFormats]::Text)
  } elseif ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
    # If the dropped data is: one or more files
    $Filenames = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)

    if ($Filenames.Count -Eq 1) {
      $Target_TextBox.Text = $Filenames[0]
    } else {
      $Target_TextBox.Text = '"' + ($Filenames -Join '" "') + '"'
    }
  }
})


$Target_TextBox.Add_TextChanged({
  # Normalize targets text string
  $targets_text = $Target_TextBox.Text
  $targets_text = $targets_text.Trim()  # Remove surrounding whitespaces

  # This script block is run if targets are deemed invalid
  $targets_invalid_scriptblock = {
    $Target_TextBox.Targets   = @()
    $Target_TextBox.BackColor = $TextBoxColors.Invalid
    GUI-Disable-RunCreate
  }

  # If empty string
  if ($targets_text -Eq "") {
    Invoke-Command -ScriptBlock $targets_invalid_scriptblock
    $Target_TextBox.BackColor = $TextBoxColors.None  # No error color on empty string
    return
  }

  # If NOT inside quotation marks
  if ($targets_text -Match '^[^"]' -And $targets_text -Match '[^"]$') {
    $targets_text = '"' + $targets_text + '"'
  }

  #TODO: Maybe this isn't needed?
  if ($targets_text -Eq '""') {
    Invoke-Command -ScriptBlock $targets_invalid_scriptblock
    return
  }

  # If uneven number of double quotes (")
  if ( ([Regex]::Matches($targets_text,'"')).Count % 2 -NE 0 ) {
    Invoke-Command -ScriptBlock $targets_invalid_scriptblock
    return
  }

  # Convert $targets_text to an array of strings
  $Targets = [Regex]::Matches($targets_text,'"([^"])+"') | ForEach {$_.Value -Replace '"',''} | Select -Uniq
  # Restore the array wrapping that PowerShell removes if it only contains one element
  if ($Targets.GetType() -Eq [String]) {
    $Targets = @($Targets)
  }

  # Check that all targets can be found
  ForEach ( $Target in $Targets ) {
    $valid_path = Test-Path -LiteralPath $Target -PathType Leaf -ErrorAction SilentlyContinue

    if (-Not $valid_path) {
      Invoke-Command -ScriptBlock $targets_invalid_scriptblock
      return
    }
  }

  # Save targets
  $Target_TextBox.FileBrowserInitialDirectoryFallback = $Targets[0]
  $Target_TextBox.Targets = $Targets

  # Clear Background Color
  $Target_TextBox.BackColor = $TextBoxColors.None

  # Try to activate the "Run" and "Create Shortcut" buttons
  GUI-Update-RunCreate
})


$Target_Button.Add_Click({
  $Filenames = Show-OpenFileDialog -InitialDirectory $Target_TextBox.Text `
                                   -InitialDirectoryFallback $Target_TextBox.FileBrowserInitialDirectoryFallback `
                                   -Filter "All files (*.*)| *.*|Applications (*.exe)| *.exe" `
                                   -FilterIndex 2 `
                                   -CheckFileExists `
                                   -Multiselect
  if ($Filenames -NE $Null) { $Target_TextBox.Text = $Filenames }
})


$Lang_Button_Edit.Add_Click({
  $Path = Join-Path $LELocation_TextBox.Text $Files.LE.Editor

  if (-Not (Test-Path $Path)) {
    $Title   = "Unable to start editor"
    $Message = "The Locale Emulator GUI seems to be missing.$([Environment]::NewLine
               )File not found:$([Environment]::NewLine
               )$Path"
    $Buttons = "OK"
    $Icon    = "Error"
    [System.Windows.MessageBox]::Show($Message,$Title,$Buttons,$Icon)
    return
  }

  # Launch the LE Editor
  #TODO: Prevent the GUI from being used until the LE Editor has exited
  #$Form.Enabled = $False
  $PS = New-Object System.Diagnostics.Process
  $PS.StartInfo.Filename = $Path
  $PS.Start()
  while (-Not $PS.HasExited) {
    [System.Windows.Forms.Application]::DoEvents()  # IMPORTANT!!!
    Start-Sleep -Milliseconds 100
  }
  #$Form.Enabled = $True

  GUI-Update-Languages
})


$SaveDir_TextBox.Add_LostFocus({
  # Test-Path arguments '-ErrorAction SilentlyContinue' is ignored when:
  # - using arguments '-PathType Container', OR
  # - receiving an empty string ("") as path.
  # Hence the try-catch.
  $dir_exists = try { Test-Path -LiteralPath $SaveDir_TextBox.Text -PathType Container } catch { $False }

  if ($dir_exists) {
    $SaveDir_TextBox.TextWas = $SaveDir_TextBox.Text
  } else {
    $SaveDir_TextBox.Text = $SaveDir_TextBox.TextWas
  }
})


$SaveDir_Button.Add_Click({
  $Directory = Show-FolderBrowserDialog -Description "Choose the default directory where to save shortcuts" `
                                        -ShowNewFolderButton $True `
                                        -InitialDirectory $SaveDir_TextBox.Text
  if ($Directory -NE $Null) { $SaveDir_TextBox.Text = $Directory }
})


$LELocation_TextBox.Add_TextChanged({
  $path                    = $LELocation_TextBox.Text
  $path_was                = $LELocation_TextBox.TextWas
  $pre_verification_result = $LELocation_TextBox.NewValid

  $LELocation_TextBox.TextWas  = $LELocation_TextBox.Text
  $LELocation_TextBox.NewValid = $Null

  if (Equal-Paths $path $path_was) {
    return
  }

  if ($pre_verification_result -NE $Null) {
    $path_is_valid = $pre_verification_result
  } else {
    $path_is_valid = Verify-LEDirectory $path
  }

  if ($path_is_valid) {
    $LELocation_TextBox.Valid     = $True
    $LELocation_TextBox.BackColor = $TextBoxColors.None
    $Lang_Button_Edit.Enabled     = Test-Path -LiteralPath (Join-Path $LELocation_TextBox.Text $Files.LE.Editor)
    GUI-Update-LE
  } else {
    $LELocation_TextBox.Valid     = $False
    $LELocation_TextBox.BackColor = $TextBoxColors.Invalid
    $Lang_Button_Edit.Enabled     = $False
    GUI-Disable-LE
  }
})


## Select a folder containing Locale Emulator (LE)
# Normally show a Folder Browser Dialog.
# BUT: If one or more LE directories has been autodetected (that is not the
#      same as $LELocation_TextBox.Text), then show a selection list instead.
$LELocation_Button.Add_Click({
  # Fetch a list of autodetected LE directories (as a list of strings)
  $Directories = $LEDirectories | Where { $_.Valid } | Select -ExpandProperty Path
  # Restore the array wrapping that PowerShell removes if an array only contains one element
  if (-Not($Directories -Is [Array])) {
    $Directories = @($Directories)
  }

  $OnlyShowFolderBrowser = $Directories.Count -Eq 0 -Or `
                          ($Directories.Count -Eq 1 -And (Equal-Paths $Directories[0] $LELocation_TextBox.Text))

  # Sets the default selection in the Selection List dialog window
  $DefaultSelection = $LELocation_TextBox.Text

  while ($True) {
    # Show-LELocationDialog
    if (-Not($OnlyShowFolderBrowser)) {
      $Directory = Show-LELocationDialog $Directories $DefaultSelection
      $DefaultSelection = $Null  # Do not use more than once.
                                 # The Dialog Window will remember the last selection used automatically.

      if ($Directory -Eq $Null)        { break }  # User cancelled
      elseif ($Directory -NE "browse") { break }  # User selected an LE directory
      # User chose "browse": The execution continues
    }

    # Show-FolderBrowserDialog
    while ($True) {
      $Directory = Show-FolderBrowserDialog -Description "Where is Locale Emulator?" `
                                            -initialDirectory $LELocation_TextBox.Text

      if ($Directory -Eq $Null)          { break }  # User cancelled
      if (Verify-LEDirectory $Directory) { break }  # User selected a valid LE directory
      # User selected an invalid directory: Display error message and try again
      $Title   = "Locale Emulator not found"
      $Message = "Locale Emulator was not found in the selected folder."
      $Buttons = "OK"
      $Icon    = "Error"
      [System.Windows.MessageBox]::Show($Message,$Title,$Buttons,$Icon)
    }

    # If only FolderBrowserDialog is used: always break
    # Else if user cancelled the FolderBrowserDialog: Return to LELocationDialog
    if ($OnlyShowFolderBrowser -Or $Directory) { break }
  }

  if (-Not($Directory)) { return }

  # Below code TRIGGERS: $LELocation_TextBox.Add_TextChanged()
  $LELocation_TextBox.NewValid = $True
  $LELocation_TextBox.Text     = $Directory
})


# Click this button to run the selected executable
$Run_Button.Add_Click({
  $LEProgram  = Join-Path $LELocation_TextBox.Text $Files.LE.Runner
  $LanguageID = $Lang_ComboBox.SelectedValue
  $Target     = $Target_TextBox.Targets[0]
  $TargetArgs = $Args_TextBox.Text

  Start-Process -FilePath $LEProgram `
                -ArgumentList "-runas `"$LanguageID`" `"$Target`" $TargetArgs"
})


# Click here to create a shortcut file to the executable
$Save_Button.Add_Click({
  $LEProgram      = Join-Path $LELocation_TextBox.Text $Files.LE.Runner
  $LanguageName   = $Lang_ComboBox.SelectedItem.Name
  $LanguageID     = $Lang_ComboBox.SelectedValue
  $TargetArgs     = $Args_TextBox.Text
  $DefaultSaveDir = $SaveDir_TextBox.Text

  if ($Target_TextBox.Targets.Count -Eq 1) {
    # One Target #
    # Let the user decide both the directory and the filename.
    $Target = $Target_TextBox.Targets[0]

    # Generate a filename as a suggestion to the user
    $InitialFilename = [IO.Path]::GetFileNameWithoutExtension($Target)
    $InitialFilename = "(LE)" + $InitialFilename

    $ShortcutFilePath = Show-SaveFileDialog -InitialFilename $InitialFilename `
                                            -InitialDirectory $DefaultSaveDir `
                                            -InitialDirectoryFallback [Environment]::GetFolderPath("Desktop") `
                                            -Filter "Shortcuts| *.lnk"

    if (-Not $ShortcutFilePath) {
      return
    }

    if (-Not ($ShortcutFilePath -Match '\.lnk$')) {
      $ShortcutFilePath += '.lnk'
    }

    Create-LEShortcutFile -Target    $Target `
                          -Args      $TargetArgs `
                          -LEPath    $LEProgram `
                          -LangID    $LanguageID `
                          -Path      $ShortcutFilePath
  } else {
    # Multiple Targets #
    # Only let the user decide the directory. Filenames are automatically generated.
    $Targets = $Target_TextBox.Targets

    $ShortcutDirectory = Show-FolderBrowserDialog -Description "Select the directory to where you want to save the shortcut files." `
                                                  -initialDirectory $DefaultSaveDir

    if (-Not $ShortcutDirectory) {
      return
    }

    ForEach ($Target in $Targets) {
      $ShortcutFilename = [IO.Path]::GetFileNameWithoutExtension($Target)
      $ShortcutFilename = "(LE)" + $ShortcutFilename + ".lnk"

      $ShortcutFilePath = Join-Path $ShortcutDirectory $ShortcutFilename

      Create-LEShortcutFile -Target    $Target `
                            -Args      $TargetArgs `
                            -LEPath    $LEProgram `
                            -LangName  $LanguageName `
                            -LangID    $LanguageID `
                            -Path      $ShortcutFilePath
    }
  }
})


# Escape key quits application
ForEach ($control in $controls) {
  $control.Add_KeyDown({
    if ($_.KeyCode -Eq "Escape") {
      Quit-Application
    }
  })
}

# Quit button quits application
$Quit_Button.Add_Click({
  Quit-Application
})
#ENDREGION GUI Events }



#REGION BEGIN GUI LELocations Events {
# Here there be event code for the LELocation window
#   For more info, check out: $LELocation_Button.Add_Click

$Form_LELocations.Add_Shown({
  $LELocations_ListBox.LastSelection = $LELocations_ListBox.SelectedValue
  $LELocations_ListBox.Focus()  # Give default focus to the ListBox
})

$LELocations_ListBox.Add_SelectedIndexChanged({
  $LELocations_ListBox.LastSelection = $LELocations_ListBox.SelectedValue
})

# Escape key clicks the Cancel button
ForEach ($lelocation_control in $lelocation_controls) {
  $lelocation_control.Add_KeyDown({
    if ($_.KeyCode -Eq "Escape") {
      $LELocations_ButtonCancel.PerformClick()
    }
  })
}

# Enter Key clicks the Select button (only if ListBox is in focus)
$LELocations_ListBox.Add_KeyDown({
  if ($_.KeyCode -Eq "Enter") {
    $LELocations_ButtonSelect.PerformClick()
  }
})

$LELocations_ButtonCancel.Add_Click({
  $Form_LELocations.ButtonAction = "Cancel"
  $Form_LELocations.Close()
})

$LELocations_ButtonBrowse.Add_Click({
  $Form_LELocations.ButtonAction = "Browse"
  $Form_LELocations.Close()
})

$LELocations_ButtonSelect.Add_Click({
  $Form_LELocations.ButtonAction = "Select"
  $Form_LELocations.Close()
})
#ENDREGION GUI LELocations Events }



#REGION BEGIN GUI Show {
# Here we finally initialize the main window and display it

$Config = Get-ConfigFile

if ($Config) {
  $SaveSettings_Checkbox.Checked = $True
  $LEDirectories = Get-LEDirectories -ConfigLEDirectory $Config.LELocation

  $dir_exists = try { Test-Path -LiteralPath $Config.DefaultSaveDirectory -PathType Container } catch { $False }
  if ($dir_exists) { $SaveDir_TextBox.text = $Config.DefaultSaveDirectory }
} else {
  $SaveSettings_Checkbox.Checked = $False
  $LEDirectories = Get-LEDirectories
}

if ($LEDirectories.Count -GT 0) {
  $LEDirectory = $LEDirectories[0]
  # Below code TRIGGERS: $LELocation_TextBox.Add_TextChanged()
  $LELocation_TextBox.NewValid = $LEDirectory.Valid
  $LELocation_TextBox.Text     = $LEDirectory.Path
}

if ($Config) {
  #NOTE: $Lang_ComboBox.DataSource was maybe populated by triggering $LELocation_TextBox.Add_TextChanged() above
  # If possible, change the selection to the same language as in the config
  $Language = $Lang_ComboBox.DataSource | Where { $_.Name -Eq $Config.Language } | Select -First 1
  if ($Language) { $Lang_ComboBox.SelectedItem = $Language }
}

# Show Application Window
#[void]$Form.ShowDialog()
[System.Windows.Forms.Application]::Run($Form)
#ENDREGION GUI Show }

#ENDREGION GUI }

