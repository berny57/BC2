#$DebugPreference='SilentlyContinue'

$log="$env:TEMP\Backcolor.log"
Start-Transcript $log -Append

'----'
$Host.name
$Title = 'BackColour'
$Version = '2.0'
$Link = 'https://github.com/berny57/BC2'
$Icon = "$PSScriptRoot\TintedWindow.ico"
$AutoStart = $False
$Script=$PSCommandPath




"$Title v$Version : $Link"

#region Build $EventAction code block

$EventAction =
{
#region P/invoke code to access Win32 API
# *** Do not indent this code! ***
$Code=@'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Desktop
{
    public class ControlPanel
    {
        //Declare constants
        private const int COLOR_WINDOW = 5;
        
        //Delare functions in unmanaged APIs
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);
        
        //Helper functions
        public static void SetWindowColor(byte r, byte g, byte b)
        {
            System.Drawing.Color color = System.Drawing.Color.FromArgb(r,g,b);
            int[] elements = { COLOR_WINDOW };
            int[] colors = { System.Drawing.ColorTranslator.ToWin32(color) };
            SetSysColors (elements.Length, elements, colors);
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Colors", true);
            key.SetValue(@"Window", string.Format("{0} {1} {2}", r, g, b));
        }
    }
}
'@

#Enable p/invoke code
$Type = Add-Type -TypeDefinition $Code -ReferencedAssemblies System.Drawing.dll -PassThru

#endregion P/invoke code to access Win32 API

Function Set-OSCWindowColor
{
    <#
        .SYNOPSIS
            Set-OSCWindowColor can be used to change the desktop background colour
        .DESCRIPTION
            Set-OSCWindowColor is an Advanced Funcion that can be used to change the background colour of the desktop
        .PARAMETER R
            Parameters R, G & B indicate the three colours (red, green, blue) in the RGB colour model.
            You can specify a value in the 0-255 range for each. If you ignore these parameters the default backgound 
            colour will be set. 
        .Example
            Set-OSCWindowsColor -R 240 -G 240 -B255
            #Change the desktop background colour to pale blue.
        .LINK
            Windows PowerShell Advanced Function
            http://technet.microsoft.com/en-us/library/dd315326.aspx
        .LINK
            SetSysColors function
            http://msdn.microsoft.com/en-us/library/windows/desktop/ms724940(v=vs.85).aspx
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $R,
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $G,
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $B
    )
    Process
    {
        #Call the P/Invoke defined helper function
        [Desktop.ControlPanel]::SetWindowColor($R,$G,$B)
        Return
    }
}


Function Get-cpColor()
{
    <#
        .SYNOPSIS
            Get-cpColor returns the value of the specified color key withing HKCU\Control Panel\Colors.
        .DESCRIPTION
            Get-cpColor returns the RGB value of the specified color key withing HKCU\Control Panel\Colors
            as a byte array. The RGB values are returned in elements [0],[1] and [2] respectivly.
        .PARAMETER Key
            Specifies the registry key to be interogated
        .EXAMPLE
            Get-cpColor('ButtonFace')
            #Returns the ButtonFace colour as a RGB byte array
        .LINK
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String] $Key
    )
    Process
    {
        #Get the string value from registy
        $Value=(Get-ItemPropertyValue -Path 'HKCU:\Control Panel\Colors' -Name $Key)

        #Convert to byte array
        [byte[]]$ColorRgb = $Value -split '\s'

        Return $ColorRgb
    }
}


#region Set Color specified in Control Panel
$Color = Get-cpColor 'Window'
$Red = $Color[0]
$Green = $Color[1]
$Blue = $Color[2]
Set-OSCWindowColor -R $Blue -G $Red -B $Green  ######  For test purposes these are rotated so the colour changes on each Event activation!
#endregion

#logging
"Colour applied by event handler"  #|out-file $Log -Append
}
#endregion Build $EventAction code block

#$EventAction|Out-File -FilePath "$PSScriptRoot\EventAction.ps1" -Width 300

#logging
"EventAction created"#|out-file $Log -Append

 
#region Event handler to detect SessionSwitch events

#Register the event handler. When activated it calls the above $EventAction codeblock
$Null = Register-ObjectEvent -InputObject ([Microsoft.Win32.SystemEvents]) -EventName "SessionSwitch" -Action $EventAction 

#logging
Get-EventSubscriber | Where-Object {$_.SourceObject -eq [Microsoft.Win32.SystemEvents]}#|out-file $Log -Append
$Job=$Events | Select-Object -ExpandProperty Action 
$Job#|out-file $Log -Append

#endregion Event handler to detect SessionSwitch events

#logging
"Eventhandler registered"#|out-file $Log -Append

#region Display form in it's own RunSpace

# A separate thread is required so event processing keeps running while the form is displayed
# 
# You could use Start-Process to display the form in a new PowerShell instance but
# a new PowerShell Window will be displayed. It briefly flashes up even if the window is specified as hidden
#
# A RunSpace is more efficient and avoids this issue

#region Create the RunSpace
$RunSpace = [runspacefactory]::CreateRunspace()
$RunSpace.ApartmentState = "STA"
$RunSpace.ThreadOptions = "ReuseThread"
$RunSpace.Open()

# Share variables with the new RunSpace
# Include the $EventAction code block so it can also be used within the form.
$RunSpace.SessionStateProxy.SetVariable("Title", $Title)
$RunSpace.SessionStateProxy.SetVariable("Version", $Version)
$RunSpace.SessionStateProxy.SetVariable("Link", $Link)
$RunSpace.SessionStateProxy.SetVariable("AutoStart", $AutoStart)
$RunSpace.SessionStateProxy.SetVariable("Script", $Script)
$RunSpace.SessionStateProxy.SetVariable("Icon", $Icon)
$RunSpace.SessionStateProxy.SetVariable("log", $log)
#endregion Create the RunSpace

#logging
"Runspace created"#|out-file $Log -Append

#  Create the $FormCcode code block. This will be passed to the RunSpace for execution
$FormCode = {



#region P/invoke code to access Win32 API
$Code=@'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Desktop
{
    public class ControlPanel
    {
        //Declare constants
        private const int COLOR_WINDOW = 5;
        
        //Delare functions in unmanaged APIs
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);
        
        //Helper functions
        public static void SetWindowColor(byte r, byte g, byte b)
        {
            System.Drawing.Color color = System.Drawing.Color.FromArgb(r,g,b);
            int[] elements = { COLOR_WINDOW };
            int[] colors = { System.Drawing.ColorTranslator.ToWin32(color) };
            SetSysColors (elements.Length, elements, colors);
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Colors", true);
            key.SetValue(@"Window", string.Format("{0} {1} {2}", r, g, b));
        }
    }
}
'@

    #Enable p/invoke code
    $Type = Add-Type -TypeDefinition $Code -ReferencedAssemblies System.Drawing.dll -PassThru

    #endregion P/invoke code to access Win32 API

Function Set-OSCWindowColor
{
    <#
        .SYNOPSIS
            Set-OSCWindowColor can be used to change the desktop background colour
        .DESCRIPTION
            Set-OSCWindowColor is an Advanced Funcion that can be used to change the background colour of the desktop
        .PARAMETER R
            Parameters R, G & B indicate the three colours (red, green, blue) in the RGB colour model.
            You can specify a value in the 0-255 range for each. If you ignore these parameters the default backgound 
            colour will be set. 
        .Example
            Set-OSCWindowsColor -R 240 -G 240 -B255
            #Change the desktop background colour to pale blue.
        .LINK
            Windows PowerShell Advanced Function
            http://technet.microsoft.com/en-us/library/dd315326.aspx
        .LINK
            SetSysColors function
            http://msdn.microsoft.com/en-us/library/windows/desktop/ms724940(v=vs.85).aspx
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $R,
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $G,
        [Parameter(Position=0, Mandatory=$True)]
        [ValidateRange(0,255)]
        [int] $B
    )
    Process
    {
        #Call the P/Invok defined helper function
        [Desktop.ControlPanel]::SetWindowColor($R,$G,$B)
        Return
    }
}


Function Get-cpColor()
{
    <#
        .SYNOPSIS
            Get-cpColor returns the value of the specified color key withing HKCU\Control Panel\Colors.
        .DESCRIPTION
            Get-cpColor returns the RGB value of the specified color key withing HKCU\Control Panel\Colors
            as a byte array. The RGB values are returned in elements [0],[1] and [2] respectivly.
        .PARAMETER Key
            Specifies the registry key to be interogated
        .EXAMPLE
            Get-cpColor('ButtonFace')
            #Returns the ButtonFace colour as a RGB byte array
        .LINK
    #>
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [String] $Key
    )
    Process
    {
        #Get the string value from registy
        $Value=(Get-ItemPropertyValue -Path 'HKCU:\Control Panel\Colors' -Name $Key)

        #Convert to byte array
        [byte[]]$ColorRgb = $Value -split '\s'

        Return $ColorRgb
    }
}


    #region Set Color specified in Control Panel - this will set the initial color when the form runs
    $ColorRGB = Get-cpColor 'Window'
    Set-OSCWindowColor -R $ColorRGB[0] -G $ColorRGB[1] -B $ColorRGB[2]  
    #logging
    "Colour sduring form load"#|out-file $Log -Append


    #endregion


Function PickColor()
{
    #Display system color picker
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $ColorDialog = New-Object System.Windows.Forms.ColorDialog
    $ColorRGB = Get-cpColor 'Window'  #RGB byte array
        

    $Color = (($ColorRGB[0] * 256 * 256) + ($ColorRGB[1] * 256) + $ColorRGB[2])
    $ColorDialog.Color = $Color
    $ColorDialog.AllowFullOpen = $True
    $ColorDialog.FullOpen = $True

    if ($ColorDialog.ShowDialog() -like 'ok')
    {
        Set-OSCWindowColor -R $ColorDialog.Color.R -G $ColorDialog.Color.G -B $ColorDialog.Color.B # You can read the colordialog RGB
        #logging
        "Colour set by colour picker"#|out-file $Log -Append
    }

}


Function Check-AutoStart($LinkFile, $LinkTarget, $LinkArg)
{
    If (Test-path -Path $LinkFile -PathType leaf)
    {
        # Check shortcut content
        $obj = New-Object -ComObject WScript.Shell
        $shortcut = $obj.CreateShortcut($LinkFile)
        If (($ShortCut.TargetPath -like $LinkTarget) -And ($ShortCut.Arguments -like $LinkArg))
        {
            Return $True
        }
    }
    Return $False
}



Function UpdateStartup($AutoStart, $LinkFile, $LinkTarget, $LinkArg)
{
    $obj = New-Object -ComObject WScript.Shell
    $shortcut = $obj.CreateShortcut($LinkFile)

    if ($AutoStart)  
    {
        #User activated AutoStart, create / update shortcut
        $ShortCut.TargetPath = "$LinkTarget"
        $ShortCut.Arguments = "$LinkArg"
        $shortcut.WorkingDirectory="$env:HOMEDRIVE$env:HOMEPATH"
        $ShortCut.IconLocation = $Icon
        $ShortCut.Description = "Created by $Title v$Version";
        $ShortCut.Save()        
    }
    else
    {
       # User de-activated AutoStart, delete the shortcut
       # SilentContinue specified in case the delete fails
       Get-ChildItem -Path "$LinkFile" -ErrorAction Continue|Remove-Item  
    }
    # Return new checkmark state depending on success
    Return (Check-AutoStart $LinkFile $LinkTarget $LinkArg)
}



    $Startup = [environment]::getfolderpath("Startup")

    $LinkFile = "$Startup\$Title.lnk"
    $LinkTarget = "$PSHome\PowerShell.exe"
    $LinkArg = "-Command $Script "  #$PSCommandPath didn't resolve within the runspace

    $AutoStart = Check-AutoStart $LinkFile $LinkTarget $LinkArg


    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    #Create the form and give it a system tray icon and context menu
    $Form = New-Object System.Windows.Forms.Form
    $NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $ContextMenu = New-Object System.Windows.Forms.ContextMenu

    # Create menu items 
    $TitleMenuItem = New-Object System.Windows.Forms.MenuItem
    $PickMenuItem = New-Object System.Windows.Forms.MenuItem
    $AutoStartMenuItem = New-Object System.Windows.Forms.MenuItem
    $ExitMenuItem = New-Object System.Windows.Forms.MenuItem

    # Attach Icons, context menu, etc. to tray icon
    $NotifyIcon.Icon = New-Object System.Drawing.Icon($Icon)
    $NotifyIcon.ContextMenu = $ContextMenu
    
    # Add menu items to context menu
    $NotifyIcon.ContextMenu.MenuItems.AddRange($TitleMenuItem)
    $NotifyIcon.ContextMenu.MenuItems.AddRange($PickMenuItem)
    $NotifyIcon.ContextMenu.MenuItems.AddRange($AutoStartMenuItem)
    $NotifyIcon.ContextMenu.MenuItems.AddRange($ExitMenuItem)

    # Hide the primary form 
    $Form.Icon = New-Object System.Drawing.Icon($Icon) 
    $Form.ShowInTaskbar = $False #$True #for debuggin, $False in production.
    $Form.WindowState = "Minimized"
 
    # Attach actions to menu items
    $TitleMenuItem.Text = "$Title v$Version"
    $TitleMenuItem.add_Click({Start-Process "$Link"})

    $PickMenuItem.text = "Pick Colour"
    $PickMenuItem.add_Click({PickColor})

    $AutoStartMenuItem.text = "Auto Start"
    $AutoStartMenuItem.Checked = Check-AutoStart $LinkFile $LinkTarget $LinkArg
    $AutoStartMenuItem.add_Click({
        # Toggle check mark 
        $AutoStartMenuItem.Checked = UpdateStartup (!($AutoStartMenuItem.Checked)) $LinkFile $LinkTarget $LinkArg
    })

    $ExitMenuItem.text = "Exit"
    $ExitMenuItem.add_Click({
        $NotifyIcon.Visible = $False
        $form.Close()
    })

    # Display the system tray item
    $NotifyIcon.Visible = $True

    #And execute the form
    [void][System.Windows.Forms.Application]::Run($Form)

    #logging
    "Form created"#|out-file $Log -Append
}




# Invoke the code in a new PowerShell thread using the defined RunSpace
$PSInstance = [powershell]::Create().AddScript($FormCode)
$PSInstance.Runspace = $RunSpace
$Job = $PSInstance.BeginInvoke()


do {} until ($Job.IsCompleted)

#endregion Display form in it's own RunSpace 


# For debugging, save $FormCode to a script file.  It can then be executed and debugged in the ISE
#$FormCode|out-file $PSScriptRoot\FormCode.ps1

#logging
"FormAction created"#|out-file $Log -Append


#region Delete the event subscription
$Events = Get-EventSubscriber | Where-Object {$_.SourceObject -eq [Microsoft.Win32.SystemEvents]}
$Jobs = $Events | Select-Object -ExpandProperty Action

$Events|Unregister-Event
$Jobs|Remove-Job
#endregion Delete the event subscription

#logging
"Event handler deleted"#|out-file $Log -Append

'-----'
Stop-Transcript
