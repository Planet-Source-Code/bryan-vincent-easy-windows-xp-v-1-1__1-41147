VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Drastic Productions - Easy Windows XP"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command37 
      Caption         =   "Clear Quick Notes"
      Height          =   375
      Left            =   4200
      TabIndex        =   43
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   4200
      MaxLength       =   5000
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "easywindows.frx":0000
      Top             =   2160
      Width           =   5655
   End
   Begin VB.CommandButton Command34 
      Caption         =   "About"
      Height          =   375
      Left            =   2160
      TabIndex        =   37
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Mouse Settings"
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Clear Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Log Off"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Restart"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Task Manager"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Lockup (5 Min)"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Lockup (15 Min)"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Internet Explorer"
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Local Disk"
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command24 
      Caption         =   "System Serial Number"
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Command Prompt"
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Windows Explorer"
      Height          =   375
      Left            =   7920
      TabIndex        =   24
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Disk Defrag"
      Height          =   375
      Left            =   6000
      TabIndex        =   23
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Scan Disk"
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Sound Recorder"
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Notepad"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Paint"
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Media Player"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Calculator"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Shutdown"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command13 
      Caption         =   "System Settings"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Sound Settings"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Regional Settings"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Password Settings"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Network Settings"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Multimedia Settings"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Modem Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Keyboard Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Controller Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Time/Date Settings "
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add/Remove Programs"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New Hardware"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Display Settings"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   0
      TabIndex        =   19
      Top             =   4680
      Width           =   5895
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   2055
         Left            =   5880
         TabIndex        =   20
         Top             =   120
         Width           =   15
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Programs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   5880
      TabIndex        =   21
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Windows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4695
      Left            =   0
      TabIndex        =   36
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   2040
      TabIndex        =   38
      Top             =   2040
      Width           =   2055
      Begin VB.CommandButton Command36 
         Caption         =   "Easy Shortcuts"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Local Registry"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   2520
      Picture         =   "easywindows.frx":000E
      Top             =   360
      Width           =   6990
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "By: drasticterror---------------------------V.1.1"
      BeginProperty Font 
         Name            =   "Alleycat ICG"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   39
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Menu windows 
      Caption         =   "&Windows"
      Begin VB.Menu addnewhardware 
         Caption         =   "&Add New Hardware"
      End
      Begin VB.Menu addremoveprograms 
         Caption         =   "&Add/Remove Programs"
      End
      Begin VB.Menu clearclipboard 
         Caption         =   "&Clear Clipboard"
         Shortcut        =   +^{F8}
      End
      Begin VB.Menu lockmouse 
         Caption         =   "&Lockup (5 Min)"
      End
      Begin VB.Menu lockkeyboard 
         Caption         =   "&Lockup (15 Min)"
      End
      Begin VB.Menu logoff 
         Caption         =   "&Log Off"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu restart 
         Caption         =   "&Restart"
         Shortcut        =   +^{F11}
      End
      Begin VB.Menu shutdown 
         Caption         =   "&Shutdown"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu closeapplications 
         Caption         =   "&Task Manager"
      End
   End
   Begin VB.Menu information 
      Caption         =   "&Information"
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu easyshortcuts 
         Caption         =   "&Easy Shortcuts"
      End
      Begin VB.Menu localdrive 
         Caption         =   "&Local Disk"
      End
      Begin VB.Menu localregistry 
         Caption         =   "&Local Registry"
      End
      Begin VB.Menu systemserialnumber 
         Caption         =   "&System Serial Number"
      End
   End
   Begin VB.Menu settings 
      Caption         =   "&Settings"
      Begin VB.Menu controlsettings 
         Caption         =   "&Controller Settings"
      End
      Begin VB.Menu displaysettings 
         Caption         =   "&Display Settings"
      End
      Begin VB.Menu keysettings 
         Caption         =   "&Keyboard Settings"
      End
      Begin VB.Menu modsettings 
         Caption         =   "&Modem Settings"
      End
      Begin VB.Menu resolution 
         Caption         =   "&Mouse Settings"
      End
      Begin VB.Menu multimedsettings 
         Caption         =   "&Multimedia Settings"
      End
      Begin VB.Menu networksettings 
         Caption         =   "&Network Settings"
      End
      Begin VB.Menu passsettings 
         Caption         =   "&Password Settings"
      End
      Begin VB.Menu regsettings 
         Caption         =   "&Regional Settings"
      End
      Begin VB.Menu soundsettings 
         Caption         =   "&Sound Settings"
      End
      Begin VB.Menu syssettings 
         Caption         =   "&System Settings"
         Shortcut        =   +^{F9}
      End
      Begin VB.Menu timedatesett 
         Caption         =   "&Time/Date Settings"
      End
   End
   Begin VB.Menu programs 
      Caption         =   "&Programs"
      Begin VB.Menu calculator 
         Caption         =   "&Calculatior"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu commandprompt 
         Caption         =   "&Command Prompt"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu diskdefrag 
         Caption         =   "&Disk Defrag"
      End
      Begin VB.Menu internetexplore 
         Caption         =   "&Internet Explorer"
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu mediaplayer 
         Caption         =   "&Media Player"
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu notepad 
         Caption         =   "&Notepad"
         Shortcut        =   +^{F5}
      End
      Begin VB.Menu paint 
         Caption         =   "&Paint"
         Shortcut        =   +^{F6}
      End
      Begin VB.Menu scandisk 
         Caption         =   "&Scan Disk"
      End
      Begin VB.Menu soundrecorder 
         Caption         =   "&Sound Recorder"
      End
      Begin VB.Menu winexplore 
         Caption         =   "&Windows Explorer"
         Shortcut        =   +^{F7}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub about_Click()
MsgBox "This was made by drasticterror as an easier way to access multiple windows, if you get errors, or something doesnt work, the most common cause is the wrong OS, or your local disk is named something other than C, but most of it should work in WinXP."
End Sub

Private Sub addnewhardware_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub addremoveprograms_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub calculator_Click()
Shell "calc", vbNormalFocus
End Sub

Private Sub clearclipboard_Click()
Clipboard.Clear
End Sub

Private Sub closeapplications_Click()
Shell "C:\WINDOWS\system32\taskmgr.exe", vbNormalFocus
End Sub

Private Sub Command1_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub Command10_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Sub

Private Sub Command11_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub

Private Sub Command12_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub

Private Sub Command13_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub

Private Sub Command14_Click()
Shell ("Shutdown -s")
End Sub

Private Sub Command15_Click()
Shell "calc", vbNormalFocus
End Sub

Private Sub Command16_Click()
Shell "C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus
End Sub

Private Sub Command17_Click()
Shell "C:\WINDOWS\System32\mspaint.exe", vbNormalFocus
End Sub

Private Sub Command18_Click()
Shell "C:\WINDOWS\system32\notepad.exe", vbNormalFocus
End Sub

Private Sub Command19_Click()
Shell "sndrec32", vbNormalFocus
End Sub

Private Sub Command2_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub Command20_Click()
Shell "C:\WINDOWS\System32\cleanmgr.exe", vbNormalFocus
End Sub

Private Sub Command21_Click()
Shell "C:\WINDOWS\System32\dfrg.msc"
End Sub

Private Sub Command22_Click()
Shell "explorer", vbNormalFocus
End Sub

Private Sub Command23_Click()
Shell "command.com", vbNormalFocus
End Sub

Private Sub Command24_Click()
Dim objs
Dim obj
Dim WMI
Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_BaseBoard")
For Each obj In objs
   MsgBox obj.SerialNumber
Next
End Sub

Private Sub Command25_Click()
driveletter$ = Left(App.Path, 1)
MsgBox "Your system is bieng run from disk drive " & driveletter, vbInformation '
End Sub

Private Sub Command26_Click()
Shell "RUNDLL32.EXE URL.DLL,FileProtocolHandler http://www.msn.com", vbNormalFocus
End Sub

Private Sub Command27_Click()
SECONDS_TO_WAIT = "900"
ORIGINAL_TIME = DateTime.Time
    
Do Until DateDiff("s", ORIGINAL_TIME, DateTime.Time, 0, 0) > _
    Val(SECONDS_TO_WAIT)
Loop
End Sub

Private Sub Command28_Click()
SECONDS_TO_WAIT = "300"
ORIGINAL_TIME = DateTime.Time
    
Do Until DateDiff("s", ORIGINAL_TIME, DateTime.Time, 0, 0) > _
    Val(SECONDS_TO_WAIT)
Loop
End Sub

Private Sub Command29_Click()
Shell "C:\WINDOWS\system32\taskmgr.exe", vbNormalFocus
End Sub

Private Sub Command3_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub Command30_Click()
Shell ("Shutdown -r")
End Sub

Private Sub Command31_Click()
Shell ("Shutdown -l")
End Sub

Private Sub Command32_Click()
Clipboard.Clear
End Sub

Private Sub Command33_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Sub

Private Sub Command34_Click()
MsgBox "This was made by drasticterror as an easier way to access multiple windows, if you get errors, or something doesnt work, the most common cause is the wrong OS, or your local disk is named something other than C, but most of it should work in WinXP."
End Sub

Private Sub Command35_Click()
Shell "regedit", vbNormalFocus
End Sub

Private Sub Command37_Click()
Text1.Text = "Quick Notes:"
End Sub

Private Sub Command36_Click()
Form2.Show
End Sub

Private Sub Command4_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub Command5_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL joy.cpl", 5)
End Sub

Private Sub Command6_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub

Private Sub Command7_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Sub

Private Sub Command8_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", 5)
End Sub


Private Sub Command9_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub


Private Sub Label3_Click()

End Sub

Private Sub commandprompt_Click()
Shell "command.com", vbNormalFocus
End Sub

Private Sub controlsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL joy.cpl", 5)
End Sub

Private Sub diskdefrag_Click()
Shell "C:\WINDOWS\System32\dfrg.msc"
End Sub

Private Sub displaysettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub easyshortcuts_Click()
Form2.Show
End Sub

Private Sub internetexplore_Click()
Shell "RUNDLL32.EXE URL.DLL,FileProtocolHandler http://www.msn.com", vbNormalFocus
End Sub

Private Sub keysettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub

Private Sub localdrive_Click()
driveletter$ = Left(App.Path, 1)
MsgBox "Your system is bieng run from disk drive " & driveletter, vbInformation '
End Sub

Private Sub localregistry_Click()
Shell "regedit", vbNormalFocus
End Sub

Private Sub lockkeyboard_Click()
SECONDS_TO_WAIT = "900"
ORIGINAL_TIME = DateTime.Time
    
Do Until DateDiff("s", ORIGINAL_TIME, DateTime.Time, 0, 0) > _
    Val(SECONDS_TO_WAIT)
Loop
End Sub

Private Sub lockmouse_Click()
SECONDS_TO_WAIT = "300"
ORIGINAL_TIME = DateTime.Time
    
Do Until DateDiff("s", ORIGINAL_TIME, DateTime.Time, 0, 0) > _
    Val(SECONDS_TO_WAIT)
Loop
End Sub

Private Sub logoff_Click()
Shell ("Shutdown -l")
End Sub

Private Sub mediaplayer_Click()
Shell "C:\Program Files\Windows Media Player\wmplayer.exe", vbNormalFocus
End Sub

Private Sub modsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Sub

Private Sub multimedsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", 5)
End Sub

Private Sub networksettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub

Private Sub notepad_Click()
Shell "C:\WINDOWS\system32\notepad.exe", vbNormalFocus
End Sub

Private Sub paint_Click()
Shell "C:\WINDOWS\System32\mspaint.exe", vbNormalFocus
End Sub

Private Sub passsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Sub

Private Sub regsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub

Private Sub resolution_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Sub

Private Sub restart_Click()
Shell ("Shutdown -r")
End Sub

Private Sub scandisk_Click()
Shell "C:\WINDOWS\System32\cleanmgr.exe", vbNormalFocus
End Sub

Private Sub shutdown_Click()
Shell ("Shutdown -s")
End Sub

Private Sub soundrecorder_Click()
Shell "sndrec32", vbNormalFocus
End Sub

Private Sub soundsettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub

Private Sub syssettings_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub

Private Sub systemserialnumber_Click()
Dim objs
Dim obj
Dim WMI

Set WMI = GetObject("WinMgmts:")
Set objs = WMI.InstancesOf("Win32_BaseBoard")
For Each obj In objs
   MsgBox obj.SerialNumber
Next
End Sub

Private Sub timedatesett_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub volume_Click()
Shell "sndvol32", vbNormalFocus
End Sub

Private Sub winexplore_Click()
Shell "explorer", vbNormalFocus
End Sub


