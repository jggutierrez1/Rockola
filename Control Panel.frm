VERSION 5.00
Begin VB.Form ControlPanel 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mi Sistema"
   ClientHeight    =   6240
   ClientLeft      =   2385
   ClientTop       =   1830
   ClientWidth     =   5460
   Icon            =   "Control Panel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6240
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   5880
   End
   Begin VB.CommandButton oBtn_Expl 
      BackColor       =   &H00FFFF00&
      Caption         =   "Windows Explorer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3090
      Width           =   2295
   End
   Begin VB.CommandButton Reg_value 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ver valor Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Reg_value 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Establecer ROCKOLA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton Reg_value 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Establecer (DEFAULT)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.PictureBox Picture1x 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   5115
      TabIndex        =   23
      Top             =   5640
      Width           =   5175
      Begin VB.DirListBox Dir1 
         Height          =   315
         Left            =   480
         TabIndex        =   37
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   6000
         Top             =   0
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   2940
         TabIndex        =   35
         Top             =   0
         Width           =   3000
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   3000
            TabIndex        =   36
            Top             =   0
            Width           =   3000
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   38
         Top             =   0
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdShow 
      BackColor       =   &H0000FF00&
      Caption         =   "Show Start Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
      Width           =   2295
   End
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H0000FF00&
      Caption         =   "Hide Start Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5040
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   2295
   End
   Begin VB.CommandButton cmdShutDown 
      BackColor       =   &H00FFFF00&
      Caption         =   "ShutDown/Restart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CommandButton cmdRun 
      BackColor       =   &H00FFFF00&
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSystemProperties 
      BackColor       =   &H00FFFF00&
      Caption         =   "Teclado en Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4080
      Width           =   2295
   End
   Begin VB.PictureBox Picture1y 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   5235
      TabIndex        =   34
      Top             =   0
      Width           =   5295
      Begin VB.ComboBox cboPanel 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.CommandButton Control 
      BackColor       =   &H00FFFF00&
      Caption         =   "Control &Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   2295
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   8
      Left            =   2400
      Picture         =   "Control Panel.frx":014A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   1530
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   7
      Left            =   1320
      Picture         =   "Control Panel.frx":058C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   1530
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      Height          =   495
      Index           =   4
      Left            =   3480
      Picture         =   "Control Panel.frx":09CE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   1320
      Picture         =   "Control Panel.frx":0E10
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   10
      Left            =   4560
      Picture         =   "Control Panel.frx":111A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   1530
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   3480
      Picture         =   "Control Panel.frx":155C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   1530
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   4560
      Picture         =   "Control Panel.frx":1866
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   6
      Left            =   240
      Picture         =   "Control Panel.frx":1CB0
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   360
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   2400
      Picture         =   "Control Panel.frx":1E46
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.PictureBox Icon 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   240
      Picture         =   "Control Panel.frx":2288
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 Restante"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3960
      TabIndex        =   40
      Top             =   6000
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Actual del SHELL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   39
      Top             =   2880
      Width           =   2070
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      Height          =   195
      Index           =   14
      Left            =   2385
      TabIndex        =   33
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sounds"
      Height          =   195
      Index           =   13
      Left            =   1290
      TabIndex        =   32
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Regional Settings"
      Height          =   195
      Index           =   12
      Left            =   3030
      TabIndex        =   31
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Passwords"
      Height          =   195
      Index           =   11
      Left            =   1170
      TabIndex        =   30
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multimedia"
      Height          =   195
      Index           =   9
      Left            =   4425
      TabIndex        =   29
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mouse"
      Height          =   195
      Index           =   8
      Left            =   3480
      TabIndex        =   28
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard"
      Height          =   195
      Index           =   6
      Left            =   4455
      TabIndex        =   27
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notepad"
      Height          =   195
      Index           =   4
      Left            =   165
      TabIndex        =   26
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
      Height          =   195
      Index           =   3
      Left            =   2325
      TabIndex        =   25
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Iconname 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date/Time"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   1080
      Width           =   795
   End
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TimeResto As Long

'Name: Ali ezzahir
'Winnipeg , Manitoba, Canada
'ezzahir@yahoo.com
'http://www.geocities/athens/aegean/6647
'http://www.geocities/athens/troy/3164
Private Declare Function SHShutDownDialog Lib "Shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SHRestartSystem Lib "Shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHRunDialog Lib "Shell32" Alias "#61" (ByVal hOwner As Long, ByVal hIcon As Long, ByVal sDir As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHFormatDrive Lib "Shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private mhOwner As Long
Private mDialogPrompt As String
Private mDialogTitle As String
Private mCancelError As Boolean
Private mhIcon As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
Alias "GetDiskFreeSpaceExA" _
(ByVal lpRootPathName As String, _
lpFreeBytesAvailableToCaller As Currency, _
lpTotalNumberOfBytes As Currency, _
lpTotalNumberOfFreeBytes As Currency) As Long

Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Dim TNB As Double
Dim TFB As Double
Dim FreeBytes As Long
Dim DriveLetter As String
Dim DLetter As String
Dim spaceInt As Integer
Dim ID
Dim uFlag As Long
Dim mFlags As Long

Const KEY_ALL_ACCESS = &H2003F

Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                    
                Else
                        GoTo SysInfoErr
                End If

        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long
        Dim rc As Long
        Dim hKey As Long
        Dim hDepth As Long
        Dim KeyValType As Long
        Dim tmpVal As String
        Dim KeyValSize As Long
       
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        

        tmpVal = String$(1024, 0)
        KeyValSize = 1024
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
                tmpVal = Left(tmpVal, KeyValSize - 1)
        Else
                tmpVal = Left(tmpVal, KeyValSize)
        End If
        
        Select Case KeyValType
        Case REG_SZ
                KeyVal = tmpVal
        Case REG_DWORD
                For i = Len(tmpVal) To 1 Step -1
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
                Next
                KeyVal = Format$("&h" + KeyVal)
        End Select
        

        GetKeyValue = True
        rc = RegCloseKey(hKey)
        Exit Function
        

GetKeyError:
        KeyVal = ""
        GetKeyValue = False
        rc = RegCloseKey(hKey)
End Function

Private Sub cboPanel_Click()
On Error GoTo errorhandler
Select Case cboPanel.ListIndex
Case 0
Call StartSysInfo
Case 1
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", vbNormalFocus)
Case 2
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,2", vbNormalFocus)
Case 3
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,3", vbNormalFocus)
Case 4
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", vbNormalFocus)
Case 5
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,1", vbNormalFocus)
Case 6
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,2", vbNormalFocus)
Case 7
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,3", vbNormalFocus)
Case 8
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", vbNormalFocus)
Case 9
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,1", vbNormalFocus)
Case 10
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,2", vbNormalFocus)
Case 11
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,3", vbNormalFocus)
Case 12
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,4", vbNormalFocus)
Case 13
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL joy.cpl", vbNormalFocus)
Case 14
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", vbNormalFocus)
Case 15
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", vbNormalFocus)
Case 16
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL main.cpl @2", vbNormalFocus)
Case 17
ID = Shell(" rundll32.exe shell32.dll,Control_RunDLL main.cpl @3", vbNormalFocus)
Case 18
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mlcfg32.cpl", vbNormalFocus)
Case 19
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0", vbNormalFocus)
Case 20
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,1", vbNormalFocus)
Case 21
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,2", vbNormalFocus)
Case 22
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,3", vbNormalFocus)
Case 23
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,4", vbNormalFocus)
Case 24
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", vbNormalFocus)
Case 25
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", vbNormalFocus)
Case 26
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", vbNormalFocus)
Case 27
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", vbNormalFocus)
Case 28
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", vbNormalFocus)
Case 29
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,1", vbNormalFocus)
Case 30
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,2", vbNormalFocus)
Case 31
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,3", vbNormalFocus)
Case 32
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", vbNormalFocus)
Case 33
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus)
End Select
Exit Sub
errorhandler:
MsgBox "The selected function is not available on this system", vbCritical, "Not available"
End Sub

Private Sub cmdShow_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
cmdHide.Enabled = True
cmdShow.Enabled = False
End Sub

Private Sub cmdHide_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar
cmdShow.Enabled = True
cmdHide.Enabled = False
End Sub

Private Sub Control_Click()
Dim sWinDir As String
sWinDir = GetWindowsDir
If Right(sWinDir, 1) <> "\" Then sWinDir = sWinDir & "\"
    'Call ControlPanels(sWinDir & "Control.Exe")
    Call ControlPanels("Control.Exe")
End Sub

Private Sub Form_Load()
TimeResto = 40
Call sCenterForm(Me)
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
cmdShow.Enabled = False
GetDiskInfo
cboPanel.AddItem "System Information"
cboPanel.AddItem "Add/Remove Programs Properties (Install/Uninstall)"
cboPanel.AddItem "Add/Remove Programs Properties (Windows Setup)"
cboPanel.AddItem "Add/Remove Programs Properties (Startup Disk)"
cboPanel.AddItem "Display Properties (Background)"
cboPanel.AddItem "Display Properties (Screen Saver)"
cboPanel.AddItem "Display Properties (Appearance)"
cboPanel.AddItem "Display Properties (Settings)"
cboPanel.AddItem "Regional Settings Properties (Regional Settings)"
cboPanel.AddItem "Regional Settings Properties (Number)"
cboPanel.AddItem "Regional Settings Properties (Currency)"
cboPanel.AddItem "Regional Settings Properties (Time)"
cboPanel.AddItem "Regional Settings Properties (Date)"
cboPanel.AddItem "Joystick Properties (Joystick)"
cboPanel.AddItem "Mouse Properties"
cboPanel.AddItem "Keyboard Properties"
cboPanel.AddItem "Printers"
cboPanel.AddItem "Fonts"
cboPanel.AddItem "Microsoft Exchange Profiles"
cboPanel.AddItem "Multimedia Properties (Audio)"
cboPanel.AddItem "Multimedia Properties (Viedo)"
cboPanel.AddItem "Multimedia Properties (MIDI)"
cboPanel.AddItem "Multimedia Properties (CD Music)"
cboPanel.AddItem "Multimedia Properties (Advanced)"
cboPanel.AddItem "Sounds Properties"
cboPanel.AddItem "Modem Properties (General)"
cboPanel.AddItem "Network (Configuration)"
cboPanel.AddItem "Password Properties (Change Passwords)"
cboPanel.AddItem "System Properties (General)"
cboPanel.AddItem "System Properties (Device Manager)"
cboPanel.AddItem "System Properties (Hardware Profiles)"
cboPanel.AddItem "System Properties (Performance)"
cboPanel.AddItem "Add New Hardware Wizard"
cboPanel.AddItem "Date/Time Properties"
Call Reg_value_Click(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
End Sub

Private Sub Icon_Click(Index As Integer)
On Error GoTo errorhandler
If Index = 0 Then
ElseIf Index = 1 Then
ElseIf Index = 2 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
ElseIf Index = 3 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")
ElseIf Index = 4 Then
   'Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0")
   Call VBA.Shell("NOTEPAD.EXE", vbNormalFocus)
ElseIf Index = 5 Then
ElseIf Index = 6 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1")
ElseIf Index = 7 Then
ElseIf Index = 8 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0")
ElseIf Index = 9 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")
ElseIf Index = 10 Then
ElseIf Index = 11 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL password.cpl")
ElseIf Index = 12 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0")
ElseIf Index = 13 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1")
ElseIf Index = 14 Then
   Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0")
End If
Exit Sub
errorhandler:
MsgBox "The selected function is not available on this system", vbCritical, "Not available"
End Sub
Public Property Let CancelError(ByVal vData As Boolean)
   mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
  CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property

Public Property Get DialogTitle() As String
   DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
   mDialogTitle = sTitle
End Property

Public Function ShowShutDown()
   SHShutDownDialog mhOwner
End Function
Public Function ShowRun()
  
  uFlag = mFlags And (&H10 Or &H20 Or &H40 Or &H80)
  uFlag = uFlag / 16
  SHRunDialog mhOwner, mhIcon, 0, mDialogTitle, mDialogPrompt, uFlag
End Function
Public Function ShowFormat(Optional ByVal iDrive As Long, Optional ByVal iCapacity As Long, Optional ByVal iFormatType As Long) As Long
  ShowFormat = SHFormatDrive(mhOwner, iDrive, iCapacity, iFormatType)
End Function

Private Sub cmdRun_Click()
Dim uFlag As Long
SHRunDialog mhOwner, mhIcon, 0, mDialogTitle, mDialogPrompt, uFlag
End Sub

Private Sub cmdShutDown_Click()
SHShutDownDialog mhOwner
End Sub

Private Sub Command5_Click()
ID = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", vbNormalFocus)
End Sub

Private Sub cmdSystemProperties_Click()
Call VBA.Shell("osk.exe")
End Sub

Private Sub Drive1_Change()
Dim DriveLetter$, DriveNumber&, DriveType&
    DriveLetter = UCase(Drive1.Drive)
    DriveNumber = (Asc(DriveLetter) - 65)
    DriveType = GetDriveType(DriveLetter)
    On Error GoTo errHandler
     DriveLetter = Drive1.Drive & "\"
     GetDiskInfo
     Dir1.Path = Drive1.Drive
     Exit Sub
errHandler:
MsgBox "Device Unavailable!. Please check drive " & DriveLetter, vbCritical, error
     Drive1.Drive = Dir1.Path
GetDiskInfo
End Sub

Public Sub GetDiskInfo()

DriveLetter = Drive1.Drive

spaceInt = InStr(DriveLetter, " ")
If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)

If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
DLetter = Left(UCase(DriveLetter), 1)

    Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
TNB = TotalBytes * 10000
    TFB = (TotalBytes - TotalFreeBytes) * 10000
DiskInfo.lblNumOfBytes.Caption = " Capacity:  " & Format$(TotalBytes * 10000, "###,###,###,##0") & " bytes"
DiskInfo.lblFreeBytes.Caption = " Free Space:  " & Format$(BytesFreeToCalller * 10000, "###,###,###,##0") & " bytes"
DiskInfo.Label3.Caption = "Disk space used:  " & Format(TFB / TNB * 100, "###.#0") & " %"
DiskInfo.Label4.Caption = "Disk space available:  " & Format(100 - TFB / TNB * 100, "###.#0") & " %"
Label5.Caption = Format(100 - TFB / TNB * 100, "###.#0") & " % of free space"
Picture1.Width = Format(100 - TFB / TNB * 100, "###.#0") * 50
End Sub

Private Sub oBtn_Expl_Click()
Call VBA.Shell("Explorer.exe", vbNormalFocus)
End Sub

Private Sub Reg_value_Click(Index As Integer)
Dim sValue As String
Dim sFile As String
sValue = Trim(Text1.Text)
sValue = sFile
Select Case Index
Case Is = 1
    sFile = App.Path & "\" & App.EXEName & ".EXE"
    Call WriteRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Shell", ValString, "" & sFile)
    Text1.Text = sFile
Case Is = 2
    Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Shell")
Case Is = 3
    Call WriteRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Shell", ValString, "Explorer.exe")
    Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", "Shell")
End Select
Text1.Refresh
End Sub

Private Sub Timer1_Timer()
If Timer1.Enabled = True Then
Picture1.Width = Format(100 - TFB / TNB * 100, "###.#0") * 50
GetDiskInfo
End If
End Sub

Private Sub Timer2_Timer()
Dim Horas As Integer, Minutos As Integer, Segundos As Integer, Cadena As String
    
Segundos = TimeResto
Horas = Int(Segundos / 3600)
Segundos = Segundos Mod 3600
Minutos = Int(Segundos / 60)
Segundos = Segundos Mod 60
Cadena = Format$(Horas, "00") & ":" & Format$(Minutos, "00") & ":" & Format$(Segundos, "00") & " Restante"
lblTime.Caption = Cadena
TimeResto = TimeResto - 1
If TimeResto = -1 Then
    Unload Me
End If
End Sub
