Attribute VB_Name = "Module1"
Option Explicit

      Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
      End Type

    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
    Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
    Private Const KEYEVENTF_EXTENDEDKEY = &H1
    Private Const KEYEVENTF_KEYUP = &H2
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private bNoClick As Boolean


Public Sub Chek_NumLockStatus()
'pon si quieres un intervalo al cronometro de 200...lo puedes disminuir si quieres...
 Dim o As OSVERSIONINFO
 Dim NumLockState As Boolean
 Dim keys(0 To 255) As Byte
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    GetKeyboardState keys(0)

  'Testeo el estado de la tecla Bloq Num...si esta apagada la enciendo...

    If GetKeyState(vbKeyNumlock) = False Then
        NumLockState = keys(vbKeyNumlock)
        If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then  'Para Win95/98/ME
            keys(vbKeyNumlock) = Abs(Not NumLockState)
            SetKeyboardState keys(0)
        ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then   'Para Windows NT, XP, 2000
            keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub


