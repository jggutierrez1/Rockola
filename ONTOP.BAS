Attribute VB_Name = "OnTop"
Option Explicit

#If Win16 Then
    Declare Sub SetWindowPos Lib "User" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#Else
    Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If


'*************************************************************************
'* Function: KeepOnTop(F As Form)
'*
'*
'*************************************************************************
'* Description: Keep form on top.
'*
'*
'*************************************************************************
'* Parameters: Form Control
'*
'*************************************************************************
'* Notes: The SetWindowPos API call gets turned off if the form is
'*        minimized.  Put this code in the resize event to make sure
'*        your form stays on top.
'*
'*************************************************************************
'* Returns:
'*************************************************************************
Sub AlwaysOnTop(FrmID As Form, Optional OnTop As Boolean = True)
    ' This function uses an argument to dete
    '     rmine whether
    ' to make the specified form always on t
    '     op or not
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2


    If OnTop Then
        Call SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    Else
        Call SetWindowPos(FrmID.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    End If
End Sub

