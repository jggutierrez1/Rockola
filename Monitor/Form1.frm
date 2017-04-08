VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Monitor&Commander [Rockola]"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   3000
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esperando Secuencia:.."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If App.PrevInstance Then
    End
End If
End Sub


Private Sub Timer1_Timer()
Dim iValue1 As Integer
Dim iValue2 As Integer
Dim iValue3 As Integer
iValue1 = VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "RELOAD_APP", "0"))
If iValue1 = 1 Then
    Me.Label1.Caption = "Reiniciando Sistemas:.."
'-----------------------------------------------------------------------------------------
    iValue2 = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "SWITCH_KAR", "0")))
    If iValue2 = 1 Then
'       Va del modo normal a karaoke.
        Me.Label1.Caption = "Configurando Créditos grátis:.."
        iValue3 = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "KEEP_SCRED", "0")))
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "BK_KEEPCRE", VBA.Trim(VBA.Str(iValue3)))
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "KEEP_SCRED", "1")
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "SAVED_CRED", "6")
    Else
'       Va del modo karaoke a normal.
        Me.Label1.Caption = "Configurando sistema en modo standard:.."
        
        iValue3 = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "BK_KEEPCRE", "0")))
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "KEEP_SCRED", VBA.Trim(VBA.Str(iValue3)))
        
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "LOAD_SCRED", "1")
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "SAVED_CRED", "0")
    End If
'-----------------------------------------------------------------------------------------
    Call Write_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "RELOAD_APP", "0")
    Call VBA.Shell(App.Path & "\Rockola.exe", vbNormalFocus)
    End
End If
End Sub


