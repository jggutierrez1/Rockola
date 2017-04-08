VERSION 5.00
Begin VB.Form Pass_Scr 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Aceeso de seguridad"
   ClientHeight    =   1125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "Pass_Scr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   6
      PasswordChar    =   "?"
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 Restante"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetos del Supervisor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "Pass_Scr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeResto As Long


Private Sub Form_Activate()
Me.Text1.SetFocus
End Sub

Private Sub Form_Load()
TimeResto = 40
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Len(Me.Text1.Text) = 6 Then
    Main_Form.Tag = VBA.Trim(Me.Text1.Text)
    Unload Me
End If
End Sub

Private Sub Timer1_Timer()
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
    Main_Form.Tag = VBA.Trim(Me.Text1.Text)
    Unload Me
End If
End Sub
