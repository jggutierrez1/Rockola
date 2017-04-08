VERSION 5.00
Begin VB.Form tSRmAIN 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TSR (ROCKOLLA TOOLS SET)"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   Icon            =   "TSR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[1]-<HERRAMIENTAS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2145
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[0]-<SALIR>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   1920
      Picture         =   "TSR.frx":014A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "tSRmAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 97
    ControlPanel2.Show vbModal
    Me.Show
Case Is = 49
    ControlPanel2.Show vbModal
    Me.Show
Case Is = 48
    End
Case Is = 96
    End
End Select
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
    End
End If
End Sub

