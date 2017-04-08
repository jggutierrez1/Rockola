VERSION 5.00
Object = "{E5821C40-7D41-11D0-943C-444553540000}#1.0#0"; "ctnedit.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BC184000-7A5A-11D2-B543-006097FAF8B8}#1.6#0"; "bbgetdir.ocx"
Begin VB.Form Act_Form 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Sistema de Avctivación"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "Form_Act.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   48627713
      CurrentDate     =   38502
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   48627713
      CurrentDate     =   38502
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   2880
      Picture         =   "Form_Act.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox otWindows_Key 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "otWindows_Key"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox otCPU_ID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "otCPU_ID"
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox otNomb_Agent 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Text            =   "otNomb_Agent"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton oCmd_Origen 
      Caption         =   "Origen de datos:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox otOrigen 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Commands 
      BackColor       =   &H8000000A&
      Caption         =   "Activar"
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Recuperar"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin CTNEDITLibCtl.ctNEdit otTokens 
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   4200
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      DropPicture     =   "Form_Act.frx":0884
      BackColor       =   -2147483633
      Alignment       =   1
   End
   Begin VB.TextBox otNomb_Local 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "otNomb_Local"
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox otSerie 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "otSerie"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Salir"
      Height          =   615
      Index           =   2
      Left            =   4200
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   48627713
      CurrentDate     =   38502
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WINDOWS XP CD-KEY:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   1770
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPU-ID:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Agente:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Última colecta:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   15
      Top             =   4200
      Width           =   1050
   End
   Begin BBGETDIRLibCtl.Bbgetdir Bbgetdir1 
      Left            =   4440
      Top             =   120
      _Version        =   65542
      _ExtentX        =   900
      _ExtentY        =   900
      _StockProps     =   0
      ExtraButtonCaption=   "New"
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   5400
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5400
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Local:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5400
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de vencimiento:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Actvación:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Tokens:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   4200
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número de serie:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
   End
End
Attribute VB_Name = "Act_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim igSw As Integer
Dim bOkCrt As Boolean
Dim sPath(60) As String
Dim sgDir_Mac As String
Dim sgDir_odb As String
Dim sgDir_Tmp As String
Dim sgDir_Fls As String
Dim sgDir_Img As String
Dim sgDir_Mp3 As String
Dim sgDir_Pub As String
Dim sgFec_iAc As String
Dim sgFec_Fac As String
Dim sgSer_Mac As String
Dim sgWin_Key As String
Dim sgNom_Loc As String * 40
Dim sgFec_Tok As String
Dim igCnt_Tok As Integer

Private Sub Command1_Click()
If FileExist(App.Path & "\ViewKeyXP.exe") = False Then
    MsgBox ("No se encuentra la utilirería " & App.Path & "\ViewKeyXP.exe")
Else
    VBA.Shell (App.Path & "\ViewKeyXP.exe")
End If
End Sub

Private Sub Command2_Click()
Call VBA.Shell("osk.exe")
End Sub

Private Sub Commands_Click(Index As Integer)
Select Case Index
Case Is = 0
    Call Get_System_Path(sPath, sgDir_Mac)
    
    sgDir_odb = sPath(1)
    sgDir_Tmp = sPath(2)
    sgDir_Fls = sPath(3)
    sgDir_Img = sPath(4)
    sgDir_Mp3 = sPath(5)
    sgDir_Pub = sPath(6)
    sgFec_iAc = sPath(7)
    sgFec_Fac = sPath(8)
    sgSer_Mac = sPath(9)
    sgNom_Loc = VBA.Trim(sPath(10))
    igCnt_Tok = sPath(11)
    sgFec_Tok = sPath(12)
    sgSer_CPU = VBA.Trim(sPath(14))
    sgWin_Key = VBA.Trim(sPath(43))
    Me.otSerie.Text = sgSer_Mac
    Me.otCPU_ID.Text = sgSer_CPU
   
    Me.otNomb_Local.Text = sgNom_Loc
    Me.otWindows_Key.Text = sgWin_Key
   
    If (sgFec_iAc = "" Or sgFec_iAc = " ") Then
        sgFec_iAc = VBA.Date()
    End If
    
    DTPicker1.Day = 1
    DTPicker1.Year = VBA.Year(VBA.DateValue(sgFec_iAc))
    DTPicker1.Month = VBA.Month(VBA.DateValue(sgFec_iAc))
    DTPicker1.Day = VBA.Day(VBA.DateValue(sgFec_iAc))
    
    If (sgFec_Fac = "" Or sgFec_Fac = " ") Then
        sgFec_Fac = VBA.DateAdd("m", 1, VBA.Date())
    End If
    
    DTPicker2.Day = 1
    DTPicker2.Year = VBA.Year(VBA.DateValue(sgFec_Fac))
    DTPicker2.Month = VBA.Month(VBA.DateValue(sgFec_Fac))
    DTPicker2.Day = VBA.Day(VBA.DateValue(sgFec_Fac))
    
    
    Me.otTokens.value = VBA.Int(VBA.Val(igCnt_Tok))
    If sgFec_Tok = "" Then
        DTPicker3.Day = 1
        DTPicker3.Month = 1
        DTPicker3.Year = 1900
    Else
        DTPicker3.Day = 1
        DTPicker3.Year = VBA.Year(VBA.DateValue(sgFec_Tok))
        DTPicker3.Month = VBA.Month(VBA.DateValue(sgFec_Tok))
        DTPicker3.Day = VBA.Day(VBA.DateValue(sgFec_Tok))
    End If
    Me.Refresh
Case Is = 1
    If otNomb_Local.Text = "" Then
        Call MsgBox("Debe suministrar el NOMBRE DEL [LOCAL]", vbInformation, "Atención")
        Exit Sub
    End If
    'If otNomb_Agent.Text = "" Then
    '    Call MsgBox("Debe suministrar el NOMBRE DEL [AGENTE]", vbInformation, "Atención")
    '    Exit Sub
    'End If
    
    Me.otTokens.value = 0
    DTPicker3.Day = VBA.Day(VBA.Date())
    DTPicker3.Month = VBA.Month(VBA.Date())
    DTPicker3.Year = VBA.Year(VBA.Date)
    
    Dim sTmp1 As String
    Dim sTmp2 As String
    
    sTmp1 = Lee_Serial
    sgSer_Mac = Left$(sTmp1, 4) & "-" & Right$(sTmp1, 4)
    sgNom_Loc = VBA.Trim(Me.otNomb_Local.Text)
    sgWin_Key = VBA.Trim(Me.otWindows_Key.Text)
    sgSer_CPU = Get_CPU_Id
        
        
    sPath(7) = DTPicker1.value
    sPath(8) = DTPicker2.value
    sPath(9) = sgSer_Mac
    sPath(10) = sgNom_Loc
    sPath(11) = Me.otTokens.value
    sPath(12) = DTPicker3.value
    sPath(14) = sgSer_CPU
    sPath(43) = sgWin_Key
    
    Call Upd_Path(sgDir_Mac, sPath)
    Me.Refresh
    Me.Commands(1).Enabled = False
    bOkCrt = True
    Call Guardar_Datos
    Call MsgBox("El producto ha sido activado...")
    Call Limpiar_Form(Me)
    End
Case Is = 2
    End
End Select
End Sub

Private Sub Form_Load()
'Me.Caption = "Activador, Ver." & VBA.Trim(VBA.Str(App.Major)) & "." & VBA.Trim(VBA.Str(App.Minor)) & "." & VBA.Trim(VBA.Str(App.Revision))
Me.Caption = Me.Caption & " [RECREATIVO VERAGUIENSE]"
sgDir_Mac = App.Path
bOkCrt = False
Call Limpiar_Form(Me)
Call Get_System_Path(sPath, sgDir_Mac)
DTPicker1.Day = 1
DTPicker1.Year = VBA.Year(VBA.Date())
DTPicker1.Month = VBA.Month(VBA.Date())
DTPicker1.Day = VBA.Day(VBA.Date())
    
DTPicker2.Day = 1
DTPicker2.Year = VBA.Year(VBA.Date())
DTPicker2.Month = IIf(VBA.Month(VBA.Date()) = 12, 1, VBA.Month(VBA.Date()) + 1)
'DTPicker2.Day = VBA.Day(VBA.Date())
sgDir_Mac = App.Path
igSw = 1
'Call oCmd_Origen_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Asegurarnos de "liberar" la memoria.
    Set Form1 = Nothing
End Sub

Private Sub oCmd_Origen_Click()
Dim lcSelectedPath As String
Call Limpiar_Form(Me)
If igSw = 0 Then
    If sgDir_Mac = "" Then
        sgDir_Mac = "C:\"
    End If
    Me.Bbgetdir1.FocusedDirectory = sgDir_Mac
    Me.Bbgetdir1.ListAutoCenter = True
    Me.Bbgetdir1.StatusText = "Origen del Sistema"
    lcSelectedPath = Me.Bbgetdir1.ShowDirectoryListEx(1) + "\"
Else
    lcSelectedPath = sgDir_Mac + "\"
End If
If lcSelectedPath <> "" Then
    If bOkCrt = True Then
        Dim iResp As Integer
        iResp = MsgBox("ya el sistema fue activado, activar nuevamente?", vbQuestion + vbYesNo, "Confirmar")
        If iResp <> 6 Then
            Exit Sub
        Else
            Me.Commands(1).Enabled = True
        End If
    End If
    Me.otOrigen.Text = lcSelectedPath
    sgDir_Mac = lcSelectedPath
    Call Commands_Click(0)
Else
    Me.otOrigen.Text = ""
    sgDir_Mac = ""
End If
End Sub

Private Sub Guardar_Datos()
Open App.Path & "\INFO.DAT" For Append As #1
Write #1, otNomb_Agent.Text, sgNom_Loc, sgSer_Mac, VBA.Format(DTPicker1.value, "DD/MM/YYYY"), VBA.Format(DTPicker2.value, "DD/MM/YYYY"), Me.otTokens.value, VBA.Format(DTPicker3.value, "DD/MM/YYYY")
Close #1
End Sub

Private Sub Limpiar_Form(poMe As Object)
Dim oCtr As Control
For Each oCtr In poMe
    If TypeOf oCtr Is TextBox Then
        oCtr.Text = ""
    End If
    If TypeOf oCtr Is ctNEdit Then
        oCtr.value = 0
    End If
    If TypeOf oCtr Is DTPicker Then
        oCtr.Day = 1
        oCtr.Month = 1
        oCtr.Year = 1900
    End If
Next
End Sub

Private Sub otNomb_Agent_Validate(Cancel As Boolean)
Me.otNomb_Agent.Text = VBA.UCase(Me.otNomb_Agent.Text)
End Sub

Private Sub otNomb_Local_Validate(Cancel As Boolean)
Me.otNomb_Local.Text = VBA.UCase(Me.otNomb_Local.Text)
End Sub

