VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BC184000-7A5A-11D2-B543-006097FAF8B8}#1.6#0"; "bbGetDir.ocx"
Begin VB.Form Act_Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Sistema de Avctivación"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   Icon            =   "Form_Act1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Periodo de Activación:"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   5415
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
         CurrentDate     =   38502
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
         CurrentDate     =   38502
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Actvación:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de vencimiento:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label olFecha_Format 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "DD/MM/YYYY:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label olFecha_Order 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "DMY"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Use el FORMATO:"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   3000
      Picture         =   "Form_Act1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox otWindows_Key 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "otWindows_Key"
      Top             =   3840
      Width           =   5175
   End
   Begin VB.TextBox otCPU_ID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "otCPU_ID"
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox otNomb_Agent 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "otNomb_Agent"
      Top             =   4200
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
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton Commands 
      BackColor       =   &H8000000A&
      Caption         =   "Activar"
      Height          =   615
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Recuperar"
      Enabled         =   0   'False
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox otNomb_Local 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "otNomb_Local"
      Top             =   3240
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
      Width           =   3855
   End
   Begin VB.CommandButton Commands 
      Caption         =   "Salir"
      Height          =   615
      Index           =   2
      Left            =   4320
      TabIndex        =   13
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WINDOWS XP CD-KEY:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1770
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPU-ID:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   720
      TabIndex        =   15
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
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1410
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
   Begin VB.Line Line3 
      X1              =   0
      X2              =   5400
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Local:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1290
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
Attribute VB_Name = "Act_Form1"
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
Dim sgDir_Fls1 As String
Dim sgDir_Fls2 As String
Dim sgDir_Img As String
Dim sgDir_Mp3 As String
Dim sgDir_Pub As String
Dim sgFec_iAc As String
Dim sgFec_Fac As String
Dim sgSer_Mac As String
Dim sgWin_Key As String
Dim sgNom_Loc As String * 40
Dim tLocaleInfo As CLocaleInfo
    
Private Sub Command2_Click()
Call VBA.Shell("osk.exe")
End Sub

Private Sub Commands_Click(Index As Integer)
Select Case Index
Case Is = 0
    Call Get_System_Path(sPath, sgDir_Mac)
    
    sgDir_odb = sPath(1)
    sgDir_Tmp = sPath(2)
    sgDir_Fls1 = sPath(3)
    sgDir_Fls2 = Read_Ini_File(App.Path & "\PathV2.ini", "PATHS", "DIR_FL2", pRuta)
    sgDir_Img = sPath(4)
    sgDir_Mp3 = sPath(5)
    sgDir_Pub = sPath(6)
    sgFec_iAc = sPath(7)
    sgFec_Fac = sPath(8)
    sgSer_Mac = sPath(9)
    sgNom_Loc = VBA.Trim(sPath(10))
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
    Err.Clear
    If FileExist(App.Path & "\xpkey.exe") = True Then
        VBA.Shell (App.Path & "\xpkey.exe")
'       ***********************************************************************
        Dim objFSO As Object, objTextStream As Object
        Dim strFileName As String, fsoForReading As Integer
        Dim sCDKEY As String
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        strFileName = (App.Path & "\xpkey.txt")
        fsoForReading = 1
        Set objTextStream = objFSO.OpenTextFile(strFileName, fsoForReading)
        sCDKEY = VBA.Trim(objTextStream.ReadLine)
        objTextStream.Close
'       ***********************************************************************
        sgWin_Key = sCDKEY
    Else
        MsgBox ("No se encuentra el archivo [xpkey.exe],. informar a soporte técnico")
        Me.otWindows_Key.Text = "WINDOWS XP CD KEY UNRECOVERY..."
        Me.otWindows_Key.Refresh
        Return
    End If
  
    Dim sTmp1 As String
    Dim sTmp2 As String
    
    sTmp1 = Lee_Serial
    sgSer_Mac = Left$(sTmp1, 4) & "-" & Right$(sTmp1, 4)
    sgNom_Loc = VBA.Trim(Me.otNomb_Local.Text)
    sgWin_Key = sgWin_Key
    sgSer_CPU = MBCPUNumber()
        
    sPath(7) = DTPicker1.value
    sPath(8) = DTPicker2.value
    sPath(9) = sgSer_Mac
    sPath(10) = sgNom_Loc
    sPath(14) = sgSer_CPU
    sPath(43) = sgWin_Key
    
    Call Upd_Path(sgDir_Mac, sPath)
    Me.Refresh
    Me.Commands(1).Enabled = False
    bOkCrt = True
    If Err.Number = 0 Then
        Call MsgBox("El producto ha sido activado...")
        Call Limpiar_Form(Me)
    Else
        Call MsgBox("huvo un error: [" & VBA.Trim(Err.Description) & "], es posible que el producto no haya sido activado...")
        Err.Clear
    End If
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
Set tLocaleInfo = New CLocaleInfo
With tLocaleInfo
    olFecha_Order.Caption = VBA.UCase(.DateFormatOrder)
    olFecha_Format.Caption = VBA.UCase(.ShortDateFormat)
End With
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

