VERSION 5.00
Object = "{120FD660-13F8-11D1-943D-444553540000}#1.0#0"; "ctmedit.ocx"
Object = "{E5821C40-7D41-11D0-943C-444553540000}#1.0#0"; "ctnedit.ocx"
Object = "{BC184000-7A5A-11D2-B543-006097FAF8B8}#1.6#0"; "bbGetDir.ocx"
Object = "{F7E69521-3C28-11D2-B3E7-00AA00B42B7C}#3.1#0"; "fpTab30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Svr_Form 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Servicio"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "Servicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1800
      Top             =   6360
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   615
      Left            =   6600
      TabIndex        =   47
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   3000
      Picture         =   "Servicio.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin TabproADOLib.fpTabProADO fpTabProADO1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _Version        =   196609
      _ExtentX        =   14208
      _ExtentY        =   10610
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignTextH      =   1
      AlignTextV      =   1
      ThreeD          =   -1  'True
      MarginLeft      =   150
      MarginRight     =   150
      ActiveTabBold   =   0   'False
      TabSeparator    =   6
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      BookRingShowHole=   -1  'True
      PageEarMarkType =   3
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      ThreeDAppearance=   0
      DataField       =   ""
      DataMember      =   ""
      TabCaption      =   "Servicio.frx":058C
      PageEarMarkPictureNext=   "Servicio.frx":0930
      PageEarMarkPicturePrev=   "Servicio.frx":094C
      EarMarkPictureNext=   "Servicio.frx":0968
      EarMarkPicturePrev=   "Servicio.frx":0984
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modificar (2)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   -16920
         TabIndex        =   55
         Top             =   -20535
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modificar (4)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   -16920
         TabIndex        =   45
         Top             =   -20535
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modificar (3)"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   -16920
         TabIndex        =   44
         Top             =   -20535
         Width           =   1680
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Modificar (1)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   5280
         Width           =   1680
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   5295
         Left            =   -22815
         TabIndex        =   42
         Top             =   -20655
         Width           =   7695
         Begin VB.Frame Frame6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Mapa de configuración del teclado:."
            ForeColor       =   &H80000008&
            Height          =   4455
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   7215
            Begin VB.CommandButton Command5 
               Caption         =   "Notepad"
               Height          =   855
               Left            =   5880
               Picture         =   "Servicio.frx":09A0
               Style           =   1  'Graphical
               TabIndex        =   93
               Top             =   360
               Width           =   1095
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   1
               Left            =   1680
               TabIndex        =   65
               Top             =   360
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":126A
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   2
               Left            =   1680
               TabIndex        =   66
               Top             =   930
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":1286
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   3
               Left            =   1680
               TabIndex        =   67
               Top             =   1500
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":12A2
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   4
               Left            =   1680
               TabIndex        =   68
               Top             =   2070
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":12BE
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   5
               Left            =   1680
               TabIndex        =   69
               Top             =   2640
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":12DA
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   6
               Left            =   1680
               TabIndex        =   70
               Top             =   3240
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":12F6
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   13
               Left            =   1680
               TabIndex        =   71
               Top             =   3720
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":1312
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   7
               Left            =   4800
               TabIndex        =   79
               Top             =   360
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":132E
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   8
               Left            =   4800
               TabIndex        =   80
               Top             =   930
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":134A
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   9
               Left            =   4800
               TabIndex        =   81
               Top             =   1500
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":1366
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   10
               Left            =   4800
               TabIndex        =   82
               Top             =   2070
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":1382
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   11
               Left            =   4800
               TabIndex        =   83
               Top             =   2640
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":139E
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   12
               Left            =   4800
               TabIndex        =   84
               Top             =   3240
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":13BA
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin CTMEDITLibCtl.ctMEdit ctMEdit4 
               Height          =   255
               Index           =   14
               Left            =   4800
               TabIndex        =   85
               Top             =   3720
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
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
               DropPicture     =   "Servicio.frx":13D6
               BackColor       =   -2147483633
               CaseType        =   1
               SelectOnFocus   =   -1  'True
               EditMask        =   "X"
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Popular:"
               Height          =   195
               Index           =   7
               Left            =   4035
               TabIndex        =   92
               Top             =   420
               Width           =   585
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VIAP:"
               Height          =   195
               Index           =   8
               Left            =   4215
               TabIndex        =   91
               Top             =   990
               Width           =   405
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página (Arriba):"
               Height          =   195
               Index           =   9
               Left            =   3540
               TabIndex        =   90
               Top             =   1560
               Width           =   1080
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Página (Abajo):"
               Height          =   195
               Index           =   10
               Left            =   3540
               TabIndex        =   89
               Top             =   2130
               Width           =   1080
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Verificar Temas:"
               Height          =   195
               Index           =   11
               Left            =   3480
               TabIndex        =   88
               Top             =   2700
               Width           =   1140
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Switch enyte Publicidad:"
               Height          =   195
               Index           =   0
               Left            =   2880
               TabIndex        =   87
               Top             =   3300
               Width           =   1740
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Switch Karaoke:"
               Height          =   195
               Index           =   12
               Left            =   3480
               TabIndex        =   86
               Top             =   3780
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borra Tema (Todos):"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   78
               Top             =   3240
               Width           =   1455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borra Tema (Actual):"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   77
               Top             =   2700
               Width           =   1455
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Retrocreder:"
               Height          =   195
               Index           =   4
               Left            =   690
               TabIndex        =   76
               Top             =   2130
               Width           =   885
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Borrar CR:"
               Height          =   195
               Index           =   3
               Left            =   840
               TabIndex        =   75
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Añadir CR.(2):"
               Height          =   195
               Index           =   2
               Left            =   585
               TabIndex        =   74
               Top             =   990
               Width           =   990
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Añadir CR.(1):"
               Height          =   195
               Index           =   1
               Left            =   585
               TabIndex        =   73
               Top             =   420
               Width           =   990
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tecla de Pausa:"
               Height          =   195
               Index           =   4
               Left            =   360
               TabIndex        =   72
               Top             =   3720
               Width           =   1170
            End
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   5175
         Left            =   -22815
         TabIndex        =   22
         Top             =   -20655
         Width           =   7695
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Acceso de supervisor:."
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   120
            TabIndex        =   60
            Top             =   3120
            Width           =   3375
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Activar Acceso"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   63
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox otAccess3 
               Height          =   375
               IMEMode         =   3  'DISABLE
               Left            =   240
               MaxLength       =   6
               PasswordChar    =   "#"
               TabIndex        =   61
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contraseña:"
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
               Left            =   240
               TabIndex        =   62
               Top             =   360
               Width           =   1035
            End
         End
         Begin VB.Frame Frame32 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Presentación"
            ForeColor       =   &H80000008&
            Height          =   4095
            Left            =   3720
            TabIndex        =   36
            Top             =   360
            Width           =   3735
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Eliminar lista de tema marcados al Iniciar la aplicación."
               Height          =   495
               Index           =   9
               Left            =   240
               TabIndex        =   107
               Top             =   3480
               Width           =   3375
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Permitir SWITCH Karaoke"
               Height          =   255
               Index           =   7
               Left            =   240
               TabIndex        =   59
               Top             =   3120
               Width           =   2295
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Permitir SWITCH publicidad"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   51
               Top             =   2640
               Width           =   2295
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Mezclar POPULAR"
               Height          =   255
               Index           =   5
               Left            =   240
               TabIndex        =   41
               Top             =   2160
               Width           =   2295
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Mantener siempre adelante"
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   40
               Top             =   1590
               Width           =   2295
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Desactivar Televisor"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   39
               Top             =   1140
               Width           =   2055
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Mostrar etiquetas en discos"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   38
               Top             =   690
               Width           =   2295
            End
            Begin VB.CheckBox Check3 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Mostrar etiquetas de VIDEO"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   37
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Frame Frame31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tiempo"
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   3375
            Begin CTNEDITLibCtl.ctNEdit ctNEdit3 
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   28
               Top             =   480
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
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
               DropPicture     =   "Servicio.frx":13F2
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit3 
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   31
               Top             =   960
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
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
               DropPicture     =   "Servicio.frx":140E
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit3 
               Height          =   255
               Index           =   3
               Left            =   1800
               TabIndex        =   34
               Top             =   1440
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
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
               DropPicture     =   "Servicio.frx":142A
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit3 
               Height          =   255
               Index           =   4
               Left            =   1800
               TabIndex        =   52
               Top             =   1920
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
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
               DropPicture     =   "Servicio.frx":1446
               BackColor       =   -2147483633
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Temas de promo cada:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   54
               Top             =   1920
               Width           =   1635
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Temas"
               Height          =   195
               Index           =   3
               Left            =   2520
               TabIndex        =   53
               Top             =   1920
               Width           =   480
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minutos"
               Height          =   195
               Index           =   2
               Left            =   2520
               TabIndex        =   35
               Top             =   1440
               Width           =   555
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bono Músical (Stand):"
               Height          =   195
               Index           =   3
               Left            =   195
               TabIndex        =   33
               Top             =   1440
               Width           =   1560
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Segundos"
               Height          =   195
               Index           =   1
               Left            =   2520
               TabIndex        =   32
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Retorno a Disco:"
               Height          =   195
               Index           =   2
               Left            =   435
               TabIndex        =   30
               Top             =   960
               Width           =   1320
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Minutos"
               Height          =   195
               Index           =   0
               Left            =   2520
               TabIndex        =   29
               Top             =   480
               Width           =   555
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Retorno a género:"
               Height          =   195
               Index           =   1
               Left            =   345
               TabIndex        =   27
               Top             =   480
               Width           =   1410
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   5175
         Left            =   -22815
         TabIndex        =   21
         Top             =   -20655
         Width           =   7695
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   "Opciones con créditos."
            ForeColor       =   &H80000008&
            Height          =   4455
            Left            =   240
            TabIndex        =   95
            Top             =   240
            Width           =   7095
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Almacenar créditos marcados."
               Height          =   255
               Index           =   1
               Left            =   3840
               TabIndex        =   98
               Top             =   840
               Width           =   2535
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Activar créditos ilimitados."
               Height          =   255
               Index           =   2
               Left            =   3840
               TabIndex        =   97
               Top             =   1560
               Width           =   2535
            End
            Begin VB.CheckBox Check2 
               BackColor       =   &H00C0FFFF&
               Caption         =   "NO permitir temas duplicados"
               Height          =   255
               Index           =   3
               Left            =   3840
               TabIndex        =   96
               Top             =   2280
               Width           =   2535
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit2 
               Height          =   255
               Index           =   1
               Left            =   2325
               TabIndex        =   99
               Top             =   840
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               DropPicture     =   "Servicio.frx":1462
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit2 
               Height          =   255
               Index           =   2
               Left            =   2325
               TabIndex        =   100
               Top             =   1560
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               DropPicture     =   "Servicio.frx":147E
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit2 
               Height          =   255
               Index           =   3
               Left            =   2325
               TabIndex        =   101
               Top             =   2280
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               DropPicture     =   "Servicio.frx":149A
               BackColor       =   -2147483633
            End
            Begin CTNEDITLibCtl.ctNEdit ctNEdit2 
               Height          =   255
               Index           =   4
               Left            =   2325
               TabIndex        =   102
               Top             =   3000
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               DropPicture     =   "Servicio.frx":14B6
               BackColor       =   -2147483633
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Límite de créditos:"
               Height          =   195
               Index           =   1
               Left            =   780
               TabIndex        =   106
               Top             =   840
               Width           =   1305
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Créditos almacenados:"
               Height          =   195
               Index           =   2
               Left            =   480
               TabIndex        =   105
               Top             =   1560
               Width           =   1605
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bonos por cada Dolar:"
               Height          =   195
               Index           =   3
               Left            =   495
               TabIndex        =   104
               Top             =   2280
               Width           =   1590
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Créditos por video:"
               Height          =   195
               Index           =   4
               Left            =   765
               TabIndex        =   103
               Top             =   3000
               Width           =   1320
            End
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7695
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   9
            Left            =   7080
            TabIndex        =   57
            Top             =   1950
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   56
            Top             =   1920
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   8
            Left            =   7080
            TabIndex        =   49
            Top             =   3975
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   1320
            TabIndex        =   48
            Top             =   3945
            Width           =   5535
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   24
            Top             =   4440
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   7
            Left            =   7080
            TabIndex        =   23
            Top             =   4470
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   14
            Top             =   480
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   7080
            TabIndex        =   13
            Top             =   510
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   12
            Top             =   1470
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   3
            Left            =   7080
            TabIndex        =   11
            Top             =   1500
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   10
            Top             =   2445
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   4
            Left            =   7080
            TabIndex        =   9
            Top             =   2475
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   8
            Top             =   2955
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   5
            Left            =   7080
            TabIndex        =   7
            Top             =   2985
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   6
            Top             =   3450
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   6
            Left            =   7080
            TabIndex        =   5
            Top             =   3480
            Width           =   375
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   4
            Top             =   975
            Width           =   5535
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   255
            Index           =   2
            Left            =   7080
            TabIndex        =   3
            Top             =   1005
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FILES SEC.:"
            Height          =   195
            Index           =   8
            Left            =   375
            TabIndex        =   58
            Top             =   2010
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PUBLICIDAD2:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   50
            Top             =   4035
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FONDO:"
            Height          =   195
            Index           =   6
            Left            =   600
            TabIndex        =   25
            Top             =   4530
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ODBC:"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FILES PRI.:"
            Height          =   195
            Index           =   2
            Left            =   375
            TabIndex        =   19
            Top             =   1560
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FOTOS:"
            Height          =   195
            Index           =   3
            Left            =   630
            TabIndex        =   18
            Top             =   2535
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CASIONERO:"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   17
            Top             =   3045
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PUBLICIDAD1:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   3540
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPORALES:"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   15
            Top             =   1065
            Width           =   1140
         End
      End
      Begin BBGETDIRLibCtl.Bbgetdir Bbgetdir1 
         Left            =   120
         Top             =   4680
         _Version        =   65542
         _ExtentX        =   900
         _ExtentY        =   661
         _StockProps     =   0
         ExtraButtonCaption=   "New"
      End
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 Restante"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5040
      TabIndex        =   94
      Top             =   6480
      Width           =   1320
   End
End
Attribute VB_Name = "Svr_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeResto As Long

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case Is = 1
    Me.Frame1.Enabled = IIf(Me.Check1(1).value = 1, True, False)
Case Is = 2
    Me.Frame2.Enabled = IIf(Me.Check1(2).value = 1, True, False)
Case Is = 3
    Me.Frame3.Enabled = IIf(Me.Check1(3).value = 1, True, False)
Case Is = 4
    Me.Frame4.Enabled = IIf(Me.Check1(4).value = 1, True, False)
End Select
End Sub

Private Sub Check3_Click(Index As Integer)
If Index = 8 Then
    If Check3(Index).value = 1 Then
        Me.otAccess3.Enabled = True
    Else
        Me.otAccess3.Enabled = False
    End If
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim strDir As String
If Index = 7 Then
    Me.CommonDialog1.InitDir = sgFle_Fon
    Me.CommonDialog1.Filter = "*.bmp;*.gif;*.jpg"
    Me.CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.ShowOpen
    strDir = Me.CommonDialog1.filename
Else
    Me.Bbgetdir1.ExtraButtonVisible = False
    Me.Bbgetdir1.FocusedDirectory = App.Path
    Me.Bbgetdir1.DisplayCurrentDirectory = True
    Me.Bbgetdir1.StatusText = "Seleccione el Directorio deseado!"
    Me.Bbgetdir1.ListCaption = "" ' use the default localized text
    strDir = Bbgetdir1.ShowDirectoryList(0)
End If
If strDir <> "" Then
    Me.Text1(Index).Text = strDir
End If
End Sub

Private Sub Command2_Click()
If Check1(1).value = 1 Then
    sgDir_odb = VBA.Trim(Me.Text1(1).Text)
    sgDir_Tmp = VBA.Trim(Me.Text1(2).Text)
    sgDir_Fls = VBA.Trim(Me.Text1(3).Text)
    sgDir_Img = VBA.Trim(Me.Text1(4).Text)
    sgDir_Mp3 = VBA.Trim(Me.Text1(5).Text)
    sgFle_Fon = VBA.Trim(Me.Text1(7).Text)
    sgDir_Pub1 = VBA.Trim(Me.Text1(6).Text)
    sgDir_Pub2 = VBA.Trim(Me.Text1(8).Text)
    sgDir_Fls2 = VBA.Trim(Me.Text1(9).Text)
End If
If Check1(2).value = 1 Then
    igLim_Cred = Me.ctNEdit2(1).value
    igCnt_CR = Me.ctNEdit2(2).value
    sgKb_BonC = Me.ctNEdit2(3).value
    sgKb_VID = Me.ctNEdit2(4).value
    igFlg_SavedCR = Me.Check2(1).value
    igKeep_Cred = Me.Check2(2).value
    igNoDuplicT = Me.Check2(3).value
End If
If Check1(3).value = 1 Then
    igDelay_Return_Gen = Me.ctNEdit3(1).value
    igDelay_Return_Dis = Me.ctNEdit3(2).value
    igDelay_Bonus_Vid = Me.ctNEdit3(3).value
    bgVideoLabel = IIf(Me.Check3(1).value = 1, True, False)
    bgDiscLabel = IIf(Me.Check3(2).value = 1, True, False)
    igScr_Alone = Me.Check3(3).value
    bgKeep_On_Top = VBA.IIf(Me.Check3(4).value = 1, True, False)
    igMixe_Popu = Me.Check3(5).value
    bgSw_Pub = IIf(VBA.Val(VBA.Int(Me.Check3(6).value)) = 1, True, False)
    sgCr_AKey = VBA.Trim(Me.otAccess3.Text)
    sgIdx_Prm = Me.ctNEdit3(4).value
    igInd_Kar = Me.Check3(7).value
    igShowPass = Me.Check3(8).value
    igStartPlayMusic = Me.Check3(9).value
End If
If Check1(4).value = 1 Then
    sgKb_Crd1 = Me.ctMEdit4(1).Text
    sgKb_Crd2 = Me.ctMEdit4(2).Text
    sgKb_Del = Me.ctMEdit4(3).Text
    sgKb_Ret = Me.ctMEdit4(4).Text
    sgKb_ResM = Me.ctMEdit4(5).Text
    sgKb_ResA = Me.ctMEdit4(6).Text
    sgKb_Pop = Me.ctMEdit4(7).Text
    sgKb_VIP = Me.ctMEdit4(8).Text
    sgKb_UP = Me.ctMEdit4(9).Text
    sgKb_DN = Me.ctMEdit4(10).Text
    sgKb_Vef = Me.ctMEdit4(11).Text
    sgKb_SwP = Me.ctMEdit4(12).Text
    sgKb_Pause = Me.ctMEdit4(13).Text
    sgKb_SwK = Me.ctMEdit4(14).Text
End If
Call Save_Defa_Path2
Unload Me
End Sub

Private Sub Command3_Click()
Call VBA.Shell("osk.exe")
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If Len(Text2.Text) = 0 And KeyAscii = 45 Then
      KeyAscii = 0
   End If
 
   If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
      KeyAscii = 0
   End If
End Sub


Private Sub Command5_Click()
VBA.Shell "NOTEPAD.EXE", vbNormalFocus
End Sub

Private Sub Form_Load()
TimeResto = (60) * 3
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
    Unload Me
End If
End Sub
