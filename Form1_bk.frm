VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{120FD660-13F8-11D1-943D-444553540000}#1.0#0"; "ctmedit.ocx"
Object = "{BE38FE43-D38D-11D0-B731-00403333B3B0}#1.0#0"; "tback.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BC184000-7A5A-11D2-B543-006097FAF8B8}#1.6#0"; "bbgetdir.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Main_Form 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8640
   ClientLeft      =   1740
   ClientTop       =   840
   ClientWidth     =   11880
   Icon            =   "Form1_bk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1_bk.frx":1CFA
   ScaleHeight     =   8640
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "S1"
      Height          =   255
      Left            =   360
      TabIndex        =   125
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S2"
      Height          =   255
      Left            =   720
      TabIndex        =   124
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin TbackLibCtl.TBack TBack4 
      Height          =   525
      Left            =   7200
      TabIndex        =   70
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   926
      _StockProps     =   224
      Appearance      =   1
      BackColor       =   65535
      GradientColorFrom=   255
      GradientColorTo =   65535
      GradientStyle   =   4
      Version         =   16777230
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   240
         TabIndex        =   123
         Top             =   1080
         Width           =   7215
      End
      Begin VB.CheckBox oChk_FndP 
         BackColor       =   &H80000009&
         Caption         =   "Portadas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   120
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CheckBox oChk_FndC 
         BackColor       =   &H80000009&
         Caption         =   "Cansionero"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox otRuteExternal2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   118
         Top             =   3600
         Width           =   4575
      End
      Begin VB.CommandButton oGetRute2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   117
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar"
         Height          =   615
         Left            =   2280
         TabIndex        =   116
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton oGetRute 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   115
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox otRuteExternal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   114
         Top             =   3240
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copiar Automaticamnte desde un origen externo"
         Height          =   195
         Left            =   240
         TabIndex        =   113
         Top             =   2760
         Width           =   3975
      End
      Begin RichTextLib.RichTextBox otNot_Found_List 
         Height          =   510
         Left            =   5520
         TabIndex        =   71
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   900
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1_bk.frx":15FFD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   4080
         TabIndex        =   122
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin BBGETDIRLibCtl.Bbgetdir Bbgetdir1 
         Left            =   6840
         Top             =   2640
         _Version        =   65542
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
         ExtraButtonCaption=   "New"
      End
      Begin VB.Label olInfo_cheker_Proc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
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
         Left            =   3600
         TabIndex        =   74
         Top             =   120
         Width           =   255
      End
      Begin VB.Label olInfo_Cheker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "olInfo_Cheker"
         Height          =   195
         Left            =   2880
         TabIndex        =   73
         Top             =   720
         Width           =   990
      End
   End
   Begin TbackLibCtl.TBack oFrame_Can 
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   224
      BackColor       =   12648384
      Version         =   16777230
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   12
         Left            =   120
         Picture         =   "Form1_bk.frx":1608A
         Stretch         =   -1  'True
         Top             =   5265
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   11
         Left            =   120
         Picture         =   "Form1_bk.frx":160D3
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   10
         Left            =   120
         Picture         =   "Form1_bk.frx":1611C
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   9
         Left            =   120
         Picture         =   "Form1_bk.frx":1629E
         Stretch         =   -1  'True
         Top             =   3855
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   8
         Left            =   120
         Picture         =   "Form1_bk.frx":16420
         Stretch         =   -1  'True
         Top             =   3390
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   7
         Left            =   120
         Picture         =   "Form1_bk.frx":17F62
         Stretch         =   -1  'True
         Top             =   2925
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   6
         Left            =   120
         Picture         =   "Form1_bk.frx":19AA4
         Stretch         =   -1  'True
         Top             =   2445
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   5
         Left            =   120
         Picture         =   "Form1_bk.frx":1B5E6
         Stretch         =   -1  'True
         Top             =   1980
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   4
         Left            =   120
         Picture         =   "Form1_bk.frx":1D128
         Stretch         =   -1  'True
         Top             =   1515
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   3
         Left            =   120
         Picture         =   "Form1_bk.frx":1EC6A
         Stretch         =   -1  'True
         Top             =   1050
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   2
         Left            =   120
         Picture         =   "Form1_bk.frx":207AC
         Stretch         =   -1  'True
         Top             =   570
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   270
         Index           =   1
         Left            =   120
         Picture         =   "Form1_bk.frx":222EE
         Stretch         =   -1  'True
         Top             =   105
         Width           =   270
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   12
         Left            =   600
         TabIndex        =   39
         Top             =   5160
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   38
         Top             =   4695
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   10
         Left            =   600
         TabIndex        =   37
         Top             =   4215
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   9
         Left            =   600
         TabIndex        =   36
         Top             =   3750
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   35
         Top             =   3285
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   34
         Top             =   2820
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   33
         Top             =   2340
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   32
         Top             =   1875
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   31
         Top             =   1410
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   30
         Top             =   945
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   29
         Top             =   465
         Width           =   4335
      End
      Begin VB.Label oLCanc 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   28
         Top             =   0
         Width           =   4335
      End
   End
   Begin TbackLibCtl.TBack oFrame_Dis 
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   630
      _Version        =   65536
      _ExtentX        =   1111
      _ExtentY        =   873
      _StockProps     =   224
      BackColor       =   14737632
      Version         =   16777230
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         HelpContextID   =   1
         Index           =   1
         Left            =   -120
         TabIndex        =   86
         Top             =   2280
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   33023
         GradientColorTo =   33023
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   88
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   87
            Top             =   240
            Width           =   2295
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   5
         Left            =   2520
         TabIndex        =   98
         Top             =   5040
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   100
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   95
         Top             =   5040
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   97
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   89
         Top             =   2280
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   91
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   7
         Left            =   0
         TabIndex        =   104
         Top             =   7800
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   112
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   111
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   8
         Left            =   2520
         TabIndex        =   105
         Top             =   7800
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   110
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   9
         Left            =   5040
         TabIndex        =   106
         Top             =   7800
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   108
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   6
         Left            =   5040
         TabIndex        =   101
         Top             =   5040
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   103
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   92
         Top             =   2280
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   94
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "19ITEMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   4
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   1
         Left            =   1365
         TabIndex        =   85
         Top             =   45
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   9
         Left            =   6405
         TabIndex        =   84
         Top             =   5575
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   8
         Left            =   3885
         TabIndex        =   83
         Top             =   5575
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   7
         Left            =   1365
         TabIndex        =   82
         Top             =   5575
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   6
         Left            =   6405
         TabIndex        =   81
         Top             =   2800
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   5
         Left            =   3885
         TabIndex        =   80
         Top             =   2805
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   4
         Left            =   1365
         TabIndex        =   79
         Top             =   2805
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   3
         Left            =   6405
         TabIndex        =   78
         Top             =   40
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label olVideo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<VIDEO>"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   330
         Index           =   2
         Left            =   3885
         TabIndex        =   77
         Top             =   45
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Tag             =   "720"
         Top             =   0
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   3
         Left            =   5760
         TabIndex        =   11
         Tag             =   "5760"
         Top             =   0
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   6
         Left            =   5760
         TabIndex        =   10
         Tag             =   "5760"
         Top             =   2775
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   9
         Left            =   5760
         TabIndex        =   9
         Tag             =   "5760"
         Top             =   5520
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   3
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   6
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   9
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         Tag             =   "3240"
         Top             =   0
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   5
         Left            =   3240
         TabIndex        =   7
         Tag             =   "3240"
         Top             =   2775
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   8
         Left            =   3240
         TabIndex        =   6
         Tag             =   "3240"
         Top             =   5525
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   2
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   5
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   8
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   5
         Tag             =   "720"
         Top             =   2775
         Width           =   600
      End
      Begin VB.Label oLNum_Disk 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "Digital dream Fat"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   3
         Tag             =   "720"
         Top             =   5525
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   1
         Left            =   0
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   7
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   2040
   End
   Begin VB.Timer oGeneral_Timer 
      Interval        =   800
      Left            =   5640
      Top             =   1560
   End
   Begin TbackLibCtl.TBack oFrame_Gen 
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   600
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   224
      BackColor       =   12632319
      Version         =   16777230
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   470
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   940
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1410
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   1880
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2350
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   3290
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   3760
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   4230
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   16
         Top             =   4700
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   15
         Top             =   5170
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   14
         Top             =   5640
         Width           =   4335
      End
      Begin VB.Label oLGenero 
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   15
         Index           =   14
         Left            =   120
         TabIndex        =   13
         Top             =   7800
         Width           =   4335
      End
   End
   Begin VB.Timer oTM_codigo2 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   1560
   End
   Begin VB.Timer oTime_Mensajes2 
      Interval        =   1200
      Left            =   6120
      Top             =   2040
   End
   Begin VB.Timer otCargador_Video 
      Interval        =   1200
      Left            =   5160
      Top             =   1560
   End
   Begin TbackLibCtl.TBack TBack3 
      Height          =   1095
      Left            =   480
      TabIndex        =   52
      Top             =   360
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   1931
      _StockProps     =   224
      BackColor       =   -2147483633
      GradientColorFrom=   0
      GradientColorTo =   0
      TransparentBackground=   -1  'True
      HasLicense      =   -1  'True
      Version         =   16777230
      Begin CTMEDITLibCtl.ctMEdit otTema_Act 
         Height          =   375
         Left            =   840
         TabIndex        =   53
         Top             =   360
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   65535
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Digital dream Fat"
            Size            =   14.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         DropPicture     =   "Form1_bk.frx":22BD5
         BackColor       =   16576
         ForeColor       =   65535
         DisabledColor   =   8454016
         UseMaskChars    =   0   'False
         EditMask        =   "##-##-##"
      End
      Begin VB.Image oImg_c_Video 
         Height          =   270
         Left            =   3000
         Picture         =   "Form1_bk.frx":22BF1
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label olTema_Act 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   840
         Width           =   3240
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   360
         Left            =   720
         TabIndex        =   54
         Top             =   0
         Width           =   2370
      End
   End
   Begin VB.Timer oTimer_Reset 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5640
      Top             =   2040
   End
   Begin VB.Timer oTime_Mensajes 
      Interval        =   1200
      Left            =   5160
      Top             =   2040
   End
   Begin VB.Timer otCargador_Music 
      Interval        =   1200
      Left            =   4680
      Top             =   1560
   End
   Begin VB.Timer oTimer_Moneda 
      Interval        =   800
      Left            =   4680
      Top             =   2040
   End
   Begin TbackLibCtl.TBack oFrame_in_Sel 
      Height          =   975
      Left            =   360
      TabIndex        =   40
      Top             =   5430
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   1720
      _StockProps     =   224
      BorderStyle     =   1
      Version         =   16777230
      Begin CTMEDITLibCtl.ctMEdit otCodigo 
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   480
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   65535
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Digital dream Fat"
            Size            =   14.24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropPicture     =   "Form1_bk.frx":22C3A
         BackColor       =   16576
         ForeColor       =   65535
         UseMaskChars    =   0   'False
         EditMask        =   "99-99-99"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   1200
         X2              =   1200
         Y1              =   480
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Género"
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
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disco"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   42
         Top             =   120
         Width           =   690
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   2160
         Y1              =   480
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   960
         X2              =   1200
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   2400
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Canción"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   41
         Top             =   120
         Width           =   990
      End
   End
   Begin TbackLibCtl.TBack TBack2 
      Height          =   345
      Left            =   960
      TabIndex        =   48
      Top             =   240
      Width           =   2610
      _Version        =   65536
      _ExtentX        =   4604
      _ExtentY        =   609
      _StockProps     =   224
      GradientColorFrom=   12632319
      GradientColorTo =   8421631
      GradientStyle   =   1
      TransparentBackground=   -1  'True
      HasLicense      =   -1  'True
      Version         =   16777230
   End
   Begin TbackLibCtl.TBack TBack1 
      Height          =   5535
      Left            =   9480
      TabIndex        =   44
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   9763
      _StockProps     =   224
      Version         =   16777230
      Begin VB.ListBox oBkList 
         DataSource      =   "oDC_Temas"
         Height          =   255
         ItemData        =   "Form1_bk.frx":22C56
         Left            =   0
         List            =   "Form1_bk.frx":22C5D
         TabIndex        =   69
         Top             =   4920
         Width           =   2415
      End
      Begin VB.ListBox oLst_A_Tocar 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   68
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ListBox oLst_Popular 
         DataSource      =   "oDC_Temas"
         Height          =   255
         ItemData        =   "Form1_bk.frx":22C6A
         Left            =   0
         List            =   "Form1_bk.frx":22C6C
         TabIndex        =   67
         Top             =   5280
         Width           =   2415
      End
      Begin VB.FileListBox oLst_Temas_Video 
         Height          =   480
         Left            =   0
         TabIndex        =   64
         Top             =   4320
         Width           =   2415
      End
      Begin VB.FileListBox oLst_Pub 
         Height          =   480
         Left            =   0
         TabIndex        =   63
         Top             =   3600
         Width           =   2415
      End
      Begin MSDataListLib.DataList oLst_Disc 
         Bindings        =   "Form1_bk.frx":22C6E
         Height          =   255
         Left            =   0
         TabIndex        =   45
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         ListField       =   "NOM_DIS"
         BoundColumn     =   "ID_DIS"
      End
      Begin MSDataListLib.DataList oLst_Gen 
         Bindings        =   "Form1_bk.frx":22C85
         Height          =   255
         Left            =   0
         TabIndex        =   46
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         ListField       =   "DESCRI"
         BoundColumn     =   "ID_GEN"
      End
      Begin MSAdodcLib.Adodc oDC_GEN 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "oDC_GEN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc oDC_DISC 
         Height          =   375
         Left            =   0
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "oDC_DISC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataList oLst_Canc 
         Bindings        =   "Form1_bk.frx":22C9B
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         ListField       =   "DE_CAN"
         BoundColumn     =   "ID_CAN"
      End
      Begin MSAdodcLib.Adodc oDC_CANC 
         Height          =   375
         Left            =   0
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Link_Odbc"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "oDC_CANC"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc oDC_Temas 
         Height          =   375
         Left            =   0
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   "D:\Rockola\ODBC\Link_Dbf.dsn"
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "oDC_Temas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label olPosCod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posición->"
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
         Left            =   0
         TabIndex        =   57
         Top             =   2520
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label olPosLst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posicion ->"
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
         Left            =   0
         TabIndex        =   56
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.TextBox oSetFocus_Codigo 
      Height          =   285
      Left            =   2400
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label olPaginas2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<PASAR PÁGINA>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   121
      Top             =   5160
      Width           =   1920
   End
   Begin VB.Image oImg_PagDn 
      Height          =   360
      Left            =   3960
      Picture         =   "Form1_bk.frx":22CB2
      Stretch         =   -1  'True
      Top             =   6000
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image oImg_PagUp 
      Height          =   360
      Left            =   3960
      Picture         =   "Form1_bk.frx":22D4F
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label olLocal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   0
      EndProperty
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
      TabIndex        =   75
      Top             =   8760
      Width           =   75
   End
   Begin VB.Image oImg_Logo1 
      Height          =   3240
      Left            =   360
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   3960
   End
   Begin VB.Label oLVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   65
      Top             =   360
      Width           =   105
   End
   Begin VB.Label olActivacion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
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
      TabIndex        =   62
      Top             =   120
      Width           =   75
   End
   Begin VB.Label olMessageVIP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No hay mensajes."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   555
      Left            =   0
      TabIndex        =   61
      Top             =   7920
      Visible         =   0   'False
      Width           =   4275
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer1 
      Height          =   240
      Left            =   3360
      TabIndex        =   60
      Top             =   6120
      Width           =   300
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   529
      _cy             =   423
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer2 
      CausesValidation=   0   'False
      Height          =   3240
      Left            =   360
      TabIndex        =   59
      Top             =   1920
      Width           =   3900
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   -1  'True
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6879
      _cy             =   5715
   End
   Begin VB.Label olMensaje_Video 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARGANDO VIDEO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   480
      TabIndex        =   58
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label olMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No hay mensajes."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   480
      Left            =   900
      TabIndex        =   51
      Top             =   6840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label olCred_Msg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INSERTE ¢ 0.25"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   660
      Left            =   720
      TabIndex        =   50
      Top             =   7320
      Width           =   3435
   End
   Begin VB.Label olCreditos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREDITOS(0)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   1080
      TabIndex        =   49
      Top             =   1440
      Width           =   2355
   End
   Begin VB.Label oLTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Haga su selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4200
      TabIndex        =   76
      Top             =   0
      Width           =   7560
      WordWrap        =   -1  'True
   End
   Begin VB.Label olPaginas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Página (1) ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   2760
      TabIndex        =   1
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2295
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ilCnt_Err As Integer

Private Sub Check1_Click()
If Check1.value = 1 Then
    Me.oChk_FndC.Enabled = True
    Me.oChk_FndP.Enabled = True
Else
    Me.oChk_FndC.Enabled = False
    Me.oChk_FndP.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Call DBCan_Cheker(True)
End Sub

Private Sub Command2_Click()
Call AlwaysOnTop(Main_Form, False)
Call Go_Service
Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
Me.otCodigo.SetFocus
End Sub

Private Sub Command3_Click()
Call AlwaysOnTop(Main_Form, False)
'Me.Hide
ControlPanel.Show vbModal
'Me.Show
Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
Me.otCodigo.SetFocus
End Sub

Private Sub Form_DblClick()
Unload Video_Form
Unload Me
End
End Sub

Private Sub Form_Deactivate()
Video_Form.MediaPlayer3.Close
Call Salvar_Temas
End Sub

Private Sub Form_Load()
ilCnt_Err = 0
sgCmdLine = VBA.Command$
If sgCmdLine = "" Then
    sgCmdLine = "NO PARAMETER"
End If
sgParms = VBA.Split(sgCmdLine, " ")
If App.PrevInstance Then
    MsgBox "La aplicacion solicitada [" & App.EXEName & "], ya se esta ejecutando!!!", vbInformation
    End
End
End If
Dim MiValor As String
VBA.Randomize
'---------------------Carga entorno de variables------------------------
On Error GoTo Solve_error
igDelay_Ret_Gen = 0
igNoDuplicT = 0
igDelay_Del_Dig_Can = 0
igLen2 = 0: igNo_RgAt = 0
igAct_PgG = 1: igTot_PgG = 0: igTot_PgC = 0
igAct_PgD = 1: igTot_PgD = 0: igTot_PgC = 0
igMax_Gen = 13: igMax_Dis = 9: igMax_Can = 12 'valores fijos su valor no pueded ser superior
igInd_Bon = 0
igNext_Bonus = 0
bgBlinkPag = False
bgPopular = False
bgVIP = False
bgIs_Video = False: bgIs_Publi = False
bgWMP_Busy = False
igCont_Sin = 0
Me.oLVersion.Caption = "Ver." & VBA.Trim(VBA.Str(App.Major)) & "." & VBA.Trim(VBA.Str(App.Minor)) & "." & VBA.Trim(VBA.Str(App.Revision))
Call Colocar_Frames
Call Registra_Dll
Call Save_Defa_Path
Call Get_System_Path(sParam)
'---------------------Carga Variables de rutas----------------------------
Dim iTmp As String
Dim sTmp As String

sgDir_odb = sParam(1)
sgDir_Tmp = sParam(2)
sgDir_Fls = sParam(3)
sgDir_Img = sParam(4)
sgDir_Mp3 = sParam(5)
sgDir_Pub = sParam(6)
sgFec_iAc = sParam(7)
sgFec_Fac = sParam(8)
sgSer_Mac = sParam(9)
sgNom_Loc = VBA.Trim(sParam(10))
igCnt_CRS = sParam(11)
sgFle_Fon = sParam(13)
sgSer_CPU = sParam(14)
igDelay_Return_Gen = VBA.Int(VBA.Val(sParam(15)))
igDelay_Return_Dis = VBA.Int(VBA.Val(sParam(16)))
igDelay_Bonus_Vid = VBA.Int(VBA.Val(sParam(17)))
bAcum_Cre = VBA.Int(VBA.Val(sParam(19)))
If bAcum_Cre = 1 Then
    igCnt_CR = VBA.Int(VBA.Val(sParam(18)))
End If
igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
igLim_Cred = VBA.Int(VBA.Val(sParam(20)))
igKeep_Cred = VBA.Int(VBA.Val(sParam(21)))
igMixe_Popu = VBA.Int(VBA.Val(sParam(22)))
sgKb_Crd1 = sParam(23)
sgKb_Crd2 = sParam(25)
sgKb_Del = sParam(27)
sgKb_Ret = sParam(29)
sgKb_ResM = sParam(31)
sgKb_ResA = sParam(33)
sgKb_Pop = sParam(35)
sgKb_VID = sParam(36)
sgKb_VIP = sParam(37)
sgKb_BonC = VBA.Int(VBA.Val(sParam(38)))
sgKb_UP = sParam(39)
sgKb_Vef = sParam(40)
sgKb_DN = sParam(41)
sgWin_Key = sParam(43)
bgVideoLabel = IIf(Val(sParam(44)) = 1, True, False)
bgDiscLabel = IIf(Val(sParam(45)) = 1, True, False)
bgKeep_On_Top = VBA.IIf(sParam(46) = 0, False, True)
igScr_Alone = VBA.Int(VBA.Val(sParam(47)))
igNoDuplicT = VBA.Int(VBA.Val(sParam(48)))
sgWin_Key = sParam(43)
igLeftDisk = 700
Call Check_Other
If sgParms(0) = "ACTIVATE" Then
    MiValor = InputBox("Inserte el código de seguridad", "Flamingo Magic Game", "", 100, 100)
    If VBA.Val(MiValor) <> 2527 Then
        End
    Else
        Act_Form.Show vbModal
        End
    End If
ElseIf sgParms(0) = "SERVICE" Then
    MiValor = InputBox("Inserte el código de seguridad", "Flamingo Magic Game", "", 100, 100)
    If VBA.Val(MiValor) <> 2527 Then
        End
    Else
        Call Go_Service
        End
    End If
End If
On Error GoTo Solve_error
'---------------------Verifica activación y serial de la pc--------------
If sgNom_Loc <> VBA.UCase("SIN ASIGNACIÓN!") Then
    Me.olLocal.Caption = "[" & sgNom_Loc & "] "
Else
    Me.olLocal.Caption = ""
End If
Me.olLocal.Caption = Me.olLocal.Caption & sgWin_Key
Dim sTmp1 As Variant
Dim sMensage As String
sMensage = "La copia del sistema no ha sido debidamente instalada o no ha sido activada"
sTmp1 = Lee_Serial
sTmp1 = Left$(sTmp1, 4) & "-" & Right$(sTmp1, 4)
If sgSer_Mac <> sTmp1 Then
    Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de DISCO")
    End
End If
sTmp1 = Get_CPU_Id
If sgSer_CPU <> sTmp1 Then
    Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de MÀQUINA")
    End
End If
sTmp1 = Null
If VBA.DateValue(sgFec_iAc) = VBA.DateValue(sgFec_Fac) Then
    Call MsgBox(sMensage, vbCritical, "La copia del sistemas debe ser activada")
    End
End If
If VBA.Date < VBA.DateValue(sgFec_iAc) Then
    Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
    End
End If
If VBA.Date > VBA.DateValue(sgFec_Fac) Then
    Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
    End
End If
Me.olActivacion.Caption = "Próxima activación: " & VBA.Format(sgFec_Fac, "dd/MM/yyyy")
Call Set_Tmp_DBF

Me.oLst_Temas_Video.Path = sgDir_Mp3
Me.oLst_Temas_Video.Pattern = "*.MPG"
Me.oLst_Temas_Video.Refresh

Me.oFrame_Dis.TransparentBackground = True
Me.oFrame_Gen.TransparentBackground = True
Me.oFrame_Can.TransparentBackground = True

Me.oTM_codigo2.Interval = igDelay_Return_Dis * 1000
Me.oTM_codigo2.Enabled = False

Call Refresh_Creditos(Me)
'---------------------PUBLICIDAD----------------------------
Call Conectar_DBPub
'----------------------GENEROS------------------------------
Call Inicial_Gen
'----------------------Otros------------------------------
otNot_Found_List.Text = ""
Call Cargar_Temas
Call Limpia_Dis
Call Limpia_Can
If igScr_Alone = 0 Then
    Video_Form.Show
End If
If bgKeep_On_Top = True Then
    Call AlwaysOnTop(Main_Form, True)
End If
Exit Sub

Solve_error:
Call ControlError
If ilCnt_Err = 0 Then
    Call MsgBox("Hay algún problema en el sistema, usualmente relacionado con las rutas del sistema, Favor ejecutar el comando \ROCKOLA.EXE SERVICE, para acceder al menú de servicio...", vbCritical, "Verificar configuración del sistema")
End If
Resume Next
ilCnt_Err = ilCnt_Err + 1
End Sub

Private Sub Conectar_DBGen()
Dim sSql As String
sSql = "SELECT * FROM File01"
With Me.oDC_GEN
    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .Refresh
End With
Me.oLst_Gen.Refresh
End Sub

Private Function Cargar_INF_Gen(Optional ByRef pMax_Gen As Integer)
Dim iNum_reg As Integer
Dim iCnt_Pag As Integer
Dim sTmp_Cad As String
Dim i As Integer
iNum_reg = 1
With Me.oDC_GEN.Recordset
    pMax_Gen = .RecordCount()
    iCnt_Pag = 1
    .MoveFirst
    Do While Not .EOF
        If iNum_reg > igMax_Gen Then
            iCnt_Pag = iCnt_Pag + 1
            iNum_reg = 1
        End If
        sTmp_Cad = .Fields("Id_Ord").value & " - " & .Fields("DESCRI").value
        aPag_Gen(iCnt_Pag).Genero(iNum_reg).De_Gen = sTmp_Cad
        aPag_Gen(iCnt_Pag).Genero(iNum_reg).ID_GEN = .Fields("ID_GEN").value
        aPag_Gen(iCnt_Pag).Genero(iNum_reg).ID_ORD = .Fields("Id_Ord").value
        aPag_Gen(iCnt_Pag).Genero(iNum_reg).No_POS = iNum_reg
        aPag_Gen(iCnt_Pag).No_Rgs = iNum_reg
        iNum_reg = iNum_reg + 1
        .MoveNext
    Loop
End With
igTot_PgG = iCnt_Pag
End Function

Private Function Desactiva_Genero(ByRef pFlag As Boolean)
Dim i As Integer
i = 0
Me.oFrame_Gen.Visible = Not pFlag
End Function

Private Function Desactiva_Disco(ByRef pFlag As Boolean)
Dim i As Integer
Me.oFrame_Dis.Visible = Not pFlag
End Function

Private Function Desactiva_Cancion(ByRef pFlag As Boolean)
Dim i As Integer
i = 0
Me.oFrame_Can.Visible = Not pFlag
End Function

Private Sub Form_Unload(Cancel As Integer)
Call Salvar_Temas
Call Upd_Cnt(igCnt_CRS)
Call ogVFP9.Set_Files_Close
Set Video_Form = Nothing
Set Main_Form = Nothing
End Sub

Private Sub Label2_Click()

End Sub

Private Sub MediaPlayer1_MediaError(ByVal pMediaObject As Object)
Main_Form.olMessage.Visible = True
Main_Form.olMessage.Caption = "TEMA NO DISPONIBLE"
Main_Form.oTime_Mensajes.Enabled = True
Call Remove_Temes
If igKeep_Cred = 0 Then
    igCnt_CR = igCnt_CR + 1
End If
Call Refresh_Creditos(Me)
Sleep 3 '*1000 'Implements a 3 second delay
VBA.SendKeys ("S")
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal NewState As Long)
Select Case NewState
Case Is = wmppsMediaEnded
    bgIs_Video = False
    bgWMP_Busy = False
    Call Muestra_Tema_Det
    If igScr_Alone = 0 Then
        Video_Form.MediaPlayer3.Close
    End If
    Main_Form.MediaPlayer2.Close
    If igDelay_Bonus_Vid > 0 Then
        igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
    End If
    If bgIs_Video = True Then
        Exit Sub
    End If
    Call Remove_Temes
    igCont_Sin = 0
Case Is = wmppsPlaying
    Me.otCargador_Music.Enabled = True
    bgIs_Video = False
    bgWMP_Busy = False
End Select
End Sub

Private Sub MediaPlayer2_MediaError(ByVal pMediaObject As Object)
If igScr_Alone = 1 Then
    Main_Form.olMessage.Visible = True
    Main_Form.olMessage.Caption = "TEMA NO DISPONIBLE"
    Main_Form.oTime_Mensajes.Enabled = True
    Call Remove_Temes
    igCnt_CR = igCnt_CR + 1
    Call Refresh_Creditos(Main_Form)
    Sleep 3 '* 1000 'Implements a 3 second delay
    VBA.SendKeys ("S")
End If
End Sub

Private Sub MediaPlayer2_PlayStateChange(ByVal NewState As Long)
If igScr_Alone = 1 Then
    Select Case NewState
    Case Is = wmppsMediaEnded
        bgWMP_Busy = False
        Call Muestra_Tema_Det
        'Video_Form.MediaPlayer3.Close
        '*Main_Form.MediaPlayer2.Close
        igCont_Sin = 0
        If igDelay_Bonus_Vid > 0 Then
            igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
        End If
        If bgIs_Video = False Then
            Exit Sub
        End If
        Call Remove_Temes
    Case Is = wmppsPlaying
        If igCont_Sin > 0 Then
            Exit Sub
        End If
        bgWMP_Busy = True
        '*Main_Form.MediaPlayer2.URL = Video_Form.MediaPlayer3.URL
        '*Main_Form.MediaPlayer2.settings.mute = True
        '*Main_Form.MediaPlayer2.Controls.currentPosition = Video_Form.MediaPlayer3.Controls.currentPosition
        '*Main_Form.MediaPlayer2.Controls.play
        igCont_Sin = igCont_Sin + 1
    End Select
End If
End Sub

Private Sub oChk_FndC_Click()
Me.otRuteExternal.Text = ""
If Me.oChk_FndC.value = 1 Then
    Me.oGetRute.Enabled = True
    Me.otRuteExternal.Enabled = True
Else
    Me.oGetRute.Enabled = False
    Me.otRuteExternal.Enabled = False
End If
End Sub

Private Sub oChk_FndP_Click()
Me.otRuteExternal2.Text = ""
If Me.oChk_FndP.value = 1 Then
    Me.oGetRute2.Enabled = True
    Me.otRuteExternal2.Enabled = True
Else
    Me.oGetRute2.Enabled = False
    Me.otRuteExternal2.Enabled = False
End If
End Sub

Private Sub oGeneral_Timer_Timer()
'------------------------------------------------------
If bgBlinkPag = True Then
    If Me.olPaginas.Visible = True Then
        If (Me.olPaginas.ForeColor) <> &HFFFF& Then
            Me.olPaginas.ForeColor = &HFFFF&
        Else
            Me.olPaginas.ForeColor = &H0&
        End If
    End If
    If olPaginas2.Visible = True Then
        If (Me.olPaginas2.ForeColor) <> &HFFFF& Then
            Me.olPaginas2.ForeColor = &HFFFF&
        Else
            Me.olPaginas2.ForeColor = &H0&
        End If
    End If
Else
    Me.olPaginas.ForeColor = &HFFFF&
    Me.olPaginas2.ForeColor = &HFFFF&
End If
'------------------------------------------------------
If VBA.Trim(Me.otCodigo.Text) = "" Or _
    igNext_Return_Gen = 0 Then
    Exit Sub
End If
If igDelay_Return_Gen > 0 Then
    'If igCan_Sel = "" Then
        'Sólo regresa a género si esta en la pantalla de discos
        If igNext_Return_Gen <= (Hour(Time()) * 60) + Minute(Time()) Then
            Call Retrocede
            Call Retrocede
            igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
        End If
    'End If
End If
End Sub

Private Sub oGetRute_Click()
Dim lcSelectedPath As String
Me.otRuteExternal.Text = ""
Me.Bbgetdir1.FocusedDirectory = "C:\"
Me.Bbgetdir1.ListAutoCenter = True
Me.Bbgetdir1.StatusText = "Origen del Cansionero?"
lcSelectedPath = Me.Bbgetdir1.ShowDirectoryListEx(1) + "\"
If lcSelectedPath <> "" Then
    Me.otRuteExternal.Text = VBA.Trim(lcSelectedPath)
Else
    Me.otRuteExternal.Text = ""
End If
Me.otRuteExternal.Refresh
End Sub

Private Sub oGetRute2_Click()
Dim lcSelectedPath As String
Me.otRuteExternal2.Text = ""
Me.Bbgetdir1.FocusedDirectory = "C:\"
Me.Bbgetdir1.ListAutoCenter = True
Me.Bbgetdir1.StatusText = "Origen de Caratulas?"
lcSelectedPath = Me.Bbgetdir1.ShowDirectoryListEx(1) + "\"
If lcSelectedPath <> "" Then
    Me.otRuteExternal2.Text = VBA.Trim(lcSelectedPath)
Else
    Me.otRuteExternal2.Text = ""
End If
Me.otRuteExternal2.Refresh
End Sub

Private Sub olCreditos_Click()
Call Show_Hide_Service
End Sub

Private Sub oSetFocus_Codigo_GotFocus()
otCodigo.Text = ""
otCodigo.SetFocus
End Sub

Private Sub otCargador_Music_Timer()
On Error GoTo Solve_error
Dim iCont As Integer
Dim sCadenas As String
If Me.oLst_A_Tocar.List(0) = "" Then
    Me.oImg_c_Video.Visible = False
    If igDelay_Bonus_Vid > 0 Then
       If igNext_Bonus <= (Hour(Time()) * 60) + Minute(Time()) Then
            igInd_Bon = igInd_Bon + 1
            Me.oLst_Temas_Video.ListIndex = igInd_Bon - 1
            sCadenas = "99999,999999,VIDEO BONUS," & Me.oLst_Temas_Video.Path & "\" & Me.oLst_Temas_Video.filename
            Me.oLst_A_Tocar.AddItem sCadenas
             igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
        End If
    End If
    If FileExist(sgDir_Img & "\Logo1.bmp") Then
        Me.oImg_Logo1.Picture = LoadPicture(sgDir_Img & "\Logo1.bmp")
        Me.oImg_Logo1.Visible = True
        Me.MediaPlayer2.Close
        Me.MediaPlayer2.Visible = False
    Else
        Me.oImg_Logo1.Picture = LoadPicture()
        Me.oImg_Logo1.Visible = False
        Me.MediaPlayer2.Visible = True
    End If
    If igScr_Alone = 0 Then
        If FileExist(sgDir_Img & "\Logo1.bmp") Then
            Video_Form.oImg_Logo1.Picture = LoadPicture(sgDir_Img & "\Logo1.bmp")
            Video_Form.oImg_Logo1.Visible = True
            Video_Form.MediaPlayer3.Close
            Video_Form.MediaPlayer3.Visible = False
        Else
            Video_Form.oImg_Logo1.Picture = LoadPicture()
            Video_Form.oImg_Logo1.Visible = False
            Video_Form.MediaPlayer3.Visible = True
        End If
    End If
    'Me.otCodigo.SetFocus
    Exit Sub
Else
    Me.oImg_Logo1.Visible = False
    Me.oImg_Logo1.Picture = LoadPicture()
    Me.MediaPlayer2.Visible = True
    
    If igScr_Alone = 0 Then
        Video_Form.oImg_Logo1.Visible = False
        Video_Form.oImg_Logo1.Picture = LoadPicture()
        Video_Form.MediaPlayer3.Visible = True
    End If
    Call Cargar_Musica
    'Me.otCodigo.SetFocus
End If
Exit Sub

Solve_error:
Call ControlError
Resume Next
End Sub

Private Sub otCargador_Video_Timer()
On Error GoTo Solve_error
'Si no hay mas temas en lista para tocar no mostrar màs videeos.
If Me.oLst_A_Tocar.List(0) = "" Then
    'Me.otCodigo.SetFocus
    Exit Sub
End If
'Si no hay mas videos que presentar, salir.
If Me.oLst_Pub.List(0) = "" Then
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "LA LISTA DE PUBLICIDAD ESTA VACÍA!"
    Me.oTime_Mensajes.Enabled = True
    Exit Sub
End If

Dim iLimCnt As Integer
Dim sFle_MpG As String
Dim sFle_Tmp As String
Dim aDet() As String
sFle_Tmp = Me.MediaPlayer1.URL
If VBA.UCase(VBA.Right(sFle_Tmp, 3)) <> "MP3" Then
    Exit Sub
End If
iLimCnt = Me.oLst_Pub.ListCount - 1
If iLimCnt <= 0 Then
    Exit Sub
End If
If igInd_Pub = 0 Then
    igInd_Pub = 1
End If
Me.oLst_Pub.ListIndex = (igInd_Pub - 1)
'sFle_MpG = VBA.Trim(Me.oLst_Pub.Text)
sFle_MpG = Me.oLst_Pub.Path & "\" & VBA.Trim(Me.oLst_Pub.filename)
If FileExist(sFle_MpG) Then
    If igScr_Alone = 0 Then
        If bgIs_Video = False Then
            If ObPlayer_Ocupado(Video_Form.MediaPlayer3) Then
                Exit Sub
            End If
        Else
            If ObPlayer_Ocupado(Me.MediaPlayer1) = True Then
                Exit Sub
            End If
        End If
        Me.olMensaje_Video.Caption = "CARGANDO VIDEO"
        Me.olMensaje_Video.Visible = True
        Video_Form.MediaPlayer3.Close
        Video_Form.MediaPlayer3.URL = VBA.Trim(sFle_MpG)
        Video_Form.MediaPlayer3.settings.mute = True
        Video_Form.MediaPlayer3.Controls.play
    Else
        If bgIs_Video = False Then
            If ObPlayer_Ocupado(Me.MediaPlayer2) Then
                Exit Sub
            End If
        Else
            If ObPlayer_Ocupado(Me.MediaPlayer1) = True Then
                Exit Sub
            End If
        End If
        Me.olMensaje_Video.Caption = "CARGANDO VIDEO"
        Me.olMensaje_Video.Visible = True
        Me.MediaPlayer2.Close
        Me.MediaPlayer2.URL = VBA.Trim(sFle_MpG)
        Me.MediaPlayer2.settings.mute = True
        Me.MediaPlayer2.Controls.play
    End If
Else
    If sFle_MpG = "" Then
        Me.olMessage.Visible = True
        Me.olMessage.Caption = "PUBLICIDAD NO ENCONTRADO!"
        Me.oTime_Mensajes.Enabled = True
    End If
End If
If igInd_Pub < igTot_Pub Then
    igInd_Pub = igInd_Pub + 1
Else
    igInd_Pub = 0
End If
Exit Sub

Solve_error:
Call ControlError
Resume Next
End Sub

Private Sub otCodigo_Change()
On Error GoTo Solve_error
Dim sValue As String, sValSel As String
Dim iNum_Pag As Integer, iNum_Pos As Integer, iLen As Integer

sValue = VBA.Trim(otCodigo.Text)
igLen = VBA.Len(sValue)

If igLen > 0 Then
    Call Show_Hide_Service(0)
    If igKeep_Cred = 0 Then
        If igCnt_CR <= 0 Then
            igCnt_CR = 0
            If Me.otCodigo.Tag = "" Then
                Me.otCodigo.EditMask = "##-##"
                Me.otCodigo.Tag = "1"
                Me.oSetFocus_Codigo.SetFocus
                Exit Sub
            End If
        Else
            If Me.otCodigo.EditMask = "##-##" Then
                Me.otCodigo.EditMask = "##-##-##"
                Me.otCodigo.Tag = "0"
            End If
        End If
    End If
Else
    igGen_Sel = "": igDis_Sel = "": igCan_Sel = ""
End If
Me.olPosCod.Caption = "Posición -> " & VBA.Trim(VBA.Str(igLen))
Select Case igLen
Case Is = 0
    igGen_Sel = ""
    igDis_Sel = ""
    igCan_Sel = ""
    igNext_Return_Gen = 0
    Me.Image2.Picture = LoadPicture()
    Call Inicial_Gen
Case 1 To 2
    Dim sGen_Ret As String
    sGen_Ret = ""
    igGen_Sel = ""
    igDis_Sel = ""
    igCan_Sel = ""
    Call Desactiva_Cancion(True)
    Call Desactiva_Disco(True)
    Call Desactiva_Genero(False)
    Me.Image2.Picture = LoadPicture()
    If igLen < 2 Then
        Exit Sub
     End If
    If Busca_Sel_1(sValue, sGen_Ret, iNum_Pag, iNum_Pos) = False Then
        If igLen = 2 Then
            Me.oLTitulo.Caption = "Género no existe..."
        Else
            Me.oLTitulo.Caption = ""
        End If
        Call Retrocede
        igAct_PgG = 1
        Exit Sub
    Else
        Me.oLst_Gen.BoundText = sGen_Ret
        Me.oLTitulo.Caption = Me.oLst_Gen.Text
    End If
    igAct_PgD = 1
    igAct_PgC = 1
    igAct_PgG = iNum_Pag
    Call Desactiva_Genero(True)
    'Call Desactiva_Disco(False)
    '***********************DISCOS************************
    Call Conectar_DBDis(sGen_Ret)
    Call Cargar_INF_Dis(sGen_Ret, igMax_RgD)
    Call Cargar_Pag_Dis(1, igMax_RgD)
    'Me.oLst_Gen.BoundText = igGen_Sel
    'Me.oLTitulo.Caption = Me.oLst_Gen.Text
    Call Desactiva_Disco(False)
    igGen_Sel = sGen_Ret
    bgVIP = False
    igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
Case 3 To 4
    Dim sDis_Ret As String
    Dim sFle_Img As String
    sDis_Ret = ""
    igDis_Sel = ""
    igCan_Sel = ""
    Call Desactiva_Cancion(True)
    Call Desactiva_Genero(True)
    Me.Image2.Picture = LoadPicture()
    If igLen < 4 Then
        Call Desactiva_Disco(False)
        Exit Sub
     End If
    sValSel = VBA.Mid(sValue, 3, 4)
    If Busca_Sel_2(sValSel, sDis_Ret, iNum_Pag, iNum_Pos, sFle_Img) = False Then
        If igLen >= 4 Then
            Me.oLTitulo.Caption = "Disco no Existe..."
        Else
            Me.oLTitulo.Caption = ""
        End If
        igAct_PgD = 1
        Call Retrocede
        'Call Limpia_Dis
        Exit Sub
    Else
        Me.oLst_Disc.BoundText = sDis_Ret
        Me.oLTitulo.Caption = Me.oLst_Disc.Text
    End If
    igAct_PgC = 1
    igAct_PgD = iNum_Pag
    If VBA.Right(VBA.Trim(sFle_Img), 1) = "\" Then
        sFle_Img = ""
    End If
    Me.Image2.Picture = LoadPicture(sFle_Img)
    Call Desactiva_Genero(True)
    Call Desactiva_Disco(True)
    Call Desactiva_Cancion(True)
    '***********************CANCION************************
    Call Conectar_DBCan(igGen_Sel, sDis_Ret)
    Call Cargar_INF_Can(igGen_Sel, sValSel, sDis_Ret, igMax_RgC)
    Call Cargar_Pag_Can(1, igMax_RgC)
    Call Desactiva_Cancion(False)
    igDis_Sel = sDis_Ret
    bgVIP = False
Case 5 To 6
    Dim sFle_Mp3 As String
    Dim sCan_Ret As String
    Dim sCad_Ato As String
    sCan_Ret = ""
    sCad_Ato = ""
    igCan_Sel = ""
    Call Desactiva_Genero(True)
    Call Desactiva_Disco(True)
    Call Desactiva_Cancion(False)
    If igLen < 6 Then
        Exit Sub
     End If
    sValSel = VBA.Mid(sValue, 5, 6)
    If Busca_Sel_3(sValSel, sCan_Ret, sFle_Mp3, iNum_Pag, iNum_Pos, sValue, sCad_Ato) = False Then
        If iLen >= 6 Then
            Me.oLTitulo.Caption = "Tema no existe..."
        Else
            Me.oLTitulo.Caption = ""
        End If
        igAct_PgC = 1
        Call Retrocede
        'Call Limpia_Can
        Exit Sub
    Else
        Me.oLst_Canc.BoundText = sCan_Ret
        Me.oLTitulo.Caption = Me.oLst_Canc.Text
    End If
    igAct_PgC = iNum_Pag
    Call Cargar_Pag_Can(iNum_Pag, igMax_RgC)
    igCan_Sel = sCan_Ret
    If igLen = 6 Then
        If FileExist(sFle_Mp3) Then
            If bgVIP = False Then
                If Not VBA.UCase(VBA.Right(VBA.Trim(sFle_Mp3), 3)) <> "MP3" Then
                    If igNoDuplicT = 1 Then
                        If Busca_ATocar(sCad_Ato) = False Then
                            Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                        End If
                    Else
                        Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                    End If
                    If igKeep_Cred = 0 Then
                        igCnt_CR = igCnt_CR - 1
                    End If
                Else
                    If sgKb_VID = 0 Then
                        If igNoDuplicT = 1 Then
                            If Busca_ATocar(sCad_Ato) = False Then
                                Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                            End If
                        Else
                            Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                        End If
                        If igKeep_Cred = 0 Then
                            igCnt_CR = igCnt_CR - 1
                        End If
                    Else
                        If sgKb_VID > igCnt_CR Then
                            Me.olMessage.Visible = True
                            Me.olMessage.Caption = "CREDITOS INSUFICIENTES!"
                            Me.oTime_Mensajes.Enabled = True
                            Call Retrocede
                            Exit Sub
                        Else
                            If igKeep_Cred = 0 Then
                                igCnt_CR = igCnt_CR - sgKb_VID
                            End If
                            If igNoDuplicT = 1 Then
                                If Busca_ATocar(sCad_Ato) = False Then
                                    Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                                End If
                            Else
                                Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                            End If
                        End If
                    End If
                End If
            Else
                If Me.oLst_A_Tocar.ListCount <= 0 Then
                    Call Me.oLst_A_Tocar.AddItem(sCad_Ato)
                Else
                    Call Me.oLst_A_Tocar.AddItem(sCad_Ato, 1)
                End If
                If igKeep_Cred = 0 Then
                    igCnt_CR = igCnt_CR - 2
                End If
                If igCnt_CR < 0 Then
                    igCnt_CR = 0
                End If
                bgVIP = False
                Me.olMessageVIP.Visible = False
            End If
            Me.oLst_A_Tocar.Refresh
            Call Refresh_Creditos(Me)
        Else
            Me.olMessage.Visible = True
            Me.olMessage.Caption = "NO ENCONTRADO!"
            Me.oTime_Mensajes.Enabled = True
            Exit Sub
        End If
        If igKeep_Cred = 0 Then
            If igCnt_CR > 0 Then
                Me.olMessage.Visible = True
                Me.olMessage.Caption = "TEMA FUE ANEXADO!"
                Me.oTime_Mensajes.Enabled = True
                'Call Retrocede
                Me.oTM_codigo2.Enabled = True
            Else
                '----------------------GENEROS------------------------------
                Call Desactiva_Cancion(True)
                Call Desactiva_Disco(True)
                Call Desactiva_Genero(False)
                Call Conectar_DBGen
                Call Cargar_INF_Gen
                Call Cargar_Pag_Gen(1, igMax_RgG)
                Me.oSetFocus_Codigo.SetFocus
            End If
        Else
            Me.olMessage.Visible = True
            Me.olMessage.Caption = "TEMA FUE ANEXADO!"
            Me.oTime_Mensajes.Enabled = True
            'Call Retrocede
            Me.oTM_codigo2.Enabled = True
        End If
    End If
End Select
Exit Sub

Solve_error:
Call ControlError
Resume Next

End Sub

Private Sub otCodigo_KeyPress(KeyAscii As Integer)
Dim iLimCnt As Integer
igKeyAscii = KeyAscii

If VBA.IsNumeric(VBA.Chr(KeyAscii)) Then
    Exit Sub
End If
If KeyAscii = 8 Then
    Me.SetFocus
    Exit Sub
End If
'133
If Inlist(VBA.UCase(VBA.Chr(KeyAscii)), sgKb_Vef) Then
    bgExit = False
    If TBack4.Visible = False Then
        TBack4.Visible = True
    Else
        TBack4.Visible = False
    End If
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_Ret) Then
    Call Retrocede
    Exit Sub
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_ResM) Then
    'Sección que se ejecuta si se preciona [S/s] (Reset single active music)
    Video_Form.MediaPlayer3.Close
    Main_Form.MediaPlayer2.Close
    Main_Form.MediaPlayer1.Close
    If igInd_Pub < igTot_Pub Then
        igInd_Pub = igInd_Pub + 1
    Else
        igInd_Pub = 0
    End If
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "TEMA IGNORADO"
    Me.oTime_Mensajes.Enabled = True
    If Me.oLst_A_Tocar.ListCount > 0 Then
        Me.oLst_A_Tocar.RemoveItem (0)
    Else
        Me.olMessage.Visible = True
        Me.olMessage.Caption = "TEMAS AGOTADOS"
        Me.oTime_Mensajes.Enabled = True
    End If
    Call Inicial_Gen
    Call Muestra_Tema_Det
    Call Refresh_Creditos(Me)
    bgWMP_Busy = False
    igCont_Sin = 0
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_ResA) Then
    'Sección que se ejecuta si se preciona [R/r] (Resert all)
    igCnt_CR = 0
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "CREDITOS ANULADOS"
    Me.oTime_Mensajes.Enabled = True
    
    Me.oLst_A_Tocar.Clear
    Video_Form.MediaPlayer3.Close
    Main_Form.MediaPlayer2.Close
    Main_Form.MediaPlayer1.Close
    Main_Form.oSetFocus_Codigo.SetFocus
    If igInd_Pub < igTot_Pub Then
        igInd_Pub = igInd_Pub + 1
    Else
        igInd_Pub = 0
    End If
    'If Me.oLst_Pub.ListCount > 0 Then
    '    Call Me.oLst_Pub.RemoveItem(0)
    'End If
    Call Inicial_Gen
    Call Muestra_Tema_Det
    Call Refresh_Creditos(Me)
    bgWMP_Busy = False
    igCont_Sin = 0
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_Crd1) Then
    'Sección que se ejecuta si se preciona [+] (Crédito)
    If igCnt_CR >= igLim_Cred Then
        Main_Form.olMessage.Visible = True
        Main_Form.olMessage.Caption = "REVISAR MONEDERO!"
        Main_Form.oTime_Mensajes.Enabled = True
        Exit Sub
    End If
    igCnt_CR = igCnt_CR + 2
    igCnt_CRS = igCnt_CRS + 2
    Call Refresh_Creditos(Me)
    If sgKb_BonC > 0 Then
        If igCnt_CR = 8 Then
            Main_Form.olMessage.Visible = True
            Main_Form.olMessage.Caption = "CRED. PROMOSIÓN [" & VBA.Trim(VBA.Str(sgKb_BonC)) & "]"
            Main_Form.oTime_Mensajes.Enabled = True
            Call Sleep(3)
            igCnt_CR = igCnt_CR + sgKb_BonC
            igCnt_CRS = igCnt_CRS + sgKb_BonC
        End If
    End If
    Call Refresh_Creditos(Me)
    Me.otCodigo.EditMask = "##-##-##"
    Me.otCodigo.Tag = ""
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_Del) Then
    'Sección que se ejecuta si se preciona [-] (Crédito)
    If igKeep_Cred = 0 Then
        If igCnt_CR > 0 Then
            igCnt_CR = igCnt_CR - 1
            igCnt_CRS = igCnt_CRS - 1
            Call Refresh_Creditos(Me)
        End If
    End If
    Exit Sub
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_Crd2) Then
    'Sección que se ejecuta si se preciona [N/n] (credtos de prueba)
    igCnt_CR = igCnt_CR + 1
    igCnt_CRP = igCnt_CRP + 1
    Call Refresh_Creditos(Me)
    Me.otCodigo.EditMask = "##-##-##"
    Me.otCodigo.Tag = ""
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_Pop) Then
    'Sección que se ejecuta si se preciona [P/p] (Popular)
    If igLen < 4 Then
        Exit Sub
    End If
    If (igLen > 3) And (igLen < 7) Then
        If igKeep_Cred = 0 Then
            If (igCnt_CR < Me.oLst_Popular.ListCount) Then
                Me.olMessage.Caption = "CRÉDITOS INSUFICIENTES"
                Me.olMessage.Visible = True
                Me.oTime_Mensajes.Enabled = True
            Else
                Me.olMessage.Caption = "CARGANDO POPULAR"
                Me.olMessage.Visible = True
                Me.oTime_Mensajes.Enabled = True
                For iLimCnt = 1 To (Me.oLst_Popular.ListCount)
                    igCnt_CR = igCnt_CR - 1
                    Call Refresh_Creditos(Me)
                    Call Me.oLst_A_Tocar.AddItem(Me.oLst_Popular.List(0))
                    Call Me.oLst_Popular.RemoveItem(0)
                Next iLimCnt
                If igMixe_Popu > 0 Then
                    Call Ramdom_List(Me.oLst_A_Tocar, Me.oBkList)
                End If
            End If
        Else
            Me.olMessage.Caption = "CARGANDO POPULAR"
            Me.olMessage.Visible = True
            Me.oTime_Mensajes.Enabled = True
            For iLimCnt = 1 To (Me.oLst_Popular.ListCount)
                If igKeep_Cred = 0 Then
                    igCnt_CR = igCnt_CR - 1
                End If
                Call Refresh_Creditos(Me)
                Call Me.oLst_A_Tocar.AddItem(Me.oLst_Popular.List(0))
                Call Me.oLst_Popular.RemoveItem(0)
            Next iLimCnt
            If igMixe_Popu > 0 Then
                Call Ramdom_List(Me.oLst_A_Tocar, Me.oBkList)
            End If
        End If
        Call Inicial_Gen
        Me.oSetFocus_Codigo.SetFocus
        Exit Sub
    End If
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_VIP) Then
    'Sección que se ejecuta si se preciona [V/v] (VIP)
    If igLen < 4 Then
        Exit Sub
    End If
    If (igLen > 3) And (igLen < 7) Then
        If igKeep_Cred = 0 Then
            If (igCnt_CR < 2) Then
                Me.olMessage.Caption = "CRÉDITOS INSUFICIENTES"
                Me.olMessage.Visible = True
                Me.oTime_Mensajes.Enabled = True
                bgVIP = False
            Else
                Me.olMessageVIP.Caption = "VIP EN PROCESO"
                Me.olMessageVIP.Visible = True
                bgVIP = True
            End If
        Else
            Me.olMessageVIP.Caption = "VIP EN PROCESO"
            Me.olMessageVIP.Visible = True
            bgVIP = True
        End If
        Exit Sub
    End If
End If
Select Case igLen
Case 0 To 1
    Me.oFrame_Dis.Visible = False
    Me.oFrame_Can.Visible = False
    'Me.oFrame_Gen.Visible = False
    Call Desplazar_Pantalla_1(KeyAscii)
Case 2 To 3
    Me.oFrame_Gen.Visible = False
    Me.oFrame_Can.Visible = False
    'bStatus = Me.oFrame_Dis.Visible
    'Me.oFrame_Dis.Visible = False
    Call Desplazar_Pantalla_2(KeyAscii)
    'Me.oFrame_Dis.Visible = bStatus
Case 4 To 6
    Me.oFrame_Gen.Visible = False
    Me.oFrame_Dis.Visible = False
    bStatus = Me.oFrame_Can.Visible
    Me.oFrame_Can.Visible = False
    Call Desplazar_Pantalla_3(KeyAscii)
    Me.oFrame_Can.Visible = bStatus
End Select
End Sub
 
Private Function Desplazar_Pantalla_1(Optional ByVal pPage_Order As Integer = 0)
Dim sChar As String
If igMax_RgG = 0 Then
    MsgBox "No hay registros que procesar en esta pàgina", vbInformation + vbOKOnly
    Exit Function
End If
Call Desactiva_Genero(True)
sChar = VBA.Chr(pPage_Order)
sChar = VBA.UCase(sChar)
If Inlist(sChar, sgKb_UP) Then
    Call Change_Page_Up_1
End If
If Inlist(sChar, sgKb_DN) Then
    Call Change_Page_Dn_1
End If
Call Desactiva_Genero(False)
End Function

Private Function Desplazar_Pantalla_2(Optional ByVal pPage_Order As Integer = 0)
Dim sChar As String
If igMax_RgD = 0 Then
    MsgBox "No hay registros que procesar en esta pàgina", vbInformation + vbOKOnly
    Exit Function
End If
Call Desactiva_Disco(True)
sChar = VBA.Chr(pPage_Order)
sChar = VBA.UCase(sChar)
If Inlist(sChar, sgKb_UP) Then
    Call Change_Page_Up_2
End If
If Inlist(sChar, sgKb_DN) Then
    Call Change_Page_Dn_2
End If
Call Desactiva_Disco(False)
End Function

Private Function Desplazar_Pantalla_3(Optional ByVal pPage_Order As Integer = 0)
Dim sChar As String
If igMax_RgC = 0 Then
    'MsgBox "No hay canciones en esta página", vbInformation + vbOKOnly
    Exit Function
End If
'If pPage_Order <> 0 Then
'    igKeyAscii = pPage_Order
'End If
sChar = VBA.Chr(pPage_Order)
sChar = VBA.UCase(sChar)
If Inlist(sChar, sgKb_UP) Then
    Call Change_Page_Up_3
End If
If Inlist(sChar, sgKb_DN) Then
    Call Change_Page_Dn_3
End If
End Function

Private Sub Change_Page_Up_1()
Dim igGenObjFocus As Integer
If igAct_PgG > 1 Then
    igAct_PgG = igAct_PgG - 1
    Call Cargar_Pag_Gen(igAct_PgG, igGenObjFocus)
Else
    Call Cargar_Pag_Gen(igAct_PgG, igMax_RgG)
End If
End Sub

Private Sub Change_Page_Up_2()
Dim igDisObjFocus As Integer
If igAct_PgD > 1 Then
    igAct_PgD = igAct_PgD - 1
    Call Cargar_Pag_Dis(igAct_PgD, igDisObjFocus)
Else
    Call Cargar_Pag_Dis(igAct_PgD, igMax_RgD)
End If
End Sub

Private Sub Change_Page_Up_3()
Dim igCanObjFocus As Integer
If igAct_PgC > 1 Then
    igAct_PgC = igAct_PgC - 1
    Call Cargar_Pag_Can(igAct_PgC, igCanObjFocus)
Else
    Call Cargar_Pag_Can(igAct_PgC, igMax_RgC)
End If
End Sub

Private Sub Change_Page_Dn_1()
If igAct_PgG < igTot_PgG Then
    igAct_PgG = igAct_PgG + 1
    Call Cargar_Pag_Gen(igAct_PgG, igMax_RgG)
Else
    Call Cargar_Pag_Gen(igAct_PgG, igMax_RgG)
End If
igGenObjFocus = 1
End Sub

Private Sub Change_Page_Dn_2()
If igAct_PgD < igTot_PgD Then
    igAct_PgD = igAct_PgD + 1
    Call Cargar_Pag_Dis(igAct_PgD, igMax_RgD)
Else
    Call Cargar_Pag_Dis(igAct_PgD, igMax_RgD)
End If
igDisObjFocus = 1
End Sub

Private Sub Change_Page_Dn_3()
If igAct_PgC < igTot_PgC Then
    igAct_PgC = igAct_PgC + 1
    Call Cargar_Pag_Can(igAct_PgC, igMax_RgC)
Else
    Call Cargar_Pag_Can(igAct_PgC, igMax_RgC)
End If
igCanObjFocus = 1
End Sub

Private Function Busca_Sel_1( _
ByVal pValor_Bus As String, ByRef pCo_Gen As String, _
ByRef pNo_Pag As Integer, ByRef pNo_Pos As Integer) As Boolean

Dim iCnt_Reg As Integer, iTot_reg As Integer, iCnt_Pag As Integer
Dim sValor   As String
For iCnt_Pag = 1 To igTot_PgG
    iTot_reg = aPag_Gen(iCnt_Pag).No_Rgs
    For iCnt_Reg = 1 To iTot_reg
        sValor = aPag_Gen(iCnt_Pag).Genero(iCnt_Reg).ID_ORD
        If pValor_Bus = sValor Then
            pCo_Gen = aPag_Gen(iCnt_Pag).Genero(iCnt_Reg).ID_GEN
            pNo_Pos = aPag_Gen(iCnt_Pag).Genero(iCnt_Reg).No_POS
            pNo_Pag = iCnt_Pag
            Busca_Sel_1 = True
            Exit Function
        End If
    Next iCnt_Reg
Next iCnt_Pag
Busca_Sel_1 = False
pCo_Gen = "99": pNo_Pos = 0: pNo_Pag = 0
End Function

Private Function Busca_Sel_2( _
ByVal pValor_Bus As String, _
ByRef pCo_Dis As String, ByRef pNo_Pag As Integer, _
ByRef pNo_Pos As Integer, ByRef pFl_IMG As String) As Boolean

Dim iCnt_Reg As Integer, iTot_reg As Integer, iCnt_Pag As Integer
Dim sValor   As String
For iCnt_Pag = 1 To igTot_PgD
    iTot_reg = aPag_Disc(iCnt_Pag).No_Rgs
    For iCnt_Reg = 1 To iTot_reg
        sValor = aPag_Disc(iCnt_Pag).Discos(iCnt_Reg).ID_ORD
        If pValor_Bus = sValor Then
            pCo_Dis = aPag_Disc(iCnt_Pag).Discos(iCnt_Reg).ID_DIS
            pNo_Pos = aPag_Disc(iCnt_Pag).Discos(iCnt_Reg).No_POS
            pFl_IMG = aPag_Disc(iCnt_Pag).Discos(iCnt_Reg).FL_IMG
            pNo_Pag = iCnt_Pag
            Busca_Sel_2 = True
            Exit Function
        End If
    Next iCnt_Reg
Next iCnt_Pag
Busca_Sel_2 = False
pCo_Dis = "99": pNo_Pos = 0: pNo_Pag = 0
End Function

Private Function Busca_Sel_3( _
ByVal pValor_Bus As String, _
ByRef pCo_Can As String, ByRef pFl_MP3 As String, _
ByRef pNo_Pag As Integer, ByRef pNo_Pos As Integer, _
ByVal pCa_Ato As String, ByRef pSt_001 As String) As Boolean

Dim iCnt_Reg As Integer, iTot_reg As Integer, iCnt_Pag As Integer
Dim sValor   As String, sDes_Can As String
For iCnt_Pag = 1 To igTot_PgC
    iTot_reg = aPag_Canc(iCnt_Pag).No_Rgs
    For iCnt_Reg = 1 To iTot_reg
        sValor = aPag_Canc(iCnt_Pag).Cancion(iCnt_Reg).ID_ORD
        If pValor_Bus = sValor Then
            pCo_Can = aPag_Canc(iCnt_Pag).Cancion(iCnt_Reg).ID_CAN
            pNo_Pos = aPag_Canc(iCnt_Pag).Cancion(iCnt_Reg).No_POS
            pFl_MP3 = aPag_Canc(iCnt_Pag).Cancion(iCnt_Reg).FL_IMG
            sDes_Can = aPag_Canc(iCnt_Pag).Cancion(iCnt_Reg).DE_CAN
            pNo_Pag = iCnt_Pag
            pSt_001 = VBA.Trim(pCa_Ato) & "," & VBA.Trim(pCo_Can) & "," & VBA.Trim(sDes_Can) & "," & VBA.Trim(pFl_MP3)
            Busca_Sel_3 = True
            Exit Function
        End If
    Next iCnt_Reg
Next iCnt_Pag
Busca_Sel_3 = False
pCo_Can = "99": pNo_Pos = 0: pNo_Pag = 0
End Function

Private Function Busca_ATocar(pCad_Ato As String) As Boolean
Dim i As Integer
Dim iCont As Integer
iCont = 0
For i = 1 To oLst_A_Tocar.ListCount - 1
    If Me.oLst_A_Tocar.List(i) = pCad_Ato Then
        Busca_ATocar = True
        Exit Function
    End If
Next i
End Function
Private Sub Conectar_DBDis(ByVal pCod_Gen As String)
Dim sSql As String
If (pCod_Gen = igGen_Sel) Then
    Exit Sub
End If
sSql = "SELECT * FROM File02 WHERE ID_GEN ='" & pCod_Gen & " ' ORDER BY ID_ORD"
With Me.oDC_DISC
    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .Refresh
End With
Me.oLst_Disc.Refresh
End Sub

Private Function Cargar_INF_Dis(pCod_Gen As String, ByRef pMax_Dis As Integer)
Dim iTot_reg As Integer
Dim iNum_reg As Integer
Dim iCnt_Pag As Integer
Dim i As Integer
iNum_reg = 1
With Me.oDC_DISC.Recordset
    iTot_reg = .RecordCount()
    If iTot_reg <= 0 Then
        oLTitulo.Caption = "No hay discos en este género."
        'Call Desactiva_Disco(True)
        Exit Function
    Else
        'Call Desactiva_Disco(False)
    End If
    pMax_Dis = iTot_reg
    .MoveFirst
    iCnt_Pag = 1
    iNum_reg = 1
    Do While Not .EOF
        If iNum_reg > igMax_Dis Then
            iCnt_Pag = iCnt_Pag + 1
            iNum_reg = 1
        End If
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).ID_ORD = .Fields("Id_Ord").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).ID_DIS = .Fields("ID_DIS").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).NOM_DIS = .Fields("NOM_DIS").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).NOM_ART = .Fields("NOM_ART").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).TX_COM = .Fields("TX_COM").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).FL_IMG = .Fields("FL_IMG").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).C_VIDEO = .Fields("C_VIDEO").value
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).No_POS = iNum_reg + 1
        aPag_Disc(iCnt_Pag).No_Rgs = iNum_reg
        iNum_reg = iNum_reg + 1
        .MoveNext
    Loop
End With
igTot_PgD = iCnt_Pag
End Function

Private Sub Conectar_DBCan(ByVal pCod_Gen As String, ByVal pCod_Dis As String)
Dim sSql As String
If (pCod_Dis = igDis_Sel) Then
    Exit Sub
End If
sSql = "SELECT * FROM File03 " & _
"WHERE ID_GEN ='" & pCod_Gen & "' " & _
"AND   ID_DIS ='" & pCod_Dis & "' "
With Me.oDC_CANC
    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .Refresh
End With
Me.oLst_Canc.Refresh
End Sub

Private Function Cargar_INF_Can(pCod_Gen As String, _
pCod_ORDD As String, pCod_Dis As String, ByRef pMax_Can As Integer)
Dim iTot_reg As Integer
Dim iNum_reg As Integer
Dim iCnt_Pag As Integer
Dim i As Integer
Dim sFld1 As String, sFld2 As String, sFld3 As String, sFld4 As String, sCadena As String
iNum_reg = 1
With Me.oDC_CANC.Recordset
    iTot_reg = .RecordCount()
    If iTot_reg <= 0 Then
        oLTitulo.Caption = "No hay canción en este disco."
        'Call Desactiva_Cancion(True)
        Exit Function
    Else
        'Call Desactiva_Cancion(False)
    End If
    pMax_Canc = iTot_reg
    .MoveFirst
    iCnt_Pag = 1
    iNum_reg = 1
    Me.oLst_Popular.Clear
    Do While Not .EOF
        If iNum_reg > igMax_Can Then
            iCnt_Pag = iCnt_Pag + 1
            iNum_reg = 1
        End If
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).ID_GEN = .Fields("ID_GEN").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).ID_DIS = .Fields("ID_DIS").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).ID_CAN = .Fields("ID_CAN").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).ID_ORD = .Fields("ID_ORD").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).DE_CAN = .Fields("DE_CAN").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).FL_IMG = .Fields("FL_MP3").value
        aPag_Canc(iCnt_Pag).Cancion(iNum_reg).No_POS = iNum_reg + 1
        aPag_Canc(iCnt_Pag).No_Rgs = iNum_reg
        iNum_reg = iNum_reg + 1
        '**************Carga los datos de l popular***********************
        iVal = VBA.Val(pCod_Gen)
        sFld0 = PADL(iVal, 2, "0")
        sFld0 = sFld0 & pCod_ORDD
        sFld1 = .Fields("ID_ORD").value
        sFld2 = .Fields("ID_CAN").value
        sFld3 = .Fields("DE_CAN").value
        sFld4 = .Fields("FL_MP3").value
        sCadena = sFld0 & sFld1 & "," & sFld2 & "," & sFld3 & "," & sFld4
        Call Me.oLst_Popular.AddItem(sCadena)
        '****************************************************************
        .MoveNext
    Loop
End With
'Call Ordenar_Popular
igTot_PgC = iCnt_Pag
End Function

Private Function Cargar_Pag_Gen(ByVal pNum_Pag As Integer, Optional ByRef pNoReg As Integer)
Dim i As Integer, iPos_Vac As Integer
Me.oLTitulo.Caption = "Seleccione el Género"
'------Se controla que no se sobrepase las cantidad de paginas permitidas----------
If igGen_Sel = "99" Then
    Call Limpia_Gen
    Exit Function
End If
If pNum_Pag > igTot_PgG Then
    pNum_Pag = igTot_PgG
End If
If pNum_Pag = 0 Then
    pNum_Pag = 1
End If
pNoReg = aPag_Gen(pNum_Pag).No_Rgs
'----------------Se Cargas lso item correspondientes a las página------------------
For i = 1 To pNoReg
    Me.oLGenero(i).Caption = aPag_Gen(pNum_Pag).Genero(i).De_Gen
Next i
'---------------------Se limpian las posiciones que no se usan---------------------
iPos_Vac = pNoReg + 1
Call Limpia_Gen(iPos_Vac)
igMax_RgG = pNoReg
Call Refresh_Paginero(igAct_PgG, igTot_PgG)
End Function

Private Sub Limpia_Gen(Optional ByVal piPos As Integer = 1)
Dim i As Integer
For i = piPos To Me.oLGenero.Count
    Me.oLGenero(i).Caption = ""
Next i
End Sub

Private Function Cargar_Pag_Dis(ByVal pNum_Pag As Integer, Optional ByRef pNoReg As Integer)
Dim sFile As String, sLabel As String
Dim i As Integer, iPos_Vac As Integer
If igDis_Sel = "99" Then
    Call Limpia_Dis
    Exit Function
End If
'------Se controla que no se sobrepase las cantidad de paginas permitidas----------
If pNum_Pag > igTot_PgD Then
    pNum_Pag = igTot_PgD
End If
If pNum_Pag = 0 Then
    pNum_Pag = 1
End If
pNoReg = aPag_Disc(pNum_Pag).No_Rgs
'----------------Se Cargas lso item correspondientes a las página------------------
Call Borra_Video_Signal
For i = 1 To pNoReg
    sFile = VBA.Trim(aPag_Disc(pNum_Pag).Discos(i).FL_IMG)
    sLabel = aPag_Disc(pNum_Pag).Discos(i).ID_ORD
    If VBA.Right(sFile, 1) = "\" Then
        sFile = ""
        sLabel = "?"
    End If
    If FileExist(sFile) Then
        Me.Image1(i).Picture = LoadPicture(sFile)
        With Me.oLNum_Disk(i)
            .Caption = sLabel
            .Visible = True
            .Left = Val(Me.oLNum_Disk(i).Tag)
        End With
        If VBA.Int(aPag_Disc(pNum_Pag).Discos(i).C_VIDEO) = 1 Then
            If bgVideoLabel = True Then
                Me.oLNum_Disk(i).Left = VBA.Val(Me.oLNum_Disk(i).Tag) - 720
                With olVideo(i)
                    .Caption = "<VIDEO>"
                    .Visible = True
                    .Tag = "1"
                End With
            End If
        Else
            Me.oLNum_Disk(i).Left = Val(Me.oLNum_Disk(i).Tag)
        End If
        If bgDiscLabel = True Then
            Me.ofLabelCont(i).Visible = True
            With Me.oDisc_Label1(i)
                .Caption = VBA.UCase(VBA.Trim(PADL(aPag_Disc(pNum_Pag).Discos(i).NOM_DIS, 19, " ")))
                .Visible = True
            End With
            With Me.oDisc_Label2(i)
                .Caption = VBA.UCase(VBA.Trim(PADL(aPag_Disc(pNum_Pag).Discos(i).NOM_ART, 19, " ")))
                .Visible = True
            End With
        End If
    Else
        Me.Image1(i).Picture = LoadPicture()
        With Me.oLNum_Disk(i)
            .Caption = "?"
            .Visible = True
            .Left = Val(Me.oLNum_Disk(i).Tag)
        End With
        Me.ofLabelCont(i).Visible = False
        If bgDiscLabel = True Then
            With Me.oDisc_Label1(i)
                .Caption = ""
                .Visible = False
            End With
            With Me.oDisc_Label2(i)
                .Caption = ""
                .Visible = False
            End With
        End If
    End If
Next i
'---------------------Se limpian las posiciones que no se usan---------------------
iPos_Vac = pNoReg + 1
Me.oFrame_Dis.Visible = False
Call Limpia_Dis(iPos_Vac)
igMax_RgD = pNoReg
Me.oFrame_Dis.Visible = True
Call Refresh_Paginero(igAct_PgD, igTot_PgD)
End Function

Private Sub Borra_Video_Signal()
Dim i As Integer
For i = 1 To 9
    With olVideo(i)
        .Caption = ""
        .Visible = False
        .Tag = "0"
    End With
Next i
End Sub

Private Sub Limpia_Dis(Optional ByVal piPos As Integer = 1)
Dim i As Integer
Dim bVisible As Boolean
bVisible = Me.oFrame_Dis.Visible
Me.oFrame_Dis.Visible = False
For i = piPos To Me.Image1.Count
    With Me.oLNum_Disk(i)
        .Visible = False
        .Caption = ""
    End With
    With Me.oDisc_Label1(i)
        .Visible = False
        .Caption = ""
    End With
    With Me.oDisc_Label2(i)
        .Visible = False
        .Caption = ""
    End With
    Me.ofLabelCont(i).Visible = False
    Me.Image1(i).Picture = LoadPicture()
Next i
Me.oFrame_Dis.Visible = bVisible
End Sub

Private Function Cargar_Pag_Can( _
ByVal pNum_Pag As Integer, Optional ByRef pNoReg As Integer)
Dim sFile As String, sLabel As String, sOrder As String
Dim i As Integer
If igGen_Sel = "99" Then
    Call Limpia_Can
    Exit Function
End If
'------Se controla que no se sobrepase las cantidad de paginas permitidas----------
If pNum_Pag > igTot_PgC Then
    pNum_Pag = igTot_PgC
End If
If pNum_Pag = 0 Then
    pNum_Pag = 1
End If
pNoReg = aPag_Canc(pNum_Pag).No_Rgs
For i = 1 To pNoReg
    sOrder = aPag_Canc(pNum_Pag).Cancion(i).ID_ORD
    sLabel = aPag_Canc(pNum_Pag).Cancion(i).DE_CAN
    sExtFl = VBA.Trim(aPag_Canc(pNum_Pag).Cancion(i).FL_IMG)
    Me.oImgVideo(i).Visible = True
    If VBA.UCase(VBA.Right(sExtFl, 3)) <> "MP3" Then
        Me.oImgVideo(i).Picture = LoadPicture(App.Path & "\icn_video_pk.GIF")
    Else
        Me.oImgVideo(i).Picture = LoadPicture(App.Path & "\THEMES.GIF")
    End If
    If Not sLabel = "" Then
        sLabel = sOrder & " - " & Proper(sLabel)
    End If
    sFile = sgDir_Mp3 & sFile
    Me.oLCanc(i).Caption = sLabel
Next i
'---------------------Se limpian las posiciones que no se usan---------------------
iPos_Vac = pNoReg + 1
Call Limpia_Can(iPos_Vac)
igMax_RgC = pNoReg
Call Refresh_Paginero(igAct_PgC, igTot_PgC)
End Function

Private Sub Limpia_Can(Optional ByVal piPos As Integer = 1)
Dim i As Integer
For i = piPos To Me.oLCanc.Count
    Me.oLCanc(i).Caption = ""
    Me.oImgVideo(i).Visible = False
Next i
End Sub

Private Sub oTime_Mensajes_Timer()
If oTimer_Reset.Enabled = False Then
    Me.oTimer_Reset.Enabled = True
End If
If (Me.olMessage.ForeColor) = &HFF& Then
    Me.olMessage.ForeColor = &HFFFF&
Else
    Me.olMessage.ForeColor = &HFF&
End If
End Sub

Private Sub oTime_Mensajes2_Timer()
Call Muestra_Tema_Det
If Me.olMessageVIP.Visible = True Then
    If bgVIP = False Then
        Me.olMessageVIP.Visible = False
    End If
    If (Me.olMessageVIP.ForeColor) = &HFF& Then
        Me.olMessageVIP.ForeColor = &HFFFF&
    Else
        Me.olMessageVIP.ForeColor = &HFF&
    End If
End If
End Sub

Private Sub oTimer_Moneda_Timer()
If igKeep_Cred >= 1 Then
    Exit Sub
End If
If igCnt_CR > 0 Then
    Exit Sub
End If
If (Me.olCred_Msg.ForeColor) = &HFF& Then
    Me.olCred_Msg.ForeColor = &HFFFF&
Else
    Me.olCred_Msg.ForeColor = &HFF&
End If
End Sub

Private Sub oTimer_Reset_Timer()
Call Muestra_Tema_Det
Me.oTime_Mensajes.Enabled = False
Me.olMessage.Visible = False
Me.olMensaje_Video.Visible = False
End Sub

Private Sub Conectar_DBTem()
Dim sSql As String
sSql = "SELECT * FROM File05"
With Me.oDC_Temas
    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    '.CommandType = adCmdTable
    '.RecordSource = "File05"
    .CommandType = adCmdText
    .RecordSource = sSql
    .Refresh
End With
End Sub

Private Sub Cargar_Musica()
Dim sFle_Mp3 As String
Dim aDet() As String
If Me.oLst_A_Tocar.List(0) <> "" Then
    aDet = VBA.Split(Me.oLst_A_Tocar.List(0), ",", , vbTextCompare)
    sFle_Mp3 = VBA.Trim(VBA.Trim(aDet(3)))
    If VBA.UCase(VBA.Right(VBA.Trim(sFle_Mp3), 3)) <> "MP3" Then
        bgIs_Video = True
        Me.oImg_c_Video.Visible = True
    Else
        bgIs_Video = False
        Me.oImg_c_Video.Visible = False
    End If
    If FileExist(sFle_Mp3) Then
        Call Cargar_Musica_P0(sFle_Mp3)
    Else
        If sFle_Mp3 <> "" Then
            Me.olMessage.Caption = "Tema no encontrado!"
            Me.olMessage.Visible = True
            Me.oTime_Mensajes.Enabled = True
            Me.oImg_c_Video.Visible = False
        End If
    End If
Else
    oImg_c_Video.Visible = False
End If
Call Muestra_Tema_Det
End Sub

Private Sub Conectar_DBPub()
oLst_Pub.Path = sgDir_Pub
oLst_Pub.Refresh
igTot_Pub = oLst_Pub.ListCount - 1
End Sub

Private Sub oTM_codigo2_Timer()
Call Retrocede
Me.oTM_codigo2.Enabled = False
igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
End Sub

Private Sub otNot_Found_List_KeyPress(KeyAscii As Integer)
If KeyAscii = VBA.Asc("V") Or KeyAscii = VBA.Asc("v") Then
    bgExit = True
End If
End Sub

Private Sub otTema_Act_GotFocus()
Me.otCodigo.SetFocus
End Sub

Private Sub Retrocede()
Dim iVal As Integer
Select Case igLen
Case Is = 1
    VBA.SendKeys "{BACKSPACE}"
Case Is = 2
    VBA.SendKeys "{BACKSPACE}"
    VBA.SendKeys "{BACKSPACE}"
Case Is = 3
    VBA.SendKeys "{BACKSPACE}"
Case Is = 4
    VBA.SendKeys "{BACKSPACE}"
    VBA.SendKeys "{BACKSPACE}"
Case Is = 5
    VBA.SendKeys "{BACKSPACE}"
Case Is = 6
    VBA.SendKeys "{BACKSPACE}"
    VBA.SendKeys "{BACKSPACE}"
Case Is > 6
    VBA.SendKeys "{BACKSPACE}"
End Select
End Sub

Private Sub Refresh_Paginero(ByVal piPag_No As Integer, ByVal ipPag_Tot As Integer)
olPaginas.Caption = "Página (" + VBA.Trim(VBA.Str(piPag_No)) + " de " + VBA.Trim(VBA.Str(ipPag_Tot)) + ")"
If piPag_No = 1 And ipPag_Tot = 1 Then
    bgBlinkPag = False
    Me.oImg_PagUp.Visible = False
    Me.oImg_PagDn.Visible = False
    Me.olPaginas2.Visible = False
    Me.olPaginas.ForeColor = &HFFFF&
    Me.olPaginas2.ForeColor = &HFFFF&
ElseIf piPag_No = 1 And ipPag_Tot > 1 Then
    bgBlinkPag = True
    Me.oImg_PagUp.Visible = False
    Me.oImg_PagDn.Visible = True
    olPaginas2.Visible = True
ElseIf piPag_No <> 1 And (ipPag_Tot = piPag_No) Then
    bgBlinkPag = True
    Me.oImg_PagUp.Visible = True
    Me.oImg_PagDn.Visible = False
    olPaginas2.Visible = True
ElseIf piPag_No > 1 And (ipPag_Tot <> piPag_No) Then
    bgBlinkPag = True
    Me.oImg_PagUp.Visible = True
    Me.oImg_PagDn.Visible = True
    olPaginas2.Visible = True
End If
End Sub

Private Sub Inicial_Gen()
Call Limpia_Dis
Call Limpia_Can
Call Desactiva_Cancion(True)
Call Desactiva_Disco(True)
Call Desactiva_Genero(False)
Call Conectar_DBGen
Call Cargar_INF_Gen
Call Cargar_Pag_Gen(1, igMax_RgG)
End Sub

Private Sub Registra_Dll()
Dim WinDir As String
Dim Cadena As String
Dim ret As Long
Dim Res As Long
Dim oFs
If Not FileExist(App.Path & "\FOXTOOLS.FLL") Then
    Call MsgBox("El archivo [FOXTOOLS.FLL], necesario para la ejecusión del programa no existe", vbCritical, "Error al buscar DLL")
    End
Else
    Cadena = String$(300, Chr$(0))
    ret = GetWindowsDirectory(Cadena, Len(Cadena))
    WinDir = Left$(Cadena, ret)
    WinDir = Left$(Cadena, InStr(Cadena, Chr$(0)) - 1)
    If Not FileExist(WinDir & "\System32\FOXTOOLS.FLL") Then
        VBA.FileCopy App.Path & "\FOXTOOLS.FLL", WinDir & "\System32"
    End If
End If
If Not FileExist(App.Path & "\LIBRARY.DLL") Then
    Call MsgBox("El archivo [LIBRARY.DLL], necesario para la ejecusión del programa no existe", vbCritical, "Error al buscar DLL")
    End
Else
    Cadena = String$(300, Chr$(0))
    ret = GetWindowsDirectory(Cadena, Len(Cadena))
    WinDir = Left$(Cadena, ret)
    WinDir = Left$(Cadena, InStr(Cadena, Chr$(0)) - 1)
    If Not FileExist(WinDir & "\System32\LIBRARY.DLL") Then
        VBA.FileCopy App.Path & "\LIBRARY.DLL", WinDir & "\System32"
        Res = VBA.Shell("REGSVR32 " & WinDir & "\System32\LIBRARY.DLL", vbNormalFocus)
    End If
End If
End Sub

Private Sub Check_Other()
If Not sgFle_Fon = "" Then
    If FileExist(sgFle_Fon) Then
        Me.Picture = LoadPicture(sgFle_Fon)
    End If
End If
On Error Resume Next
Call MkDir(sgDir_Tmp)
On Error Resume Next
Call MkDir(sgDir_odb)
'---------------------------------------Link_Tab.dsn---------------------------------------
If Not FileExist(sgDir_odb & "\Link_Tab.dsn") Then
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DRIVER", "Driver da Microsoft para arquivos texto (*.txt; *.csv)")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "UID", "admin")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "UserCommitSync", "Yes")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "Threads", "3")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "SafeTransactions", "0")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "PageTimeout", "5")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "MaxBufferSize", "2048")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "FIL", "Text")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "Extensions", "txt,csv,tab,asc")
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DriverId", "27")
End If
Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DefaultDir", sgDir_Fls)
Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DBQ", sgDir_Fls)

'---------------------------------------Link_Dbf.dsn---------------------------------------
If Not FileExist(sgDir_odb & "\Link_Dbf.dsn") Then
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "DRIVER", "Microsoft Visual FoxPro Driver")
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "UID", "")
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "Collate", "Machine")
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "BackgroundFetch", "Yes")
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "Exclusive", "No")
    Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "SourceType", "DBF")
End If
Call Write_Ini_File(sgDir_odb & "\Link_Dbf.dsn", "ODBC", "SourceDB", sgDir_Tmp)

'---------------------------------------schema.ini---------------------------------------
If Not FileExist(sgDir_Fls & "\schema.ini") Then
    '---------------------------------------[file01.tab]---------------------------------------
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "ColNameHeader", "False")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Format", "CSVDelimited")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "CharacterSet", "OEM")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col1", "ID_GEN Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col2", "ID_ORD Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col3", "DESCRI Char Width 20")
    If Not FileExist(sgDir_Fls & "\file01.tab") Then
        Open sgDir_Fls & "\file01.tab" For Output As #1
        Write #1, "001", "01", "NECESITA CARGAR LA INFORMACIÓN..."
        Close #1
    End If
    '---------------------------------------[file02.tab]---------------------------------------
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "ColNameHeader", "False")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Format", "CSVDelimited")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "CharacterSet", "OEM")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col1", "ID_GEN Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col2", "ID_DIS Char Width 5")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col3", "ID_ORD Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col4", "NOM_DIS Char Width 40")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col5", "NOM_ART Char Width 40")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col6", "FL_IMG Char Width 80")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col7", "TX_COM Char Width 40")
    If Not FileExist(sgDir_Fls & "\file02.tab") Then
        Open sgDir_Fls & "\file02.tab" For Output As #1
        Write #1, "001", "00001", "01", "NO HAY DISCOS!!!", "NO ARTISTA!!!", "", ""
        Close #1
    End If
    
    '---------------------------------------[file03.tab]---------------------------------------
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "ColNameHeader", "False")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Format", "CSVDelimited")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "CharacterSet", "OEM")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col1", "ID_GEN Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col2", "ID_DIS Char Width 5")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col3", "ID_CAN Integer")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col4", "ID_ORD Char Width 2")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col5", "DE_CAN Char Width 50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col6", "FL_MP3 Char Width 80")
    If Not FileExist(sgDir_Fls & "\file03.tab") Then
        Open sgDir_Fls & "\file03.tab" For Output As #1
        Write #1, "001", "00001", "1", "01", "NO HAY CANSIÓN!!!", ""
        Close #1
    End If
    
    '---------------------------------------[file05.tab]--------------------------------------
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "ColNameHeader", "False")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Format", "CSVDelimited")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "CharacterSet", "OEM")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col1", "ID_CAN Integer")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col2", "ID_COD Char Width 65")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col3", "DE_CAN Char Width 50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col4", "FL_MP3 Char Width 80")
    If Not FileExist(sgDir_Fls & "\file05.tab") Then
        Open sgDir_Fls & "\file05.tab" For Output As #1
        Write #1, ""
        Close #1
    End If
End If
 End Sub

Private Sub Cargar_Temas()
Dim iTot_reg As Integer
Dim iNum_reg As Integer
Dim sCadena As String
Dim sFld1 As String, sFld2 As String
Dim sFld3 As String, sFld4 As String
Dim aDet() As String
Call ogVFP9.Create_Table(1)
Call ogVFP9.Add_From_Tab(1, sgDir_Fls)
Call ogVFP9.xGoTop(1)
iTot_reg = ogVFP9.Tot_Reg(1)
For iNum_reg = 1 To iTot_reg
    sFld1 = VBA.Trim(ogVFP9.xFields(1, "ID_COD"))
    sFld2 = VBA.Trim(VBA.Str(ogVFP9.xFields(1, "ID_CAN")))
    sFld3 = VBA.Trim(ogVFP9.xFields(1, "DE_CAN"))
    sFld4 = VBA.Trim(ogVFP9.xFields(1, "FL_MP3"))
    sCadena = sFld1 & "," & sFld2 & "," & sFld3 & "," & sFld4
    oLst_A_Tocar.AddItem sCadena
    Call ogVFP9.xNext(1)
Next iNum_reg
End Sub

Private Sub Salvar_Temas()
Dim iNum_reg As Integer
Dim sFlds() As String
Call ogVFP9.Create_Table(1)
If oLst_A_Tocar.ListCount <= 0 Then
    Call ogVFP9.Reset_table(1)
Else
    For iNum_reg = 0 To oLst_A_Tocar.ListCount - 1
        sFlds = VBA.Split(oLst_A_Tocar.List(iNum_reg), ",", , vbTextCompare)
        Call ogVFP9.Add_Data_1(sFlds(0), sFlds(1), sFlds(2), sFlds(3))
    Next iNum_reg
End If
Call ogVFP9.Save_Data(1, sgDir_Fls)
Call ogVFP9.Close_Table(1)
End Sub

Private Sub Colocar_Frames()
With Me.oFrame_Dis
    .Height = 8535
    .Left = 4440
    .Top = 600
    .Width = 7455
End With
With Me.oFrame_Can
    .Height = 8535
    .Left = 4920
    .Top = 1560
    .Width = 7095
End With
With Me.oFrame_Gen
    .Height = 8535
    .Left = 4920
    .Top = 1560
    .Width = 5295
End With
With Me.TBack4
    .Height = 4605
    .Left = 2400
    .Top = 3840
    .Width = 7695
End With
End Sub

Private Sub Set_Tmp_DBF()
Set ogVFP9 = CreateObject("library.VFP_txt_Utils")
Call ogVFP9.Set_Files_Tmp(sgDir_Fls, sgDir_Tmp, sgDir_Img, sgDir_Mp3, sgDir_Pub)
Call ogVFP9.Set_Files_Close
End Sub

Private Sub Cargar_Musica_P0(ByVal spFle_Mp3 As String)
If bgIs_Video = True Then
    'If Not ((Video_Form.MediaPlayer3.playState = wmppsReady) Or (Video_Form.MediaPlayer3.playState = wmppsUndefined)) Then
    '    Exit Sub
    'End If
    If Video_Form.MediaPlayer3.URL <> spFle_Mp3 Then
        Main_Form.MediaPlayer2.Close
    Else
        If ObPlayer_Ocupado(Video_Form.MediaPlayer3) = True Then
            Exit Sub
        End If
    End If
    '********************VISOR DE VIDEO GRANDE****************************
    Video_Form.MediaPlayer3.URL = spFle_Mp3
    Video_Form.MediaPlayer3.settings.mute = False
    Video_Form.MediaPlayer3.settings.volume = 120
    Video_Form.MediaPlayer3.Controls.play
    '********************TOCADOR DE MUSICA SOLA***************************
    Me.MediaPlayer1.URL = ""
    Me.MediaPlayer1.settings.mute = True
    '*********************VISOR DE VIDEO CHICO****************************
    Me.MediaPlayer2.URL = ""
    Me.MediaPlayer2.settings.mute = True
Else
    If ObPlayer_Ocupado(Me.MediaPlayer1) = True Then
        Exit Sub
    End If
    '********************VISOR DE VIDEO GRANDE****************************
    Video_Form.MediaPlayer3.URL = ""
    Video_Form.MediaPlayer3.settings.mute = True
    '********************TOCADOR DE MUSICA SOLA***************************
    Me.MediaPlayer1.URL = spFle_Mp3
    Me.MediaPlayer1.settings.mute = False
    Me.MediaPlayer1.settings.volume = 120
    '*********************VISOR DE VIDEO CHICO****************************
    Me.MediaPlayer2.URL = ""
    Me.MediaPlayer2.settings.mute = True
    '***********************************************
End If
End Sub

Function Ramdom_List(oList1 As Object, oList2 As Object)
Dim iIndex As Integer
Dim D(50) As Integer
Dim k As Integer
Dim iTot_Lst As Integer
Dim oColRandom As New Collection
iTot_Lst = oList1.ListCount
If iTot_Lst <= 0 Then
    Exit Function
End If
Do Until oColRandom.Count = iTot_Lst
    k = Int((Rnd * iTot_Lst - 0 + 1) + 0)
    lCount = lCount + 1
    On Error Resume Next
    oColRandom.Add k, Chr(k)
Loop
oList2.Clear
For iIndex = 0 To oList1.ListCount - 1
    oList2.AddItem (oList1.List(iIndex))
Next iIndex
oList1.Clear
For k = 1 To oColRandom.Count
    D(k) = oColRandom.Item(k)
    oList1.AddItem oList2.List(D(k) - 1)
Next k
oList2.Clear
End Function

Private Function Inlist(sEntrada As String, Optional sPar1 As String = "", Optional sPar2 As String = "", Optional sPar3 As String = "") As Boolean
Dim sCadenas As String
sCadenas = VBA.Trim(sPar1) & VBA.Trim(sPar2) & VBA.Trim(sPar3)
Inlist = IIf(InStr(1, sCadenas, sEntrada, vbTextCompare) > 0, True, False)
End Function

Private Sub DBCan_Cheker(pVal As Boolean)
Dim sRes As Integer
Dim sSql1 As String
Dim sSql2 As String
Dim iErr_Fnd As Integer
Dim iErr_Cnt As Integer
Dim iCop_Cnt As Integer
Dim iNCop_Cnt As Integer
Dim iTot_Cnt As Integer
Dim sArr() As String
iErr_Fnd = 0: iErr_Cnt = 0: iCop_Cnt = 0: iTot_Cnt = 0: iNCop_Cnt = 0
If pVal = True Then
    Call otCodigo_KeyPress(VBA.Asc("S"))
    otNot_Found_List.Text = ""
    otNot_Found_List.Refresh
    olInfo_Cheker.Caption = "Recuperando información del cansionero..."
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.value = 0
Else
    otNot_Found_List.Text = ""
    otNot_Found_List.Refresh
    olInfo_Cheker.Caption = "Recuperando información del cansionero..."
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.value = 0
    Exit Sub
End If
otNot_Found_List.Text = ""
otNot_Found_List.Refresh
List1.AddItem "Recuperando información del cansionero..."
List1.ListIndex = List1.ListCount - 1
sSql1 = "SELECT * FROM File03 ORDER BY ID_GEN,ID_DIS,ID_CAN,ID_ORD"
With Me.oDC_CANC
    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql1
    .Refresh
End With
Dim iNumReg As Integer
Dim iTotReg As Integer
Dim i__Proc As Integer
Dim sCadena As String
Dim sMp3Fle As String
Dim sTmp As String
Dim sValue As String
Me.List1.Clear
If Me.oChk_FndC.value = 1 Then
    With Me.oDC_CANC.Recordset
        If .RecordCount <= 0 Then
            Call MsgBox("No se encontraron temas en el cansionero actual...", vbCritical, "Error")
            iTotReg = 0
            Exit Sub
        Else
            iTotReg = .RecordCount
        End If
        olInfo_Cheker.Caption = VBA.Str(iTotReg) & " Registros encontrados [CANSIONES"
        Me.List1.AddItem VBA.Str(iTotReg) & " Registros [CANSIONES] por verificar..."
        Me.List1.ListIndex = List1.ListCount - 1
        .MoveFirst
        iNumReg = 0
        olInfo_Cheker.Caption = "Procesando.."
        Me.ProgressBar1.Min = 0
        Me.ProgressBar1.Max = 100
        Do While Not .EOF
            If bgExit = True Then
                otNot_Found_List.Text = ""
                otNot_Found_List.Refresh
                olInfo_Cheker.Caption = "Recuperando información del cansionero..."
                Me.ProgressBar1.Min = 0
                Me.ProgressBar1.Max = 100
                Me.ProgressBar1.value = 0
                Exit Sub
            End If
            iNumReg = iNumReg + 1
            sMp3Fle = VBA.Trim(.Fields("FL_MP3").value)
            sArr = VBA.Split(sMp3Fle, "\")
            iErr_Fnd = 0
            On Error GoTo Solve_error
            For X = 1 To 100
                sTmp = VBA.Right(VBA.UCase(VBA.Trim(sArr(X))), 3)
                If sTmp = "MP3" Or sTmp = "MPEG" Or sTmp = "MPG" Then
                    iErr_Fnd = 0
                    sExt_Mp3 = VBA.Trim(VBA.Trim(sArr(X)))
                    Exit For
                Else
                    sTmp = sArr(X)
                    If iErr_Fnd = 1 Then
                        Exit For
                    End If
                End If
            Next X
            iErr_Fnd = 0
            On Error GoTo 0
            i__Proc = (iNumReg / iTotReg) * 100
            olInfo_cheker_Proc.Caption = Str(i__Proc) & "%"
            Me.ProgressBar1.value = i__Proc
        
            olInfo_cheker_Proc.Refresh
            Me.otRuteExternal.Text = VBA.Trim(Me.otRuteExternal.Text)
            Me.otRuteExternal.Text = Me.otRuteExternal.Text & ""
            If FileExist(sMp3Fle) = False Then
                iTot_Cnt = iTot_Cnt + 1
                If Me.Check1.value = 1 Then
                    iErr_Fnd = 0
                    sValue = Me.otRuteExternal & sExt_Mp3
                    If CopyFast(sValue, sMp3Fle, ProgressBar2) = True Then
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal & "\" & sExt_Mp3 & " ->OK..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iCop_Cnt = iCop_Cnt + 1
                        Me.ProgressBar2.Visible = False
                    Else
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal & "\" & sExt_Mp3 & " ->FAILED..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iNCop_Cnt = iNCop_Cnt + 1
                    End If
                Else
                    vString = "" & sMp3Fle
                    Me.List1.AddItem PADR(.Fields("ID_GEN").value, 7, " ") & " " & PADR(.Fields("ID_DIS").value, 10, " ") & " " & .Fields("ID_ORD").value & " " & PADR(VBA.Trim(.Fields("DE_CAN").value), 50, " ") & " " & Chr(10) & VBA.Trim(.Fields("FL_MP3").value)
                    Me.List1.ListIndex = List1.ListCount - 1
                End If
            End If
            .MoveNext
        Loop
        Me.List1.AddItem "----------------------------------------------------------------"
        Me.List1.AddItem "Archivos Faltantes  : " + VBA.Trim(VBA.Str(iTot_Cnt))
        Me.List1.AddItem "Archivos Copiados   : " + VBA.Trim(VBA.Str(iCop_Cnt))
        Me.List1.AddItem "Archivos No Copiados: " + VBA.Trim(VBA.Str(iNCop_Cnt))
        Me.List1.AddItem "----------------------------------------------------------------"
        Me.List1.AddItem " "
        Me.List1.AddItem " "
        Me.List1.ListIndex = List1.ListCount - 1
    End With
End If
If Me.oChk_FndP.value = 1 Then
    Dim sImgFle As String
    Dim sExt_Img As String
    sSql2 = "SELECT * FROM File02 ORDER BY ID_GEN,ID_DIS,ID_ORD"
    With Me.oDC_DISC
        .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
        .CommandType = adCmdText
        .RecordSource = sSql2
        .Refresh
    End With
    With Me.oDC_DISC.Recordset
        If .RecordCount <= 0 Then
            Call MsgBox("No se encontraron portadas de discos en en sistema actual...", vbCritical, "Error")
            iTotReg = 0
            Exit Sub
        Else
            iTotReg = .RecordCount
        End If
        olInfo_Cheker.Caption = VBA.Str(iTotReg) & " Registros encontrados [PORTADAS]..."
        Me.List1.AddItem VBA.Str(iTotReg) & " Registros de [PORTADAS] por verificar..."
        Me.List1.ListIndex = List1.ListCount - 1
        .MoveFirst
        iNumReg = 0
        olInfo_Cheker.Caption = "Procesando.."
        Me.ProgressBar1.Min = 0
        Me.ProgressBar1.Max = 100
        Do While Not .EOF
            If bgExit = True Then
                otNot_Found_List.Text = ""
                otNot_Found_List.Refresh
                olInfo_Cheker.Caption = "Recuperando información de DISCOS."
                Me.ProgressBar1.Min = 0
                Me.ProgressBar1.Max = 100
                Me.ProgressBar1.value = 0
                Exit Sub
            End If
            iNumReg = iNumReg + 1
            sImgFle = VBA.Trim(.Fields("FL_IMG").value)
            sArr = VBA.Split(sImgFle, "\")
            iErr_Fnd = 0
            On Error GoTo Solve_error
            For X = 1 To 100
                sTmp = VBA.Right(VBA.UCase(VBA.Trim(sArr(X))), 3)
                If sTmp = "JPG" Or sTmp = "JPEG" Or sTmp = "BMP" Then
                    iErr_Fnd = 0
                    sExt_Img = VBA.Trim(VBA.Trim(sArr(X)))
                    Exit For
                Else
                    sTmp = sArr(X)
                    If iErr_Fnd = 1 Then
                        Exit For
                    End If
                End If
            Next X
            iErr_Fnd = 0
            On Error GoTo 0
            i__Proc = (iNumReg / iTotReg) * 100
            olInfo_cheker_Proc.Caption = Str(i__Proc) & "%"
            Me.ProgressBar1.value = i__Proc
        
            olInfo_cheker_Proc.Refresh
            Me.otRuteExternal2.Text = VBA.Trim(Me.otRuteExternal2.Text)
            Me.otRuteExternal2.Text = Me.otRuteExternal2.Text & ""
            If FileExist(sImgFle) = False And sImgFle <> "" Then
                iTot_Cnt = iTot_Cnt + 1
                If Me.Check1.value = 1 Then
                    sValue = Me.otRuteExternal2 & sExt_Img
                    If CopyFast(sValue, sImgFle, Me.ProgressBar2) = True Then
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal2 & sExt_Img & " ->OK..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iCop_Cnt = iCop_Cnt + 1
                        Me.ProgressBar2.Visible = False
                    Else
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal2 & sExt_Img & " ->FAILED..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iNCop_Cnt = iNCop_Cnt + 1
                    End If
                End If
            End If
        .MoveNext
    Loop
    Me.List1.AddItem "----------------------------------------------------------------"
    Me.List1.AddItem "Archivos Faltantes  : " + VBA.Trim(VBA.Str(iTot_Cnt))
    Me.List1.AddItem "Archivos Copiados   : " + VBA.Trim(VBA.Str(iCop_Cnt))
    Me.List1.AddItem "Archivos No Copiados: " + VBA.Trim(VBA.Str(iNCop_Cnt))
    Me.List1.AddItem "----------------------------------------------------------------"
    Me.List1.AddItem ""
    Me.List1.ListIndex = List1.ListCount - 1
    End With
End If
Me.otNot_Found_List.Text = ""
For i = 1 To List1.ListCount - 1
    Me.otNot_Found_List.Text = Me.otNot_Found_List.Text & Me.List1.List(i) & Chr(10)
    Debug.Print Me.List1.List(i)
Next i
Me.otNot_Found_List.Refresh
Me.otNot_Found_List.SaveFile (App.Path & "\MP3_Not_Found.RTF")
Exit Sub

Solve_error:
iErr_Cnt = iErr_Cnt + 1
iErr_Fnd = 1
Resume Next
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
For i = 1 To 9
    If Me.olVideo(i).Visible = True Then
        If (Me.olVideo(i).ForeColor) = &H80FF80 Then
            Me.olVideo(i).BackColor = &HFF&
            Me.olVideo(i).ForeColor = &HFFFF&
        Else
            Me.olVideo(i).BackColor = &H0&
            Me.olVideo(i).ForeColor = &H80FF80
        End If
    End If
Next i
End Sub

Private Sub Go_Service()
Call Load(Svr_Form)
With Svr_Form
    .Text1(1).Text = sgDir_odb
    .Text1(2).Text = sgDir_Tmp
    .Text1(3).Text = sgDir_Fls
    .Text1(4).Text = sgDir_Img
    .Text1(5).Text = sgDir_Mp3
    .Text1(6).Text = sgDir_Pub
    .Text1(7).Text = sgFle_Fon
'   ----------------------------------
    .ctNEdit2(1).value = igLim_Cred
    .ctNEdit2(2).value = igCnt_CR
    .ctNEdit2(3).value = sgKb_BonC
    .ctNEdit2(4).value = sgKb_VID
    .Check2(1).value = bAcum_Cre
    .Check2(2).value = igKeep_Cred
    .Check2(3).value = igNoDuplicT
'   ----------------------------------
    .ctNEdit3(1).value = igDelay_Return_Gen
    .ctNEdit3(2).value = igDelay_Return_Dis
    .ctNEdit3(3).value = igDelay_Bonus_Vid
    .Check3(1).value = VBA.IIf(bgVideoLabel = True, 1, 0)
    .Check3(2).value = VBA.IIf(bgDiscLabel = True, 1, 0)
    .Check3(3).value = igScr_Alone
    .Check3(4).value = VBA.IIf(bgKeep_On_Top = True, 1, 0)
    .Check3(5).value = igMixe_Popu
'   ----------------------------------
    .ctMEdit4(1).Text = sgKb_Crd1
    .ctMEdit4(2).Text = sgKb_Crd2
    .ctMEdit4(3).Text = sgKb_Del
    .ctMEdit4(4).Text = sgKb_Ret
    .ctMEdit4(5).Text = sgKb_ResM
    .ctMEdit4(6).Text = sgKb_ResA
    .ctMEdit4(7).Text = sgKb_Pop
    .ctMEdit4(8).Text = sgKb_VIP
    .ctMEdit4(9).Text = sgKb_UP
    .ctMEdit4(10).Text = sgKb_DN
    .ctMEdit4(11).Text = sgKb_Vef
    .Show vbModal
End With
End Sub

Private Sub Show_Hide_Service(Optional iPar As Integer = 1)
If iPar = 1 Then
    If Command2.Visible = False Then
        Command2.Visible = True
    Else
        Command2.Visible = False
    End If
    If Command3.Visible = False Then
        Command3.Visible = True
    Else
        Command3.Visible = False
    End If
Else
    Command2.Visible = False
    Command3.Visible = False
End If
End Sub


