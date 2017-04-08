VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{120FD660-13F8-11D1-943D-444553540000}#1.0#0"; "ctmedit.ocx"
Object = "{BE38FE43-D38D-11D0-B731-00403333B3B0}#1.0#0"; "TBack.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BC184000-7A5A-11D2-B543-006097FAF8B8}#1.6#0"; "bbGetDir.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Begin VB.Form Main_Form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   1740
   ClientTop       =   840
   ClientWidth     =   12000
   Icon            =   "Mini_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GorditoButton.Boton oImg_PagDn 
      Height          =   615
      Left            =   10320
      TabIndex        =   123
      Top             =   8400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      PicturePosition =   4
      Caption         =   ""
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":08CA
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TbackLibCtl.TBack oFrame_Gen 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   1402
      _StockProps     =   224
      BackColor       =   12632319
      Version         =   16777230
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   8
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":0BBD
         MousePointer    =   99  'Custom
         TabIndex        =   121
         Top             =   6720
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   7
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":0D0F
         MousePointer    =   99  'Custom
         TabIndex        =   120
         Top             =   5760
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   6
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":0E61
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   4800
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   5
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":0FB3
         MousePointer    =   99  'Custom
         TabIndex        =   118
         Top             =   3840
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   4
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":1105
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   2880
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   3
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":1257
         MousePointer    =   99  'Custom
         TabIndex        =   116
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   2
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":13A9
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label oLGenero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   1
         Left            =   0
         MouseIcon       =   "Mini_Main.frx":14FB
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   0
         Width           =   6975
      End
   End
   Begin TbackLibCtl.TBack oFrame_Can 
      Height          =   7275
      Left            =   7560
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   12832
      _StockProps     =   224
      BackColor       =   12648384
      Version         =   16777230
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   1
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":164D
         MousePointer    =   99  'Custom
         TabIndex        =   100
         Top             =   0
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   2
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":179F
         MousePointer    =   99  'Custom
         TabIndex        =   99
         Top             =   960
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   3
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":18F1
         MousePointer    =   99  'Custom
         TabIndex        =   98
         Top             =   1905
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   4
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":1A43
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Top             =   2865
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   5
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":1B95
         MousePointer    =   99  'Custom
         TabIndex        =   96
         Top             =   3825
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   6
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":1CE7
         MousePointer    =   99  'Custom
         TabIndex        =   95
         Top             =   4785
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   7
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":1E39
         MousePointer    =   99  'Custom
         TabIndex        =   94
         Top             =   5730
         Width           =   6975
      End
      Begin VB.Label oLCanc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   525
         Index           =   8
         Left            =   360
         MouseIcon       =   "Mini_Main.frx":1F8B
         MousePointer    =   99  'Custom
         TabIndex        =   93
         Top             =   6690
         Width           =   6975
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   1
         Left            =   0
         Picture         =   "Mini_Main.frx":20DD
         Stretch         =   -1  'True
         Top             =   120
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   2
         Left            =   0
         Picture         =   "Mini_Main.frx":29C4
         Stretch         =   -1  'True
         Top             =   1065
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   3
         Left            =   0
         Picture         =   "Mini_Main.frx":4506
         Stretch         =   -1  'True
         Top             =   2010
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   4
         Left            =   0
         Picture         =   "Mini_Main.frx":6048
         Stretch         =   -1  'True
         Top             =   2955
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   5
         Left            =   0
         Picture         =   "Mini_Main.frx":7B8A
         Stretch         =   -1  'True
         Top             =   3900
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   6
         Left            =   0
         Picture         =   "Mini_Main.frx":96CC
         Stretch         =   -1  'True
         Top             =   4845
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   7
         Left            =   0
         Picture         =   "Mini_Main.frx":B20E
         Stretch         =   -1  'True
         Top             =   5790
         Width           =   270
      End
      Begin VB.Image oImgVideo 
         Height          =   405
         Index           =   8
         Left            =   0
         Picture         =   "Mini_Main.frx":CD50
         Stretch         =   -1  'True
         Top             =   6735
         Width           =   270
      End
   End
   Begin TbackLibCtl.TBack oFrame_Dis 
      Height          =   7275
      Left            =   7200
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   12832
      _StockProps     =   224
      BackColor       =   14737632
      Version         =   16777230
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         HelpContextID   =   1
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   2880
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   33023
         GradientColorTo =   33023
         Version         =   16777230
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
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
            TabIndex        =   61
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   60
            Top             =   0
            Width           =   2055
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   62
         Top             =   6600
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   64
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   63
            Top             =   240
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   65
         Top             =   2880
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   67
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   66
            Top             =   240
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   6
         Left            =   4920
         TabIndex        =   68
         Top             =   6600
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   70
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   69
            Top             =   240
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   5
         Left            =   2520
         TabIndex        =   71
         Top             =   6600
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   73
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   72
            Top             =   240
            Width           =   2175
         End
      End
      Begin TbackLibCtl.TBack ofLabelCont 
         Height          =   495
         Index           =   3
         Left            =   4920
         TabIndex        =   74
         Top             =   2880
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   224
         BackColor       =   16576
         GradientColorFrom=   12632256
         GradientColorTo =   0
         Version         =   16777230
         Begin VB.Label oDisc_Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "19ITEMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   76
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label oDisc_Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "oDisc_Label2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
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
            TabIndex        =   75
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   6
         Left            =   4920
         Picture         =   "Mini_Main.frx":E892
         Top             =   3960
         Width           =   420
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   5
         Left            =   2520
         Picture         =   "Mini_Main.frx":E90D
         Top             =   3960
         Width           =   420
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   4
         Left            =   120
         Picture         =   "Mini_Main.frx":E988
         Top             =   3960
         Width           =   420
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   3
         Left            =   4920
         Picture         =   "Mini_Main.frx":EA03
         Top             =   240
         Width           =   420
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   2
         Left            =   2520
         Picture         =   "Mini_Main.frx":EA7E
         Top             =   240
         Width           =   420
      End
      Begin VB.Image olNuevo 
         Height          =   165
         Index           =   1
         Left            =   120
         Picture         =   "Mini_Main.frx":EAF9
         Top             =   240
         Width           =   420
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   6
         Left            =   4920
         MouseIcon       =   "Mini_Main.frx":EB74
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   1
         Left            =   120
         MouseIcon       =   "Mini_Main.frx":ECC6
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   480
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
         Index           =   6
         Left            =   5640
         TabIndex        =   88
         Tag             =   "5640"
         Top             =   3720
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   4
         Left            =   120
         MouseIcon       =   "Mini_Main.frx":EE18
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   2
         Left            =   2520
         MouseIcon       =   "Mini_Main.frx":EF6A
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   480
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
         TabIndex        =   87
         Tag             =   "720"
         Top             =   3720
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
         Index           =   2
         Left            =   3240
         TabIndex        =   86
         Tag             =   "3240"
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   5
         Left            =   2520
         MouseIcon       =   "Mini_Main.frx":F0BC
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   2415
         Index           =   3
         Left            =   4920
         MouseIcon       =   "Mini_Main.frx":F20E
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         Top             =   480
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
         Index           =   5
         Left            =   3240
         TabIndex        =   85
         Tag             =   "3240"
         Top             =   3735
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
         Left            =   5640
         TabIndex        =   84
         Tag             =   "5640"
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
         Index           =   1
         Left            =   840
         TabIndex        =   83
         Tag             =   "840"
         Top             =   0
         Width           =   600
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
         TabIndex        =   82
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
         Index           =   3
         Left            =   6285
         TabIndex        =   81
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
         Index           =   4
         Left            =   1365
         TabIndex        =   80
         Top             =   3765
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
         TabIndex        =   79
         Top             =   3765
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
         Left            =   6285
         TabIndex        =   78
         Top             =   3780
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
         Index           =   1
         Left            =   1485
         TabIndex        =   77
         Top             =   45
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin TbackLibCtl.TBack TBack3 
      Height          =   1335
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   2355
      _StockProps     =   224
      BackColor       =   -2147483633
      GradientColorFrom=   0
      GradientColorTo =   0
      TransparentBackground=   -1  'True
      HasLicense      =   -1  'True
      Version         =   16777230
      Begin CTMEDITLibCtl.ctMEdit otTema_Act 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   65535
         BackColor       =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Digital dream Fat"
            Size            =   15.74
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         DropPicture     =   "Mini_Main.frx":F360
         BackColor       =   16576
         ForeColor       =   65535
         DisabledColor   =   8454016
         UseMaskChars    =   0   'False
         EditMask        =   "##-##-##"
      End
      Begin VB.Image oImg_c_Video 
         Height          =   375
         Left            =   2520
         Picture         =   "Mini_Main.frx":F37C
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label olTema_Act 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   15
         TabIndex        =   13
         Top             =   960
         Width           =   3075
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   435
         Left            =   120
         TabIndex        =   12
         Top             =   -80
         Width           =   2925
      End
   End
   Begin RichTextLib.RichTextBox oService_Info 
      Height          =   525
      Left            =   10800
      TabIndex        =   101
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   926
      _Version        =   393217
      BackColor       =   65535
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      OLEDropMode     =   0
      TextRTF         =   $"Mini_Main.frx":F459
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer oTM_Mouse 
      Interval        =   20000
      Left            =   5880
      Top             =   4320
   End
   Begin VB.Timer Timer3 
      Interval        =   800
      Left            =   5400
      Top             =   4320
   End
   Begin VB.Timer oTM_ScreenSaver 
      Interval        =   3000
      Left            =   4920
      Top             =   4320
   End
   Begin VB.Timer oTM_Box 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6360
      Top             =   3840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S1"
      Height          =   255
      Left            =   0
      TabIndex        =   47
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "S2"
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin TbackLibCtl.TBack TBack4 
      Height          =   525
      Left            =   6960
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   926
      _StockProps     =   224
      Appearance      =   1
      BackColor       =   8454143
      GradientColorTo =   16777215
      Version         =   16777230
      Begin VB.TextBox otRuteFiles 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1680
         TabIndex        =   56
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CommandButton oGetdFiles 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   55
         Top             =   3840
         Width           =   615
      End
      Begin VB.CheckBox oChk_FndF 
         BackColor       =   &H80000009&
         Caption         =   "DBContainer"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<-Regresar"
         Height          =   615
         Left            =   240
         TabIndex        =   53
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox otOrigen 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   52
         Top             =   3360
         Width           =   6015
      End
      Begin VB.CommandButton oGetOgen 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   51
         Top             =   3360
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   240
         TabIndex        =   45
         Top             =   1440
         Width           =   7215
      End
      Begin VB.CheckBox oChk_FndP 
         BackColor       =   &H80000009&
         Caption         =   "Portadas"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CheckBox oChk_FndC 
         BackColor       =   &H80000009&
         Caption         =   "Cansionero"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox otRuteExternal2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   4560
         Width           =   4575
      End
      Begin VB.CommandButton oGetRute2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   39
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar"
         Height          =   615
         Left            =   5760
         TabIndex        =   38
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton oGetRute 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   6360
         TabIndex        =   37
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox otRuteExternal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Top             =   4200
         Width           =   4575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Copiar Automaticamnte desde un origen externo"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   3120
         Width           =   3975
      End
      Begin RichTextLib.RichTextBox otNot_Found_List 
         Height          =   390
         Left            =   5640
         TabIndex        =   29
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Mini_Main.frx":F4DD
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
         TabIndex        =   30
         Top             =   720
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   4080
         TabIndex        =   44
         Top             =   1080
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de actualización de información"
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
         Left            =   1560
         TabIndex        =   50
         Top             =   120
         Width           =   4860
      End
      Begin BBGETDIRLibCtl.Bbgetdir Bbgetdir1 
         Left            =   6840
         Top             =   3000
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
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.Label olInfo_Cheker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<:>"
         Height          =   195
         Left            =   2880
         TabIndex        =   31
         Top             =   1080
         Width           =   225
      End
   End
   Begin VB.Timer otCargador_Video 
      Interval        =   1200
      Left            =   5400
      Top             =   2880
   End
   Begin VB.Timer otCargador_Music 
      Interval        =   1200
      Left            =   4920
      Top             =   2880
   End
   Begin TbackLibCtl.TBack TBack2 
      Height          =   345
      Left            =   960
      TabIndex        =   6
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
      Height          =   6735
      Left            =   8760
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   2670
      _Version        =   65536
      _ExtentX        =   4710
      _ExtentY        =   11880
      _StockProps     =   224
      Version         =   16777230
      Begin VB.FileListBox oLst_Promo 
         Height          =   285
         Left            =   0
         TabIndex        =   49
         Top             =   5160
         Width           =   2415
      End
      Begin VB.FileListBox oLst_Pub2 
         Height          =   285
         Left            =   0
         TabIndex        =   48
         Top             =   4800
         Width           =   2415
      End
      Begin VB.ListBox oBkList 
         DataSource      =   "oDC_Temas"
         Height          =   255
         ItemData        =   "Mini_Main.frx":F56A
         Left            =   0
         List            =   "Mini_Main.frx":F571
         TabIndex        =   27
         Top             =   6120
         Width           =   2415
      End
      Begin VB.ListBox oLst_A_Tocar 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ListBox oLst_Popular 
         DataSource      =   "oDC_Temas"
         Height          =   255
         ItemData        =   "Mini_Main.frx":F57E
         Left            =   0
         List            =   "Mini_Main.frx":F580
         TabIndex        =   25
         Top             =   6480
         Width           =   2415
      End
      Begin VB.FileListBox oLst_Temas_Video 
         Height          =   480
         Left            =   0
         TabIndex        =   22
         Top             =   5520
         Width           =   2415
      End
      Begin VB.FileListBox oLst_Pub1 
         Height          =   285
         Left            =   0
         TabIndex        =   21
         Top             =   4440
         Width           =   2415
      End
      Begin MSDataListLib.DataList oLst_Disc 
         Bindings        =   "Mini_Main.frx":F582
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         ListField       =   "NOM_DIS"
         BoundColumn     =   "ID_DIS"
      End
      Begin MSDataListLib.DataList oLst_Gen 
         Bindings        =   "Mini_Main.frx":F599
         Height          =   255
         Left            =   0
         TabIndex        =   4
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
         Width           =   2655
         _ExtentX        =   4683
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
         Bindings        =   "Mini_Main.frx":F5AF
         Height          =   255
         Left            =   0
         TabIndex        =   5
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
         Top             =   2760
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
      Begin MSAdodcLib.Adodc oDC_Promos 
         Height          =   375
         Left            =   0
         Top             =   3240
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
         Caption         =   "oDC_Promos"
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5400
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   3840
   End
   Begin VB.Timer oTime_Mensajes2 
      Interval        =   1200
      Left            =   6360
      Top             =   3360
   End
   Begin VB.Timer oTimer_Reset 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5880
      Top             =   3360
   End
   Begin VB.Timer oTime_Mensajes 
      Interval        =   1200
      Left            =   5400
      Top             =   3360
   End
   Begin VB.Timer oTimer_Moneda 
      Interval        =   800
      Left            =   4920
      Top             =   3360
   End
   Begin VB.Timer oTM_codigo2 
      Enabled         =   0   'False
      Left            =   6360
      Top             =   2880
   End
   Begin VB.Timer oGeneral_Timer 
      Interval        =   800
      Left            =   5880
      Top             =   2880
   End
   Begin VB.Timer oTimer_Srv 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   5880
      Top             =   3840
   End
   Begin CTMEDITLibCtl.ctMEdit otCodigo 
      Height          =   375
      Left            =   1125
      TabIndex        =   106
      Top             =   5880
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   65535
      BackColor       =   16576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Digital dream Fat"
         Size            =   15.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropPicture     =   "Mini_Main.frx":F5C6
      BackColor       =   16576
      ForeColor       =   65535
      UseMaskChars    =   0   'False
      EditMask        =   "99-99-99"
   End
   Begin VB.TextBox oSetFocus_Codigo 
      Height          =   285
      Left            =   2640
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   5880
      Width           =   615
   End
   Begin GorditoButton.Boton oImg_PagUp 
      Height          =   615
      Left            =   10320
      TabIndex        =   124
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      PicturePosition =   4
      Caption         =   ""
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":F5E2
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GorditoButton.Boton oControl 
      Height          =   615
      Index           =   3
      Left            =   4800
      TabIndex        =   125
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "VIP"
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":F894
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GorditoButton.Boton oControl 
      Height          =   615
      Index           =   2
      Left            =   6480
      TabIndex        =   126
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Cancel"
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":FB7F
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GorditoButton.Boton oControl 
      Height          =   615
      Index           =   1
      Left            =   8160
      TabIndex        =   127
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Reset"
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":FDE4
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GorditoButton.Boton oControl 
      Height          =   615
      Index           =   4
      Left            =   3120
      TabIndex        =   128
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "Popular"
      BackColor       =   255
      ResalteColor    =   16711680
      Picture         =   "Mini_Main.frx":145C5
      IntensityColor  =   4
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer2 
      CausesValidation=   0   'False
      Height          =   3240
      Left            =   300
      TabIndex        =   17
      Top             =   1920
      Width           =   4020
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
      _cx             =   7091
      _cy             =   5715
   End
   Begin VB.Label olPaginas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Página (1) ->"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   2760
      TabIndex        =   122
      Top             =   5160
      Width           =   1110
   End
   Begin VB.Label olT_Mant 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[00:00:00] Restante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   113
      Top             =   8400
      Visible         =   0   'False
      Width           =   3540
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2760
      TabIndex        =   112
      Top             =   5520
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   208
      X2              =   208
      Y1              =   400
      Y2              =   384
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Disco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   111
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Género"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   110
      Top             =   5520
      Width           =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   96
      X2              =   96
      Y1              =   400
      Y2              =   384
   End
   Begin VB.Label olTimer1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Digital dream Fat"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   109
      Top             =   6360
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label olTimer1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Digital dream Fat"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   108
      Top             =   6360
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "----TIEMPO-----"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1560
      TabIndex        =   107
      Top             =   6315
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   152
      X2              =   152
      Y1              =   400
      Y2              =   384
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   3600
      MouseIcon       =   "Mini_Main.frx":14BEA
      MousePointer    =   99  'Custom
      Picture         =   "Mini_Main.frx":14D3C
      Stretch         =   -1  'True
      ToolTipText     =   "Retroceder [Borra Selección]"
      Top             =   5880
      Width           =   840
   End
   Begin VB.Label olTest 
      BackColor       =   &H80000009&
      Caption         =   "TEST:"
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
      Left            =   4920
      TabIndex        =   105
      Top             =   6240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Caption         =   "METRO ENTRADA:"
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
      Left            =   4920
      TabIndex        =   104
      Top             =   5880
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label olMetros2 
      BackColor       =   &H80000009&
      Caption         =   "Label8"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#####0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Digital dream Fat"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   103
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label olMetros 
      BackColor       =   &H80000009&
      Caption         =   "Label8"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#####0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6154
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Digital dream Fat"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6840
      TabIndex        =   102
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label olAct_Pos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3840
      TabIndex        =   92
      Top             =   1590
      Width           =   480
   End
   Begin VB.Label oLDuracion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1200
      TabIndex        =   91
      Top             =   1590
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3150
      TabIndex        =   90
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duración:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   345
      TabIndex        =   89
      Top             =   1590
      Width           =   825
   End
   Begin VB.Label olMensajeSis 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No hay mensajes."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   720
      TabIndex        =   57
      Top             =   6720
      Width           =   2925
   End
   Begin VB.Image oInd_VideoSW 
      Height          =   270
      Left            =   0
      Picture         =   "Mini_Main.frx":14E23
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label olPaginas2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<PASAR PÁGINA>"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   360
      TabIndex        =   43
      Top             =   5175
      Visible         =   0   'False
      Width           =   1755
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
      TabIndex        =   33
      Top             =   8760
      Visible         =   0   'False
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
      TabIndex        =   23
      Top             =   0
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
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label olMessageVIP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No hay mensajes."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   7200
      Visible         =   0   'False
      Width           =   2925
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer1 
      Height          =   240
      Left            =   1200
      TabIndex        =   18
      Top             =   5880
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
   Begin VB.Label olMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No hay mensajes."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   3855
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Label olCred_Msg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INSERTE ¢ 0.25"
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
      Left            =   840
      TabIndex        =   8
      Tag             =   "INSERTE ¢ 0.25"
      Top             =   7680
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label olCreditos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CREDITOS(0)"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label oLTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Haga su selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   34
      Top             =   0
      Width           =   6480
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   960
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2535
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
      Left            =   600
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iShow As Integer
Dim DriverOpened As Boolean
Dim piCnt_Canc As Integer
Dim bFlagChek1 As Boolean
Dim bFlagChek2 As Boolean
Dim bFlagChek3 As Boolean
Dim gbServ_Mode As Boolean
Dim bfTm As Boolean
Dim iPosPagD As Integer
Dim pTimeResto As Long

Private Sub Check1_Click()
If Check1.value = 1 Then
    Me.otOrigen.Enabled = True
    Me.oGetOgen.Enabled = True
    Me.oChk_FndF.Enabled = True
    Me.oChk_FndC.Enabled = True
    Me.oChk_FndP.Enabled = True
Else
    Me.otOrigen.Enabled = False
    Me.oGetOgen.Enabled = False
    Me.oChk_FndF.Enabled = False
    Me.oChk_FndC.Enabled = False
    Me.oChk_FndP.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Dim sSourceDir As String
Dim sTargetDir As String
Dim sErr As String
Dim iP As Integer
Dim lRet As Long
If (bFlagChek1 = True Or bFlagChek2 = True Or bFlagChek3 = True) Then
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "Directorio en rojo, NO VALIDO. CORREGIR!"
    Me.oTime_Mensajes.Enabled = True
    Exit Sub
End If
'-----------------------------------------------------------
sSourceDir = VBA.Trim(Me.otRuteFiles.Text & "*.*")
sTargetDir = VBA.Trim(sgDir_Fls) & "\"
lRet = CopyFileWindowsWay(sSourceDir, sTargetDir, sErr, 1)
Me.olMessage.Visible = True
Me.olMessage.Caption = VBA.IIf(lRet = True, "Los archivos fueron copiados satisfactoriamente", "Hubo un error al tratar de copiar los archivos.")
'-------------------------(1)-------------------------------
Call Conectar_DBPub
Call Conectar_DBPro
Call Cargar_Gen
Call Limpia_Dis
Call Limpia_New
Call Limpia_Can
'-----------------------Otros-------------------------------
Call DBCan_Cheker(True)
Me.oTime_Mensajes.Enabled = True
Call Cargar_Temas
End Sub

Private Sub Command2_Click()
Call AlwaysOnTop(Main_Form, False)
Call Go_Service
Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
Call Show_Hide_Service
Call Refresh_Creditos(Main_Form)
Me.otCodigo.SetFocus
End Sub

Private Sub Command3_Click()
Call AlwaysOnTop(Main_Form, False)
'Me.Hide
ControlPanel.Show vbModal
'Me.Show
Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
Call Show_Hide_Service
Me.otCodigo.SetFocus
End Sub

Private Sub Command4_Click()
Me.otCodigo.SetFocus
VBA.SendKeys (sgKb_Vef)
End Sub

Private Sub Form_Click()
otCodigo.SetFocus
End Sub

Private Sub Form_Deactivate()
Video_Form.MediaPlayer3.Close
End Sub

Public Sub Form_Load()
sgCmdLine = VBA.Command$
sgParms = VBA.Split(sgCmdLine, ",")
If App.PrevInstance Then
    MsgBox "La aplicacion solicitada [" & App.EXEName & "], ya se esta ejecutando!!!", vbInformation
    End
End If
Call Write_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "APPRUNNING", "1")
Dim MiValor As String
VBA.Randomize
If UBound(sgParms) > -1 Then
    If sgParms(0) = "ACTIVATE" Then
        If UBound(sgParms) = 0 Then
            MiValor = InputBox("Inserte el código de seguridad", "Recreativos Veraguenses", "", 100, 100)
            If VBA.Val(MiValor) <> 2527 Then
                End
            Else
                Act_Form1.Show vbModal
                End
            End If
        ElseIf UBound(sgParms) = 1 Then
            If sgParms(1) = 2527 Then
                Act_Form1.Show vbModal
                End
            End If
        End If
    End If
End If
'---------------------Carga entorno de variables------------------------
On Error GoTo Solve_error
bfTm = False
igCnt_CR = 0: igCnt_CRP = 0: igCnt_CRG = 0
igDelay_Ret_Gen = 0
gbServ_Mode = False
igFlg_TCR = 0
igNoDuplicT = 0
igInd_Kar = 0
igInd_Pub = 0
igDelay_Del_Dig_Can = 0
igLen2 = 0: igNo_RgAt = 0
igAct_PgG = 1: igTot_PgG = 0: igTot_PgC = 0
igAct_PgD = 1: igTot_PgD = 0: igTot_PgC = 0
igMax_Gen = 8: igMax_Dis = 6: igMax_Can = 8 'valores fijos su valor no pueded ser superior
igInd_Bon = 0
igNext_Bonus = 0
piCnt_Canc = 0
igShowPass = 0
bFlagFoc = True
bgBlinkPag = False
bgPopular = False
bgVIP = False
bgIs_Video = False: bgIs_Publi = False
bgWMP_Busy = False
igCont_Sin = 0
Me.oLVersion.Caption = "Ver." & VBA.Trim(VBA.Str(App.Major)) & "." & VBA.Trim(VBA.Str(App.Minor)) & "." & VBA.Trim(VBA.Str(App.Revision))
'Call Chek_NumLockStatus
Call Colocar_Frames
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
sgDir_Pub1 = sParam(6)
sgFec_iAc = sParam(7)
sgFec_Fac = sParam(8)
sgSer_Mac = sParam(9)
sgNom_Loc = VBA.Trim(sParam(10))
sgFle_Fon = sParam(13)
sgSer_CPU = sParam(14)
igCnt_CRG = VBA.Val(UnScramble(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", "0")))
igFlg_TCR = VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "FLG_TESTCR", "0"))
igDelay_Return_Gen = VBA.Int(VBA.Val(sParam(15)))
igDelay_Return_Dis = VBA.Int(VBA.Val(sParam(16)))
igDelay_Bonus_Vid = VBA.Int(VBA.Val(sParam(17)))
igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
igLim_Cred = VBA.Int(VBA.Val(sParam(20)))
igKeep_Cred = VBA.Int(VBA.Val(sParam(21)))
If (igKeep_Cred = 0) Then
    If igFlg_SavedCR = 1 Then
        igCnt_CR = VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "ACU_SAVECR", "#####0"))
    Else
        igCnt_CR = 0
    End If
Else
    igCnt_CR = 6
End If
If igFlg_TCR = 1 Then
    Me.olMetros2.Visible = True
    Me.olMetros2.Caption = PADL(0, 6, "0")
    Me.olTest.Visible = True
Else
    Me.olMetros2.Visible = False
    Me.olTest.Visible = False
End If
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
sgKb_SwP = sParam(51)
sgWin_Key = sParam(43)
sgCr_AKey = sParam(52)
sgIdx_Prm = sParam(53)
bgSw = True
bgVideoLabel = IIf(Val(sParam(44)) = 1, True, False)
bgDiscLabel = IIf(Val(sParam(45)) = 1, True, False)
bgKeep_On_Top = VBA.IIf(sParam(46) = 0, False, True)
igScr_Alone = VBA.Int(VBA.Val(sParam(47)))
igNoDuplicT = VBA.Int(VBA.Val(sParam(48)))
bgSw_Pub = IIf(Val(VBA.Int(sParam(49))) = 1, True, False)
If bgSw_Pub = False Then
    Me.oInd_VideoSW.Picture = LoadPicture(App.Path & "\" & "grnled.gif")
Else
    Me.oInd_VideoSW.Picture = LoadPicture(App.Path & "\" & "redled.gif")
End If
Me.oInd_VideoSW.Visible = True
sgDir_Pub2 = sParam(50)
sgWin_Key = sParam(43)
igLeftDisk = 700
If UBound(sgParms) > -1 Then
    Select Case sgParms(0)
    Case Is = "CONVERSION"
        Call Set_Open_Dbf
        sPath = IIf(igInd_Kar = 0, sgDir_Fls, sgDir_Fls2)
        Call ogVFP9.EXPORT__TABFILES(sPath, sPath, sgDir_Tmp)
        Call MsgBox("La conversión se realizó:..")
        End
    Case Is = "SERVICE"
        If UBound(sgParms) = 0 Then
            MiValor = InputBox("Inserte el código de seguridad", "Flamingo Magic Game", "", 100, 100)
            If VBA.Val(MiValor) <> 2527 Then
                End
            Else
                Call Go_Service
                End
            End If
        Else
            If sgParms(1) = 2527 Then
                Call Go_Service
                End
            End If
        End If
    End Select
End If
Call Check_Other
On Error GoTo Solve_error
'---------------------Verifica activación y serial de la pc--------------
If sgNom_Loc <> VBA.UCase("SIN ASIGNACIÓN!") Then
    Me.olLocal.Caption = "[" & sgNom_Loc & "] "
Else
    Me.olLocal.Caption = ""
End If
Me.olMensajeSis.Caption = ""
Me.olLocal.Caption = Me.olLocal.Caption & sgWin_Key
Dim iReadFlg As Integer
iReadFlg = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "RUNNINGSEC", "0")))
If iReadFlg = 1 Then
    Dim sVal1 As String
    Dim sVal2 As String
    If CheckForKL = False Then
        Call KTASK(TERMINATE, 0, 0, 0)
        End
    Else
        Call KTASK(READAUTH, READCODE1, READCODE2, READCODE3)
        Call Write_Ini_File(App.Path & "\PathV2.ini", "SERIAL", "ID", Scramble("5C01001080000000000666413036"))
    End If
    sVal1 = Scramble(VBA.Trim(ReadText()))
    sVal2 = VBA.Trim(Read_Ini_File(App.Path & "\PathV2.ini", "SERIAL", "ID", ""))
    If (sVal1 <> sVal1) Then
        MsgBox "USB-KEY device no match.", vbCritical
        End
    End If
    Me.olActivacion.Caption = "Próxima activación: UNLIMITED WITH USB-KEY"
Else
    Dim sTmp1 As Variant
    Dim iValue As Integer
    Dim sValue As String
    Dim sMensage As String
    sMensage = "La copia del sistema no ha sido debidamente instalada o no ha sido activada"
    iValue = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "NDISCVALID", "0")))
    If iValue = 0 Then
        sTmp1 = Lee_Serial
        sTmp1 = Left$(sTmp1, 4) & "-" & Right$(sTmp1, 4)
        If sgSer_Mac <> sTmp1 Then
            Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de DISCO")
            Call Call_Tsr
            End
        End If
    End If
    iValue = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "NMOTHVALID", "0")))
    If iValue = 0 Then
        ''sTmp1 = Get_CPU_Id
        sTmp1 = MBCPUNumber()
        If sgSer_CPU <> sTmp1 Then
            Call MsgBox(sMensage, vbCritical, "El sistema a sido movido de MÁQUINA")
            Call Call_Tsr
            End
        End If
    End If
    iValue = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "NDATEVALID", "0")))
    If iValue = 0 Then
        If VBA.DateValue(sgFec_iAc) = VBA.DateValue(sgFec_Fac) Then
            Call MsgBox(sMensage, vbCritical, "La copia del sistemas debe ser activada")
            Call Call_Tsr
            End
        End If
        If VBA.Date < VBA.DateValue(sgFec_iAc) Then
            Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
            Call Call_Tsr
            End
        End If
        If VBA.Date > VBA.DateValue(sgFec_Fac) Then
            Call MsgBox(sMensage, vbCritical, "La copia del sistema ha perdido vigencia")
            Call Call_Tsr
            End
        End If
    End If
    Me.olActivacion.Caption = "Próxima activación: " & VBA.Format(sgFec_Fac, "dd/MM/yyyy")
End If
iShow = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "SHOW_MOUSE", "0")))
Call ShowCursor(IIf(iShow = 1, True, False))
iValue = VBA.Int(VBA.Val(Read_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "DELETEACTI", "0")))
If iValue = 0 Then
    Dim sFile As String
    Set oFsys = CreateObject("Scripting.FileSystemObject")

    sFile = App.Path & "\ActivadorR.exe"
    If FileExist(sFile) Then
        Call oFsys.DeleteFile(sFile)
    End If
    For i = 1 To 9
        sFile = App.Path & "\ActivadorR" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
        sFile = App.Path & "\ActivadorR0" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
    Next i
    sFile = "C:\ActivadorR.exe"
    If FileExist(sFile) Then
        Call oFsys.DeleteFile(sFile)
    End If
    For i = 1 To 9
        sFile = "C:\ActivadorR" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
        sFile = "C:\ActivadorR0" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
    Next i
    sFile = "D:\Activador.exe"
    If FileExist(sFile) Then
        Call oFsys.DeleteFile(sFile)
    End If
    For i = 1 To 9
        sFile = "D:\ActivadorR" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
        sFile = "D:\ActivadorR0" & i & ".exe"
        If FileExist(sFile) Then
            Call oFsys.DeleteFile(sFile)
        End If
    Next i
End If
'------------------------------------------------------------
Dim oFs As Object
Dim sWinDir As String
Dim oA As Object
Set ogVFP9 = Nothing
sWinDir = GetTheWindowsDirectory()
Set oFs = CreateObject("Scripting.FileSystemObject")
If Set_Open_Dbf() = False Then
    If ((oFs.FileExists(App.Path & "\LIBRARY.DLL") = True) And (oFs.FileExists(App.Path & "\FOXTOOLS.FLL") = True)) Then
        Call MsgBox("No existen alguna de las librerías [LIBRARY.DLL,FOXTOOLS.FLL] ", vbCritical, "Error")
    End If
End If
Err.Clear

Call Set_Open_Dbf
Call Set_Tmp_DBF
Call Set_DBF_To_Tmp
If Err.Number > 0 Then
    End
End If
'Desordenar_array
Dim vFound As Boolean
vFound = False
   
Me.oLst_Temas_Video.Path = sgDir_Mp3
Me.oLst_Temas_Video.Pattern = "*.MPG"
Me.oLst_Temas_Video.Refresh
If Me.oLst_Temas_Video.ListCount <= 0 Then
    vFound = False
Else
    vFound = True
End If
If vFound = False Then
    Me.oLst_Temas_Video.Pattern = "*.MPG"
    Me.oLst_Temas_Video.Path = sgDir_Mp3
'    Me.oLst_Temas_Video.Refresh
End If

Me.oFrame_Dis.TransparentBackground = True
Me.oFrame_Gen.TransparentBackground = True
Me.oFrame_Can.TransparentBackground = True

Me.oTM_codigo2.Interval = igDelay_Return_Dis * 1000
Me.oTM_codigo2.Enabled = False
Call Refresh_Creditos(Me)
'----------------------GENEROS------------------------------
Call Cargar_Gen
'---------------------PUBLICIDAD----------------------------
Call Conectar_DBPub
'-----------------------PROMOS------------------------------
Call Conectar_DBPro
'-----------------------Otros-------------------------------
otNot_Found_List.Text = ""
Call Cargar_Temas
If igScr_Alone = 0 Then
    Video_Form.Show
End If
If bgKeep_On_Top = True Then
    Call AlwaysOnTop(Main_Form, True)
End If
Exit Sub

Solve_error:
Call ControlError
Resume Next
End Sub

Private Sub Conectar_DBGen(Optional iClose As Integer = 0)
Dim sSql As String
Dim sConnectionString As String
sSql = "SELECT File01.* FROM File01 WHERE File01.gen_st=0 ORDER BY File01.id_ord "
Me.oDC_GEN.ConnectionString = ""
Me.oDC_GEN.RecordSource = ""
With Me.oDC_GEN
    sConnectionString = "Provider=VFPOLEDB.1;Data Source=" & sgDir_Tmp & ";Password='';Collating Sequence=MACHINE"
    .ConnectionString = sConnectionString
'   .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    If iClose = 1 Then
        .Recordset.Close
    End If
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ShowCursor(True)
Me.oTM_Mouse.Enabled = True
End Sub

Public Sub Form_Unload(Cancel As Integer)
Call Write_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "APPRUNNING", "0")
Call Upd_Cnt
Call ogVFP9.Set_Files_Close
End Sub

Private Sub Image1_Click(Index As Integer)
Dim sVal As String
sVal = Me.Cap_Label_ID(Me.oLNum_Disk.Item(Index).Caption)
If VBA.Trim(sVal) = "?" Then
    sVal = ""
End If
Call VBA.SendKeys(sVal)
End Sub

Private Sub Image3_Click()
Call VBA.SendKeys(sgKb_Ret)
otCodigo.SetFocus
End Sub

Private Sub Label7_DblClick()
Call Show_Hide_Service
End Sub

Private Sub LabelBox1_Click()

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
    bfTm = False
Case Is = wmppsPlaying
    Me.otCargador_Music.Enabled = True
    bgIs_Video = False
    bgWMP_Busy = False
    bfTm = True
    Me.oTM_Box.Enabled = True
End Select
End Sub

Private Sub MediaPlayer2_MediaError(ByVal pMediaObject As Object)
If igScr_Alone = 1 Then
    Main_Form.olMessage.Visible = True
    Main_Form.olMessage.Caption = "TEMA NO DISPONIBLE"
    Main_Form.oTime_Mensajes.Enabled = True
    Call Remove_Temes
    If igKeep_Cred = 0 Then
        igCnt_CR = igCnt_CR + 1
    End If
    Call Refresh_Creditos(Main_Form)
    Sleep 3 '* 1000 'Implements a 3 second delay
    VBA.SendKeys ("S")
End If
End Sub

Private Sub MediaPlayer2_PlayStateChange(ByVal NewState As Long)
If igScr_Alone = 1 Then
    Select Case NewState
    Case Is = wmppsMediaEnded
        If bgIs_Video = True Then
            If igScr_Alone = 1 Then
                bfTm = False
            End If
        End If
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
        If bgIs_Video = True Then
            If igScr_Alone = 1 Then
                bfTm = True
                Me.oTM_Box.Enabled = True
            End If
        End If
        If igCont_Sin > 0 Then
            Exit Sub
        End If
        bgWMP_Busy = True
        '*Main_Form.MediaPlayer2.URL = Video_Form.MediaPlayer3.URL
        'Main_Form.MediaPlayer2.settings.mute = False
        '*Main_Form.MediaPlayer2.Controls.currentPosition = Video_Form.MediaPlayer3.Controls.currentPosition
        '*Main_Form.MediaPlayer2.Controls.play
        igCont_Sin = igCont_Sin + 1
    End Select
End If
End Sub

Private Sub oChk_FndC_Click()
If bFlagChek2 = True Then
    Me.otRuteExternal.Text = ""
End If
If Me.oChk_FndC.value = 1 Then
    Me.oGetRute.Enabled = True
    Me.otRuteExternal.Enabled = True
Else
    Me.oGetRute.Enabled = False
    Me.otRuteExternal.Enabled = False
End If
End Sub

Private Sub oChk_FndF_Click()
If bFlagChek1 = True Then
    Me.otRuteFiles.Text = ""
End If
If Me.oChk_FndF.value = 1 Then
    Me.oGetdFiles.Enabled = True
    Me.otRuteFiles.Enabled = True
Else
    Me.oGetdFiles.Enabled = False
    Me.otRuteFiles.Enabled = False
End If
End Sub

Private Sub oChk_FndP_Click()
If bFlagChek3 = True Then
    Me.otRuteExternal2.Text = ""
End If
If Me.oChk_FndP.value = 1 Then
    Me.oGetRute2.Enabled = True
    Me.otRuteExternal2.Enabled = True
Else
    Me.oGetRute2.Enabled = False
    Me.otRuteExternal2.Enabled = False
End If
End Sub

Private Sub oControl_Click(Index As Integer)
otCodigo.SetFocus
Select Case Index
Case Is = 1
    Call VBA.SendKeys(sgKb_ResA)
Case Is = 2
    Call VBA.SendKeys(sgKb_ResM)
Case Is = 3
    Call VBA.SendKeys(sgKb_VIP)
Case Is = 4
    Call VBA.SendKeys(sgKb_Pop)
Case Is = 1
    Call VBA.SendKeys(sgKb_ResM)
End Select
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
    'If oImg_PagUp.Visible = True Then
        'If oImg_PagUp.BorderStyle = 0 Then
        '    oImg_PagUp.BorderStyle = 1
        'Else
        '    oImg_PagUp.BorderStyle = 0
        'End If
    'End If
    'If oImg_PagDn.Visible = True Then
        'If oImg_PagDn.BorderStyle = 0 Then
        '    oImg_PagDn.BorderStyle = 1
        'Else
        '    oImg_PagDn.BorderStyle = 0
        'End If
    'End If
    'AQUI
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
    'Me.oImg_PagDn.BorderStyle = 0
    'Me.oImg_PagUp.BorderStyle = 0
    'AQUI
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

Private Sub oGetdFiles_Click()
Dim lcSelectedPath As String
Me.otRuteFiles.Text = ""
Me.Bbgetdir1.FocusedDirectory = "C:\"
Me.Bbgetdir1.ListAutoCenter = True
Me.Bbgetdir1.StatusText = "Origen del DBContainer?"
lcSelectedPath = Me.Bbgetdir1.ShowDirectoryListEx(1) + "\"
If lcSelectedPath <> "" Then
    Me.otRuteFiles.Text = VBA.Trim(lcSelectedPath)
Else
    Me.otRuteFiles.Text = ""
End If
Me.otRuteFiles.Refresh
End Sub

Private Sub oGetOgen_Click()
Dim lcSelectedPath As String
Dim sRuta As String
Me.otRuteExternal.Text = ""
Me.Bbgetdir1.FocusedDirectory = "C:\"
Me.Bbgetdir1.ListAutoCenter = True
Me.Bbgetdir1.StatusText = "Origen de datos?"
lcSelectedPath = Me.Bbgetdir1.ShowDirectoryListEx(1)
If lcSelectedPath <> "" Then
    Me.otOrigen.Text = VBA.Trim(lcSelectedPath)
Else
    Me.otOrigen.Text = ""
    Me.otRuteExternal.Text = ""
    Me.otRuteExternal2.Text = ""
End If
Me.otOrigen.Refresh
sRuta = VBA.Trim(otOrigen.Text)
Me.otRuteExternal.Text = sRuta & "\CANCIONERO\"
Me.otRuteExternal2.Text = sRuta & "\FOTOS\"
Me.otRuteFiles.Text = sRuta & "\Files\"
If DirExists(VBA.Trim(Me.otRuteFiles.Text)) = False Then
    bFlagChek1 = True
    Me.oChk_FndF.ForeColor = &HFF&
    Me.oChk_FndF.value = 0
Else
    Me.oChk_FndF.ForeColor = &HFF00&
    Me.oChk_FndF.value = 1
End If
If DirExists(VBA.Trim(Me.otRuteExternal.Text)) = False Then
    bFlagChek2 = True
    Me.oChk_FndC.ForeColor = &HFF&
    Me.oChk_FndC.value = 0
Else
    Me.oChk_FndC.ForeColor = &HFF00&
    Me.oChk_FndC.value = 1
End If
If DirExists(VBA.Trim(Me.otRuteExternal2.Text)) = False Then
    bFlagChek3 = True
    Me.oChk_FndP.ForeColor = &HFF&
    Me.oChk_FndP.value = 0
Else
    Me.oChk_FndP.ForeColor = &HFF00&
    Me.oChk_FndP.value = 1
End If
End Sub

Private Sub oGetRute_Click()
Dim lcSelectedPath As String
Me.otRuteExternal.Text = ""
Me.Bbgetdir1.FocusedDirectory = "C:\"
Me.Bbgetdir1.ListAutoCenter = True
Me.Bbgetdir1.StatusText = "Origen del cancionero?"
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

Private Sub oImg_PagDn_Click()
otCodigo.SetFocus
Call VBA.SendKeys(sgKb_DN)
End Sub

Private Sub oImg_PagUp_Click()
otCodigo.SetFocus
Call VBA.SendKeys(sgKb_UP)
otCodigo.SetFocus
End Sub

Private Sub oLCanc_Click(Index As Integer)
Dim sVal As String
sVal = Cap_Label_ID(Me.oLCanc.Item(Index).Caption)
Call VBA.SendKeys(sVal)
otCodigo.SetFocus
End Sub

Private Sub oLGenero_Click(Index As Integer)
Dim sVal As String
sVal = Cap_Label_ID(Me.oLGenero.Item(Index).Caption)
Call VBA.SendKeys(sVal)
otCodigo.SetFocus
End Sub

Public Function Cap_Label_ID(ByVal sLabel As String) As String
Dim sValue As String
sValue = VBA.Left(sLabel, 2)
Cap_Label_ID = sValue
otCodigo.SetFocus
End Function

Private Sub oSetFocus_Codigo_GotFocus()
otCodigo.Text = ""
otCodigo.SetFocus
End Sub

Private Sub otCargador_Music_Timer()
On Error GoTo Solve_error
Dim iCont As Integer
Dim pCa_Ato As String, sCadenas As String
If Me.oLst_A_Tocar.List(0) = "" Then
    Me.oImg_c_Video.Visible = False
    If igDelay_Bonus_Vid > 0 Then
       If igNext_Bonus <= (Hour(Time()) * 60) + Minute(Time()) Then
            If Me.oDC_Promos.Recordset.RecordCount <= 0 Then
                igInd_Bon = igInd_Bon + 1
                Me.oLst_Temas_Video.ListIndex = igInd_Bon - 1
                sCadenas = "999999,999999,VIDEO BONUS," & Me.oLst_Temas_Video.Path & "\" & Me.oLst_Temas_Video.filename & ",.,."
            Else
                sCadenas = Add_Promos
            End If
            If Not (VBA.Trim(sCadenas) = "") Then
                Me.oLst_A_Tocar.AddItem sCadenas
                igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
            End If
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

Private Function Add_Promos()
Dim pOrd_Gen As String, pOrd_Dis As String, pOrd_Can As String
Dim pCo_Can As String, pFl_MP3 As String, sDes_Can As String
Dim pFl_Dis As String, sCad2 As String
If Me.oDC_Promos.Recordset.RecordCount <= 0 Then
    Add_Promos = ""
    Exit Function
End If
pOrd_Gen = Me.oDC_Promos.Recordset.Fields("ID_ORD1").value
pOrd_Dis = Me.oDC_Promos.Recordset.Fields("ID_ORD2").value
pOrd_Can = Me.oDC_Promos.Recordset.Fields("ID_ORD").value
pCo_Can = Me.oDC_Promos.Recordset.Fields("ID_CAN").value
pFl_MP3 = Me.oDC_Promos.Recordset.Fields("FL_mp3").value
pFl_Dis = Me.oDC_Promos.Recordset.Fields("FL_IMG").value
sDes_Can = Me.oDC_Promos.Recordset.Fields("DE_CAN").value
Add_Promos = VBA.Trim(pOrd_Gen & pOrd_Dis & pOrd_Can) & "," & VBA.Trim(pCo_Can) & "," & VBA.Trim(sDes_Can) & "," & VBA.Trim(pFl_MP3) & "," & "*" & "," & VBA.Trim(pFl_Dis)
If Not Me.oDC_Promos.Recordset.EOF() = True Then
    Me.oDC_Promos.Recordset.MoveNext
Else
    Me.oDC_Promos.Recordset.MoveFirst
End If
End Function

Private Sub otCargador_Video_Timer()
On Error GoTo Solve_error
Dim iLimCnt As Integer
'Si no hay mas temas en lista para tocar no mostrar màs videos.
If Me.oLst_A_Tocar.List(0) = "" Then
    'Me.otCodigo.SetFocus
    Exit Sub
End If
'Si no hay mas videos que presentar, salir.
iLimCnt = UBound(agArr_Pub1)
'If Me.oLst_Pub1.List(0) = "" Then
'    Me.olMessage.Visible = True
'    Me.olMessage.Caption = "LA LISTA DE PUBLICIDAD ESTA VACÍA!"
'    Me.oTime_Mensajes.Enabled = True
'    Exit Sub
'End If
If iLimCnt = 0 Then
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "LA LISTA DE PUBLICIDAD ESTA VACÍA!"
    Me.oTime_Mensajes.Enabled = True
    Exit Sub
End If
Dim sFle_MpG As String
Dim sFle_Tmp As String
Dim aDet() As String
sFle_Tmp = Me.MediaPlayer1.URL
If VBA.UCase(VBA.Right(sFle_Tmp, 3)) <> "MP3" Then
    Exit Sub
End If
'iLimCnt = Me.oLst_Pub1.ListCount - 1
If igInd_Pub = 0 Then
    igInd_Pub = 1
End If
'Me.oLst_Pub1.ListIndex = (igInd_Pub - 1)
'sFle_MpG = VBA.Trim(Me.oLst_Pub1.Text)
'sFle_MpG = Me.oLst_Pub1.Path & "\" & VBA.Trim(Me.oLst_Pub1.filename)
sFle_MpG = IIf(bgSw_Pub = False, Me.oLst_Pub1.Path, Me.oLst_Pub2.Path) & "\" & agArr_Pub1(igInd_Pub)
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
Dim sValue As String, sValSel As String, sCad_Ato2 As String
Dim iNum_Pag As Integer, iNum_Pos As Integer, iLen As Integer

sValue = VBA.Trim(otCodigo.Text)
igLen = VBA.Len(sValue)

If gbServ_Mode = True Then
    If Me.otCodigo.EditMask = "##-##" Then
        Me.otCodigo.EditMask = "##-##-##"
        Me.otCodigo.Tag = "0"
    End If
Else
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
        Else
            If Me.otCodigo.EditMask = "##-##" Then
                Me.otCodigo.EditMask = "##-##-##"
                Me.otCodigo.Tag = "0"
            End If
        End If
    Else
        igGen_Sel = "": igDis_Sel = "": igCan_Sel = ""
    End If
End If
Me.olPosCod.Caption = "Posición -> " & VBA.Trim(VBA.Str(igLen))
'   *******************************************************************************
If Trim(Me.otCodigo.Text) = "99" Then
    If igShowPass = 1 Then
        If gbServ_Mode = False Then
            Main_Form.Tag = ""
            Main_Form.otCodigo.Text = ""
            Call AlwaysOnTop(Main_Form, False)
            Pass_Scr.Show vbModal
            Call AlwaysOnTop(Main_Form, bgKeep_On_Top)
            If VBA.Trim(Main_Form.Tag) <> sgCr_AKey Then
                Me.olMessage.Caption = "ACCESO DENEGADO!!!"
                Me.olMessage.Visible = True
                Me.oTime_Mensajes.Enabled = True
                Me.olMessage.Tag = ""
                Me.oSetFocus_Codigo.SetFocus
                gbServ_Mode = False
                Me.oService_Info.filename = ""
                Me.oService_Info.Visible = False
                Exit Sub
            End If
            Me.olMessage.Caption = "ACCESO CONSEDIDO!"
            Me.olMessage.Tag = ""
            Me.olMessage.Visible = True
            'Me.oTime_Mensajes.Enabled = True
            pTimeResto = (2 * 30)
            gbServ_Mode = True
            Me.Timer3.Enabled = True
            Me.oService_Info.filename = App.Path & "\Service_Info.RTF"
            Me.olT_Mant.Visible = True
            Me.oService_Info.Top = 2160
            Me.oService_Info.Width = 6735
            Me.oService_Info.Left = 5760
            Me.oService_Info.Height = 4455
            Me.oService_Info.Visible = True
        End If
    End If
End If
'   *******************************************************************************
Select Case igLen
Case Is = 0
    olTimer1(1).Caption = 0
    olTimer1(2).Caption = 0
   
    olTimer1(1).Visible = False
    olTimer1(2).Visible = False
    
    igGen_Sel = ""
    igDis_Sel = ""
    igCan_Sel = ""

    igAct_PgG = 1
    igAct_PgD = 1
    igAct_PgC = 1

    igNext_Return_Gen = 0
    Me.Image2.Visible = False
    'Me.Image2.Picture = LoadPicture()
    Call Cargar_Gen
    'Call Desactiva_Disco(True)
    'Call Desactiva_Cancion(True)
 Case 1 To 2
    If gbServ_Mode = True Then
            Call Desactiva_Cancion(True)
            Call Desactiva_Disco(True)
            Call Desactiva_Genero(True)
            Me.oService_Info.filename = App.Path & "\Service_Info.RTF"
            Me.olT_Mant.Visible = True
            Me.oService_Info.Visible = True
        Exit Sub
    End If
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
        Me.oLTitulo.Caption = "[Seleccione le Disco]..."
    End If
    'igAct_PgD = 1
    igAct_PgC = 1
    igAct_PgG = iNum_Pag
    Call Desactiva_Genero(True)
    '***********************DISCOS************************
    Call Conectar_DBDis(sGen_Ret)
    Call Cargar_INF_Dis(sGen_Ret, igMax_RgD)
    Call Cargar_Pag_Dis(igAct_PgD, igMax_RgD)

    'Me.oLst_Gen.BoundText = igGen_Sel
    'Me.oLTitulo.Caption = Me.oLst_Gen.Text
    Call Desactiva_Disco(False)
    igGen_Sel = sGen_Ret
    bgVIP = False
    igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
    If Val(olTimer1(1).Caption) = 0 Then
        olTimer1(1).Caption = Format(igDelay_Return_Gen * 60, "#0")
        olTimer1(1).Visible = True
    End If
Case 3 To 4
    If gbServ_Mode = True Then
        Exit Sub
    End If
   
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
            Me.oLTitulo.Caption = "[Seleccione le Disco]..."
        End If
        'igAct_PgD = iNum_Pos
        Call Retrocede
        Call Limpia_Dis
        Exit Sub
    Else
        Me.oLst_Disc.BoundText = sDis_Ret
        Me.oLTitulo.Caption = Me.oLst_Disc.Text
    End If
    igAct_PgC = 1
    X = iNum_Pag
    igAct_PgD = iNum_Pag
    If VBA.Right(VBA.Trim(sFle_Img), 1) = "\" Then
        sFle_Img = ""
    End If
    Me.Image2.Picture = LoadPicture(sFle_Img)
    Me.Image2.Visible = True
    Call Desactiva_Genero(True)
    Call Desactiva_Disco(True)
    Call Desactiva_Cancion(True)
    '***********************CANCION************************
    Call Conectar_DBCan(igGen_Sel, sDis_Ret)
    Call Cargar_INF_Can(igGen_Sel, sValSel, sDis_Ret, igMax_RgC)
    Call Cargar_Pag_Can(igAct_PgC, igMax_RgC)
    Call Desactiva_Cancion(False)
    igDis_Sel = sDis_Ret
    bgVIP = False
Case 5 To 6
    If gbServ_Mode = True Then
        Select Case VBA.Trim(Me.otCodigo.Text)
        Case Is = "990001"
            Call VBA.SendKeys(sgKb_SwP)
            Call VBA.SendKeys(".")
            Call VBA.SendKeys(".")
            Call VBA.SendKeys(".")
        Case Is = "990002"
            Call VBA.SendKeys(sgKb_Crd2)
        Case Is = "990003"
            Call VBA.SendKeys(sgKb_Del)
            Call VBA.SendKeys(sgKb_Del)
            Call VBA.SendKeys(sgKb_Del)
        Case Is = "990004"
            Call VBA.SendKeys(sgKb_ResM)
        Case Is = "990005"
            Call VBA.SendKeys(sgKb_ResA)
        Case Is = "990006"
            pTimeResto = 1
        Case Is = "990007"
            Call VBA.Shell(App.Path & "\osk.exe")
        Case Is = "990008"
            Call Show_Hide_Service
        Case Is = "990009"
            Call VBA.SendKeys(sgKb_SwK)
        Case Is = "990010"
            Call Go_Service
        Case Is = "990011"
            VBA.Shell (App.Path & "\TSR.exe")
        Case Is = "990012"
            VBA.Shell ("EXPLORER.EXE")
        Case Is = "990013"
            Unload Video_Form
            End
        Case Is = "990014"
           If igFlg_TCR = 1 Then
               igFlg_TCR = 0
            Else
               igFlg_TCR = 1
            End If
            Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "FLG_TESTCR", VBA.Format(igFlg_TCR, "#####0"))
        Case Else
            If igLen = 6 Then
                Me.otCodigo.Text = ""
            End If
        End Select
        olTimer1(2).Caption = Format(Me.oTM_codigo2.Interval / 1000, "###0")
        Me.oTM_codigo2.Enabled = True
        
        Me.olMessage.Caption = ""
        Me.olMessage.Tag = "1"
        Me.olMessage.Visible = True
        Me.oTimer_Srv.Enabled = True
        Exit Sub
    End If
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
                        If igCnt_CR > 0 Then
                            igCnt_CR = igCnt_CR - 1
                        End If
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
                            If igCnt_CR > 0 Then
                                igCnt_CR = igCnt_CR - 1
                            End If
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
            '---Segmento que añade los bonus de promociones--
            If sgIdx_Prm > 0 Then
                piCnt_Canc = piCnt_Canc + 1
                If piCnt_Canc >= sgIdx_Prm Then
                    piCnt_Canc = 0
                    sCad_Ato2 = Add_Promos
                    If Not (VBA.Trim(sCad_Ato2) = "") Then
                        Call Me.oLst_A_Tocar.AddItem(sCad_Ato2)
                        Me.olMessage.Visible = True
                        Me.olMessage.Caption = "DISCO PROMO EN COLA!"
                        Me.oTime_Mensajes.Enabled = True
                    End If
                End If
            End If
            '------------------------------------------------
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
                olTimer1(2).Caption = Format(Me.oTM_codigo2.Interval / 1000, "###0")
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
            olTimer1(2).Caption = Format(Me.oTM_codigo2.Interval / 1000, "###0")
            Me.oTM_codigo2.Enabled = True
        End If
        Call Salvar_Temas
    End If
olTimer1(2).Visible = True
End Select
Exit Sub

Solve_error:
Call ControlError
Resume Next

End Sub

Private Sub otCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 123) Then 'F12 PARA SALIR DEL SISTEMA
    Call Form_Unload(1)
    Unload Video_Form
    End
End If
If (KeyCode = 122) Then 'F11 PARA MENÚ DE SERVICIO
    Call Show_Hide_Service
End If
End Sub

Private Sub otCodigo_KeyPress(KeyAscii As Integer)
Dim iLimCnt As Integer
Dim sTmp As String
igKeyAscii = KeyAscii
If VBA.IsNumeric(VBA.Chr(KeyAscii)) Then
    Exit Sub
End If
If KeyAscii = 8 Then
    Me.SetFocus
    Exit Sub
End If
If Inlist(VBA.UCase(VBA.Chr(KeyAscii)), sgKb_SwK) Then
    igInd_Kar = IIf(igInd_Kar = 0, 1, 0)
    If igInd_Kar = 0 Then
        Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DefaultDir", sgDir_Fls)
        Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DBQ", sgDir_Fls)
    Else
        Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DefaultDir", sgDir_Fls2)
        Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DBQ", sgDir_Fls2)
    End If
    Call Write_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "SWITCH_KAR", VBA.Format(igInd_Kar, "#####0"))
    Call Write_Ini_File(App.Path & "\PathV2.ini", "ROCKOLA", "RELOAD_APP", "1")
    Call VBA.Shell(App.Path & "\R_monitor.exe", vbNormalNoFocus)
    End
    Exit Sub
End If
If Inlist(VBA.UCase(VBA.Chr(KeyAscii)), sgKb_Pause) Then
    If Me.oLst_A_Tocar.List(0) = "" Then
        Exit Sub
    End If
    'If igScr_Alone = 1 Then
        If bgIs_Video = True Then
            If Me.MediaPlayer2.playState = wmppsPaused Then
                Me.MediaPlayer2.Controls.play
                Video_Form.MediaPlayer3.Controls.play
                Me.otCargador_Music.Enabled = True
            Else
                Me.MediaPlayer2.Controls.pause
                Video_Form.MediaPlayer3.Controls.pause
                Me.otCargador_Music.Enabled = False
            End If
       End If
    'Else
    'End If
    Exit Sub
End If
'133
If Inlist(VBA.UCase(VBA.Chr(KeyAscii)), sgKb_Vef) Then
    bgExit = False
    bFlagChek = False
    If TBack4.Visible = False Then
        TBack4.Visible = True
    Else
        TBack4.Visible = False
    End If
    Exit Sub
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_SwP) Then
    Me.olMessage.Visible = True
    Me.olMessage.Caption = "SWITCHING VIDEO..."
    Me.oTime_Mensajes.Enabled = True
    igInd_Pub = 0
    If bgSw_Pub = False Then
        Me.oInd_VideoSW.Picture = LoadPicture(App.Path & "\" & "redled.gif")
        Me.oInd_VideoSW.Visible = True
    Else
        'Me.oInd_VideoSW.Picture = LoadPicture(App.Path & "\" & "grnled.GIF")
        Me.oInd_VideoSW.Visible = False
   End If
   bgSw_Pub = IIf(bgSw_Pub = True, False, True)
   Call Conectar_DBPub
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
    
    Me.oLDuracion.Caption = "00:00"
    Me.olAct_Pos.Caption = "00:00"
    bfTm = False

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
    Call Cargar_Gen
    Call Muestra_Tema_Det
    Call Refresh_Creditos(Me)
    bgWMP_Busy = False
    igCont_Sin = 0
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_ResA) Then
    'Sección que se ejecuta si se preciona [R/r] (Resert all)
    
    Me.oLDuracion.Caption = "00:00"
    Me.olAct_Pos.Caption = "00:00"
    bfTm = False
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
    'If Me.oLst_Pub1.ListCount > 0 Then
    '    Call Me.oLst_Pub1.RemoveItem(0)
    'End If
    Call Cargar_Gen
    Call Muestra_Tema_Det
    Call Refresh_Creditos(Me)
    bgWMP_Busy = False
    igCont_Sin = 0
    If (igKeep_Cred = 0) Then
        igCnt_CR = 0
        If igFlg_SavedCR = 1 Then
            Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "ACU_SAVECR", VBA.Format(igCnt_CR, "#####0"))
        End If
    Else
        igCnt_CR = 6
    End If
    Me.oSetFocus_Codigo.SetFocus
    Exit Sub
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_Crd1) Then
    Call Add_CR1
    Call Add_CR2
    Call Salvar_Temas
End If

If Inlist(VBA.Chr(KeyAscii), sgKb_Del) Then
    'Sección que se ejecuta si se preciona [-] (Crédito)
    If igKeep_Cred = 0 Then
        If igCnt_CR > 0 Then
            igCnt_CR = igCnt_CR - 1
            Call Refresh_Creditos(Me)
        End If
    End If
    Exit Sub
End If
If Inlist(VBA.Chr(KeyAscii), sgKb_Crd2) Then
    If VBA.Trim(sgCr_AKey) = "" Then
        Me.olMessage.Caption = "NO SE HA CONDIGURADO ACCESO"
        Me.olMessage.Visible = True
        Me.oTime_Mensajes.Enabled = True
        Me.oSetFocus_Codigo.SetFocus
        Exit Sub
    End If
    igCnt_CR = igCnt_CR + 6
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
        Call Cargar_Gen
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
    
    igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
    olTimer1(1).Caption = Format(igDelay_Return_Gen * 60, "#0")
    olTimer1(1).Visible = True
Case 4 To 6
    Me.oFrame_Gen.Visible = False
    Me.oFrame_Dis.Visible = False
    bStatus = Me.oFrame_Can.Visible
    Me.oFrame_Can.Visible = False
    Call Desplazar_Pantalla_3(KeyAscii)
    Me.oFrame_Can.Visible = bStatus
    
    igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
    olTimer1(1).Caption = Format(igDelay_Return_Gen * 60, "#0")
    olTimer1(1).Visible = True
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
        'VBA.DoEvents
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
        'VBA.DoEvents
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
            pSt_001 = VBA.Trim(pCa_Ato) & "," & VBA.Trim(pCo_Can) & "," & VBA.Trim(sDes_Can) & "," & VBA.Trim(pFl_MP3) & ",.,."
            Busca_Sel_3 = True
            Exit Function
        End If
        'VBA.DoEvents
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
    'VBA.DoEvents
Next i
End Function
Private Sub Conectar_DBDis(ByVal pCod_Gen As String, Optional iClose As Integer = 0)
Dim sConnectionString As String
Dim sSql As String
If (pCod_Gen = igGen_Sel) Then
    Exit Sub
End If
sSql = "SELECT File02.* FROM File02 WHERE  File02.ID_GEN =" & pCod_Gen & " AND File02.dis_st=0 ORDER BY File02.ID_ORD "
With Me.oDC_DISC
    sConnectionString = "Provider=VFPOLEDB.1;Data Source=" & sgDir_Tmp & ";Password='';Collating Sequence=MACHINE"
    .ConnectionString = sConnectionString
'   .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    If iClose = 1 Then
        .Recordset.Close
    End If
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
        Exit Function
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
        aPag_Disc(iCnt_Pag).Discos(iNum_reg).FL_NEW = .Fields("FL_NEW").value
        aPag_Disc(iCnt_Pag).No_Rgs = iNum_reg
        iNum_reg = iNum_reg + 1
        .MoveNext
    Loop
End With
igTot_PgD = iCnt_Pag
End Function

Private Sub Conectar_DBCan(ByVal pCod_Gen As String, ByVal pCod_Dis As String, Optional iClose As Integer = 0)
Dim sConnectionString As String
Dim sSql As String
If (pCod_Dis = igDis_Sel) Then
    Exit Sub
End If
sSql = "SELECT File01.id_gen, File03.* " & _
"FROM file01 " & _
"INNER JOIN file02 ON  File01.id_gen = File02.id_gen " & _
"INNER JOIN file03 ON  File02.id_dis = File03.id_dis " & _
"WHERE file01.ID_GEN =" & pCod_Gen & " " & _
"AND   File03.ID_DIS =" & pCod_Dis & " " & _
"AND   File01.gen_st =0 AND File02.dis_st=0 " & _
"ORDER BY File03.id_ord "
With Me.oDC_CANC
    sConnectionString = "Provider=VFPOLEDB.1;Data Source=" & sgDir_Tmp & ";Password='';Collating Sequence=MACHINE"
    .ConnectionString = sConnectionString
'   .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    .CommandType = adCmdText
    .RecordSource = sSql
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    If iClose = 1 Then
        .Recordset.Close
    End If
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
        Exit Function
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
        sCadena = sFld0 & "**" & "," & VBA.Trim(sFld2) & "," & VBA.Trim(sFld3) & "," & VBA.Trim(sFld4) & ",.,."
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
    Me.oLGenero(i).ToolTipText = "Seleccionar el GÉNERO [" & VBA.Left(Me.oLGenero(i).Caption, 2) & "]"
    Me.oLGenero(i).MousePointer = 99
    Me.oLGenero(i).BorderStyle = 1
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
    Me.oLGenero(i).ToolTipText = ""
    Me.oLGenero(i).MousePointer = 0
    Me.oLGenero(i).BorderStyle = 0
Next i
End Sub

Private Function Cargar_Pag_Dis(ByVal pNum_Pag As Integer, Optional ByRef pNoReg As Integer)
Dim sFile As String, sLabel As String
Dim i As Integer, iPos_Vac As Integer
If igDis_Sel = "99" Then
    Call Limpia_Dis
    Call Limpia_New
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
Me.olMensajeSis.Caption = "Cargando portadas de discos..."
For i = 1 To pNoReg
    sFile = VBA.Trim(aPag_Disc(pNum_Pag).Discos(i).FL_IMG)
    sLabel = aPag_Disc(pNum_Pag).Discos(i).ID_ORD
    If VBA.Right(sFile, 1) = "\" Then
        sFile = ""
'        sLabel = "?"
    End If
    Me.Image1(i).MousePointer = 99
    If FileExist(sFile) Then
        Me.Image1(i).Picture = LoadPicture(sFile)
        With Me.oLNum_Disk(i)
            .Caption = sLabel
            .Visible = True
            .Left = Val(Me.oLNum_Disk(i).Tag)
        End With
        If VBA.Int(aPag_Disc(pNum_Pag).Discos(i).C_VIDEO) = 1 Then
            If bgVideoLabel = True Then
                Me.oLNum_Disk(i).Left = VBA.Val(Me.oLNum_Disk(i).Tag) - 600
                With olVideo(i)
                    .Caption = "<VIDEO>"
                    .Visible = True
                    .Tag = "1"
                End With
            End If
        Else
            Me.oLNum_Disk(i).Left = Val(Me.oLNum_Disk(i).Tag)
            With olVideo(i)
                .Caption = ""
                .Visible = False
                .Tag = "1"
            End With
        End If
        If aPag_Disc(pNum_Pag).Discos(i).FL_NEW = 1 Then
            Me.olNuevo(i).Visible = True
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
        Me.Image1.Item(i).ToolTipText = "Seleccionar el DISCO [" & Me.oLNum_Disk.Item(i).Caption & "]"
    Else
        Me.Image1(i).Picture = LoadPicture()
        With Me.oLNum_Disk(i)
            .Caption = sLabel
            .Visible = True
            .Left = Val(Me.oLNum_Disk(i).Tag)
        End With
        Me.olNuevo(i).Visible = False
        Me.ofLabelCont(i).Visible = False
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
        Me.Image1.Item(i).ToolTipText = "DISCO NO DISPONIBLE."
    End If
    'VBA.DoEvents
Next i
'---------------------Se limpian las posiciones que no se usan---------------------
iPos_Vac = pNoReg + 1
Me.oFrame_Dis.Visible = False
Call Limpia_Dis(iPos_Vac)
Call Limpia_New(iPos_Vac)
igMax_RgD = pNoReg
Me.oFrame_Dis.Visible = True
Call Refresh_Paginero(igAct_PgD, igTot_PgD)
Me.olMensajeSis.Caption = ""
End Function

Private Sub Borra_Video_Signal()
Dim i As Integer
For i = 1 To olVideo.Count
    With olVideo(i)
        .Caption = ""
        .Visible = False
        .Tag = "0"
    End With
    'VBA.DoEvents
Next i
End Sub

Private Sub Limpia_New(Optional ByVal piPos As Integer = 1)
For i = piPos To Me.olNuevo.Count
    With Me.olNuevo(i)
        .Visible = False
    End With
Next
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
    With Me.olNuevo(i)
        .Visible = False
    End With
    Me.ofLabelCont(i).Visible = False
    Me.Image1(i).Picture = LoadPicture()
    Me.Image1(i).MousePointer = 0
    Me.Image1.Item(i).ToolTipText = ""
    'VBA.DoEvents
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
Me.olMensajeSis.Caption = "Cargando temas del discos..."
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
    Me.oLCanc(i).Caption = VBA.Trim(sLabel)
    Me.oLCanc(i).ToolTipText = "Seleccionar el TEMA [" & VBA.Left(Me.oLCanc(i).Caption, 2) & "]"
    Me.oLCanc(i).MousePointer = 99
    Me.oLCanc(i).BorderStyle = 1
    'VBA.DoEvents
Next i
'---------------------Se limpian las posiciones que no se usan---------------------
iPos_Vac = pNoReg + 1
Call Limpia_Can(iPos_Vac)
igMax_RgC = pNoReg
Call Refresh_Paginero(igAct_PgC, igTot_PgC)
Me.olMensajeSis.Caption = ""
End Function

Private Sub Limpia_Can(Optional ByVal piPos As Integer = 1)
Dim i As Integer
For i = piPos To Me.oLCanc.Count
    Me.oLCanc(i).Caption = ""
    Me.oLCanc(i).MousePointer = 0
    Me.oLCanc(i).ToolTipText = ""
    Me.oLCanc(i).BorderStyle = 0
    Me.oImgVideo(i).Visible = False
    'VBA.DoEvents
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
VBA.DoEvents
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
VBA.DoEvents
End Sub


Private Sub oTimer_Moneda_Timer()
If igKeep_Cred = 1 Then
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
VBA.DoEvents
End Sub

Private Sub oTimer_Reset_Timer()
Call Muestra_Tema_Det
Me.oTime_Mensajes.Enabled = False
Me.olMessage.Visible = False
Me.olMensaje_Video.Visible = False
VBA.DoEvents
End Sub

Private Sub Conectar_DBTem(Optional iClose As Integer = 0)
Dim sConnectionString As String
Dim sSql As String
sSql = "SELECT File05.* FROM File05 "
With Me.oDC_Temas
    sConnectionString = "Provider=VFPOLEDB.1;Data Source=" & sgDir_Tmp & ";Password='';Collating Sequence=MACHINE"
    .ConnectionString = sConnectionString
'   .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
'   .CommandType = adCmdTable
'   .RecordSource = "File05"
    .CommandType = adCmdText
    .RecordSource = sSql
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    If iClose = 1 Then
        .Recordset.Close
    End If
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

Private Sub Conectar_DBPro(Optional iClose As Integer = 0)
Dim sConnectionString As String
Dim sSql As String
sSql = "SELECT " & _
        "file01.ID_GEN, file01.ID_ORD AS ID_ORD1, " & _
        "file02.ID_DIS ,file02.ID_ORD AS ID_ORD2, " & _
        "file02.FL_IMG, " & _
        "file03.* FROM file01,file02,file03 " & _
        "WHERE  file01.ID_GEN=file02.ID_GEN " & _
        "AND    file03.ID_DIS=file02.ID_DIS " & _
        "HAVING file03.FL_PRC = 1 "
With Me.oDC_Promos
    sConnectionString = "Provider=VFPOLEDB.1;Data Source=" & sgDir_Tmp & ";Password='';Collating Sequence=MACHINE"
    .ConnectionString = sConnectionString
'    .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
    '.CommandType = adCmdTable
    '.RecordSource = "File05"
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .CommandType = adCmdText
    .RecordSource = sSql
    .Refresh
    If .Recordset.RecordCount > 0 Then
        .Recordset.MoveFirst
    End If
    If iClose = 1 Then
        .Recordset.Close
        .Refresh
    End If
End With
End Sub

Private Sub Conectar_DBPub()
Dim i As Integer
Dim iTot As Integer
'Dim agArr_Pub1() As String
Dim arr_Pub2() As String
Me.oLst_Pub1.Path = sgDir_Pub1
Me.oLst_Pub1.Refresh
Me.oLst_Pub2.Path = sgDir_Pub2
Me.oLst_Pub2.Refresh
If bgSw_Pub = False Then
    igTot_Pub = Me.oLst_Pub1.ListCount
Else
    igTot_Pub = Me.oLst_Pub2.ListCount
End If
ReDim agArr_Pub1(igTot_Pub)
ReDim arr_Pub2(igTot_Pub)
If bgSw_Pub = False Then
    For i = 0 To igTot_Pub - 1
        agArr_Pub1(i) = Me.oLst_Pub1.List(i)
        arr_Pub2(i) = Me.oLst_Pub1.List(i)
    Next i
Else
    For i = 0 To igTot_Pub - 1
        agArr_Pub1(i) = Me.oLst_Pub2.List(i)
        arr_Pub2(i) = Me.oLst_Pub2.List(i)
    Next i
End If
If igTot_Pub > 1 Then
    Call Shuffle_Array(agArr_Pub1, arr_Pub2)
End If
Me.oLst_Pub1.Path = ""
Me.oLst_Pub1.Refresh
Me.oLst_Pub1.Enabled = False
Me.oLst_Pub2.Path = ""
Me.oLst_Pub2.Refresh
Me.oLst_Pub2.Enabled = False
ReDim arr_Pub2(0)
End Sub

Private Sub oTimer_Srv_Timer()
If Me.olMessage.Tag = "1" Then
    Me.oTime_Mensajes.Tag = ""
    'gbServ_Mode = False
'aqui
End If
Me.oTimer_Srv.Enabled = False
End Sub

Private Sub oTM_Box_Timer()
If bfTm = True Then
    If bgIs_Video = False Then
        Me.oLDuracion.Caption = Me.MediaPlayer1.currentMedia.durationString
        Me.olAct_Pos.Caption = Me.MediaPlayer1.Controls.currentPositionString
    Else
        If igScr_Alone = 1 Then
            Me.oLDuracion.Caption = Me.MediaPlayer2.currentMedia.durationString
            Me.olAct_Pos.Caption = Me.MediaPlayer2.Controls.currentPositionString
        Else
            Me.oLDuracion.Caption = Video_Form.MediaPlayer3.currentMedia.durationString
            Me.olAct_Pos.Caption = Video_Form.MediaPlayer3.Controls.currentPositionString
        End If
    End If
Else
    Me.oLDuracion.Caption = "00:00"
    Me.olAct_Pos.Caption = "00:00"
    Me.oTM_Box.Enabled = False
End If
VBA.DoEvents
End Sub

Private Sub oTM_codigo2_Timer()
Call Retrocede
If VBA.Mid(VBA.Trim(Me.otCodigo.Text), 1, 2) = "99" Then
    Call Retrocede
    Call Retrocede
End If
Me.oTM_codigo2.Enabled = False
igNext_Return_Gen = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Return_Gen)
End Sub

Private Sub oTM_Mouse_Timer()
Call SetCursorPosition(Me, 0, 0)
Call ShowCursor(False)
Me.oTM_Mouse.Enabled = False
End Sub

Private Sub oTM_ScreenSaver_Timer()
Dim bActive As Boolean
SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, bActive, False
If bActive Then
    Main_Form.olMessage.Visible = True
    Main_Form.olMessage.Caption = "SCREEN SAVER ACTIVO..."
    Main_Form.oTime_Mensajes.Enabled = True
    Beep
Else
    Me.oTM_ScreenSaver.Enabled = False
End If
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
Case Else
'VBA.SendKeys "{BACKSPACE}"
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

Private Sub Cargar_Gen()
Call Limpia_Gen
Call Limpia_Dis
Call Limpia_Can
Call Desactiva_Cancion(True)
Call Desactiva_Disco(True)
Call Desactiva_Genero(False)
Call Conectar_DBGen
Call Cargar_INF_Gen
Call Cargar_Pag_Gen(1, igMax_RgG)
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
If igInd_Kar = 0 Then
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DefaultDir", sgDir_Fls)
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DBQ", sgDir_Fls)
Else
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DefaultDir", sgDir_Fls2)
    Call Write_Ini_File(sgDir_odb & "\Link_Tab.dsn", "ODBC", "DBQ", sgDir_Fls2)
End If
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

'------------------------------------------------schema.ini---------------------------------------
If Not FileExist(sgDir_Fls & "\schema.ini") Then
'------------------------------------------------schema.ini---------------------------------------
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "ColNameHeader", "False")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Format", "CSVDelimited")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "MaxScanRows", "50")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "CharacterSet", "OEM")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col1", "ID_GEN Char Width 2")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col2", "ID_ORD Char Width 2")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col3", "DESCRI Char Width 20")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file01.tab", "Col4", "FL_KAG Integer")
'   If Not FileExist(sgDir_Fls & "\file01.tab") Then
'        Open sgDir_Fls & "\file01.tab" For Output As #1
'        Write #1, "001", "01", "NECESITA CARGAR LA INFORMACIÓN..."
'        Close #1
'   End If
'------------------------------------------------schema.ini---------------------------------------
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "ColNameHeader", "False")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Format", "CSVDelimited")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "MaxScanRows", "50")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "CharacterSet", "OEM")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col1", "ID_GEN Char Width 2")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col2", "ID_DIS Integer")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col3", "ID_ORD Char Width 2")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col4", "NOM_DIS Char Width 40")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col5", "NOM_ART Char Width 40")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col6", "FL_IMG Char Width 80")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col7", "TX_COM Char Width 40")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col8", "C_VIDEO Integer")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file02.tab", "Col9", "FL_PRD Integer")
'   If Not FileExist(sgDir_Fls & "\file02.tab") Then
'       Open sgDir_Fls & "\file02.tab" For Output As #1
'       Write #1, "001", "00001", "01", "NO HAY DISCOS!!!", "NO ARTISTA!!!", "", ""
'       Close #1
'   End If
'-----------------------------------------------[file03.tab]---------------------------------------
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "ColNameHeader", "False")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Format", "CSVDelimited")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "MaxScanRows", "50")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "CharacterSet", "OEM")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col1", "ID_GEN Char Width 2")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col2", "ID_DIS Integer")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col3", "ID_CAN Integer")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col4", "ID_ORD Char Width 3")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col5", "DE_CAN Char Width 50")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col6", "FL_MP3 Char Width 80")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col7", "FL_PRC Integer")
'   Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file03.tab", "Col8", "FL_KAC Integer")
'   If Not FileExist(sgDir_Fls & "\file03.tab") Then
'       Open sgDir_Fls & "\file03.tab" For Output As #1
'       Write #1, "001", "00001", "1", "01", "NO HAY CANSIÓN!!!", ""
'       Close #1
'   End If
    
    '---------------------------------------[file05.tab]--------------------------------------
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "ColNameHeader", "False")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Format", "CSVDelimited")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "MaxScanRows", "50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "CharacterSet", "OEM")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col1", "ID_CAN Integer")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col2", "ID_COD Char Width 65")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col3", "DE_CAN Char Width 50")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col4", "FL_MP3 Char Width 80")
    Call Write_Ini_File(sgDir_Fls & "\schema.ini", "file05.tab", "Col5", "FL_MP3 Char Width 80")
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
    sFld5 = VBA.Trim(ogVFP9.xFields(1, "ID_PRO"))
    sFld6 = VBA.Trim(ogVFP9.xFields(1, "FL_DIS"))
    sCadena = sFld1 & "," & sFld2 & "," & sFld3 & "," & sFld4 & "," & sFld5 & "," & sFld6
    oLst_A_Tocar.AddItem sCadena
    Call ogVFP9.xNext(1)
    'VBA.DoEvents
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
        Call ogVFP9.Add_Data_1(sFlds(0), sFlds(1), sFlds(2), sFlds(3), sFlds(4), sFlds(5))
    Next iNum_reg
End If
Call ogVFP9.Save_Data(1, sgDir_Fls)
Call ogVFP9.Close_Table(1)
Call Write_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "IDX_PUBLIC", VBA.Format(igInd_Pub, "#####0"))
Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", Scramble(VBA.Format(igCnt_CRG, "#####0")))
End Sub

Private Sub Colocar_Frames()
With Me.oFrame_Gen
    .Height = 485
    .Left = 304
    .Top = 64
    .Width = 489
End With
With Me.oFrame_Dis
    .Height = 485
    .Left = 304
    .Top = 64
    .Width = 489
End With
With Me.oFrame_Can
    .Height = 485
    .Left = 304
    .Top = 64
    .Width = 489
End With
With Me.TBack4
    .Height = 200
    .Left = 304
    .Top = 72
    .Width = 200
End With
End Sub

Private Function Set_Open_Dbf() As Boolean
Set_Open_Dbf = True
Err.Clear
On Error GoTo mierror
Set ogVFP9 = CreateObject("library.VFP_txt_Utils")

Exit Function

mierror:
MsgBox "Error al invocar la Librería  [LIBRARY.DLL], rutina VFP_txt_Utils:.."
Set_Open_Dbf = False
'End
Exit Function

End Function

Private Sub Set_Tmp_DBF()
Dim sPath As String
Err.Clear
On Error GoTo mierror

sPath = IIf(igInd_Kar = 0, sgDir_Fls, sgDir_Fls2)
Call ogVFP9.check_integ_01(sPath)
Call ogVFP9.check_integ_02(sPath)
Call ogVFP9.check_integ_03(sPath)
Call ogVFP9.check_integ_05(sPath)
Call ogVFP9.SINC_PRO(sPath)
Call ogVFP9.Set_Files_Close
Call ogVFP9.SINC_VID(sPath)
Call ogVFP9.Set_Files_Close
'If iFlag = 1 Then
'    Call ogVFP9.Set_Files_Tmp(sPath, sgDir_Tmp, sgDir_Img, sgDir_Mp3, sgDir_Pub1, App.Path & "\FOXTOOLS.FLL")
'    Call ogVFP9.Set_Files_Close
'End If
Call ogVFP9.Set_Files_Verif(sPath, sgDir_Img, sgDir_Mp3, App.Path & "\FOXTOOLS.FLL")
Call ogVFP9.Set_Files_Close
Exit Sub

mierror:
MsgBox "Error al invocar la las tablas desde [" & IIf(igInd_Kar = 0, sgDir_Fls, sgDir_Fls2) & "]:.."
End
Exit Sub

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
    'VBA.DoEvents
Loop
oList2.Clear
For iIndex = 0 To oList1.ListCount - 1
    oList2.AddItem (oList1.List(iIndex))
    'VBA.DoEvents
Next iIndex
oList1.Clear
For k = 1 To oColRandom.Count
    D(k) = oColRandom.Item(k)
    oList1.AddItem oList2.List(D(k) - 1)
    'VBA.DoEvents
Next k
oList2.Clear
End Function

Private Function Inlist(sEntrada As String, Optional sPar1 As String = "", Optional sPar2 As String = "", Optional sPar3 As String = "") As Boolean
Dim sCadenas As String
sCadenas = VBA.Trim(sPar1) & VBA.Trim(sPar2) & VBA.Trim(sPar3)
Inlist = IIf(InStr(1, sCadenas, sEntrada, vbTextCompare) > 0, True, False)
End Function

Private Sub DBCan_Cheker(pVal As Boolean)
Dim sErr As String
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
    olInfo_Cheker.Caption = "Recuperando información del cancionero..."
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.value = 0
Else
    olInfo_Cheker.Caption = "Recuperando información del cancionero..."
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = 100
    Me.ProgressBar1.value = 0
    Exit Sub
End If
otNot_Found_List.Text = ""
otNot_Found_List.Refresh
List1.ListIndex = List1.ListCount - 1
Dim iNumReg As Integer
Dim iTotReg As Integer
Dim i__Proc As Integer
Dim sCadena As String
Dim sMp3Fle As String
Dim sTmp As String
Dim sValue As String
Me.List1.Clear
iTot_Cnt = 0: iCop_Cnt = 0: iNCop_Cnt = 0: iTotReg = 0
If Me.oChk_FndP.value = 1 Then
    Call Regarga_Foto
End If
If Me.oChk_FndC.value = 1 Then
    Call Recarga_Canc
End If
olInfo_Cheker.Caption = "Salvando LOG  de eventos..."
Me.otNot_Found_List.Refresh
Me.otNot_Found_List.SaveFile (App.Path & "\MP3_Not_Found.RTF")
olInfo_Cheker.Caption = "Listo..."
End Sub

Private Sub Recarga_Canc()
    Dim sErr As String
    Dim sRes As Integer
    Dim sSql1 As String
    Dim sSql2 As String
    Dim iErr_Fnd As Integer
    Dim iErr_Cnt As Integer
    Dim iCop_Cnt As Integer
    Dim iNCop_Cnt As Integer
    Dim iTot_Cnt As Integer
    Dim sValue As String
    iTot_Cnt = 0: iCop_Cnt = 0: iNCop_Cnt = 0: iTotReg = 0
    sSql1 = "SELECT * FROM File03 ORDER BY ID_GEN,ID_DIS,ID_CAN,ID_ORD"
    With Me.oDC_CANC
        .ConnectionString = "FILE NAME=" & sgDir_odb & "\Link_Dbf.dsn"
        .CommandType = adCmdText
        .RecordSource = sSql1
        .Refresh
    End With
    With Me.oDC_CANC.Recordset
        If .RecordCount <= 0 Then
            Call MsgBox("No se encontraron temas en el cancionero actual...", vbCritical, "Error")
            iTotReg = 0
            Exit Sub
        Else
            iTotReg = .RecordCount
        End If
        olInfo_Cheker.Caption = VBA.Str(iTotReg) & " Registros encontrados [CANCIONES]"
        Me.List1.AddItem VBA.Str(iTotReg) & " Registros [CANCIONES] por verificar..."
        Call Add_Log
        Me.List1.ListIndex = List1.ListCount - 1
        .MoveFirst
        iNumReg = 0
        olInfo_Cheker.Caption = "Procesando.."
        Me.ProgressBar1.Min = 0
        Me.ProgressBar1.Max = 100
        Do While Not .EOF
            If bgExit = True Then
                olInfo_Cheker.Caption = "Recuperando información del cancionero..."
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
            i__Proc = VBA.FormatNumber((iNumReg / iTotReg) * 100, 2, vbTrue)
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
                    If CopyFileWindowsWay(sValue, sgDir_Mp3, sErr) = True Then
                        Me.List1.AddItem "Copiando: " & sValue & " ->OK..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iCop_Cnt = iCop_Cnt + 1
                        Me.ProgressBar2.Visible = False
                    Else
                        Me.List1.AddItem "Copiando " & sValue & " ->FAILED..."
                        Call Add_Log
                        Me.List1.AddItem sErr
                        Call Add_Log
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
            VBA.DoEvents
        Loop
        Me.List1.AddItem "----------------------------------------------------------------"
        Call Add_Log
        Me.List1.AddItem "Archivos Faltantes  : " + VBA.Trim(VBA.Str(iTot_Cnt))
        Call Add_Log
        Me.List1.AddItem "Archivos Copiados   : " + VBA.Trim(VBA.Str(iCop_Cnt))
        Call Add_Log
        Me.List1.AddItem "Archivos No Copiados: " + VBA.Trim(VBA.Str(iNCop_Cnt))
        Call Add_Log
        Me.List1.AddItem "----------------------------------------------------------------"
        Call Add_Log
        Me.List1.AddItem " "
        Call Add_Log
        Me.List1.AddItem " "
        Call Add_Log
        Me.List1.ListIndex = List1.ListCount - 1
    End With
Exit Sub

Solve_error:
iErr_Cnt = iErr_Cnt + 1
iErr_Fnd = 1
Resume Next
End Sub
Sub Regarga_Foto()
    Dim sErr As String
    Dim sRes As Integer
    Dim sSql1 As String
    Dim sSql2 As String
    Dim iErr_Fnd As Integer
    Dim iErr_Cnt As Integer
    Dim iCop_Cnt As Integer
    Dim iNCop_Cnt As Integer
    Dim iTot_Cnt As Integer
    Dim sValue As String
    Dim sImgFle As String
    Dim sExt_Img As String
    iTot_Cnt = 0: iCop_Cnt = 0: iNCop_Cnt = 0: iTotReg = 0
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
        Call Add_Log
        Me.List1.ListIndex = List1.ListCount - 1
        .MoveFirst
        iNumReg = 0
        olInfo_Cheker.Caption = "Procesando.."
        Me.ProgressBar1.Min = 0
        Me.ProgressBar1.Max = 100
        Do While Not .EOF
            If bgExit = True Then
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
            i__Proc = VBA.FormatNumber((iNumReg / iTotReg) * 100, 2, vbTrue)
            olInfo_cheker_Proc.Caption = Str(i__Proc) & "%"
            Me.ProgressBar1.value = i__Proc
        
            olInfo_cheker_Proc.Refresh
            Me.otRuteExternal2.Text = VBA.Trim(Me.otRuteExternal2.Text)
            Me.otRuteExternal2.Text = Me.otRuteExternal2.Text & ""
            If FileExist(sImgFle) = False And sImgFle <> "" Then
                iTot_Cnt = iTot_Cnt + 1
                If Me.Check1.value = 1 Then
                    sValue = Me.otRuteExternal2 & sExt_Img
                    If CopyFileWindowsWay(sValue, sgDir_Img, sErr) = True Then
                    'If CopyFast(sValue, sImgFle, Me.ProgressBar2, sErr) = True Then
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal2 & sExt_Img & " ->OK..."
                        Me.List1.ListIndex = List1.ListCount - 1
                        iCop_Cnt = iCop_Cnt + 1
                        Me.ProgressBar2.Visible = False
                    Else
                        Me.List1.AddItem "Copiando " & Me.otRuteExternal2 & sExt_Img & " ->FAILED..."
                        Call Add_Log
                        Me.List1.AddItem sErr
                        Call Add_Log
                        Me.List1.ListIndex = List1.ListCount - 1
                        iNCop_Cnt = iNCop_Cnt + 1
                    End If
                End If
            End If
        .MoveNext
        VBA.DoEvents
    Loop
    Me.List1.AddItem "----------------------------------------------------------------"
    Call Add_Log
    Me.List1.AddItem "Archivos Faltantes  : " + VBA.Trim(VBA.Str(iTot_Cnt))
    Call Add_Log
    Me.List1.AddItem "Archivos Copiados   : " + VBA.Trim(VBA.Str(iCop_Cnt))
    Call Add_Log
    Me.List1.AddItem "Archivos No Copiados: " + VBA.Trim(VBA.Str(iNCop_Cnt))
    Call Add_Log
    Me.List1.AddItem "----------------------------------------------------------------"
    Call Add_Log
    Me.List1.AddItem ""
    Call Add_Log
    Me.List1.ListIndex = List1.ListCount - 1
    End With
Exit Sub

Solve_error:
iErr_Cnt = iErr_Cnt + 1
iErr_Fnd = 1
Resume Next
End Sub

Private Sub Add_Log()
Me.otNot_Found_List.Text = Me.otNot_Found_List.Text & Me.List1.List(List1.ListCount - 1) & Chr(10)
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
For i = 1 To 6
    If Me.olVideo(i).Visible = True Then
        If (Me.olVideo(i).ForeColor) = &H80FF80 Then
            Me.olVideo(i).BackColor = &HFF&
            Me.olVideo(i).ForeColor = &HFFFF&
        Else
            Me.olVideo(i).BackColor = &H0&
            Me.olVideo(i).ForeColor = &H80FF80
        End If
    End If
    VBA.DoEvents
Next i
End Sub

Private Sub Go_Service()
Me.otCodigo.Text = ""
Call Load(Svr_Form)
With Svr_Form
    .Text1(1).Text = sgDir_odb
    .Text1(2).Text = sgDir_Tmp
    .Text1(3).Text = sgDir_Fls
    .Text1(4).Text = sgDir_Img
    .Text1(5).Text = sgDir_Mp3
    .Text1(6).Text = sgDir_Pub1
    .Text1(7).Text = sgFle_Fon
    .Text1(8).Text = sgDir_Pub2
    .Text1(9).Text = sgDir_Fls2
'   ----------------------------------
    .ctNEdit2(1).value = igLim_Cred
    .ctNEdit2(2).value = igCnt_CR
    .ctNEdit2(3).value = sgKb_BonC
    .ctNEdit2(4).value = sgKb_VID
    .Check2(1).value = igFlg_SavedCR
    .Check2(2).value = igKeep_Cred
    .Check2(3).value = igNoDuplicT
'   ----------------------------------
    .ctNEdit3(1).value = igDelay_Return_Gen
    .ctNEdit3(2).value = igDelay_Return_Dis
    .ctNEdit3(3).value = igDelay_Bonus_Vid
    .ctNEdit3(4).value = sgIdx_Prm
    .Check3(1).value = VBA.IIf(bgVideoLabel = True, 1, 0)
    .Check3(2).value = VBA.IIf(bgDiscLabel = True, 1, 0)
    .Check3(3).value = igScr_Alone
    .Check3(4).value = VBA.IIf(bgKeep_On_Top = True, 1, 0)
    .Check3(5).value = igMixe_Popu
    .Check3(6).value = IIf(bgSw_Pub = False, 0, 1)
    .otAccess3.Text = sgCr_AKey
    .Check3(7).value = igInd_Kar
    .Check3(8).value = igShowPass
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
    .ctMEdit4(12).Text = sgKb_SwP
    .ctMEdit4(13).Text = sgKb_Pause
    .ctMEdit4(14).Text = sgKb_SwK
    .Show vbModal
End With
If Not sgFle_Fon = "" Then
    If FileExist(sgFle_Fon) Then
        Me.TBack1.TransparentBackground = False
        Me.TBack2.TransparentBackground = False
        Me.TBack3.TransparentBackground = False
        Me.TBack4.TransparentBackground = False
        Me.oFrame_Dis.TransparentBackground = False
        Me.oFrame_Gen.TransparentBackground = False
        Me.oFrame_Can.TransparentBackground = False
        Me.Picture = LoadPicture(sgFle_Fon)
        Me.oFrame_Dis.TransparentBackground = True
        Me.oFrame_Gen.TransparentBackground = True
        Me.oFrame_Can.TransparentBackground = True
        Me.TBack1.TransparentBackground = True
        Me.TBack2.TransparentBackground = True
        Me.TBack3.TransparentBackground = True
        Me.TBack4.TransparentBackground = True
        'aqui
    End If
End If
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
'if VBA.AppActivate (App.Title)

End Sub

Private Sub Timer2_Timer()
If igLen > 0 Then
    If Val(Me.olTimer1(1).Caption) > 0 Then
        Me.olTimer1(1).Caption = Val(olTimer1(1).Caption) - 1
    Else
        Me.olTimer1(1).Caption = 0
        Me.olTimer1(1).Visible = False
    End If
    If Val(Me.olTimer1(2).Caption) > 0 Then
        Me.olTimer1(2).Caption = Val(olTimer1(2).Caption) - 1
    Else
        Me.olTimer1(2).Caption = 0
        Me.olTimer1(2).Visible = False
    End If
End If
VBA.DoEvents
End Sub

Public Function Shuffle_Array(iArray() As String, iTempArray() As String)
Dim iCtr As Integer
Dim iCtr2 As Integer
Dim iTemp As Integer
Dim iMaxElement As Integer
Dim iMinElement As Integer

Dim bArray() As Boolean
iMaxElement = UBound(iArray)
iMinElement = LBound(iArray)

ReDim bArray(iMinElement To iMaxElement) As Boolean

For iCtr = iMinElement To iMaxElement
    Do
    Randomize Timer
    iTemp = Int((iMaxElement - iMinElement + 1) * Rnd) + _
        iMinElement
 
        If bArray(iTemp) = False Then
            iArray(iTemp) = iTempArray(iCtr)
            bArray(iTemp) = True
            Exit Do
        End If
        'VBA.DoEvents
    Loop
Next
End Function

Function Es_Video(sValue As String) As Boolean
Dim bValue As Boolean
bValue = False
Select Case sValue
Case Is = "MPEG"
    bValue = True
Case Is = "MPEG4"
    bValue = True
Case Is = "MPG"
    bValue = True
Case Is = "MPG4"
    bValue = True
End Select
Es_Video = bValue
End Function

Function DirExists(ByVal sDirName As String) As Boolean
   Dim sDir As String
   On Error Resume Next
   DirExists = False
   sDir = Dir$(sDirName, vbDirectory)
   If (Len(sDir) > 0) And (Err = 0) Then
      DirExists = True
   End If
End Function

Public Function GetWindowsDir() As String
    Const MAX_PATH2 = 255

    Dim sRet As String, lngRet As Long
    sRet = String$(MAX_PATH2, 0)
    lngRet = GetWindowsDirectory(sRet, MAX_PATH2)
    GetWindowsDir = Left(sRet, lngRet)
End Function

Public Sub ControlPanels(filename As String)
Dim rtn As Double
On Error Resume Next
rtn = Shell(filename, 5)
End Sub

Public Sub Call_Tsr()
Call VBA.Shell(App.Path & "\tsr.exe", vbNormalFocus)
End Sub

Public Sub Add_CR1()
If igKeep_Cred = 1 Then
    Exit Sub
End If
If igFlg_TCR = 1 Then
    igCnt_CRP = igCnt_CRP + 1
    Me.olMetros2.Visible = True
    Me.olMetros2.Caption = PADL(igCnt_CRP, 6, "0")
    Me.olTest.Visible = True
    Exit Sub
Else
    Me.olMetros2.Visible = False
    Me.olTest.Visible = False
    igCnt_CRP = 0
End If
If igCnt_CR >= igLim_Cred Then
    Main_Form.olMessage.Visible = True
    Main_Form.olMessage.Caption = "REVISAR MONEDERO!"
    Main_Form.oTime_Mensajes.Enabled = True
Else
    igCnt_CR = igCnt_CR + 2
    igCnt_CRG = igCnt_CRG + 1
    If igCnt_CRG > 999999 Then
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAN", Scramble(VBA.Format(igCnt_CRG, "#####0")))
        igCnt_CRG = 0
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", Scramble(VBA.Format(igCnt_CRG, "#####0")))
        igCnt_CRG = VBA.Val(UnScramble(Read_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", "0")))
    Else
        Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", Scramble(VBA.Format(igCnt_CRG, "#####0")))
    End If
    If igFlg_SavedCR = 1 Then
       Call Write_Ini_File(App.Path & "\PathV2.ini", "CREDITS", "ACU_SAVECR", VBA.Format(igCnt_CR, "#####0"))
    End If
End If
End Sub

Public Sub Add_CR2()
    'Call Refresh_Creditos(Me)
    If sgKb_BonC > 0 Then
        If igCnt_CR = 8 Then
            Main_Form.olMessage.Visible = True
            Main_Form.olMessage.Caption = "CRED. PROMOSIÓN [" & VBA.Trim(VBA.Str(sgKb_BonC)) & "]"
            Main_Form.oTime_Mensajes.Enabled = True
            Call Sleep(3)
            igCnt_CR = igCnt_CR + sgKb_BonC
        End If
    End If
    Call Refresh_Creditos(Me)
    If Me.otCodigo.EditMask <> "##-##-##" Then
        Me.otCodigo.EditMask = "##-##-##"
        Me.otCodigo.Tag = ""
        Me.oSetFocus_Codigo.SetFocus
    End If
    Exit Sub
End Sub

Private Sub Timer3_Timer()
Dim Horas As Integer, Minutos As Integer, Segundos As Integer, Cadena As String
    
Segundos = pTimeResto
Horas = Int(Segundos / 3600)
Segundos = Segundos Mod 3600
Minutos = Int(Segundos / 60)
Segundos = Segundos Mod 60
Cadena = Format$(Horas, "00") & ":" & Format$(Minutos, "00") & ":" & Format$(Segundos, "00") & " Restante"
Me.olT_Mant.Caption = "Opción de servicio activada" & Chr(13) & "[" & Cadena & "] Restante " & Chr(13) & "99-00-06 DESACTIVAR"
pTimeResto = pTimeResto - 1
If pTimeResto = -1 Then
    gbServ_Mode = False
    Me.Timer3.Enabled = False
    Me.oService_Info.filename = ""
    Me.olT_Mant.Visible = False
    Me.oService_Info.Visible = False
   
    Call Retrocede
    Call Retrocede
End If
End Sub

Public Sub Set_DBF_To_Tmp()
Dim oFs As Object
Dim sDir_Org As String
sDir_Org = IIf(igInd_Kar = 0, sgDir_Fls, sgDir_Fls2)
Set oFs = CreateObject("Scripting.FileSystemObject")
Call oFs.CopyFile(sDir_Org & "\FILE01.DBF", sgDir_Tmp & "\FILE01.DBF", True)
Call oFs.CopyFile(sDir_Org & "\FILE02.DBF", sgDir_Tmp & "\FILE02.DBF", True)
Call oFs.CopyFile(sDir_Org & "\FILE03.DBF", sgDir_Tmp & "\FILE03.DBF", True)
If oFs.FileExists(sgDir_Tmp & "\FILE05.DBF") = False Then
    Call oFs.CopyFile(sDir_Org & "\FILE05.DBF", sgDir_Tmp & "\FILE05.DBF", True)
End If
End Sub
'***KEY LOCK
Private Function CheckForKL() As Boolean
    Dim RotateCount1 As Integer, RotateCount2 As Integer
    Dim Argument1 As Integer, Argument2 As Integer, Argument3 As Integer
    Dim error As Variant
    'On Error Resume Next
    On Error GoTo 0
    error = KTASK(KLCHECK, ValidateCode1, ValidateCode2, ValidateCode3)
    If error = KEY_ERROR_NOKEYLOK Or error = KEY_ERROR_NOERROR Or error = KEY_ERROR_NOKEYLOK_ALSO Then
        RotateCount1 = ReturnValue2 And 7
        RotateCount2 = ReturnValue1 And 15
        Argument3 = ReturnValue1 Xor ReturnValue2
        Call RotateLeft(ReturnValue1, RotateCount1)
        Argument1 = ReturnValue1 Xor READCODE3 Xor ReturnValue2
        Call RotateLeft(ReturnValue2, RotateCount2)
        Argument2 = ReturnValue2
        Call KTASK(Argument1, Argument2, Argument3, 0)
        If ((ReturnValue1 <> ClientIDCode1) Or (ReturnValue2 <> ClientIDCode2)) Then
            MsgBox "No USB-KEY device attached.", vbCritical
            CheckForKL = False
        Else
            CheckForKL = True
        End If
    End If
End Function

Private Function KTASK(ByVal Arg1 As Integer, ByVal Arg2 As Integer, ByVal Arg3 As Integer, ByVal Arg4 As Integer) As Variant
  Const STARTANTIDEBUGGER As Long = 0
  
  Dim ReturnValue1Long As Long
  Dim KeybdRet As Integer, KfuncRet As Long
  Dim LgArg1 As Long, LgArg2 As Long, LgArg3 As Long, LgArg4 As Long
  
  LgArg1 = Arg1
  If LgArg1 < 0 Then LgArg1 = LgArg1 + 2 ^ 16
  LgArg2 = Arg2
  If LgArg2 < 0 Then LgArg2 = LgArg2 + 2 ^ 16
  LgArg3 = Arg3
  If LgArg3 < 0 Then LgArg3 = LgArg3 + 2 ^ 16
  LgArg4 = Arg4
  If LgArg4 < 0 Then LgArg4 = LgArg4 + 2 ^ 16

  KfuncRet = KFUNC(LgArg1, LgArg2, LgArg3, LgArg4)
  KTASK = ShowLastKeyError(Err.LastDLLError)
  ReturnValue1Long = KfuncRet Mod 65536
  If ReturnValue1Long > 32767 Then ReturnValue1Long = ReturnValue1Long - 65536
  If ReturnValue1Long < -32768 Then ReturnValue1Long = ReturnValue1Long + 65536
  ReturnValue1 = ReturnValue1Long
  ReturnValue2 = Int(KfuncRet / 65536)

End Function

Private Sub RotateLeft(ByRef Argument As Integer, ByVal Count As Integer)
  Dim i As Integer
  Dim HighBit As Long
  Dim LocalTarget As Integer, LocalTargetLong As Long
  
  
  LocalTargetLong = Argument
  For i = 1 To Count
  If LocalTargetLong < 0 Then LocalTargetLong = LocalTargetLong - &HFFFF0000
    HighBit = LocalTargetLong And 32768
    If HighBit = 32768 Then HighBit = 1 Else HighBit = 0
    LocalTargetLong = LocalTargetLong * 2
    If LocalTargetLong < -32768 Then LocalTargetLong = LocalTargetLong + 2 ^ 16
    If LocalTargetLong > 32767 Then LocalTargetLong = LocalTargetLong - 2 ^ 16
  LocalTargetLong = LocalTargetLong + HighBit
  Next
  If LocalTargetLong > 32767 Then LocalTargetLong = LocalTargetLong - 65536
  LocalTarget = LocalTargetLong
  Argument = LocalTarget
End Sub

Private Function ShowLastKeyError(LastDLLError As Long) As Variant
    Dim error As Long
    Dim msgResponse As Long
    
    error = GETLASTKEYERROR()
  
    If error = KEY_ERROR_NOERROR Then
        error = LastDLLError
    End If
  
    If error = KEY_ERROR_NOKEYLOK Or _
       error = KEY_ERROR_NOERROR Or _
       error = KEY_ERROR_NOKEYLOK_ALSO Or _
       error = KEY_ERROR_NOLEASEDATE Or _
       error = KEY_ERROR_LEASEDATEBAD Or _
       error = KEY_ERROR_FSDATEBAD Then
        msgResponse = 0
    ElseIf error = KEY_ERROR_NO_SESSIONS Then
        msgResponse = MsgBox("The session limit has been reached.", , "Security Device Error")
    ElseIf error = KEY_ERROR_WRONGKEYLOK Then
        msgResponse = MsgBox("The Security Device failed to Authenticate.", , "Security Device Error")
    ElseIf error = KEY_ERROR_BADVERSION Then
        msgResponse = MsgBox("This Security Device found old version of Driver.", , "Security Device Error")
    ElseIf error = KEY_ERROR_BADFUNC Then
        msgResponse = MsgBox("This Security Device KFUNC Command was not recognized by the driver.", , "Security Device Error")
    ElseIf error = KEY_ERROR_NOREADAUTH Then
        msgResponse = MsgBox("The Security Device failed the Read Authorization call.", , "Security Device Error")
    ElseIf error = KEY_ERROR_NOWRITEAUTH Then
        msgResponse = MsgBox("The Security Device Memory Write has not been authorized.", , "Security Device Error")
    ElseIf error = KEY_ERROR_INVALIDADDRESS Then
        msgResponse = MsgBox("The Security Device Address is out of range.", , "Security Device Error")
    ElseIf error = KEY_ERROR_NOCOUNTSLEFT Then
        msgResponse = MsgBox("The counter was already fully counted down to zero.", , "Security Device Error")
    ElseIf error = KEY_ERROR_WRITETIMEOUT Then
        msgResponse = MsgBox("This Security Device failed the internal Write Timeout Test.", , "Security Device Error")
    ElseIf error = KEY_ERROR_NOLEASEDATE Then
        msgResponse = MsgBox("This Security Device failed the internal Write Timeout Test.", , "Security Device Error")
    Else
        msgResponse = MsgBox("Unrecognized Security Device error has occured - Error " & error, , "Security Device Error")
    End If
    
    ShowLastKeyError = error
End Function

Private Function ReadText() As String
   Dim ReadDataArray%(56)
   Dim sText As String
   Dim Contents As Integer
   Dim TestString As String
   Dim i As Integer
   Dim LongRV As Long
   
   For i = 0 To 56 - 1
        Call KTASK(GETVARWORD, i, 0, 0)
        If ReturnValue1 < 0 Then
            LongRV = ReturnValue1 + 2 ^ 16
        Else
            LongRV = ReturnValue1
        End If
        
        If LongRV >= 48 And LongRV <= 57 Or _
           LongRV >= 65 And LongRV <= 90 Or _
           LongRV >= 97 And LongRV <= 122 Or _
           LongRV = 32 Then
              sText = sText + Chr(LongRV)
        End If
            
   Next
   ReadText = sText
End Function
'***KEY LOCK***

Private Sub Transparencia()
Dim iIndex As Integer
For iIndex = 1 To igMax_Gen
    oLGenero(iIndex).Transparentia Main_Form.hDC, 150, True, True, oLGenero(iIndex).Left, oLGenero(iIndex).Top
Next iIndex
End Sub

Sub TileBackground()
Dim bgdImage    As StdPicture
Dim X           As Single
Dim Y           As Single

Set bgdImage = LoadPicture(App.Path & "\Fondos\Shakira-800_600.jpg")
Y = 0
While Y < Me.ScaleHeight
    X = 0
    While X < Me.ScaleWidth
        PaintPicture bgdImage, X, Y
       
        X = X + (bgdImage.Width / Screen.TwipsPerPixelX) / 1.8
    Wend
        Y = Y + (bgdImage.Height / Screen.TwipsPerPixelY) / 1.8
Wend
End Sub
