VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form EmpQtlyFTR 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   4215
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   930
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   660
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   390
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   375
      Left            =   4823
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6038
      TabIndex        =   5
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3638
      TabIndex        =   4
      Top             =   8160
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8160
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   720
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11033
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Form 941 for 2008 - Page 1"
      TabPicture(0)   =   "frmEmpQtlyFTR.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBNumber1"
      Tab(0).Control(1)=   "Text2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Form 941 for 2008 - Page 2"
      TabPicture(1)   =   "frmEmpQtlyFTR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Schedule B (Form 941)"
      TabPicture(2)   =   "frmEmpQtlyFTR.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label8"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   375
         Left            =   -68520
         TabIndex        =   7
         Top             =   4680
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Calculator      =   "frmEmpQtlyFTR.frx":0054
         Caption         =   "frmEmpQtlyFTR.frx":0074
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpQtlyFTR.frx":00DA
         Keys            =   "frmEmpQtlyFTR.frx":00F8
         Spin            =   "frmEmpQtlyFTR.frx":0142
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   30081025
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -66240
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -74400
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   $"frmEmpQtlyFTR.frx":016A
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   7815
      End
      Begin VB.Label Label7 
         Caption         =   "1   Number of employees who received wages, tips, or other components for the pay period"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   7815
      End
      Begin VB.Label Label6 
         Caption         =   "Part 1: Answer these questions for this quarter."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   5895
      End
   End
End
Attribute VB_Name = "EmpQtlyFTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

