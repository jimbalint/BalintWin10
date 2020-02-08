VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmSupplemental 
   Caption         =   "Supplemental"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   556
      Calendar        =   "frmSupplemental.frx":0000
      Caption         =   "frmSupplemental.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSupplemental.frx":0184
      Keys            =   "frmSupplemental.frx":01A2
      Spin            =   "frmSupplemental.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "08/22/2008"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39682
      CenturyMode     =   0
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   2760
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Text            =   "cmbYear"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox cmbQtr 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Text            =   "cmbQtr"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Quarter"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Title"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Employer's Report of Wages - Supplemental"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1343
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmSupplemental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

