VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmDEDate 
   Caption         =   "GL Data Entry"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin TDBNumber6Ctl.TDBNumber Period 
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2160
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      Calculator      =   "DEDate.frx":0000
      Caption         =   "DEDate.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DEDate.frx":0084
      Keys            =   "DEDate.frx":00A2
      Spin            =   "DEDate.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   12
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   1
      MaxValueVT      =   5242885
      MinValueVT      =   3014661
   End
   Begin TDBNumber6Ctl.TDBNumber FiscalYear 
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1200
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      Calculator      =   "DEDate.frx":0114
      Caption         =   "DEDate.frx":0134
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DEDate.frx":0198
      Keys            =   "DEDate.frx":01B6
      Spin            =   "DEDate.frx":0200
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   2005
      MinValue        =   2000
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   2004
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Period:"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fiscal Year:"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1020
      TabIndex        =   4
      Top             =   240
      Width           =   9375
   End
End
Attribute VB_Name = "frmDEDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    frmDEGrid.Show vbModal
    Unload Me
End Sub

Private Sub Form_Load()
    Me.lblCompName = MainMenu.lblCompanyName
End Sub
