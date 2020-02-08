VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmReportA 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin TDBNumber6Ctl.TDBNumber txtHiAcct 
      Height          =   375
      Left            =   6030
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmReportA.frx":0000
      Caption         =   "frmReportA.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportA.frx":008C
      Keys            =   "frmReportA.frx":00AA
      Spin            =   "frmReportA.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   24117249
      Value           =   0
      MaxValueVT      =   5636101
      MinValueVT      =   3342341
   End
   Begin TDBNumber6Ctl.TDBNumber txtLoAcct 
      Height          =   375
      Left            =   2310
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmReportA.frx":011C
      Caption         =   "frmReportA.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportA.frx":01A8
      Keys            =   "frmReportA.frx":01C6
      Spin            =   "frmReportA.frx":0210
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245185
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Frame fraRegCon 
      Height          =   615
      Left            =   1950
      TabIndex        =   10
      Top             =   2040
      Width           =   4335
      Begin VB.OptionButton optCon 
         Caption         =   "&Consolidated"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optReg 
         Caption         =   "&Regular"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4710
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2550
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.CheckBox chkAllAccts 
      Caption         =   "&Print All Accounts"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label lblHiAcct 
      Caption         =   "High Acct #::"
      Height          =   255
      Left            =   4350
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   750
      TabIndex        =   8
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblLoAcct 
      Caption         =   "Low Acct #:"
      Height          =   255
      Left            =   870
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmReportA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAllAccts_Click()
   If Me.chkAllAccts.Value = 0 Then
      Me.txtLoAcct.Enabled = True
      Me.txtHiAcct.Enabled = True
      Me.lblLoAcct.Enabled = True
      Me.lblHiAcct.Enabled = True
      Me.txtLoAcct = "1"
      Me.txtHiAcct = "999999999"
   Else
      Me.txtLoAcct.Enabled = False
      Me.txtHiAcct.Enabled = False
      Me.lblLoAcct.Enabled = False
      Me.lblHiAcct.Enabled = False
      Me.txtLoAcct = ""
      Me.txtHiAcct = ""
   End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    If chkAllAccts.Value = 1 Then
       txtLoAcct.Value = "0"
       txtHiAcct.Value = "0"
    End If
    
    Response = True

    Me.Hide

End Sub

Private Sub Form_Load()
    Response = False
    
    Me.chkAllAccts.Value = 1
    Me.txtLoAcct.Enabled = False
    Me.txtHiAcct.Enabled = False
    Me.lblLoAcct.Enabled = False
    Me.lblHiAcct.Enabled = False
    Me.txtLoAcct = ""
    Me.txtHiAcct = ""

End Sub
