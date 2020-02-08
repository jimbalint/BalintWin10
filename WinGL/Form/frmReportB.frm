VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmReportB 
   Caption         =   "Print GLAccount File"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
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
   ScaleHeight     =   5145
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin TDBNumber6Ctl.TDBNumber txtHiBranch 
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   3600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":0000
      Caption         =   "frmReportB.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":008C
      Keys            =   "frmReportB.frx":00AA
      Spin            =   "frmReportB.frx":00F4
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtLoBranch 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":011C
      Caption         =   "frmReportB.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":01A8
      Keys            =   "frmReportB.frx":01C6
      Spin            =   "frmReportB.frx":0210
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtHiMain 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":0238
      Caption         =   "frmReportB.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":02C4
      Keys            =   "frmReportB.frx":02E2
      Spin            =   "frmReportB.frx":032C
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtLoMain 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":0354
      Caption         =   "frmReportB.frx":0374
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":03E0
      Keys            =   "frmReportB.frx":03FE
      Spin            =   "frmReportB.frx":0448
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin VB.CheckBox chkAllBranches 
      Caption         =   "Print All &Branches"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox chkAllAccounts 
      Caption         =   "&Print All Accounts"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin TDBNumber6Ctl.TDBNumber txtHiAcct 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":0470
      Caption         =   "frmReportB.frx":0490
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":04FC
      Keys            =   "frmReportB.frx":051A
      Spin            =   "frmReportB.frx":0564
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtLoAcct 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calculator      =   "frmReportB.frx":058C
      Caption         =   "frmReportB.frx":05AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmReportB.frx":0618
      Keys            =   "frmReportB.frx":0636
      Spin            =   "frmReportB.frx":0680
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblHiBranch 
      Caption         =   "High Branch:"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblLoBranch 
      Caption         =   "Low Branch:"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblHiMain 
      Caption         =   "High Main Acct:"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblLoMain 
      Caption         =   "Low Main Acct:"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblHiAcct 
      Caption         =   "High Acct:"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblLoAcct 
      Caption         =   "Low Acct:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "frmReportB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAllAccounts_Click()
    If chkAllAccounts.Value = 1 Then
       AcctSet (False)
    Else
       AcctSet (True)
       txtLoAcct = "1"
       txtHiAcct = "999999999"
    End If
End Sub

Private Sub chkAllBranches_Click()
    If chkAllBranches.Value = 1 Then
       BranchSet (False)
    Else
       BranchSet (True)
       txtLoMain = "1"
       txtHiMain = "999999999"
       txtLoBranch = "1"
       txtHiBranch = "99"
    End If
End Sub

Private Sub cmdOk_Click()
    
    If chkAllAccounts.Value = 1 Then
       txtLoAcct = "0"
       txtHiAcct = "0"
    End If
    
    If chkAllBranches.Value = 1 Then
       txtLoMain = "0"
       txtHiMain = "0"
       txtLoBranch = "0"
       txtHiBranch = "0"
    End If
    
    Response = True
    Me.Hide
End Sub

Private Sub Form_Load()
    
    AcctSet (False)
    BranchSet (False)
    Response = False
    NumSet

End Sub

Private Sub AcctSet(ByVal TF As Boolean)

    lblLoAcct.Enabled = TF
    txtLoAcct.Enabled = TF
    
    lblHiAcct.Enabled = TF
    txtHiAcct.Enabled = TF

End Sub


Private Sub BranchSet(ByVal TF As Boolean)

    lblLoMain.Enabled = TF
    txtLoMain.Enabled = TF
    
    lblHiMain.Enabled = TF
    txtHiMain.Enabled = TF
    
    lblLoBranch.Enabled = TF
    txtLoBranch.Enabled = TF
    
    lblHiBranch.Enabled = TF
    txtHiBranch.Enabled = TF

End Sub


Private Sub NumSet()

    txtLoAcct.MinValue = 0
    txtLoAcct.MaxValue = 999999999
    txtLoAcct.EditMode = dbiOverwrite
    txtLoAcct.HighlightText = True
    txtLoAcct.Format = "########0"
    txtLoAcct.DisplayFormat = "########0;;;0"

    txtHiAcct.MinValue = 0
    txtHiAcct.MaxValue = 999999999
    txtHiAcct.EditMode = dbiOverwrite
    txtHiAcct.HighlightText = True
    txtHiAcct.Format = "########0"
    txtHiAcct.DisplayFormat = "########0;;;0"

    txtLoMain.MinValue = 0
    txtLoMain.MaxValue = 999999999
    txtLoMain.EditMode = dbiOverwrite
    txtLoMain.HighlightText = True
    txtLoMain.Format = "########0"
    txtLoMain.DisplayFormat = "########0;;;0"

    txtHiMain.MinValue = 0
    txtHiMain.MaxValue = 999999999
    txtHiMain.EditMode = dbiOverwrite
    txtHiMain.HighlightText = True
    txtHiMain.Format = "########0"
    txtHiMain.DisplayFormat = "########0;;;0"

    txtLoBranch.MinValue = 0
    txtLoBranch.MaxValue = 999999999
    txtLoBranch.EditMode = dbiOverwrite
    txtLoBranch.HighlightText = True
    txtLoBranch.Format = "########0"
    txtLoBranch.DisplayFormat = "########0;;;0"

    txtHiBranch.MinValue = 0
    txtHiBranch.MaxValue = 999999999
    txtHiBranch.EditMode = dbiOverwrite
    txtHiBranch.HighlightText = True
    txtHiBranch.Format = "########0"
    txtHiBranch.DisplayFormat = "########0;;;0"

End Sub
