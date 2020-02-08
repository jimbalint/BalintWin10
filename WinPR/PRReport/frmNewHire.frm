VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmNewHire 
   Caption         =   "New Hire Date Range"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
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
   ScaleHeight     =   4950
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbState 
      Height          =   390
      Left            =   2280
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2760
      Width           =   4095
   End
   Begin TDBDate6Ctl.TDBDate tdbdtStartDate 
      Height          =   375
      Left            =   2198
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   661
      Calendar        =   "frmNewHire.frx":0000
      Caption         =   "frmNewHire.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmNewHire.frx":0174
      Keys            =   "frmNewHire.frx":0192
      Spin            =   "frmNewHire.frx":01F0
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
      Text            =   "04/04/2009"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39907
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4898
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1778
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin TDBDate6Ctl.TDBDate tdbdtEndHireDate 
      Height          =   375
      Left            =   2198
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   661
      Calendar        =   "frmNewHire.frx":0218
      Caption         =   "frmNewHire.frx":0318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmNewHire.frx":0388
      Keys            =   "frmNewHire.frx":03A6
      Spin            =   "frmNewHire.frx":0404
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
      Text            =   "04/04/2009"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39907
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "State to report (or all): "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   638
      TabIndex        =   5
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmNewHire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.tdbdtStartDate = DateSerial(Year(Now()), Month(Now()), 1)
    Me.tdbdtEndHireDate = Now()
    
    With Me.cmbState
        .AddItem "All"
        ' Populate state dropdown box
        PRState.GetBySQL ("SELECT * FROM PRState order by PRState.StateAbbrev")
        Do
            Me.cmbState.AddItem PRState.StateAbbrev
            If Not PRState.GetNext Then
               Exit Do
            End If
        Loop
        .Text = "OH"
    End With
        
    
    Me.lblCompanyName.Caption = PRCompany.Name
    Me.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
    StartDate = Me.tdbdtStartDate
    EndDate = Me.tdbdtEndHireDate
    NewHireReport
End Sub

