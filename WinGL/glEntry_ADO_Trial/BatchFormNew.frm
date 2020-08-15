VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BatchFormNew 
   Caption         =   " BATCH RECORD"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BatchFormNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5855.564
   ScaleMode       =   0  'User
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraQBBasis 
      Caption         =   "  Basis  "
      Height          =   735
      Left            =   7680
      TabIndex        =   41
      Top             =   4680
      Width           =   2775
      Begin VB.OptionButton optQBCashBasis 
         Caption         =   "Cash"
         Height          =   375
         Left            =   1560
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optQBAccrualBasis 
         Caption         =   "Accrual"
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraAcctTranslate 
      Height          =   735
      Left            =   4080
      TabIndex        =   40
      Top             =   7800
      Width           =   6135
      Begin TDBNumber6Ctl.TDBNumber tdbAcctTranslateValue 
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   240
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   661
         Calculator      =   "BatchFormNew.frx":030A
         Caption         =   "BatchFormNew.frx":032A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BatchFormNew.frx":0392
         Keys            =   "BatchFormNew.frx":03B0
         Spin            =   "BatchFormNew.frx":03FA
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
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.OptionButton optMultiply 
         Caption         =   "Multiply"
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optDivide 
         Caption         =   "Divide"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkAcctTranslate 
      Caption         =   "Account Number Translation"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   8040
      Width           =   3015
   End
   Begin VB.ComboBox cmbCheckingAcct 
      Height          =   360
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   7320
      Width           =   5175
   End
   Begin VB.ComboBox cmbSuspAcct 
      Height          =   360
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5640
      Width           =   4455
   End
   Begin VB.Frame fraQBOption 
      Height          =   735
      Left            =   4560
      TabIndex        =   39
      Top             =   4680
      Width           =   2775
      Begin VB.OptionButton optQBSummary 
         Caption         =   "Summary"
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optQBDetail 
         Caption         =   "Detail"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraQBType 
      Caption         =   "  Select QB data to import  "
      Height          =   735
      Left            =   360
      TabIndex        =   38
      Top             =   4680
      Width           =   3975
      Begin VB.OptionButton optQBCheck 
         Caption         =   "Check Detail"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optQBGL 
         Caption         =   "General Ledger"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkQBChecking 
      Caption         =   "Use QB Checking Acct"
      Height          =   495
      Left            =   600
      TabIndex        =   22
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton cmdFileOpen 
      Height          =   375
      Left            =   7800
      Picture         =   "BatchFormNew.frx":0422
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtQBFile 
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   6720
      Width           =   8775
   End
   Begin VB.CheckBox chkUseCurrentQB 
      Caption         =   "&Use currently opened QuickBooks File"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   6240
      Width           =   4455
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calendar        =   "BatchFormNew.frx":072C
      Caption         =   "BatchFormNew.frx":0844
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchFormNew.frx":08B0
      Keys            =   "BatchFormNew.frx":08CE
      Spin            =   "BatchFormNew.frx":092C
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
      FirstMonth      =   1
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   2
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "12/14/2005"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   38700
      CenturyMode     =   0
   End
   Begin VB.CheckBox chkQB 
      Caption         =   "Acquire from &QuickBooks"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CheckBox chkBudget 
      Caption         =   "&Budget Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cmbJournal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   0
      Text            =   " Pick Journal Source"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      TabIndex        =   2
      Text            =   "cmbPeriod"
      Top             =   3240
      Width           =   3975
   End
   Begin VB.ComboBox cmbFiscalYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      TabIndex        =   1
      Text            =   "cmbFiscalYear"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   21
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   9000
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calendar        =   "BatchFormNew.frx":0954
      Caption         =   "BatchFormNew.frx":0A6C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchFormNew.frx":0AD8
      Keys            =   "BatchFormNew.frx":0AF6
      Spin            =   "BatchFormNew.frx":0B54
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
      FirstMonth      =   1
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   2
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "12/14/2005"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   38700
      CenturyMode     =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10920
      Y1              =   2345.891
      Y2              =   2345.891
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10920
      Y1              =   5351.564
      Y2              =   5351.564
   End
   Begin VB.Label lblSuspAcct 
      Caption         =   "Suspense Acct #:"
      Height          =   375
      Left            =   480
      TabIndex        =   37
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label lblChkAcct 
      Caption         =   "Checking Acct #:"
      Height          =   375
      Left            =   3240
      TabIndex        =   36
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label lblQBFile 
      Caption         =   "QuickBooks File Name"
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label lblEndDate 
      Alignment       =   1  'Right Justify
      Caption         =   "&End Date:"
      Height          =   375
      Left            =   7560
      TabIndex        =   34
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblStartDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   375
      Left            =   3720
      TabIndex        =   33
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label txtCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   360
      TabIndex        =   32
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label7 
      Caption         =   "JOURNAL SOURCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label txtCredits 
      Caption         =   "CREDITS"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label lblUpdated 
      Caption         =   "Update User and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblCreated 
      Caption         =   "Created User and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label txtDebits 
      Caption         =   "DEBITS"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Label txtRecord 
      Caption         =   "RECORDS IN BATCH"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "FISCAL PERIOD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "FISCAL YEAR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblBatchNumber 
      Caption         =   "BATCH NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "BatchFormNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PRGlobalID As Long
Dim SQLStr As String

Public BatchNumber As Long
Dim bat As New rBatch
Public userOK As Boolean
Dim jou As New XArrayDB
Dim xdbGLAccount As New XArrayDB

' Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2
Dim SessMgr As New QBSessionManager
' Dim SessMgr As New QBXMLRP2Lib.RequestProcessor2
Dim strTicket As String

Dim RepQ As QBFC13Lib.ICustomDetailReportQuery
Dim RepQ2 As QBFC13Lib.IGeneralDetailReportQuery

Dim RequestSet As QBFC13Lib.IMsgSetRequest
Dim ResponseSet As QBFC13Lib.IMsgSetResponse
Dim qResponse As QBFC13Lib.IResponse
Dim RepRet As QBFC13Lib.IReportRet
Dim orReportData As QBFC13Lib.IORReportData

' for chart of account query
Dim AccQ As QBFC13Lib.IAccountQuery
Dim RetList As IAccountRetList
Dim ItemRet As QBFC13Lib.IAccountRet

Dim xdbAccts As New XArrayDB

Dim nRequest As Long
Dim index As Long
Dim ct As Long

Dim i, j, k, l, m As Long
Dim x, y, z As String

Dim GLHAccount As Long
Dim GLHFiscalYear As Integer
Dim GLHPeriod As Byte
Dim GLHBatchNumber As Long
Dim GLHAmount As Currency
Dim GLHReference As String
Dim GLHDescription As String
Dim GLHSourceCode As Byte
Dim GLHJournalSource As Byte
Dim GLHHistType As String
Dim GLHUpdateFlag As Boolean

Dim RecCount As Long
Dim RefNum As String
Dim Desc As String
Dim TotalAmount As Currency

Dim TotalDebits As Currency
Dim TotalCredits As Currency
Dim RecordCount As Long

Dim IconType As Integer

Dim rsAcctSum As New ADODB.Recordset

Dim TextChannel As Integer
Dim Quote, Comma As String

Dim P1, RunBal As Currency
Dim TextLine As String


Private Sub chkAcctTranslate_Click()
    If Me.chkAcctTranslate Then
        Me.fraAcctTranslate.Visible = True
    Else
        Me.fraAcctTranslate.Visible = False
    End If
End Sub

Private Sub Form_Load()

'    Set jou = xFactory.GetJournals(FileName)
'    Set JournalList.Array = jou
'    JournalList.Columns(0).Width = 500
'    JournalList.Columns(1).Width = 3500
     
Dim xRow, AcctLen As Long
     
    PRGlobalID = 0
    Me.tdbAcctTranslateValue.Format = "#####0"
    Me.tdbAcctTranslateValue.DisplayFormat = "#####0"
    
    Response = False
    
    ' *** moved from init to here so default selection works ***
    Set xdbGLAccount = xFactory.GetAccounts(FileName, "0")
    With Me.cmbSuspAcct
        .AddItem ""
        .ItemData(.NewIndex) = 0
        ' loop thru the xArray
        For xRow = 0 To xdbGLAccount.UpperBound(1)
            AcctLen = Len(CStr(xdbGLAccount.Value(xRow, 0)))
            .AddItem Space(12 - AcctLen) & xdbGLAccount.Value(xRow, 0) & " " & _
                     xdbGLAccount.Value(xRow, 1)
            .ItemData(.NewIndex) = xdbGLAccount.Value(xRow, 0)
        Next xRow
        .ListIndex = 0
    End With
    With Me.cmbCheckingAcct
        .AddItem ""
        .ItemData(.NewIndex) = 0
        ' loop thru the xArray
        For xRow = 0 To xdbGLAccount.UpperBound(1)
            .AddItem Format(xdbGLAccount.Value(xRow, 0), "#########0") & " " & _
                     xdbGLAccount.Value(xRow, 1)
            .ItemData(.NewIndex) = xdbGLAccount.Value(xRow, 0)
        Next xRow
        .ListIndex = 0
    End With
    
'    With Me.tdbcmbSuspAcct
'        .Array = xdbGLAccount
'        .ScrollBars = dblVertical
'        .ColumnHeaders = False
'        .Caption = "Department Name/#"
'        .AutoCompletion = True
'        .AutoDropDown = True
'        .LimitToList = True
'        .Columns(0).Width = 1200
'        .Columns(1).Width = 3000
'        .AlternatingRowStyle = True
'        .EvenRowStyle.BackColor = &H8000000F
'    End With
'    ' set to default?
'    If Not IsNull(com.SuspAcct) Then
'        If com.SuspAcct <> 0 Then
'            j = xdbGLAccount.Find(0, 0, com.SuspAcct, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
'            If j >= 0 Then
'                Me.tdbcmbSuspAcct.SelectedItem = j
'            End If
'        End If
'    End If
'
'    ' use for the checking acct field also
'    With Me.tdbcmbCheckingAcct
'        .Array = xdbGLAccount
'        .ScrollBars = dblVertical
'        .ColumnHeaders = False
'        .Caption = "Department Name/#"
'        .AutoCompletion = True
'        .AutoDropDown = True
'        .LimitToList = True
'        .Columns(0).Width = 1200
'        .Columns(1).Width = 3000
'        .AlternatingRowStyle = True
'        .EvenRowStyle.BackColor = &H8000000F
'    End With
    
    ' hide the QuickBook fields by default
    QBShow False

End Sub

Private Sub chkQB_Click()
    If Me.chkQB.Value = 1 Then
        QBShow True
    Else
        QBShow False
    End If
End Sub

Private Sub chkQBChecking_Click()
    If Me.chkQBChecking.Value = 1 Then
       Me.lblChkAcct.Enabled = False
       Me.cmbCheckingAcct.Enabled = False
    Else
       Me.lblChkAcct.Enabled = True
       Me.cmbCheckingAcct.Enabled = True
    End If
End Sub

Private Sub chkUseCurrentQB_Click()
    
    If Me.chkUseCurrentQB.Value = 1 Then
       Me.lblQBFile.Enabled = False
       Me.cmdFileOpen.Enabled = False
       Me.txtQBFile.Enabled = False
    Else
       Me.lblQBFile.Enabled = True
       Me.cmdFileOpen.Enabled = True
       Me.txtQBFile.Enabled = True
    End If
    
End Sub

Private Sub cmbFiscalYear_Click()
    
Dim i As Integer
Dim v As Variant
Dim fy As Integer

    Me.cmbPeriod.Clear
    fy = CInt(cmbFiscalYear)
      
    If com.FirstPeriod = 1 Then
       v = DateSerial(fy, com.FirstPeriod, 1)
    Else
       v = DateSerial(fy - 1, com.FirstPeriod, 1)
    End If

    cmbPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    
    For i = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
    Next i
    
    cmbPeriod.ListIndex = 0
    
    
'    cmbPeriod.Clear
'    Dim ndx, fy As Integer
'
'    fy = CInt(cmbFiscalYear)
'    For ndx = 1 To com.NumberPds
'        cmbPeriod.AddItem com.MonthName(ndx, fy)
'    Next ndx
'    cmbPeriod.ListIndex = bat.period - 1

End Sub

Private Sub CmdExit_Click()
    Response = False
    Me.Hide
End Sub

Private Sub cmdFileOpen_Click()
    Me.CommonDialog1.DialogTitle = "QB File to open"
    Me.CommonDialog1.Filter = "QB Data Files (*.qbw)|*.qbw"
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
       Me.txtQBFile = Me.CommonDialog1.FileName
    End If
End Sub

Private Sub cmdOK_Click()
    
'    On Error GoTo glErr
    
    If Me.optQBCheck Then
        If Me.chkQBChecking = 0 And Me.cmbCheckingAcct.ListIndex = 0 Then
            MsgBox "Must pick checking account!!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
            Me.cmbCheckingAcct.SetFocus
            Exit Sub
        End If
    End If
    
    If cmbJournal.ListIndex = -1 Then
       MsgBox "Must pick Journal Source !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       cmbJournal.SetFocus
       Exit Sub
    End If
    
    If Me.cmbPeriod.ListIndex = -1 Then
       MsgBox "Must pick period !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       cmbPeriod.SetFocus
       Exit Sub
    End If
    
    If Me.cmbFiscalYear.ListIndex = -1 Then
       MsgBox "Must pick Fiscal Year !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       Me.cmbFiscalYear.SetFocus
       Exit Sub
    End If
    
    If Me.chkAcctTranslate And IsNull(Me.tdbAcctTranslateValue.Value) Then
        MsgBox "You must pick an account translate value!", vbExclamation + vbOKOnly
        Me.tdbAcctTranslateValue.SetFocus
        Exit Sub
    End If
        
    ' save QB account translate info?
    If Me.chkQB And Me.chkAcctTranslate Then
        
        If PRGlobalID <> 0 Then
            If Me.optDivide = True Then
                PRGlobal.Byte1 = 1
            Else
                PRGlobal.Byte1 = 2
            End If
            PRGlobal.Var1 = CLng(Me.tdbAcctTranslateValue.Value)
            PRGlobal.Save (False)       ' put
        Else
            PRGlobal.Clear
            PRGlobal.Description = "AcctTranslate"
            PRGlobal.UserID = CompanyID
            If Me.optDivide = True Then
                PRGlobal.Byte1 = 1
            Else
                PRGlobal.Byte1 = 2
            End If
            PRGlobal.Var1 = CLng(Me.tdbAcctTranslateValue.Value)
            PRGlobal.Save (True)        ' add
        End If
    
    End If
    
    userOK = True
    bat.fiscalYear = CLng(cmbFiscalYear)
    bat.period = cmbPeriod.ListIndex + 1
    bat.JournalSource = jou.Value(cmbJournal.ListIndex + 1, 0)
    
    ' add 100 to the journal source number if budget entry
    If chkBudget Then bat.JournalSource = bat.JournalSource + 100
    
'    bat.debits = CCur(txtDebits)
'    bat.credits = CCur(txtCredits)
'    bat.nRecords = CLng(txtRecords)
    
    bat.Updated = Now
    bat.updateUser = curUser
    bat.PutRecord bat.BatchNumber, FileName
    Response = True
    
    Me.Hide
    
    If Me.chkQB Then QBInit
    
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub Init()

Dim ndx As Long
Dim CurFY As Integer
    
    userOK = False
    bat.GetBatch BatchNumber, FileName
    txtCompanyName = com.Name
    lblBatchNumber = "Batch # " & bat.BatchNumber
    lblCreated = "Created by " & UserName(bat.createUser) & " on " & ShowDate(bat.Created)
    lblUpdated = "Record is OPEN (Not Updated)"
    txtRecord = "RECORD COUNT = " & CStr(bat.nRecords)
    txtDebits = "DEBITS = " & Format(bat.debits, "#.00")
    txtCredits = "CREDITS = " & Format(bat.credits, "#.00")
    
    
'    For ndx = com.FirstFiscalYear To Year(Now) + 1
'        cmbFiscalYear.AddItem ndx
'    Next ndx
'
'    'if bat.fiscalYear=0 then
'    cmbFiscalYear = bat.fiscalYear
    
    CurFY = Int(com.LastClose / 10 ^ 4)
    If Int(com.LastClose / 100) Mod 100 <> 1 Then CurFY = CurFY + 1
    If CurFY < 1990 Or CurFY > 2020 Then CurFY = Year(Now())
    
    For ndx = CurFY + 1 To CurFY - 5 Step -1
        cmbFiscalYear.AddItem ndx
    Next ndx
    cmbFiscalYear.ListIndex = 1
    
'    For ndx = 1 To com.NumberPds
'        cmbPeriod.AddItem com.MonthName(ndx, bat.fiscalYear)
'        cmbPeriod.AddItem com.MonthName(ndx, CurFY)
'    Next ndx
''    cmbPeriod.ListIndex = bat.period - 1
'    cmbPeriod.ListIndex = 0
    
    Set jou = xFactory.GetJournals(FileName)
    
    For ndx = 1 To jou.UpperBound(1)
        cmbJournal.AddItem (CStr(jou.Value(ndx, 0)) & "-" & jou.Value(ndx, 1))
        If jou.Value(ndx, 0) = bat.JournalSource Then
            cmbJournal.ListIndex = ndx - 1
        End If
    Next ndx

    Response = False

End Sub

'Private Sub cmdPrint_Click()
'    ReviewReport.BatchNumber = BatchNumber
'    ReviewReport.Show vbModal
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        ' Case vbKeyF6: OnPrint
    End Select
End Sub

Private Sub OnPrint()
    ' ReviewReport.BatchNumber = BatchNumber
    ' ReviewReport.Show vbModal
End Sub

Private Sub QBShow(ByVal ShowIt As Boolean)

Dim Mo As Byte
Dim yr As Integer
    
    If Not ShowIt Then         ' hide them
           
        Me.lblChkAcct.Visible = False
        Me.lblEndDate.Visible = False
        Me.lblQBFile.Visible = False
        Me.lblStartDate.Visible = False
        Me.lblSuspAcct.Visible = False
       
        Me.txtQBFile.Visible = False
       
        Me.cmbSuspAcct.Visible = False
        Me.cmbCheckingAcct.Visible = False
        
        Me.TDBDate1.Visible = False
        Me.TDBDate2.Visible = False
       
        Me.chkUseCurrentQB.Visible = False
        Me.cmdFileOpen.Visible = False
    
        Me.fraQBType.Visible = False
        Me.fraQBOption.Visible = False
        Me.fraQBBasis.Visible = False
    
        Me.chkQBChecking.Enabled = False
    
        Me.chkAcctTranslate.Visible = False
        Me.fraAcctTranslate.Visible = False
    
    Else
               
        Me.lblChkAcct.Visible = True
        Me.lblEndDate.Visible = True
        Me.lblStartDate.Visible = True
        Me.lblSuspAcct.Visible = True
        Me.lblQBFile.Visible = True
       
        Me.TDBDate1.Visible = True
        Me.TDBDate2.Visible = True
        
        Me.TDBDate1.Visible = True
        Me.TDBDate2.Visible = True
       
        Me.chkUseCurrentQB.Visible = True
        Me.chkUseCurrentQB.Value = 1
       
        Me.cmdFileOpen.Visible = True
        Me.txtQBFile.Visible = True
       
        If Me.cmbFiscalYear.ListIndex <> -1 And Me.cmbPeriod.ListIndex <> -1 Then
              
            If com.FirstPeriod = 1 Then        ' Jan is the first period
                Mo = Me.cmbPeriod.ListIndex + 1
                Me.TDBDate1 = DateSerial(Me.cmbFiscalYear, Mo, 1)
                Me.TDBDate2 = LastDay(Me.cmbFiscalYear, Mo)
            Else
                Mo = Me.cmbPeriod.ListIndex + com.FirstPeriod    ' the list index is zero based
                If Mo <= 12 Then
                    yr = Me.cmbFiscalYear - 1
                Else
                    yr = Me.cmbFiscalYear
                    Mo = Me.cmbPeriod.ListIndex - 12 + com.FirstPeriod
                End If
                Me.TDBDate1 = DateSerial(yr, Mo, 1)
                Me.TDBDate2 = LastDay(yr, Mo)
            End If
       
        End If
        
        Me.fraQBType.Visible = True
        Me.fraQBOption.Visible = True
        Me.fraQBBasis.Visible = True
        Me.optQBGL = True
        Me.optQBDetail = True
        Me.optQBAccrualBasis = True
        
        Me.cmbCheckingAcct.Visible = True
        Me.chkQBChecking.Enabled = True
        Me.cmbSuspAcct.Visible = True
    
'        Me.TDBDate1 = DateSerial(2009, 6, 1)
'        Me.TDBDate2 = DateSerial(2009, 6, 30)
    
        Me.chkAcctTranslate.Visible = True
        Me.fraAcctTranslate.Visible = True
    
        SQLStr = "SELECT * FROM PRGlobal WHERE Description = 'AcctTranslate' " & _
                 "AND UserID = " & CompanyID
                    
        If PRGlobal.GetBySQL(SQLStr) = True Then
            PRGlobalID = PRGlobal.GlobalID
            If PRGlobal.Byte1 = 1 Then
                Me.optDivide = True
                Me.optMultiply = False
            Else
                Me.optDivide = False
                Me.optMultiply = True
            End If
            Me.tdbAcctTranslateValue.Value = CLng(PRGlobal.Var1)
            Me.chkAcctTranslate = 1
            Me.fraAcctTranslate.Visible = True
        End If
        
        If Me.chkAcctTranslate = 0 Then
            Me.fraAcctTranslate.Visible = False
            If Me.optDivide = False And Me.optMultiply = False Then
                Me.optDivide = True
            End If
        End If
    
    End If
        
End Sub

Private Function LastDay(ByVal yr As Integer, ByVal Mo As Byte) As Variant

    ' add one month
    If Mo = 12 Then
       yr = yr + 1
       Mo = 1
    Else
       Mo = Mo + 1
    End If
    
    ' subtract one day
    LastDay = DateAdd("d", -1, DateSerial(yr, Mo, 1))

End Function

Private Sub QBInit()

    Me.Hide
    MainMenu.Hide

    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg2 = "Now loading QuickBooks Data ..."
    frmProgress.Show
    
    ' connect to the data base with ADO
    ' x = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
    ' open the company database
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
    Else
        x = BalintFolder & "\Data\" & mdbName(com.FileName)
    End If
    CNOpen x, Password

    ' open a record set to the GLHistory file
    rsInit "SELECT * FROM GLHistory", Cn, rs
    
    frmProgress.Caption = "Opening QB Session"
    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg2 = "Now opening QuickBooks Session .... "
    frmProgress.Show
    
    SessMgr.OpenConnection2 "", "Windows GL Entry", ctLocalQBD
    
'    Dim connPref As QBXMLRP2Lib.QBXMLRPConnectionType
'    connPref = localQBD
'    SessMgr.OpenConnection2 "", "Windows GL Entry", connPref
'MsgBox "a"

    frmProgress.Caption = "Begin QB Session"
    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg2 = "Now Beginning QuickBooks Session .... "
    frmProgress.Show
    
    SessMgr.BeginSession Me.txtQBFile, omDontCare
    
'    Dim openMode As QBXMLRP2Lib.QBFileMode
'    openMode = qbFileOpenDoNotCare
'    strTicket = SessMgr.BeginSession("", openMode)
    
    frmProgress.Caption = "Get QB Chart of Accounts"
    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg2 = "Now Getting QB Chart of Accounts"
    frmProgress.Show

    QBGetAccounts
    
    frmProgress.Caption = "Gather QB Detail Data"
    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg2 = "Now GatheringQB Detail Data"
    frmProgress.Show
    
    If Me.optQBCheck Then
        QBLoadCheckData
    Else
        QBLoadGLData
    End If
    
    ' close the QB connection
    SessMgr.CloseConnection
    Set SessMgr = Nothing
    
    ' hide the progress screen
    frmProgress.Hide
    
    MainMenu.Show
    Me.Show
    
End Sub
Private Sub QBLoadGLData()

    ' ***** create dump text file
'    If BalintFolder = "" Then
'        x = Mid(App.Path, 1, 2) & "\Balint\Data\qbgl.txt"
'    Else
'        x = BalintFolder & "\Data\qbgl.txt"
'    End If
'    TextChannel = FreeFile
'
'    Open x For Output As #TextChannel
    
    Quote = """"
    Comma = ","
    
    ' temp record set for summ by account option
    rsAcctSum.CursorLocation = adUseClient
    rsAcctSum.Fields.Append "Account", adDouble
    rsAcctSum.Fields.Append "Amount", adCurrency
    rsAcctSum.Open , , adOpenDynamic, adLockOptimistic

    frmProgress.lblMsg2 = "Now loading QuickBooks GL Detail .... "
    frmProgress.Show
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    Set RepQ2 = RequestSet.AppendGeneralDetailReportQueryRq
    RepQ2.IncludeAccounts.SetValue (iaAll)
    RepQ2.GeneralDetailReportType.SetValue (gdrtGeneralLedger)
    
    If Me.optQBCashBasis = True Then
        RepQ2.ReportBasis.SetValue (rbCash)
    End If
    
    ' xxxxx repq2.ReportDetailLevelFilter.SetValue ???????
    ' ///RepQ2.ReportAccountFilter.ORReportAccountFilter.AccountTypeFilter.SetValue (atfBank)
    
    RepQ2.ORReportPeriod.ReportPeriod.FromReportDate.SetValue (Format(Me.TDBDate1, "mm / dd / yyyy"))
    RepQ2.ORReportPeriod.ReportPeriod.ToReportDate.SetValue (Format(Me.TDBDate2, "mm / dd / yyyy"))
    
    ' MsgBox " perform the request "
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
        
    ' MsgBox " interpret the response"
    Set qResponse = ResponseSet.ResponseList.GetAt(nRequest)
    
    ' check for errors
    If qResponse.StatusCode <> 0 Then
                                       
        If qResponse.StatusCode <= 499 Then
            IconType = vbInformation
        ElseIf qResponse.StatusCode <= 999 Then
            IconType = vbExclamation
        Else
            IconType = vbCritical
        End If
       
        MsgBox qResponse.StatusMessage & vbCrLf & _
              "Status Code: " & qResponse.StatusCode, IconType
              
        If qResponse.StatusCode >= 1000 Then  ' exit completely
            SessMgr.EndSession
            SessMgr.CloseConnection
            End
        End If
    
    End If
            
    Set RepRet = qResponse.Detail
        
    If RepRet Is Nothing Then
        MsgBox "Nothing in Report", vbExclamation, "QB Import"
        SessMgr.EndSession
        SessMgr.CloseConnection
        Exit Sub
    End If
    
    If (Not RepRet.ReportData Is Nothing) Then
        
        If Not (RepRet.ColDescList Is Nothing) Then
            For index = 0 To RepRet.ColDescList.Count - 1
                Dim colDesc As IColDesc
                Set colDesc = RepRet.ColDescList.GetAt(index)
                Dim index2 As Integer
                For index2 = 0 To colDesc.ColTitleList.Count - 1
                    Dim colTitle8 As IColTitle
                    Set colTitle8 = colDesc.ColTitleList.GetAt(index2)
                    Dim titlerow9  As Long
                    titlerow9 = colTitle8.titleRow.GetValue
                    If (Not colTitle8.Value Is Nothing) Then
                        Dim Value10 As String
                        Value10 = colTitle8.Value.GetValue
                    
'                        TextLine = Quote & index & Quote & Comma & Quote
'                        TextLine = TextLine & Value10 & Quote
'                        Print #TextChannel, TextLine
                        
                        ' MsgBox Value10 & vbCr & index
                        
                    End If
                Next index2
            Next index
        End If
        
        If (Not RepRet.ReportData.ORReportDataList Is Nothing) Then
            ct = RepRet.ReportData.ORReportDataList.Count
            For index = 0 To ct - 1
                Set orReportData = RepRet.ReportData.ORReportDataList.GetAt(index)
                If (Not orReportData Is Nothing) Then
                    QBProcessGLLine orReportData
                End If
            Next index
        End If
    End If
    
    ' update the batch record
    bat.debits = TotalDebits
    bat.credits = TotalCredits
    bat.nRecords = RecordCount
    bat.PutRecord bat.BatchNumber, FileName
    
    ' update to GLH if summary option was chosen
    If Me.optQBSummary And rsAcctSum.RecordCount > 0 Then
        rsAcctSum.MoveFirst
        Do
            GLHAccount = rsAcctSum!Account
            GLHReference = Me.TDBDate1 & " " & Me.TDBDate2
            GLHDescription = "QB Update"
            GLHAmount = rsAcctSum!Amount
            QBAddGLH
            rsAcctSum.MoveNext
        Loop Until rsAcctSum.EOF
    End If

    ' close the ADO record set and connection
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing

    rsAcctSum.Close
    Set rsAcctSum = Nothing

'    Close #TextChannel

End Sub

Private Sub QBProcessGLLine(orReportData As QBFC13Lib.IORReportData)
    
Dim colDataList As QBFC13Lib.IColDataList
Dim colData As QBFC13Lib.IColData
Dim colType As QBFC13Lib.IColDataList

Dim ndxEmp As Integer
Dim SSN As String
Dim dTxn As String
Dim enm As String
Dim CurIncome As Currency
Dim CurWageBase As Currency
Dim cc As Currency

Dim lngID As Long
Dim AcctNum As Long
Dim DrCr As String
    
Dim ii, jj As Long
    
    ' init the GLH variables
    GLHFiscalYear = Me.cmbFiscalYear
    GLHPeriod = bat.period
    GLHBatchNumber = bat.BatchNumber
    GLHSourceCode = 0
    GLHJournalSource = bat.JournalSource
    GLHHistType = "A"
    GLHUpdateFlag = True
    GLHReference = ""
    GLHDescription = ""
    GLHAmount = 0
    
    TextLine = ""
    DrCr = ""
    
    AcctNum = 0
    
    If (orReportData Is Nothing) Then Exit Sub
    
    Select Case orReportData.ortype
        
        Case QBFC13Lib.ENORReportData.orrdDataRow
            
            If (orReportData.DataRow Is Nothing) Then Exit Sub
            If (orReportData.DataRow.rowNumber Is Nothing) Then Exit Sub
            
'            If (orReportData.DataRow.RowData Is Nothing) Then Exit Sub
'            If (orReportData.DataRow.colDataList Is Nothing) Then Exit Sub
            
            Set colDataList = orReportData.DataRow.colDataList
            
            ' skip checking offset
            Set colData = colDataList.GetAt(0)
            If colData.Value Is Nothing Then Exit Sub
'            If colData.Value.GetValue() = "Checking" Then Exit Sub
                
            For i = 0 To colDataList.Count - 1
                Set colData = colDataList.GetAt(i)
                If (Not colData.Value Is Nothing) Then
                       
                    j = colData.colID.GetValue
                    y = Trim(colData.Value.GetValue)
                           
                    ' *************************************************
                    ' translation for Richlak
                    ' xxxxx If j = 8 Then j = 1     ' account number xxxxx

                    If j = 9 Then       ' DEBIT amount
                        j = 8
                        DrCr = "Dr"
                    End If

                    If j = 10 Then      ' CREDIT amount
                        j = 8
                        DrCr = "Cr"
                    End If

                    If j = 11 Then j = 9    ' balance - last field
                    ' *************************************************
                    
'                    ' jb
'                    TextLine = Quote & i & Quote & Comma & Quote
'                    TextLine = TextLine & j & Quote & Comma & Quote
'                    TextLine = TextLine & y & Quote & Comma & Quote
'                    TextLine = TextLine & x & Quote
'                    Print #TextChannel, TextLine
                   
                    ' 11/29/10 - addl message tracking
                    frmProgress.lblMsg3 = "QB Status: " & i & "/" & j & "/" & y
                    frmProgress.Refresh
                   
                    If j = 1 Then           ' start of a new account
                        
                        ' look for first space in string
                        x = ""
                        ii = 1
                        Do
                            If Mid(y, ii, 1) = " " Then Exit Do
                            x = Trim(x) & Mid(y, ii, 1)
                            ii = ii + 1
                            If ii > Len(Trim(y)) Then Exit Do
                            
                        Loop
                                                                        
                        If IsNumeric(x) Then        ' QB uses accounts numbers
                            k = xdbAccts.Find(1, 4, x, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                        Else                        ' find by unique QB account name - acct# in description
                            k = xdbAccts.Find(1, 0, x, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                        End If
                        
                        If k = -1 Then     ' not found in QB COA ???
                            If Me.cmbSuspAcct.ListIndex = 0 Then
                                GLHAccount = 0
                            Else
                                GLHAccount = Me.cmbSuspAcct.ItemData(Me.cmbSuspAcct.ListIndex)
                            End If
                        Else
                            
                            l = xdbAccts.Value(k, 4)
                            
                            ' acct number translation if selected
                            If Me.chkAcctTranslate And Me.tdbAcctTranslateValue.Value > 0 Then
                                If Me.optDivide = True Then
                                    l = l / Me.tdbAcctTranslateValue.Value
                                Else
                                    l = l * Me.tdbAcctTranslateValue.Value
                                End If
                            End If
                            
                            ' is it in the WinGL chart of accts?
                            m = xdbGLAccount.Find(0, 0, l, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
                            
                            If m >= 0 Then
                                
                            
                                GLHAccount = l
                            
                            Else
                                
                                If Me.cmbSuspAcct.ListIndex = 0 Then
                                    GLHAccount = 0
                                Else
                                    GLHAccount = Me.cmbSuspAcct.ItemData(Me.cmbSuspAcct.ListIndex)
                                End If
                            End If
                        End If
                    
                    ElseIf j = 2 Then       ' type
                    ElseIf j = 3 Then       ' date
                    ElseIf j = 4 Then       ' check/inv num
                        GLHReference = y
                    ElseIf j = 5 Then       ' name
                        GLHDescription = y
                    ElseIf j = 6 Then       ' memo
                        ' use it if the name is not assigned
                        If GLHDescription = "" Then
                            GLHDescription = y
                        End If
                    ElseIf j = 7 Then       ' split
                    ElseIf j = 8 Then       ' amount
                        If IsNumeric(y) Then
                            GLHAmount = CCur(y)
                        End If
                    ElseIf j = 9 Then       ' balance
                        If GLHAmount <> 0 Then
                            
'                            ' richlak - compare the balance to determing dr or cr
'                            If IsNumeric(y) Then
'                                P1 = CCur(y)
'                                If P1 > RunBal Then     '  this balace greater than prior - debit
'                                    GLHAmount = Abs(GLHAmount)
'                                Else
'                                    GLHAmount = -(Abs(GLHAmount))       ' credit
'                                End If
'                            End If
                            
                            ' Richlak translation
                            If DrCr = "Cr" Then
                                GLHAmount = -(Abs(GLHAmount))
                            End If
                            
                            If Me.optQBDetail Then      ' add each entry
                                QBAddGLH
                            Else                        ' sum per acct
                                rsAcctSum.Find "Account = " & GLHAccount, 0, adSearchForward, 1
                                If rsAcctSum.EOF Then
                                    rsAcctSum.AddNew
                                    rsAcctSum!Account = CLng(x)
                                    rsAcctSum!Amount = 0
                                End If
                                rsAcctSum!Amount = rsAcctSum!Amount + GLHAmount
                                rsAcctSum.Update
                            End If
                        End If
                        
                        ' store the run bal for next time
                        If IsNumeric(y) Then
                            RunBal = CCur(y)
                        End If
                        
                        GLHAmount = 0
                    End If
                
                End If
            Next i
            
    End Select
        
End Sub
Private Sub QBLoadCheckData()
    
    frmProgress.lblMsg2 = "Now loading QuickBooks Check Detail .... "
    frmProgress.Refresh
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 4, 0)
    Set RepQ2 = RequestSet.AppendGeneralDetailReportQueryRq
    RepQ2.IncludeAccounts.SetValue (iaAll)
    RepQ2.GeneralDetailReportType.SetValue (gdrtCheckDetail)
    RepQ2.ReportAccountFilter.ORReportAccountFilter.AccountTypeFilter.SetValue (atfBank)
    
    If Me.optQBCashBasis = True Then
        RepQ2.ReportBasis.SetValue (rbCash)
    End If
    
    RepQ2.ORReportPeriod.ReportPeriod.FromReportDate.SetValue (Format(Me.TDBDate1, "mm / dd / yyyy"))
    RepQ2.ORReportPeriod.ReportPeriod.ToReportDate.SetValue (Format(Me.TDBDate2, "mm / dd / yyyy"))
    
    ' MsgBox " perform the request "
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
        
    ' MsgBox " interpret the response"
    Set qResponse = ResponseSet.ResponseList.GetAt(nRequest)
    
    ' check for errors
    If qResponse.StatusCode <> 0 Then
               
        If qResponse.StatusCode <= 499 Then
            IconType = vbInformation
        ElseIf qResponse.StatusCode <= 999 Then
            IconType = vbExclamation
        Else
            IconType = vbCritical
        End If
       
        MsgBox qResponse.StatusMessage & vbCrLf & _
              "Status Code: " & qResponse.StatusCode, IconType
              
        If qResponse.StatusCode >= 1000 Then  ' exit completely
            SessMgr.EndSession
            SessMgr.CloseConnection
            End
        End If
    
    End If
            
    Set RepRet = qResponse.Detail
        
    If RepRet Is Nothing Then
        MsgBox ("Nothing in Report")
        SessMgr.EndSession
        SessMgr.CloseConnection
        Exit Sub
    End If
    
    If (Not RepRet.ReportData Is Nothing) Then
        If (Not RepRet.ReportData.ORReportDataList Is Nothing) Then
            ct = RepRet.ReportData.ORReportDataList.Count
            For index = 0 To ct - 1
                Set orReportData = RepRet.ReportData.ORReportDataList.GetAt(index)
                If (Not orReportData Is Nothing) Then
                    QBProcessCheckLine orReportData
                End If
            Next index
        End If
    End If
    
    ' create the offset entry
    If Me.chkQBChecking Then                ' use the bank accounts from QB
        
        For i = 1 To xdbAccts.UpperBound(1)
            If xdbAccts(i, 3) = "Bank" Then
               GLHAccount = xdbAccts(i, 4)
               
                ' acct number translation if selected
                If Me.chkAcctTranslate And Me.tdbAcctTranslateValue.Value > 0 Then
                    If Me.optDivide = True Then
                        GLHAccount = GLHAccount / Me.tdbAcctTranslateValue.Value
                    Else
                        GLHAccount = GLHAccount * Me.tdbAcctTranslateValue.Value
                    End If
                End If
               
               GLHAmount = xdbAccts(i, 2)
               GLHReference = xdbAccts(i, 0)
               GLHDescription = "QB " & Format(Now(), "mm/dd/yyyy")
               GLHUpdateFlag = True
               QBAddGLH
            End If
        Next i
    
    Else            ' total amount to the account specified
        
        With Me.cmbCheckingAcct
            If .ListIndex = 0 Then
                GLHAccount = 0      ' ???
            Else
                GLHAccount = .ItemData(.ListIndex)
            End If
        End With
        GLHAmount = TotalAmount * (-1)
        GLHReference = "Checking"
        GLHDescription = "QB " & Format(Now(), "mm/dd/yyyy")
        GLHUpdateFlag = True
        QBAddGLH
    End If
        
    ' update the batch record
    bat.debits = TotalDebits
    bat.credits = TotalCredits
    bat.nRecords = RecordCount
    bat.PutRecord bat.BatchNumber, FileName
    
    ' close the ADO record set and connection
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing

End Sub


Private Sub QBProcessCheckLine(orReportData As QBFC13Lib.IORReportData)
    
Dim colDataList As QBFC13Lib.IColDataList
Dim colData As QBFC13Lib.IColData
Dim colType As QBFC13Lib.IColDataList

Dim ndxEmp As Integer
Dim SSN As String
Dim dTxn As String
Dim enm As String
Dim CurIncome As Currency
Dim CurWageBase As Currency
Dim cc As Currency

Dim lngID As Long
Dim AcctNum As Long
    
    ' init the GLH variables
    GLHFiscalYear = Me.cmbFiscalYear
    GLHPeriod = bat.period
    GLHBatchNumber = bat.BatchNumber
    GLHSourceCode = 0
    GLHJournalSource = bat.JournalSource
    GLHHistType = "A"
    GLHUpdateFlag = True
    GLHReference = ""
    GLHDescription = ""
    GLHAmount = 0
    
    If (orReportData Is Nothing) Then Exit Sub
    
    Select Case orReportData.ortype
        
        Case QBFC13Lib.ENORReportData.orrdDataRow
            
            If (orReportData.DataRow Is Nothing) Then Exit Sub
            If (orReportData.DataRow.rowNumber Is Nothing) Then Exit Sub
            
'            If (orReportData.DataRow.RowData Is Nothing) Then Exit Sub
'            If (orReportData.DataRow.colDataList Is Nothing) Then Exit Sub
            
            Set colDataList = orReportData.DataRow.colDataList

            ' skip checking offset
            Set colData = colDataList.GetAt(0)
            If colData.Value Is Nothing Then Exit Sub
'            If colData.Value.GetValue() = "Checking" Then Exit Sub

            For i = 0 To colDataList.Count - 1
                Set colData = colDataList.GetAt(i)
                If (Not colData.Value Is Nothing) Then
                   
                   j = colData.colID.GetValue
                   y = colData.Value.GetValue
                   
                   If j = 2 Then    ' check number
                      
                      GLHReference = y
                   
                   ElseIf j = 4 Then       ' payee
                      
                      GLHDescription = y
                   
                   ElseIf j = 6 Then        ' account
                     
                      ' see if the account number is embedded
                      k = InStr(1, y, Chr(183), vbTextCompare)
                      If k <> 0 Then
                         x = Mid(y, k + 2, Len(y) - k + 2)
                      Else
                         x = y
                      End If
                     
                      k = xdbAccts.Find(1, 0, x, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                                                 
                      If k = -1 Then     ' not found - put to 0 and suspense - not found in QB COA
                         
                            GLHAccount = Me.cmbSuspAcct.ItemData(Me.cmbSuspAcct.ListIndex)
                      
                      Else
                         
                         ' bank amounts totaled and written as one entry
                         If xdbAccts(k, 3) = "Bank" Then
                            GLHUpdateFlag = False
                         End If
                         
                         ' is it in the WinGL chart of accts?
                         l = xdbAccts.Value(k, 4)
                         ' acct number translation if selected
                         If Me.chkAcctTranslate = 1 And Me.tdbAcctTranslateValue.Value > 0 Then
                            If Me.optDivide = True Then
                                l = l / Me.tdbAcctTranslateValue.Value
                            Else
                                l = l * Me.tdbAcctTranslateValue.Value
                            End If
                         End If
                         m = xdbGLAccount.Find(0, 0, l, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
                         If m >= 0 Then
                            GLHAccount = l
                         Else
                             GLHAccount = Me.cmbSuspAcct.ItemData(Me.cmbSuspAcct.ListIndex)
                         End If
                      
                      End If
                         
                   ElseIf j = 7 Then
                      
                      GLHAmount = CCur(y) * (-1)
                   
                   ElseIf j = 8 Then   ' column 8 for the banking line
                   
                      If GLHUpdateFlag = False Then GLHAmount = CCur(y)
                   
                   End If
                
                End If
            
            Next i

            ' add to the history file
            If GLHUpdateFlag Then
               GLHReference = RefNum
               GLHDescription = Desc
               QBAddGLH
               TotalAmount = TotalAmount + GLHAmount
            Else
               RefNum = GLHReference                 ' save info from the bank line
               Desc = GLHDescription
            End If
    
            ' update amount totals
            If k = -1 Then
               i = 0
            Else
               i = k
            End If
            xdbAccts(i, 2) = xdbAccts(i, 2) + GLHAmount
    
    End Select

End Sub

Private Sub QBGetAccounts()
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 4, 0)
    
    Set AccQ = RequestSet.AppendAccountQueryRq
    
'    Set AccQ = RequestSet.AppendAccountQueryRq.ORAccountListQuery.FullNameList
    
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
    Set qResponse = ResponseSet.ResponseList.GetAt(nRequest)

    ' check for errors
    If qResponse.StatusCode <> 0 Then
       
       If qResponse.StatusCode <= 499 Then
          IconType = vbInformation
       ElseIf qResponse.StatusCode <= 999 Then
          IconType = vbExclamation
       Else
          IconType = vbCritical
       End If
       
       MsgBox qResponse.StatusMessage & vbCrLf & _
              "Status Code: " & qResponse.StatusCode, IconType
              
       If qResponse.StatusCode >= 1000 Then  ' exit completely
          SessMgr.EndSession
          SessMgr.CloseConnection
          End
       End If
    
    End If

    Set RetList = qResponse.Detail
    
    If RetList Is Nothing Then Exit Sub   ' no accounts ???
    
    j = RetList.Count
    
    ' setup the xdb array
    ' row for each account
    ' col 0 = Account Name
    ' col 1 = Account description (account number)
    ' col 2 = Amount total
    ' col 3 = Account Type
    ' col 4 = Account Number
    ' start data w/ row #1
    xdbAccts.ReDim 0, 0, 0, 4
    xdbAccts.DefaultColumnType(2) = XTYPE_CURRENCY
    xdbAccts.DefaultColumnType(4) = XTYPE_LONG
    xdbAccts(0, 0) = ""
    xdbAccts(0, 1) = ""
    xdbAccts(0, 2) = 0
    xdbAccts(0, 3) = ""
    
    k = 0
        
    For i = 0 To j - 1
                
        Set ItemRet = RetList.GetAt(i)
        If (Not ItemRet Is Nothing) Then
            If (Not ItemRet.Name Is Nothing) Then
                k = k + 1
                xdbAccts.AppendRows (1)
                xdbAccts(k, 0) = ItemRet.Name.GetValue
                            
                If Not (ItemRet.Desc Is Nothing) Then
                    xdbAccts(k, 1) = ItemRet.Desc.GetValue
                Else
                    xdbAccts(k, 1) = ""
                End If
              
                xdbAccts(k, 2) = 0
                xdbAccts(k, 3) = ItemRet.AccountType.GetAsString
              
                ' assign the account number
                xdbAccts(k, 4) = CLng(GetNumber(xdbAccts(k, 1)))    ' from the QB account description
              
                ' use the QB acct number if it is there
                If Not (ItemRet.AccountNumber Is Nothing) Then
                    x = ItemRet.AccountNumber.GetValue
                    If IsNumeric(x) Then
                        xdbAccts(k, 4) = CLng(ItemRet.AccountNumber.GetValue)
                    End If
                End If
            
            End If
        End If
    Next i

    Set ItemRet = Nothing
    Set RetList = Nothing
    Set qResponse = Nothing
    Set ResponseSet = Nothing
    Set RequestSet = Nothing
    
End Sub

Private Function GetNumber(ByVal InString As String) As Long

' return a long from the digits at the beginning of a string

Dim x1, x2 As String
Dim ln, i1, i2 As Long

    GetNumber = 0
    If IsNull(InString) Then Exit Function
       
    x2 = ""
    ln = Len(InString)
    If ln = 0 Then Exit Function
    i1 = 0
    
    Do
       i1 = i1 + 1
       If i1 > ln Then Exit Do
       x1 = Mid(InString, i1, 1)
       If InStr(1, "0123456789", x1, vbTextCompare) = 0 Then Exit Do
       x2 = x2 & Mid(InString, i1, 1)
    Loop

    If IsNumeric(x2) Then GetNumber = CLng(x2)

End Function


Private Sub QBAddGLH()
               
    ' If GLHAccount = 0 Then GLHAccount = Me.tdbSuspAcct
               
    rs.AddNew
    rs.Fields("Account") = GLHAccount
    rs.Fields("FiscalYear") = GLHFiscalYear
    rs.Fields("Period") = GLHPeriod
    rs.Fields("BatchNumber") = GLHBatchNumber
    rs.Fields("Amount") = GLHAmount
    rs.Fields("Reference") = Mid(GLHReference, 1, 20)
    rs.Fields("Description") = Mid(GLHDescription, 1, 20)
    rs.Fields("SourceCode") = GLHSourceCode
    rs.Fields("JournalSource") = GLHJournalSource
    rs.Fields("HisType") = GLHHistType
    rs.Fields("UpdateFlag") = GLHUpdateFlag
    rs.Fields("PostDate") = Now()
    rs.Update

    ' update totals for the batch file
    If GLHAmount > 0 Then
       TotalDebits = TotalDebits + GLHAmount
    Else
       TotalCredits = TotalCredits + GLHAmount
    End If
    RecordCount = RecordCount + 1

    If RecordCount = 1 Or RecordCount Mod 10 = 0 Then
       frmProgress.lblMsg2 = "Loading Detail GL Line # " & RecordCount
       frmProgress.Refresh
    End If

End Sub

Private Sub optQBCheck_Click()
    If Me.optQBGL = False Then
        Me.fraQBOption.Enabled = False
        Me.optQBDetail.Enabled = False
        Me.optQBSummary.Enabled = False
        Me.cmbCheckingAcct.Enabled = True
        Me.chkQBChecking.Enabled = True
        Me.lblChkAcct.Enabled = True
    Else
        Me.fraQBOption.Enabled = True
        Me.optQBDetail.Enabled = True
        Me.optQBSummary.Enabled = True
        Me.cmbCheckingAcct.Enabled = False
        Me.chkQBChecking.Enabled = False
        Me.lblChkAcct.Enabled = False
    End If
End Sub

Private Sub optQBGL_Click()
    If Me.optQBGL = False Then
        Me.fraQBOption.Enabled = False
        Me.optQBDetail.Enabled = False
        Me.optQBSummary.Enabled = False
        
        Me.cmbCheckingAcct.Enabled = True
        
        Me.chkQBChecking.Enabled = True
        Me.lblChkAcct.Enabled = True
    Else
        Me.fraQBOption.Enabled = True
        Me.optQBDetail.Enabled = True
        Me.optQBSummary.Enabled = True
        Me.cmbCheckingAcct.Enabled = False
        Me.chkQBChecking.Enabled = False
        Me.lblChkAcct.Enabled = False
    End If
End Sub
