VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form BatchForm 
   Caption         =   " BATCH RECORD"
   ClientHeight    =   8250
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
   Icon            =   "BatchForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleMode       =   0  'User
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkQBChecking 
      Caption         =   "Use QB Checking Acct"
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   6600
      Width           =   2415
   End
   Begin TDBNumber6Ctl.TDBNumber tdbChkAcct 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      Calculator      =   "BatchForm.frx":030A
      Caption         =   "BatchForm.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchForm.frx":0396
      Keys            =   "BatchForm.frx":03B4
      Spin            =   "BatchForm.frx":03FE
      AlignHorizontal =   0
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdFileOpen 
      Height          =   375
      Left            =   3000
      Picture         =   "BatchForm.frx":0426
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
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
      TabIndex        =   9
      Top             =   5880
      Width           =   8775
   End
   Begin VB.CheckBox chkUseCurrentQB 
      Caption         =   "&Use currently opened QuickBooks File"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4800
      Width           =   6975
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
      Calendar        =   "BatchForm.frx":0730
      Caption         =   "BatchForm.frx":0848
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchForm.frx":08B4
      Keys            =   "BatchForm.frx":08D2
      Spin            =   "BatchForm.frx":0930
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
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox cmbJournal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   " Pick Journal Source"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Text            =   "cmbPeriod"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox cmbFiscalYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Text            =   "cmbFiscalYear"
      Top             =   3120
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
      Left            =   6030
      TabIndex        =   14
      Top             =   7560
      Width           =   2055
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
      Left            =   2790
      TabIndex        =   13
      Top             =   7560
      Width           =   2055
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
      Calendar        =   "BatchForm.frx":0958
      Caption         =   "BatchForm.frx":0A70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchForm.frx":0ADC
      Keys            =   "BatchForm.frx":0AFA
      Spin            =   "BatchForm.frx":0B58
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
   Begin TDBNumber6Ctl.TDBNumber tdbSuspAcct 
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      Top             =   6600
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      Calculator      =   "BatchForm.frx":0B80
      Caption         =   "BatchForm.frx":0BA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "BatchForm.frx":0C0C
      Keys            =   "BatchForm.frx":0C2A
      Spin            =   "BatchForm.frx":0C74
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
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
      Y1              =   4471.854
      Y2              =   4471.854
   End
   Begin VB.Label lblSuspAcct 
      Caption         =   "Suspense Acct #:"
      Height          =   375
      Left            =   7320
      TabIndex        =   29
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label lblChkAcct 
      Caption         =   "Checking Acct #:"
      Height          =   375
      Left            =   3240
      TabIndex        =   28
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblQBFile 
      Caption         =   "QuickBooks File Name"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblEndDate 
      Alignment       =   1  'Right Justify
      Caption         =   "&End Date:"
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblStartDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Start Date:"
      Height          =   375
      Left            =   3720
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label txtCredits 
      Caption         =   "CREDITS"
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label txtDebits 
      Caption         =   "DEBITS"
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
      TabIndex        =   19
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label txtRecord 
      Caption         =   "RECORDS IN BATCH"
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
      TabIndex        =   18
      Top             =   1680
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
      Left            =   6000
      TabIndex        =   17
      Top             =   2880
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
      TabIndex        =   16
      Top             =   2880
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
      TabIndex        =   15
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "BatchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BatchNumber As Long
Public userOK As Boolean
' Dim jou As New XArrayDB

Dim SessMgr As New QBSessionManager

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

Dim I, j, k, l As Long
Dim x, Y, z As String

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
Dim rs As New ADODB.Recordset

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
       Me.tdbChkAcct.Enabled = False
    Else
       Me.lblChkAcct.Enabled = True
       Me.tdbChkAcct.Enabled = True
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
    
Dim I As Integer
Dim v As Variant
Dim FY As Integer

    Me.cmbPeriod.Clear
    FY = CInt(cmbFiscalYear)
      
    If GLCompany.FirstPeriod = 1 Then
       v = DateSerial(FY, GLCompany.FirstPeriod, 1)
    Else
       v = DateSerial(FY - 1, GLCompany.FirstPeriod, 1)
    End If

    cmbPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    
    For I = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbPeriod.AddItem "Pd. #:" & I + 1 & " - " & Format(v, "mmmm-yyyy")
    Next I
    
    cmbPeriod.ListIndex = 0
    
    
'    cmbPeriod.Clear
'    Dim ndx, fy As Integer
'
'    fy = CInt(cmbFiscalYear)
'    For ndx = 1 To glcompany.NumberPds
'        cmbPeriod.AddItem glcompany.MonthName(ndx, fy)
'    Next ndx
'    cmbPeriod.ListIndex = glbatch.period - 1

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
    
    userOK = True
    
    GLBatch.FiscalYear = CLng(cmbFiscalYear)
    GLBatch.Period = cmbPeriod.ListIndex + 1
'    GLBatch.JournalSource = jou.Value(cmbJournal.ListIndex + 1, 0)
    GLBatch.JournalSource = Me.cmbJournal.ItemData(cmbJournal.ListIndex)
    
    ' add 100 to the journal source number if budget entry
    If chkBudget Then GLBatch.JournalSource = GLBatch.JournalSource + 100
    
'    glbatch.debits = CCur(txtDebits)
'    glbatch.credits = CCur(txtCredits)
'    glbatch.recct = CLng(txtRecords)
    
    GLBatch.Updated = Now
    GLBatch.UpdateUser = GLUser.ID
    
    BatchNumber = GLBatch.BatchNumber
    
    GLBatch.Save (Equate.RecPut)
    Response = True
    
    ' re-get the batch
    If GLBatch.GetBatch(BatchNumber) = False Then
        MsgBox "GL Batch error: " & BatchNumber, vbExclamation
        GoBack
    End If
    
    Me.Hide
    
    If Me.chkQB Then LoadQBData

    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub Init()

Dim ndx As Long
Dim CurFY As Integer
    
    userOK = False
    
    If GLBatch.GetBatch(BatchNumber) = False Then
        MsgBox "Batch not found?: ", vbExclamation
        GoBack
    End If
    
    txtCompanyName = GLCompany.Name
    lblBatchNumber = "Batch # " & GLBatch.BatchNumber
    lblCreated = "Created by " & GLBatch.CreateUser & " on " & ShowDate(GLBatch.Created)
    lblUpdated = "Record is OPEN (Not Updated)"
    txtRecord = "RECORD COUNT = " & CStr(GLBatch.RecCt)
    txtDebits = "DEBITS = " & Format(GLBatch.Debits, "#.00")
    txtCredits = "CREDITS = " & Format(GLBatch.Credits, "#.00")
    

'    For ndx = glcompany.FirstFiscalYear To Year(Now) + 1
'        cmbFiscalYear.AddItem ndx
'    Next ndx
'
'    'if glbatch.fiscalYear=0 then
'    cmbFiscalYear = glbatch.fiscalYear
    
    CurFY = Int(GLCompany.LastClose / 10 ^ 4)
    If Int(GLCompany.LastClose / 100) Mod 100 <> 1 Then CurFY = CurFY + 1
    If CurFY < 1990 Or CurFY > 2040 Then CurFY = Year(Now())
    
    For ndx = CurFY + 1 To CurFY - 5 Step -1
        cmbFiscalYear.AddItem ndx
    Next ndx
    cmbFiscalYear.ListIndex = 1
    
'    For ndx = 1 To glcompany.NumberPds
'        cmbPeriod.AddItem glcompany.MonthName(ndx, glbatch.fiscalYear)
'        cmbPeriod.AddItem glcompany.MonthName(ndx, CurFY)
'    Next ndx
''    cmbPeriod.ListIndex = glbatch.period - 1
'    cmbPeriod.ListIndex = 0
    
    SQLString = " SELECT * FROM GLJournal ORDER BY JournalSource "
    If GLJournal.GetBySQL(SQLString) = False Then
        MsgBox "No GLJournal records found?: ", vbExclamation
        GoBack
    End If
    
    ndx = 0
    Do
        ndx = ndx + 1
        cmbJournal.AddItem (CStr(GLJournal.JournalSource) & "-" & GLJournal.JournalName)
        cmbJournal.ItemData(cmbJournal.NewIndex) = GLJournal.JournalSource
'        If jou.Value(ndx, 0) = GLBatch.JournalSource Then
'            cmbJournal.ListIndex = ndx - 1
'        End If
        If GLJournal.GetNext = 0 Then Exit Do
    Loop

    Response = False

End Sub

Private Sub cmdPrint_Click()
    ReviewReport.BatchNumber = BatchNumber
    ReviewReport.Show vbModal
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        ' Case vbKeyF6: OnPrint
    End Select
End Sub

Private Sub Form_Load()

'    Set jou = xFactory.GetJournals(FileName)
'    Set JournalList.Array = jou
'    JournalList.Columns(0).Width = 500
'    JournalList.Columns(1).Width = 3500
     
     Response = False

     ' hide the QuickBook fields by default
     QBShow False


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
       
       Me.tdbChkAcct.Visible = False
       Me.TDBDate1.Visible = False
       Me.TDBDate2.Visible = False
       Me.tdbSuspAcct.Visible = False
       
       Me.chkUseCurrentQB.Visible = False
    
       Me.cmdFileOpen.Visible = False
    
    Else
       
       Me.lblChkAcct.Visible = True
       Me.lblEndDate.Visible = True
       Me.lblStartDate.Visible = True
       Me.lblSuspAcct.Visible = True
       Me.lblQBFile.Visible = True
       
       Me.tdbChkAcct.Visible = True
       Me.TDBDate1.Visible = True
       Me.TDBDate2.Visible = True
       Me.tdbSuspAcct.Visible = True
       
       Me.chkUseCurrentQB.Visible = True
       Me.chkUseCurrentQB.Value = 1
       
       Me.cmdFileOpen.Visible = True
       Me.txtQBFile.Visible = True
       
       ' set defaults
       Me.tdbSuspAcct = GLCompany.SuspAcct
       
       If Me.cmbFiscalYear.ListIndex <> -1 And Me.cmbPeriod.ListIndex <> -1 Then
          
          If GLCompany.FirstPeriod = 1 Then        ' Jan is the first period
             Mo = Me.cmbPeriod.ListIndex + 1
             Me.TDBDate1 = DateSerial(Me.cmbFiscalYear, Mo, 1)
             Me.TDBDate2 = LastDay(Me.cmbFiscalYear, Mo)
          Else
             Mo = Me.cmbPeriod.ListIndex + GLCompany.FirstPeriod    ' the list index is zero based
             If Mo <= 12 Then
                yr = Me.cmbFiscalYear - 1
             Else
                yr = Me.cmbFiscalYear
                Mo = Me.cmbPeriod.ListIndex - 12 + GLCompany.FirstPeriod
             End If
             Me.TDBDate1 = DateSerial(yr, Mo, 1)
             Me.TDBDate2 = LastDay(yr, Mo)
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

Private Sub LoadQBData()
    
    frmProgress.lblMsg1 = GLCompany.Name
    frmProgress.lblMsg2 = "Now loading QuickBooks Data ..."
    frmProgress.Show

    ' connect to the data base with ADO
    ' 2020 - not needed - connection already open
    ' x = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    ' CNOpen x, Password

    ' open a record set to the GLHistory file
    rsInit "SELECT * FROM GLHistory", cn, rs
    
    SessMgr.OpenConnection2 "", "Windows GL Entry", ctLocalQBD
    SessMgr.BeginSession Me.txtQBFile, omDontCare

    frmProgress.lblMsg2 = "Now loading QuickBooks Chart of Accounts .... "
    frmProgress.Refresh

    GetAccounts
    
    frmProgress.lblMsg2 = "Now loading QuickBooks Check Detail .... "
    frmProgress.Refresh
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 4, 0)
    Set RepQ2 = RequestSet.AppendGeneralDetailReportQueryRq
    RepQ2.IncludeAccounts.SetValue (iaAll)
    RepQ2.GeneralDetailReportType.SetValue (gdrtCheckDetail)
    RepQ2.ReportAccountFilter.ORReportAccountFilter.AccountTypeFilter.SetValue (atfBank)
    
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
                    ProcessLine orReportData
                End If
            Next index
        End If
    End If
    
    ' create the offset entry
    If Me.chkQBChecking Then                ' use the bank accounts from QB
        For I = 1 To xdbAccts.UpperBound(1)
            If xdbAccts(I, 3) = "Bank" Then
               GLHAccount = xdbAccts(I, 4)
               GLHAmount = xdbAccts(I, 2)
               GLHReference = xdbAccts(I, 0)
               GLHDescription = "QB " & Format(Now(), "mm/dd/yyyy")
               GLHUpdateFlag = True
               AddGLH
            End If
        Next I
    Else            ' total amount to the account specified
        GLHAccount = Me.tdbChkAcct
        GLHAmount = TotalAmount * (-1)
        GLHReference = "Checking"
        GLHDescription = "QB " & Format(Now(), "mm/dd/yyyy")
        GLHUpdateFlag = True
        AddGLH
    End If

    ' update the batch record
    GLBatch.Debits = TotalDebits
    GLBatch.Credits = TotalCredits
    GLBatch.RecCt = RecordCount
    GLBatch.Save (Equate.RecPut)
    
    ' close the ADO record set and connection
    rs.Close
    Set rs = Nothing
    ' 2020
    ' cn.Close
    ' Set cn = Nothing
    
    ' close the QB connection
    SessMgr.CloseConnection
    Set SessMgr = Nothing
    
    ' hide the progress screen
    frmProgress.Hide
    
End Sub

Private Sub ProcessLine(orReportData As QBFC13Lib.IORReportData)
    
Dim colDataList As QBFC13Lib.IColDataList
Dim colData As QBFC13Lib.IColData
Dim ColType As QBFC13Lib.IColDataList

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
    GLHPeriod = GLBatch.Period
    GLHBatchNumber = GLBatch.BatchNumber
    GLHSourceCode = 0
    GLHJournalSource = GLBatch.JournalSource
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

            For I = 0 To colDataList.Count - 1
                Set colData = colDataList.GetAt(I)
                If (Not colData.Value Is Nothing) Then
                   
                   j = colData.colID.GetValue
                   Y = colData.Value.GetValue
                   
                   If j = 2 Then    ' check number
                      
                      GLHReference = Y
                   
                   ElseIf j = 4 Then       ' payee
                      
                      GLHDescription = Y
                   
                   ElseIf j = 6 Then        ' account
                     
                      ' see if the account number is embedded
                      k = InStr(1, Y, Chr(183), vbTextCompare)
                      If k <> 0 Then
                         x = Mid(Y, k + 2, Len(Y) - k + 2)
                      Else
                         x = Y
                      End If
                     
                      k = xdbAccts.Find(1, 0, x, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                         
                      If k = -1 Then     ' not found - put to 0 and suspense
                         
                         GLHAccount = CLng(Me.tdbSuspAcct)
                      
                      Else
                         
                         ' bank amounts totaled and written as one entry
                         If xdbAccts(k, 3) = "Bank" Then
                            GLHUpdateFlag = False
                         End If
                         
                         GLHAccount = xdbAccts(k, 4)
                      
                      End If
                   
                   ElseIf j = 7 Then
                      
                      GLHAmount = CCur(Y) * (-1)
                   
                   ElseIf j = 8 Then   ' column 8 for the banking line
                   
                      If GLHUpdateFlag = False Then GLHAmount = CCur(Y)
                   
                   End If
                
                End If
            
            Next I

            ' add to the history file
            If GLHUpdateFlag Then
               GLHReference = RefNum
               GLHDescription = Desc
               AddGLH
               TotalAmount = TotalAmount + GLHAmount
            Else
               RefNum = GLHReference                 ' save info from the bank line
               Desc = GLHDescription
            End If
    
            ' update amount totals
            If k = -1 Then
               I = 0
            Else
               I = k
            End If
            xdbAccts(I, 2) = xdbAccts(I, 2) + GLHAmount
    
    End Select

End Sub

Private Sub GetAccounts()
    
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
    
    For I = 0 To j - 1
        Set ItemRet = RetList.GetAt(I)
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
              
              If Not (ItemRet.AccountNumber Is Nothing) Then
                 xdbAccts(k, 4) = CLng(ItemRet.AccountNumber.GetValue)
              End If
              
           End If
        End If
    Next I

    Set ItemRet = Nothing
    Set RetList = Nothing
    Set qResponse = Nothing
    Set ResponseSet = Nothing
    Set RequestSet = Nothing
    
End Sub

Private Function GetNumber(ByVal InString As String) As Long

' return a long from the digits at the beginning of a string

Dim x1, x2 As String
Dim Ln, i1, i2 As Long

    GetNumber = 0
    If IsNull(InString) Then Exit Function
       
    x2 = ""
    Ln = Len(InString)
    If Ln = 0 Then Exit Function
    i1 = 0
    
    Do
       i1 = i1 + 1
       If i1 > Ln Then Exit Do
       x1 = Mid(InString, i1, 1)
       If InStr(1, "0123456789", x1, vbTextCompare) = 0 Then Exit Do
       x2 = x2 & Mid(InString, i1, 1)
    Loop
    
    If IsNumeric(x2) Then GetNumber = CLng(x2)

End Function


Private Sub AddGLH()
               
    If GLHAccount = 0 Then GLHAccount = Me.tdbSuspAcct
               
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
    
    ' 2020
    rs.Fields("PostDate") = Date
    
    rs.Update

    ' update totals for the batch file
    If GLHAmount > 0 Then
       TotalDebits = TotalDebits + GLHAmount
    Else
       TotalCredits = TotalCredits + GLHAmount
    End If
    RecordCount = RecordCount + 1

    If RecordCount = 1 Or RecordCount Mod 10 = 0 Then
       frmProgress.lblMsg2 = "Loading Check Detail Line # " & RecordCount
       frmProgress.Refresh
    End If

End Sub

