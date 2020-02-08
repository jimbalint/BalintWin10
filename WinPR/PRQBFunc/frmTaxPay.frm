VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTaxPay 
   Caption         =   "Payroll Tax Payments in QuickBooks"
   ClientHeight    =   11040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   11040
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoName 
      Caption         =   "Don't include Employee Name in QB check memo field"
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   2280
      Width           =   5175
   End
   Begin VB.ComboBox cmbQBChk 
      Height          =   360
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1800
      Width           =   3975
   End
   Begin VB.ComboBox cmbQBAP 
      Height          =   360
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1200
      Width           =   3975
   End
   Begin VSFlex8Ctl.VSFlexGrid fgDist 
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   7920
      Width           =   12975
      _cx             =   22886
      _cy             =   2778
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   7080
      TabIndex        =   12
      Top             =   10200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Display Option  "
      Height          =   855
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   4575
      Begin VB.OptionButton optQBSetup 
         Caption         =   "QuickBooks Setup"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optAmounts 
         Caption         =   "Payroll Amounts"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "CLEAR ALL"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "SELECT ALL"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "&UPDATE TO QB"
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton cmdQBRefresh 
      Caption         =   "&REFRESH QB ACCOUNTS"
      Height          =   615
      Left            =   5040
      TabIndex        =   3
      Top             =   10200
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   12975
      _cx             =   22886
      _cy             =   8916
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   10680
      TabIndex        =   1
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Net Pay Bank Account:"
      Height          =   495
      Left            =   7560
      TabIndex        =   17
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "A/P Account:"
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblBatchInfo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   12855
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   9720
      Width           =   12495
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "frmTaxPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset

Dim rsQBCount As Long

Dim rsQBRecID As Long
Dim rsQB As New ADODB.Recordset
Dim rsQBDist As New ADODB.Recordset
Dim rsCity As New ADODB.Recordset
Dim rsState As New ADODB.Recordset
Dim rsNet As New ADODB.Recordset
Dim rsItem As New ADODB.Recordset
Dim rsGross As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsPRHist As New ADODB.Recordset

Dim DueDrop, PayDrop As String
Dim QBPayeeDrop As String
Dim QBExpAcctDrop As String
Dim QBLiabAcctDrop As String
Dim JobDrop As String

Dim LoadFlag As Boolean
Dim Flg As Boolean
Dim boo As Boolean

Dim i, j, k As Long
Dim X, Y, z As String
Dim P1, P2, P3 As Currency
Dim D2 As Date

Dim SSTax, FWTTax, MedTax, NetPay As Currency

' track 6.2% er match for 2011
Dim SSTax62 As Currency
Dim TaxYear As Long

Dim FUNWage, SUNWage As Currency
Dim TotalGross As Currency

Dim StartYM, EndYM As Long

Dim GlobalNet, GlobalFed, GlobalSWT As Long

Dim billAdd As IBillAdd
Dim expenseLineAdd1 As IExpenseLineAdd
Dim responseMsgSet As IMsgSetResponse
Dim ResponseList As IResponseList
Dim requestMsgSet As IMsgSetRequest
Dim Response As IResponse

Dim JnlAddReq As IJournalEntryAdd
Dim orJournalLine1 As IORJournalLine

Dim DistRecs As Long
Dim QBCount As Long

Dim QBCheckingAcct, QBAPAcct As Long
Dim CompanyGlobalID As Long
Dim HiCheckDate As Date
Dim DirDepTotal As Currency

Private Sub Form_Load()
    
    ' *** To Do 05/26/10 ***
    ' multi batch test - sep update for each batch - check net pay and dir dep
    '       sep due date calc ....
    ' print - show pay to
    ' net pay by check / dir dep - show ee brkdwn in lower fg
    
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' TO DO
    '   >>> Multi State Unemp %
    '
    '   no edit columns
    '   store transactions
    '       use for report to pay mthly/qtrly ???
    '       after the fact edits ....
    '       def macro - QB lookup???
    '   update flags ....
    '   QB Refresh - re-init drops
    '
    '   Net Pay store
    '       expand for each check on update / print
    '   Direct Deposit Test
    '
    '   qtrly rounding calcs
    '
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    ' **********************************************************
    ' tax pay defn stored in PRGlobal
    ' Fed / State / City / Wkc / FUN / SUN / Items
    '
    ' UserID = CompanyID
    ' TypeCode - tax type
    ' Byte1 = Payment Due - 0=JnlEntry / PREquate.PeriodType - Pay / Mthly / Qtrly / Annual
    ' Byte2 = Select
    ' Byte3 = Pay Type - PREquate.PayType Check / EFT
    ' Byte4 = DueDays
    ' Byte5 - 1000000 added to the related id for matching amounts
    ' Var1 = QBPayee
    ' Var2 = QBAccount - EXPENSE
    
    ' *** NOT USED *** Var3 = QBCheckingAccount  *** NOT USED ***
    
    ' Var4 = RelatedID - StateID / CityID
    ' Var5 = QBAccount - Accrued LIABILITY
    ' **********************************************************
    
    ' ----------------------------------------------------------
    
    ' record set of PRHist ID's -
    ' use to set PRHist.QBUpdateFlag
    rsPRHist.CursorLocation = adUseClient
    rsPRHist.Fields.Append "HistID", adDouble
    rsPRHist.Open , , adOpenDynamic, adLockOptimistic
    
    ' get the date range
    If BatchNumbr = 0 Then
        frmDateRange.chkQBOverride.Visible = True
        frmDateRange.Show vbModal
        If InitFlag = False Then GoBack
        If frmDateRange.optCheckDate = True Then
            OptDate = "CHECK DATE"
        ElseIf frmDateRange.optPEDate = True Then
            OptDate = "P/E DATE"
        End If
    Else
        If PRBatch.GetByID(BatchNumbr) = False Then
            MsgBox "PR Batch not found: " & BatchNumbr, vbExclamation
            GoBack
        End If
        PEDate = PRBatch.PEDate
        CheckDt = PRBatch.CheckDate
    End If
    
    If BatchNumbr > 0 Then
        Me.lblBatchInfo = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        PEDate = PRBatch.PEDate
        PRBatchID = BatchNumbr
        CheckDate = PRBatch.CheckDate
        OptDate = " "
    Else
        If OptDate = "CHECK DATE" Then
            Me.lblBatchInfo = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            Me.lblBatchInfo = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    
    ' ----------------------------------------------------------
    
    Me.lblMsg1 = ""
    
    LoadFlag = True
    
    ' >>>>>>>>>>>>>>>>>
    ' Clear PRGlobal records
    ' SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBPayCity
    ' cnDes.Execute SQLString
    ' SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBPayState
    ' cnDes.Execute SQLString
    ' >>>>>>>>>>>>>>>>>
    
    Me.lblCompanyName = PRCompany.Name
    
    DueDrop = "|#0;Jnl Entry|#4;Pay|#1;Month|#2;Quarter|#3;Year"
    PayDrop = "|#0;---|#1;Check|#2;EFT"
    
    ' >>>> drop for Payee / Accounts
    ' >>>> filter QB Accounts for type = checking ???
    
    LoadQBAccounts
    DefineRS
    
    ' loop thru the batches
    If BatchNumbr > 0 Then
        If PRBatch.GetByID(BatchNumbr) = False Then
            MsgBox "PR Batch Not Found: " & BatchNumbr, vbExclamation
            GoBack
        End If
        HiCheckDate = PRBatch.CheckDate
        ScanHist PRBatch.BatchID
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRBatch WHERE CheckDate >= " & CLng(StartDate) & _
                        " AND CheckDate <= " & CLng(EndDate) & _
                        " ORDER BY CheckDate"
        Else
            SQLString = "SELECT * FROM PRBatch WHERE PEDate >= " & CLng(StartDate) & _
                        " AND PEDate <= " & CLng(EndDate) & _
                        " ORDER BY CheckDate"
        End If
        If PRBatch.GetBySQL(SQLString) = False Then
            MsgBox "No batches found in the date range", vbInformation
            GoBack
        End If
        Do
            ScanHist PRBatch.BatchID
            HiCheckDate = PRBatch.CheckDate
            If PRBatch.GetNext = False Then Exit Do
        Loop
    End If
    
    ' ---- fill in jobid unassigned ----
    If rsJob.RecordCount > 0 Then
        rsJob.MoveFirst
        Do
            If rsJob!JobID = 0 Then
                rsJob!JobID = 999999999
                rsJob.Update
            End If
            rsJob.MoveNext
        Loop Until rsJob.EOF
    End If
    
    rsUpdate
    
    rsQBDist.Sort = "RecID, JobID"
    
    GridInit
    
    If rsQB.RecordCount = 0 Then
        MsgBox "No History found!", vbExclamation
        GoBack
    End If
    
    Me.optAmounts = True
    ColDisplay
    
    Me.KeyPreview = True

    ' *********
    ' Me.cmdPay.Enabled = False

    LoadFlag = False

    fgDistDisplay

End Sub

Private Sub DefineRS()
    
    ' tax payment definitions
    ' used to store the total amounts per category
    rsQB.CursorLocation = adUseClient
    rsQB.Fields.Append "Select", adBoolean
    rsQB.Fields.Append "Desc", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "DuePeriod", adInteger
    rsQB.Fields.Append "DueDays", adVarChar, 5, adFldIsNullable
    rsQB.Fields.Append "QBExpenseAcct", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "QBLiabilityAcct", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "QBPayTo", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "DueDate", adVarChar, 15, adFldIsNullable
    rsQB.Fields.Append "PayType", adInteger
    rsQB.Fields.Append "Amount", adCurrency
    rsQB.Fields.Append "GlobalID", adDouble
    rsQB.Fields.Append "SortOrder", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "TypeCode", adInteger
    rsQB.Fields.Append "RelatedID", adDouble
    rsQB.Fields.Append "RecID", adDouble
    rsQB.Open , , adOpenDynamic, adLockOptimistic
    
    ' splits by Job
    rsQBDist.CursorLocation = adUseClient
    rsQBDist.Fields.Append "RecID", adDouble
    rsQBDist.Fields.Append "JobID", adDouble
    rsQBDist.Fields.Append "Amount", adCurrency
    rsQBDist.Open , , adOpenDynamic, adLockOptimistic
    
    ' recordset for the unique State and City ID's
    rsState.CursorLocation = adUseClient
    rsState.Fields.Append "RecID", adDouble
    rsState.Fields.Append "Amount", adCurrency
    rsState.Fields.Append "SUNWage", adCurrency
    rsState.Fields.Append "Gross", adCurrency
    rsState.Open , , adOpenDynamic, adLockOptimistic
    
    rsCity.CursorLocation = adUseClient
    rsCity.Fields.Append "RecID", adDouble
    rsCity.Fields.Append "Amount", adCurrency
    rsCity.Open , , adOpenDynamic, adLockOptimistic
    
    ' PRGlobal record for net pay
    rsNet.CursorLocation = adUseClient
    rsNet.Fields.Append "EmployeeID", adDouble
    rsNet.Fields.Append "NetPay", adCurrency
    rsNet.Fields.Append "CheckNumber", adDouble
    rsNet.Open , , adOpenDynamic, adLockOptimistic
    
    rsItem.CursorLocation = adUseClient
    rsItem.Fields.Append "ItemID", adDouble
    rsItem.Fields.Append "Amount", adCurrency
    rsItem.Fields.Append "Desc", adVarChar, 30, adFldIsNullable
    rsItem.Open , , adOpenDynamic, adLockOptimistic
    
    ' for wage by Dept
    rsGross.CursorLocation = adUseClient
    rsGross.Fields.Append "DepartmentID", adDouble
    rsGross.Fields.Append "Amount", adCurrency
    rsGross.Open , , adOpenDynamic, adLockOptimistic
    
    ' for amounts by Job
    rsJob.CursorLocation = adUseClient
    rsJob.Fields.Append "TypeCode", adDouble
    rsJob.Fields.Append "JobID", adDouble
    rsJob.Fields.Append "RelatedID", adDouble
    rsJob.Fields.Append "Amount", adCurrency
    rsJob.Fields.Append "Gross", adCurrency
    rsJob.Open , , adOpenDynamic, adLockOptimistic

End Sub

Private Sub GridInit()
    
    If rsQB.RecordCount = 0 Then Exit Sub
    
    ' grid setups
    With Me.fg
        
        ' from Set Grid in PRGlobal Module - no alt row color
        .FixedCols = 0
        .FocusRect = flexFocusSolid
        .DataMode = flexDMBound
        .Editable = flexEDKbdMouse
        Set .DataSource = rsQB.DataSource
        .DataMember = rsQB.DataMember
        .TabBehavior = flexTabCells
        
        .ColComboList(.ColIndex("DuePeriod")) = DueDrop
        .ColComboList(.ColIndex("PayType")) = PayDrop
        .ColComboList(.ColIndex("QBPayTo")) = QBPayeeDrop
        .ColComboList(.ColIndex("QBExpenseAcct")) = QBExpAcctDrop
        .ColComboList(.ColIndex("QBLiabilityAcct")) = QBLiabAcctDrop
    
        .ColHidden(.ColIndex("PayType")) = True
        .ColHidden(.ColIndex("GlobalID")) = True
        .ColHidden(.ColIndex("SortOrder")) = True
        .ColHidden(.ColIndex("TypeCode")) = True
        .ColHidden(.ColIndex("RelatedID")) = True
        .ColHidden(.ColIndex("RecID")) = True
    
        .ColWidth(.ColIndex("QBPayTo")) = 2200
        .ColWidth(.ColIndex("QBExpenseAcct")) = 2200
        .ColWidth(.ColIndex("QBLiabilityAcct")) = 2200
        .ColWidth(.ColIndex("DueDate")) = 1300
        .ColWidth(.ColIndex("DueDays")) = 1000
        
        .ColAlignment(.ColIndex("DueDate")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("DuePeriod")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("DueType")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("QBPayTo")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("QBExpenseAcct")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("Amount")) = flexAlignRightCenter
    
        ' description column
        .ColKey(.ColIndex("Desc")) = "Desc"
        .TextMatrix(0, .ColIndex("Desc")) = "Item Description"
        .ColWidth(.ColIndex("Desc")) = 3500
        .ColAlignment(.ColIndex("Desc")) = flexAlignLeftCenter
    
        ' row colors by type code
        j = 0
        k = 0
        For i = 1 To fg.Rows - 1
            If .TextMatrix(i, .ColIndex("TypeCode")) <> k And k <> 0 Then
                If j = 0 Then
                    j = 1
                Else
                    j = 0
                End If
            End If
            If j = 0 Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = RGB(192, 192, 192)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbYellow
            Else
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
            End If
            k = .TextMatrix(i, .ColIndex("TypeCode"))
        Next i
    
        .AllowSelection = False
    
    End With

    SetGrid rsQBDist, Me.fgDist

    With Me.fgDist
        
        .ColHidden(.ColIndex("RecID")) = True
        
        .ColKey(.ColIndex("JobID")) = "JobID"
        .TextMatrix(0, .ColIndex("JobID")) = "Job Name"
        .ColComboList(.ColIndex("JobID")) = JobDrop
        
        .AutoSize 0, .Cols - 1, False, 200
        
        .Editable = flexEDNone
    
    End With

    ' init due dates
    With fg
        For i = 1 To .Rows - 1
            .Row = i
            SetDueDate
        Next i
        .Row = 1
        .TopRow = 1
    End With

    ' scan the recordset and assign N/A values
    If rsQB.RecordCount > 0 Then
        rsQB.MoveFirst
        Do
            With fg
                
                If rsQB!Desc = "Net Pay by Check" Or rsQB!Desc = "Direct Deposit" Then
                
                    .TextMatrix(.Row, .ColIndex("QBLiabilityAcct")) = "0"
                    .TextMatrix(.Row, .ColIndex("DuePeriod")) = PREquate.PeriodTypePay
                    .TextMatrix(.Row, .ColIndex("DueDays")) = ""
                    .TextMatrix(.Row, .ColIndex("DueDate")) = PRBatch.CheckDate
                
                ElseIf rsQB!DuePeriod = 0 Then
                    
                    ' jnl entry
                    ' .TextMatrix(.Row, .ColIndex("DuePeriod")) = "0"
                    .TextMatrix(.Row, .ColIndex("DueDays")) = ""
                    .TextMatrix(.Row, .ColIndex("DueDate")) = ""
                    .TextMatrix(.Row, .ColIndex("QBPayTo")) = "0"
                    
                Else
                    
                    ' payable
                    .TextMatrix(.Row, .ColIndex("QBLiabilityAcct")) = "0"
                
                End If
            End With
            rsQB.Update
            rsQB.MoveNext
        Loop Until rsQB.EOF
        rsQB.MoveFirst
    End If

End Sub

Private Sub ScanHist(ByVal BtchID As Long)
    
    ' *******************************************************
    ' create the record sets
    
    BatchNumbr = BtchID
    
    ' ***************************************************
    ' scan the data
    
    ' outer loop for the history - check the update flag
    SQLString = "SELECT * FROM PRHist WHERE BatchID = " & BatchNumbr
    If PRHist.GetBySQL(SQLString) = False Then Exit Sub
    
    Do
    
        If frmDateRange.chkQBOverride = 0 And PRHist.QBUpdateFlag = 1 Then GoTo NxtPrHist
        
        ' keep a record of each PRHistID
        rsPRHist.AddNew
        rsPRHist!HistID = PRHist.HistID
        rsPRHist.Update
        
        ' update the rs for State and City ID's
        rsState.Find "RecID = " & PRHist.StateID, 0, adSearchForward, 1
        If rsState.EOF Then
            rsState.AddNew
            rsState!RecID = PRHist.StateID
            rsState!Amount = 0
            rsState!SUNWage = 0
            rsState!Gross = 0
            rsState.Update
        End If
        
        ' >> move to PRDist for multi state pay checks <<
        rsState!Gross = rsState!Gross + PRHist.Gross
        rsState!Amount = rsState!Amount + PRHist.SWTTax
        rsState!SUNWage = rsState!SUNWage + PRHist.SUNWage
        rsState.Update
        
        ' net pay by check - include for Dir Dep?
        If PRHist.Net <> 0 Then
            NetPay = NetPay + PRHist.Net
            rsNet.AddNew
            rsNet!EmployeeID = PRHist.EmployeeID
            rsNet!NetPay = PRHist.Net
            rsNet!CheckNumber = PRHist.CheckNumber
            rsNet.AddNew
        End If
        
        SSTax = SSTax + PRHist.SSTax
        
        SSTax62 = SSTax62 + Round(PRHist.SSWage * 0.062, 2)
        TaxYear = Int(PRHist.YearMonth / 100)
        
        MedTax = MedTax + PRHist.MedTax
        FWTTax = FWTTax + PRHist.FWTTax
        FUNWage = FUNWage + PRHist.FUNWage
        
        ' ----------------------------------------------------------------------
        ' scan PRDist for the PRHist record
        SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
        If PRDist.GetBySQL(SQLString) = True Then
            Do
        
                rsCity.Find "RecID = " & PRDist.CityID, 0, adSearchForward, 1
                If rsCity.EOF Then
                    rsCity.AddNew
                    rsCity!RecID = PRDist.CityID
                    rsCity!Amount = 0
                    rsCity.Update
                End If
                rsCity!Amount = rsCity!Amount + PRDist.CityTax
                rsCity.Update
            
                ' gross wage by dept
                SQLString = "DepartmentID = " & PRDist.DepartmentID
                rsGross.Find SQLString, 0, adSearchForward, 1
                If rsGross.EOF = True Then
                    rsGross.AddNew
                    rsGross!DepartmentID = PRDist.DepartmentID
                    rsGross!Amount = 0
                End If
                rsGross!Amount = rsGross!Amount + PRDist.Amount
                rsGross.Update
                
                ' gross by job / dept
                rsJob.Filter = adFilterNone
                SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayGrossPay & _
                            " AND JobID = " & PRDist.JobID & _
                            " AND RelatedID = " & PRDist.DepartmentID
                rsJob.Filter = SQLString
                If rsJob.RecordCount = 0 Then
                    rsJob.AddNew
                    rsJob!TypeCode = PREquate.GlobalTypeQBPayGrossPay
                    rsJob!JobID = PRDist.JobID
                    rsJob!RelatedID = PRDist.DepartmentID
                End If
                rsJob!Amount = rsJob!Amount + PRDist.Amount
                rsJob.Update
                rsJob.Filter = adFilterNone
                
                ' gross wage by job / state - for state unemp split
                rsJob.Filter = adFilterNone
                SQLString = "TypeCode = " & PREquate.GlobalTypeQBPaySUN & _
                            " AND JobID = " & PRDist.JobID & _
                            " AND RelatedID = " & PRDist.StateID
                rsJob.Filter = SQLString
                If rsJob.RecordCount = 0 Then
                    rsJob.AddNew
                    rsJob!TypeCode = PREquate.GlobalTypeQBPaySUN
                    rsJob!JobID = PRDist.JobID
                    rsJob!RelatedID = PRDist.StateID
                End If
                rsJob!Gross = rsJob!Gross + PRDist.Amount
                rsJob!Amount = rsJob!Amount + PRDist.SUNWage
                rsJob.Update
                rsJob.Filter = adFilterNone
                
                ' gross wage by job - for Fed Unemp and SS/MED
                rsJob.Filter = adFilterNone
                SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayFUN & _
                            " AND JobID = " & PRDist.JobID
                rsJob.Filter = SQLString
                If rsJob.RecordCount = 0 Then
                    rsJob.AddNew
                    rsJob!TypeCode = PREquate.GlobalTypeQBPayFUN
                    rsJob!JobID = PRDist.JobID
                    rsJob!Amount = 0
                End If
                rsJob!Amount = rsJob!Amount + PRDist.Amount
                rsJob.Update
                rsJob.Filter = adFilterNone
                
                TotalGross = TotalGross + PRDist.Amount
    
                If PRDist.GetNext = False Then Exit Do
            Loop
        End If
    
        ' ----------------------------------------------------------------------
        ' scan PRItemHist
        SQLString = "SELECT * FROM PRItemHist WHERE HistID = " & PRHist.HistID
        If PRItemHist.GetBySQL(SQLString) = True Then
            Do
            
                ' lump direct deposit together
                If PRItemHist.ItemType = PREquate.ItemTypeDirDepDed Then
                    i = 999999
                Else
                    i = PRItemHist.EmployerItemID
                End If
                rsItem.Find "ItemID = " & i, 0, adSearchForward, 1
                If rsItem.EOF Then
                    rsItem.AddNew
                    rsItem!ItemID = i
                    rsItem!Amount = 0
                    If i = 999999 Then
                        rsItem!Desc = "Direct Deposit"
                    Else
                        If PRItem.GetByID(PRItemHist.EmployerItemID) = False Then
                            MsgBox "Employer Item not found: " & PRItemHist.EmployerItemID, vbExclamation
                            GoBack
                        End If
                        rsItem!Desc = Mid(PRItem.Title, 1, 30)
                    End If
                    rsItem!Amount = 0
                    rsItem.Update
                End If
                rsItem!Amount = rsItem!Amount + PRItemHist.Amount
                rsItem.Update
            
                ' include match amount? bozo
                '   add offset entry also?????
                '   review save routine
                '   set flag in PRGlobal to subtract 100000 from the id
                '   review the match logic in the item detail report
                If PRItem.GetByID(PRItemHist.EmployerItemID) = True Then
                    If PRItem.MatchPct <> 0 Then
                    
                        ' calculate the matching amount
                        Dim MatchAmt As Currency
                        MatchAmt = 0
                        P1 = Round((PRHist.Gross - PRItemHist.WageExcluded) * PRItem.MaxPct / 100, 2) ' wage base x max pct
                        If P1 <= PRItemHist.Amount Then
                            MatchAmt = Round(P1 * PRItem.MatchPct / 100, 2)
                        Else
                            MatchAmt = Round(PRItemHist.Amount * PRItem.MatchPct / 100, 2)
                        End If
                    
                        ' add to er item id
                        i = PRItem.EmployerItemID + 1000000
                        rsItem.Find "ItemID = " & i, 0, adSearchForward, 1
                        If rsItem.EOF Then
                            rsItem.AddNew
                            rsItem!ItemID = i
                            rsItem!Amount = 0
                            rsItem!Desc = Mid(PRItem.Title & " Match", 1, 30)
                            rsItem!Amount = 0
                            rsItem.Update
                        End If
                        rsItem!Amount = rsItem!Amount + MatchAmt
                        rsItem.Update
                    End If
                End If
            
                If PRItemHist.GetNext = False Then Exit Do
            Loop
        End If
    
NxtPrHist:
        If PRHist.GetNext = False Then Exit Do
    
    Loop
    
End Sub

Private Sub rsUpdate()
    
    ' **************************************************************************
    ' add to main rs
    ' ?? sep PREquates for each ??
    If SSTax <> 0 Then
        
        ItemAdd PREquate.GlobalTypeQBPayFED, "Employee SS Tax", SSTax, "A1", 1
    
        ' Employer SS Tax by job
        ' !!!! 2011 - employer match is still 6.2%
        If TaxYear >= 2011 Then
            P1 = SSTax62
        Else
            P1 = SSTax
        End If
        
        i = 0
        rsJob.Filter = adFilterNone
        SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayFUN
        rsJob.Filter = SQLString
        If rsJob.RecordCount > 0 Then
            rsJob.MoveFirst
            Do
                i = i + 1
                If i = rsJob.RecordCount Then
                    P2 = P1
                Else
                    If TaxYear >= 2011 Then
                        P2 = Round(rsJob!Amount / TotalGross * SSTax62, 2)
                    Else
                        P2 = Round(rsJob!Amount / TotalGross * SSTax, 2)
                    End If
                    P1 = P1 - P2
                End If
                ItemAdd PREquate.GlobalTypeQBPayFED, "Employer SS Tax", P2, "A11", 2, rsJob!JobID
                rsJob.MoveNext
            Loop Until rsJob.EOF
        End If
        rsJob.Filter = adFilterNone
    
    End If
    
    If MedTax <> 0 Then
        
        ItemAdd PREquate.GlobalTypeQBPayFED, "Employee MED Tax", MedTax, "A2", 3
    
        ' Employer MED Tax by job
        P1 = MedTax
        i = 0
        rsJob.Filter = adFilterNone
        SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayFUN
        rsJob.Filter = SQLString
        If rsJob.RecordCount > 0 Then
            rsJob.MoveFirst
            Do
                i = i + 1
                If i = rsJob.RecordCount Then
                    P2 = P1
                Else
                    P2 = Round(rsJob!Amount / TotalGross * MedTax, 2)
                    P1 = P1 - P2
                End If
                ItemAdd PREquate.GlobalTypeQBPayFED, "Employer MED Tax", P2, "A21", 4, rsJob!JobID
                rsJob.MoveNext
            Loop Until rsJob.EOF
        End If
        rsJob.Filter = adFilterNone
    
    End If
    
    If FWTTax <> 0 Then
        ItemAdd PREquate.GlobalTypeQBPayFED, "FWT Tax", FWTTax, "A3", 5
    End If
    
    If FUNWage <> 0 Then
        
        P2 = 0
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFUNPct
        If PRGlobal.GetBySQL(SQLString) Then
            P2 = PRGlobal.Amount
        End If
        
        P1 = SuperRound(FUNWage, P2 / 100)
        P3 = P1
    
        ' split by job
        rsJob.Filter = adFilterNone
        SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayFUN
        rsJob.Filter = SQLString
        
        i = 0
        If rsJob.RecordCount > 0 Then
            rsJob.MoveFirst
            Do
                i = i + 1
                If i = rsJob.RecordCount Then
                    P2 = P3
                Else
                    P2 = Round(P1 * rsJob!Amount / TotalGross, 2)
                    P3 = P3 - P2
                End If
                ItemAdd PREquate.GlobalTypeQBPayFUN, "FUTA Tax", P2, "ZZ2", 0, rsJob!JobID
                rsJob.MoveNext
            Loop Until rsJob.EOF
        End If
    
    End If
    
    ' deductions from PRItemHist
    If rsItem.RecordCount > 0 Then
        rsItem.Sort = "ItemID"
        rsItem.MoveFirst
        Do
            ItemAdd PREquate.GlobalTypeQBPayItem, rsItem!Desc, rsItem!Amount, "Y", rsItem!ItemID
            rsItem.MoveNext
        Loop Until rsItem.EOF
    End If
    
    ' net pay by check used?
    If NetPay <> 0 Then
        ItemAdd PREquate.GlobalTypeQBPayNetPay, "Net Pay by Check", NetPay, "Z", 0
        ' force to due each pay
        PRGlobal.Byte1 = PREquate.PeriodTypePay
        PRGlobal.Save (Equate.RecPut)
        
        ' update the proper rsqb record 12/05/2010
        SQLString = "RecID = " & PRGlobal.GlobalID
        rsQB.Find SQLString, 0, adSearchForward, 1
        If rsQB.EOF = False Then
            rsQB!DuePeriod = PRGlobal.Byte1
            rsQB.Update
        End If
    
    End If
    
    If rsCity.RecordCount > 0 Then
    
        rsCity.MoveFirst
        
        Do
            
            ' get the city record
            If PRCity.GetByID(rsCity!RecID) = False Then
                X = "CWT: " & rsCity!RecID
            Else
                X = "CWT: " & PRCity.ShortName
            End If
            
            ItemAdd PREquate.GlobalTypeQBPayCity, X, rsCity!Amount, "G " & X, rsCity!RecID
            
            rsCity.MoveNext
        
        Loop Until rsCity.EOF
        
    End If
    
    If rsState.RecordCount > 0 Then
    
        rsState.Sort = "RecID"
        rsState.MoveFirst
        Do
            
            ' get the state record
            If PRState.GetByID(rsState!RecID) = False Then
                X = "SWT:" & rsState!RecID
                Y = "SUTA: " & rsState!RecID
            Else
                X = "SWT: " & PRState.StateName
                Y = "SUTA: " & PRState.StateName
            End If
            ItemAdd PREquate.GlobalTypeQBPayState, X, rsState!Amount, "H " & X, rsState!RecID
            
            ' state unemployment - by state / job
            rsJob.Filter = adFilterNone
            SQLString = "TypeCode = " & PREquate.GlobalTypeQBPaySUN & _
                        " AND RelatedID = " & rsState!RecID
            rsJob.Filter = SQLString
            If rsJob.RecordCount > 0 Then
                rsJob.MoveFirst
                Do
                    P1 = SuperRound(rsJob!Gross / rsState!Gross * rsState!SUNWage, PRCompany.StateUnempPct / 100)
                    ItemAdd PREquate.GlobalTypeQBPaySUN, Y, P1, "ZZ1", rsState!RecID, rsJob!JobID
                    rsJob.MoveNext
                Loop Until rsJob.EOF
            End If
    
            rsState.MoveNext
        
        Loop Until rsState.EOF
    
    End If
    
    ' add gross by dept \  dept/job
    If rsGross.RecordCount > 0 Then
    
        rsGross.Sort = "DepartmentID"
        rsJob.Sort = "TypeCode, RelatedID"
        rsGross.MoveFirst
        Do
            
            ' add gross by dept
            If PRDepartment.GetByID(rsGross!DepartmentID) = True Then
                X = "Gross: " & PRDepartment.Name
            Else
                X = "Gross: " & rsGross!DepartmentID
            End If
            
            ItemAdd PREquate.GlobalTypeQBPayGrossPay, X, rsGross!Amount, "#", rsGross!DepartmentID
        
            ' loop thru jobs of dept
            rsJob.Filter = adFilterNone
            SQLString = "TypeCode = " & PREquate.GlobalTypeQBPayGrossPay & _
                        " AND RelatedID = " & rsGross!DepartmentID
            rsJob.Filter = SQLString
            If rsJob.RecordCount > 0 Then
                rsJob.MoveFirst
                Do
                    rsQBDist.AddNew
                    rsQBDist!RecID = rsQBRecID
                    rsQBDist!JobID = rsJob!JobID
                    rsQBDist!Amount = rsJob!Amount
                    rsQBDist.AddNew
                    rsJob.MoveNext
                Loop Until rsJob.EOF
            End If
            rsJob.Filter = adFilterNone
            
            rsGross.MoveNext
        
        Loop Until rsGross.EOF
    
    End If

'    ' ************************************************************************
'    ' splits by Job
'
'    ' >>>>>>>>> FED SPLIT <<<<<<<<<<<< - optionally(?) split SS/MED/FWT
'    '  unemp - multi state - split amongst jobs no matter what state job is in?
'    '  if updating unemp tax per pay - rounding issues if report by quarter?
'
'    rsQB.MoveFirst
'    Do
'
'        ' >>>> SS & MED not FED <<<<
'
'        If rsQB!TypeCode = PREquate.GlobalTypeQBPayFED Or _
'           rsQB!TypeCode = PREquate.GlobalTypeQBPayFUN Or _
'           rsQB!TypeCode = PREquate.GlobalTypeQBPaySUN Then
'
'            i = 0
'            P2 = rsQB!Amount
'            rsJob.MoveFirst
'            Do
'
'                i = i + 1
'                If i = rsJob.RecordCount Then
'                    P1 = P2
'                Else
'                    P1 = Round(rsJob!Amount / TotalGross * rsQB!Amount, 2)  ' <<<< per gross per state???
'                    P2 = P2 - P1
'                End If
'
'                If JCJob.GetByID(rsJob!JobID) Then
'                    X = JCJob.FullName
'                Else
'                    X = JCJob.JobID
'                End If
'
'                rsQBDist.AddNew
'                rsQBDist!Desc = Mid(X & " " & rsQB!Desc, 1, 50)
'                rsQBDist!QBPayTo = rsQB!QBPayTo
'                rsQBDist!QBExpenseAcct = rsQB!QBExpenseAcct
'                rsQBDist!QBCheckingAcct = rsQB!QBCheckingAcct
'                rsQBDist!DuePeriod = rsQB!DuePeriod
'                ' rsqbdist!DueDate = ....
'                rsQBDist!PayType = rsQB!PayType
'                rsQBDist!Amount = P1
'
'                ' rsQBDist!DeptID = rsJob!JobID
'
'                rsQBDist!JobID = 0
'                rsQBDist!SortOrder = rsQB!SortOrder
'                rsQBDist!TypeCode = rsQB!TypeCode
'                rsQBDist.Update
'
'                rsJob.MoveNext
'
'            Loop Until rsJob.EOF
'
'        Else
'
'            rsQBDist.AddNew
'            rsQBDist!Desc = rsQB!Desc
'            rsQBDist!QBPayTo = rsQB!QBPayTo
'            rsQBDist!QBExpenseAcct = rsQB!QBExpenseAcct
'            rsQBDist!QBCheckingAcct = rsQB!QBCheckingAcct
'            rsQBDist!DuePeriod = rsQB!DuePeriod
'            ' rsqbdist!DueDate = ....
'            rsQBDist!PayType = rsQB!PayType
'            rsQBDist!Amount = rsQB!Amount
'
'            ' rsQBDist!DeptID = rsQB!DeptID
'
'            rsQBDist!JobID = 0
'            rsQBDist!SortOrder = rsQB!SortOrder
'            rsQBDist!TypeCode = rsQB!TypeCode
'            rsQBDist.Update
'
'        End If
'
'        rsQB.MoveNext
'
'    Loop Until rsQB.EOF

    rsQB.Sort = "SortOrder"
    
    fgDistDisplay

End Sub

Private Sub ItemAdd(ByVal TypeCode As Long, _
                    ByVal Desc As String, _
                    ByVal Amount As Currency, _
                    ByVal SortOrder As String, _
                    ByVal RelatedID As Long, _
                    Optional JobID As Long)
                    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & TypeCode & _
                " AND UserID = " & PRCompany.CompanyID
    
    If RelatedID <> 0 Then
        SQLString = SQLString & " AND Var4 = '" & RelatedID & "'"
    End If
    
    If PRGlobal.GetBySQL(SQLString) = False Then
        
        PRGlobal.Clear
        PRGlobal.TypeCode = TypeCode
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Var4 = RelatedID
        
        ' used to designate matching amounts for 401k deductions
        If RelatedID > 1000000 Then
            PRGlobal.Var5 = 1
        End If
        
        PRGlobal.Byte2 = 1          ' set the select flag !!!
        PRGlobal.Save (Equate.RecAdd)
    
        ' notify user of new city
        If TypeCode = PREquate.GlobalTypeQBPayCity Then
            boo = PRCity.GetByID(RelatedID)
            X = "New city added: " & PRCity.CityName & vbCr & vbCr & _
                "Be sure to assign the QB accounts"
            MsgBox X, vbOKOnly + vbInformation
        End If
    
        ' notify user of new city
        If TypeCode = PREquate.GlobalTypeQBPayItem Then
            boo = PRItem.GetByID(RelatedID)
            X = "New deduction added: " & PRItem.Title & vbCr & vbCr & _
                "Be sure to assign the QB accounts"
            MsgBox X, vbOKOnly + vbInformation
        End If
    
    End If
    GlobalNet = PRGlobal.GlobalID
    
    rsQB.Filter = adFilterNone
    SQLString = "TypeCode = " & TypeCode & " AND RelatedID = " & RelatedID
    rsQB.Filter = SQLString
    If rsQB.RecordCount = 0 Then
        rsQB.AddNew
        
        rsQB!Select = PRGlobal.Byte2
        
        ' 02/23/2011 - always select each item!!!
        rsQB!Select = 1
        
        rsQB!Desc = Desc
        rsQB!QBPayTo = PRGlobal.Var1
        rsQB!QBExpenseAcct = PRGlobal.Var2
        rsQB!QBLiabilityAcct = PRGlobal.Var5
        rsQB!DuePeriod = PRGlobal.Byte1
        If PRGlobal.Byte4 <> 0 Then
            rsQB!DueDays = PRGlobal.Byte4
        End If
        rsQB!PayType = PRGlobal.Byte3
        rsQB!GlobalID = PRGlobal.GlobalID
        rsQB!TypeCode = TypeCode
        rsQB!SortOrder = SortOrder
        rsQB!RelatedID = RelatedID
        rsQBCount = rsQBCount + 1
        rsQB!RecID = rsQBCount
        rsQB.Update
    End If
    
    rsQBRecID = rsQB!RecID
    rsQB!Amount = rsQB!Amount + Amount
    rsQB.Update
    
    rsQB.Filter = adFilterNone

    ' by job?
    If JobID <> 0 Then
        rsQBDist.Filter = adFilterNone
        SQLString = " JobID = " & JobID & " AND RecID = " & rsQB!RecID
        rsQBDist.Filter = SQLString
        If rsQBDist.RecordCount = 0 Then
            rsQBDist.AddNew
            rsQBDist!JobID = JobID
            rsQBDist!RecID = rsQBRecID
            rsQBDist.Update
        End If
        rsQBDist!Amount = rsQBDist!Amount + Amount
        rsQBDist.Update
    
        rsQBDist.Filter = adFilterNone
    
    End If

End Sub
Private Sub fgDistDisplay()

    If LoadFlag = True Then Exit Sub

    rsQBDist.Filter = adFilterNone
    SQLString = "RecID = " & rsQB!RecID
    rsQBDist.Filter = SQLString

    With fgDist
        .AutoSize 0, .Cols - 1, False, 200
    End With

    If rsQBDist.RecordCount = 0 Or Me.optAmounts = False Then
        fgDist.Visible = False
    Else
        fgDist.Visible = True
    End If

End Sub
                    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub ColDisplay()

    ' display the appropriate columns

Dim CD As Long

    With fg
        
        For CD = 0 To .Cols - 1
            
            If Me.optAmounts = True Then
                .ColHidden(.ColIndex("Select")) = False
                .ColHidden(.ColIndex("QBPayTo")) = True
                .ColHidden(.ColIndex("QBExpenseAcct")) = True
                .ColHidden(.ColIndex("QBLiabilityAcct")) = True
                ' .ColHidden(.ColIndex("DuePeriod")) = False
                .ColHidden(.ColIndex("DueDays")) = False
                .ColHidden(.ColIndex("DueDate")) = False
                ' .ColHidden(.ColIndex("PayType")) = False
                .ColHidden(.ColIndex("Amount")) = False
                If rsQBDist.RecordCount > 0 Then
                    fgDist.Visible = True
                Else
                    fgDist.Visible = False
                End If
            Else
                .ColHidden(.ColIndex("Select")) = True
                .ColHidden(.ColIndex("QBPayTo")) = False
                .ColHidden(.ColIndex("QBExpenseAcct")) = False
                .ColHidden(.ColIndex("QBLiabilityAcct")) = False
                ' .ColHidden(.ColIndex("DuePeriod")) = True
                .ColHidden(.ColIndex("DueDays")) = True
                .ColHidden(.ColIndex("DueDate")) = True
                ' .ColHidden(.ColIndex("PayType")) = True
                .ColHidden(.ColIndex("Amount")) = True
                fgDist.Visible = False
            End If
        
        Next CD
    
    End With

End Sub

Private Function StrValue(ByVal Str As String) As Long

    StrValue = 0
    If IsNull(Str) Then Exit Function
    If Str = "" Then Exit Function
    On Error Resume Next
    StrValue = CLng(Str)
    On Error GoTo 0

End Function


Private Sub LoadQBAccounts()

    QBPayeeDrop = "|#0;N/A"
    QBExpAcctDrop = QBPayeeDrop
    QBLiabAcctDrop = QBPayeeDrop

    SQLString = "SELECT * FROM QBAccount ORDER BY Name "
    If QBAccount.GetBySQL(SQLString) = False Then Exit Sub
    
    Me.cmbQBAP.Clear
    Me.cmbQBChk.Clear
    
    Do
        
        If QBAccount.AccountType = "AccountsPayable" Then
            With Me.cmbQBAP
                .AddItem QBAccount.Name
                .ItemData(.NewIndex) = QBAccount.QBAccountID
            End With
        ElseIf QBAccount.AccountType = "Bank" Then
            With Me.cmbQBChk
                .AddItem QBAccount.Name
                .ItemData(.NewIndex) = QBAccount.QBAccountID
            End With
            QBLiabAcctDrop = QBLiabAcctDrop & "|#" & QBAccount.QBID & ";" & QBAccount.Name
            QBExpAcctDrop = QBExpAcctDrop & "|#" & QBAccount.QBID & ";" & QBAccount.Name
        ElseIf QBAccount.AccountType = "VENDOR" Then
            QBPayeeDrop = QBPayeeDrop & "|#" & QBAccount.QBID & ";" & QBAccount.Name
        Else
            QBLiabAcctDrop = QBLiabAcctDrop & "|#" & QBAccount.QBID & ";" & QBAccount.Name
            QBExpAcctDrop = QBExpAcctDrop & "|#" & QBAccount.QBID & ";" & QBAccount.Name
        End If
            
        If QBAccount.GetNext = False Then Exit Do
    
    Loop

    JobDrop = "|#0; |#999999999;" & PRCompany.Name
    SQLString = "SELECT * FROM JCJob"
    If JCJob.GetBySQL(SQLString) = True Then
        Do
            JobDrop = JobDrop & "|#" & JCJob.JobID & ";" & JCJob.Name
            If JCJob.GetNext = False Then Exit Do
        Loop
    End If

    ' init Combos
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBPayCompany & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeQBPayCompany
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    CompanyGlobalID = PRGlobal.GlobalID
    
    QBAPAcct = StrValue(PRGlobal.Var1)
    SetCombo Me.cmbQBAP, QBAPAcct
    
    QBCheckingAcct = StrValue(PRGlobal.Var2)
    SetCombo Me.cmbQBChk, QBCheckingAcct

    Me.chkNoName = PRGlobal.Byte1

    With Me.fg
        .ColComboList(.ColIndex("QBPayTo")) = QBPayeeDrop
        .ColComboList(.ColIndex("QBExpenseAcct")) = QBExpAcctDrop
        .ColComboList(.ColIndex("QBLiabilityAcct")) = QBLiabAcctDrop
    End With

End Sub

Private Sub SetCombo(ByRef cmb As ComboBox, ByVal QBAcctID As Long)

    With cmb
        If .ListCount > 0 Then
            For i = 0 To .ListCount - 1
                If .ItemData(i) = QBAcctID Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End With

End Sub
Private Sub cmbQBAP_Click()
    If PRGlobal.GetByID(CompanyGlobalID) = True Then
        If Me.cmbQBAP.ListIndex >= 0 Then
            PRGlobal.Var1 = Me.cmbQBAP.ItemData(Me.cmbQBAP.ListIndex)
        Else
            PRGlobal.Var1 = ""
        End If
        PRGlobal.Save (Equate.RecPut)
    End If
End Sub
Private Sub cmbQBChk_Click()
    If PRGlobal.GetByID(CompanyGlobalID) = True Then
        If Me.cmbQBChk.ListIndex >= 0 Then
            PRGlobal.Var2 = Me.cmbQBChk.ItemData(Me.cmbQBChk.ListIndex)
        Else
            PRGlobal.Var2 = ""
        End If
        PRGlobal.Save (Equate.RecPut)
    End If
End Sub
Private Sub chkNoName_Click()
    If PRGlobal.GetByID(CompanyGlobalID) = True Then
        PRGlobal.Byte1 = Me.chkNoName
        PRGlobal.Save (Equate.RecPut)
    End If
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With fg
        
        ' Journal Entry
        If Col = .ColIndex("DuePeriod") Then
            If .TextMatrix(Row, Col) = "0" Then
                .TextMatrix(Row, .ColIndex("DueDays")) = ""
                .TextMatrix(Row, .ColIndex("DueDate")) = ""
                .TextMatrix(Row, .ColIndex("QBPayTo")) = "0"
            Else                ' payable
                .TextMatrix(Row, .ColIndex("QBLiabilityAcct")) = "0"
            End If
        End If
        
        UpdateGlobal
    
    End With

End Sub

Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    UpdateGlobal
    fgDistDisplay

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With fg
        
        ' description entries can not be changes
        If Col = .ColIndex("Desc") Then
            Cancel = True
            Exit Sub
        End If
        If Col = .ColIndex("Amount") Then
            Cancel = True
            Exit Sub
        End If
        
        ' DuePeriod - none - don't allow DueDays/DueDate/PayTo edits
        If .TextMatrix(Row, .ColIndex("Desc")) = "Net Pay by Check" Or _
           .TextMatrix(Row, .ColIndex("Desc")) = "Direct Deposit" Then
            ' Net Pay MUST be each pay
            Cancel = True
            If Col = .ColIndex("QBExpenseAcct") Then
                Cancel = False
            ElseIf Col = .ColIndex("QBPayTo") Then
                Cancel = False
            End If
        ElseIf .TextMatrix(Row, .ColIndex("DuePeriod")) = "0" Then
            If Col = .ColIndex("DueDays") Then Cancel = True
            If Col = .ColIndex("DueDate") Then Cancel = True
            If Col = .ColIndex("QBPayTo") Then Cancel = True
        Else
            If Col = .ColIndex("QBLiabilityAcct") Then Cancel = True
        End If
        
    End With

End Sub

Private Sub UpdateGlobal()
    
    ' /////////////
    If LoadFlag = True Then Exit Sub
    
    ' update to PRGlobal
    If PRGlobal.GetByID(rsQB!GlobalID) = False Then
        MsgBox "PRGlobal Not Found: " & rsQB!GlobalID, vbExclamation
        GoBack
    End If

    PRGlobal.Byte1 = nNull(rsQB!DuePeriod)
    If rsQB!Select = True Then
        PRGlobal.Byte2 = 1
    Else
        PRGlobal.Byte2 = 0
    End If
    PRGlobal.Byte3 = nNull(rsQB!PayType)
    PRGlobal.Byte4 = CByte(StrValue(nNull(rsQB!DueDays)))
    PRGlobal.Var1 = rsQB!QBPayTo & ""
    PRGlobal.Var2 = rsQB!QBExpenseAcct & ""
    PRGlobal.Var5 = rsQB!QBLiabilityAcct & ""
    PRGlobal.Save (Equate.RecPut)

    SetDueDate

End Sub
Private Sub SetDueDate()
    
Dim DueDays As Long
    
    ' calc due date if necessary
    With fg
        
        .TextMatrix(.Row, .ColIndex("DueDate")) = ""
        DueDays = StrValue(.TextMatrix(.Row, .ColIndex("DueDays")))
        
        ' PayType not being used
        ' If rsQB!PayType <> 0 And rsQB!DuePeriod <> 0 Then
        If rsQB!DuePeriod <> 0 Then
            
            If rsQB!DuePeriod = PREquate.PeriodTypePay Then
                D2 = HiCheckDate
            ElseIf rsQB!DuePeriod = PREquate.PeriodTypeMonth Then
                D2 = DateSerial(Year(HiCheckDate), Month(HiCheckDate) + 1, 1) - 1
            ElseIf rsQB!DuePeriod = PREquate.PeriodTypeQuarter Then
                Select Case PRBatch.YearMonth Mod 100
                    Case Is <= 3:   D2 = DateSerial(Year(HiCheckDate), 4, 1) - 1
                    Case Is <= 6:   D2 = DateSerial(Year(HiCheckDate), 7, 1) - 1
                    Case Is <= 9:   D2 = DateSerial(Year(HiCheckDate), 10, 1) - 1
                    Case Is <= 12:  D2 = DateSerial(Year(HiCheckDate) + 1, 1, 1) - 1
                End Select
            ElseIf rsQB!DuePeriod = PREquate.PeriodTypeYear Then
                D2 = DateSerial(Year(HiCheckDate) + 1, 1, 1) - 1
            End If
            .TextMatrix(.Row, .ColIndex("DueDate")) = Format(D2 + DueDays, "mm/dd/yyyy")
        End If
    End With

End Sub

Private Sub optAmounts_Click()
    ColDisplay
End Sub

Private Sub optQBSetup_Click()
    ColDisplay
End Sub

Private Sub cmdClearAll_Click()
    If rsQB.RecordCount = 0 Then Exit Sub
    rsQB.MoveFirst
    Do
        rsQB!Select = False
        rsQB.Update
        UpdateGlobal
        rsQB.MoveNext
    Loop Until rsQB.EOF
    rsQB.MoveFirst
End Sub

Private Sub cmdSelectAll_Click()
    If rsQB.RecordCount = 0 Then Exit Sub
    rsQB.MoveFirst
    Do
        rsQB!Select = True
        rsQB.Update
        UpdateGlobal
        rsQB.MoveNext
    Loop Until rsQB.EOF
    rsQB.MoveFirst
End Sub

Private Sub cmdQBRefresh_Click()
    frmQBAccts.Show vbModal
    LoadQBAccounts
End Sub
Private Sub cmdPrint_Click()

    PrtInit ("Land")    ' "Port" = Portrait
    SetFont 9, Equate.LandScape
    LandSw = 1
    Ln = 0

    rsQB.MoveFirst

    Do

        If Ln = 0 Or Ln > MaxLines Then
            PrintHeader
        End If
        
        If rsQB!Select = True Then

            PrintValue(1) = rsQB!Desc:          FormatString(1) = "a30"
            PrintValue(2) = " ":                FormatString(2) = "a3"

            ' Due Period
            Select Case rsQB!DuePeriod
                Case 0:                             X = "Jnl Entry"
                Case PREquate.PeriodTypePay:        X = "Pay"
                Case PREquate.PeriodTypeMonth:      X = "Monthly"
                Case PREquate.PeriodTypeQuarter:    X = "Quarterly"
                Case PREquate.PeriodTypeYear:       X = "Annually"
                Case Else:                          X = ""
            End Select
            PrintValue(3) = X:                  FormatString(3) = "a10"
            
            ' due date
            If rsQB!DueDate = 0 Then
                X = ""
            Else
                X = Format(rsQB!DueDate, " mm/dd/yyyy ")
            End If
            PrintValue(4) = X:                      FormatString(4) = "a12"
            
            PrintValue(5) = rsQB!Amount:            FormatString(5) = "d12"
            
            PrintValue(6) = " ":                    FormatString(6) = "a3"
            
            If rsQB!QBExpenseAcct <> 0 Then
                boo = QBAccount.GetByQBID(rsQB!QBExpenseAcct)
                X = QBAccount.Name
            Else
                X = ""
            End If
            PrintValue(7) = X:                      FormatString(7) = "a20"
            PrintValue(8) = " ":                    FormatString(8) = "a2"
            
            If rsQB!QBLiabilityAcct <> 0 Then
                boo = QBAccount.GetByQBID(rsQB!QBLiabilityAcct)
                X = QBAccount.Name
            ElseIf rsQB!QBPayTo <> 0 Then
                X = Me.cmbQBAP.Text
            Else
                X = ""
            End If
            PrintValue(9) = X:                       FormatString(9) = "a20"
            PrintValue(10) = " ":                    FormatString(10) = "a2"
            
            If rsQB!QBPayTo <> 0 Then
                boo = QBAccount.GetByQBID(rsQB!QBPayTo)
                X = QBAccount.Name
            Else
                X = ""
            End If
            PrintValue(11) = X:                     FormatString(11) = "a20"
            PrintValue(12) = " ":                   FormatString(12) = "a2"
                                                    
            PrintValue(13) = " ":                   FormatString(13) = "~"
            FormatPrint
            Ln = Ln + 1

            ' job distn ?
            rsQBDist.Filter = adFilterNone
            rsQBDist.Filter = "RecID = " & rsQB!RecID
            If rsQBDist.RecordCount > 0 Then
                rsQBDist.MoveFirst
                Do
                    If rsQBDist!JobID = 999999999 Then
                        X = PRCompany.Name
                    Else
                        boo = JCJob.GetByID(rsQBDist!JobID)
                        X = JCJob.FullName
                    End If
                    PrintValue(1) = X:                      FormatString(1) = "a55"
                    PrintValue(2) = " ":                    FormatString(2) = "a2"
                    PrintValue(3) = rsQBDist!Amount:        FormatString(3) = "d10"
                    PrintValue(4) = " ":                    FormatString(4) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    If Ln > MaxLines Then PrintHeader
                    rsQBDist.MoveNext
                Loop Until rsQBDist.EOF
            End If
            Ln = Ln + 1
            If Ln > MaxLines Then PrintHeader

        End If

        rsQB.MoveNext

    Loop Until rsQB.EOF
    rsQB.MoveFirst
    PrvwReturn = True
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub PrintHeader()
            
    If Ln <> 0 Then FormFeed
    X = "Payroll update to QuickBooks"
    Y = "A/P Account: " & Me.cmbQBAP.Text
    PageHeader X, Me.lblBatchInfo, Y
    Ln = Ln + 1
    
    PrintValue(1) = "Item Description":     FormatString(1) = "a33"
    PrintValue(2) = "Due Period":           FormatString(2) = "a10"
    PrintValue(3) = "  Due Date":           FormatString(3) = "a12"
    PrintValue(4) = "Amount ":              FormatString(4) = "r12"
    PrintValue(5) = " ":                    FormatString(5) = "a3"
    PrintValue(6) = "QB Expense Acct":      FormatString(6) = "a20"
    PrintValue(7) = " ":                    FormatString(7) = "a2"
    PrintValue(8) = "QB Liability Acct":    FormatString(8) = "a20"
    PrintValue(9) = " ":                    FormatString(9) = "a2"
    PrintValue(10) = "QB Vendor":           FormatString(10) = "a20"
    PrintValue(11) = " ":                   FormatString(11) = "~"
    FormatPrint
    Ln = Ln + 2

End Sub

Private Sub cmdPay_Click()

Dim qbFlag As Boolean
Dim SelAll As Boolean

    If MsgBox("OK to proceed with the QuickBooks Update?", vbQuestion + vbYesNo, _
        "Payroll to QuickBooks Update") = vbNo Then Exit Sub

    ' >>>> verify all entries are set properly <<<<
    
    ' udpate flag override warning
    If frmDateRange.chkQBOverride = 1 Then
        If MsgBox("OK to RE-UPDATE history records?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
    End If
    
    QBCount = 0
    qbFlag = False
    LoadFlag = False
    
    If rsQB.RecordCount = 0 Then Exit Sub
  
    LoadFlag = True
    
    ' verify all settings are complete!
    
    ' give warning if any entries are not checked
    SelAll = True
    
    rsQB.MoveFirst
    Do
        
        If rsQB!Select = False Then SelAll = False
        
        ' QBExpense Acct must always be filled in
        If rsQB!QBExpenseAcct = "" Or rsQB!QBExpenseAcct = "0" Then
            MsgBox "QB Expense Account not assigned for: " & vbCr & _
                   rsQB!Desc, vbExclamation
            Me.optQBSetup = True
            Exit Sub
        End If
        
        If rsQB!DuePeriod = 0 Then
            
            If rsQB!QBLiabilityAcct = "" Or rsQB!QBLiabilityAcct = "0" Then
                MsgBox "QB Liability Account not assigned for: " & vbCr & _
                       rsQB!Desc, vbExclamation
                Me.optQBSetup = True
                Exit Sub
            End If
        
        Else
            
            If rsQB!QBPayTo = "" Or rsQB!QBPayTo = "0" Then
                MsgBox "QB Vendor not assigned for: " & vbCr & _
                       rsQB!Desc, vbExclamation
                       Me.optQBSetup = True
                Exit Sub
            End If
        
        End If
        
        ' extra warning for net pay update
        If rsQB!TypeCode = PREquate.GlobalTypeQBPayNetPay And rsQB!Select = False Then
            i = MsgBox("Net Pay Update not enabled! OK to turn it on?", vbCritical + vbYesNo)
            If i = vbYes Then
                rsQB!Select = 1
                rsQB.Update
            End If
        End If
        
        rsQB.MoveNext
    Loop Until rsQB.EOF
    
    If SelAll = False Then
        X = "All entries are not selected for update to QuickBooks"
        X = X & vbCr & vbCr & "OK to continue?"
        If MsgBox(X, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    
    ' this button disabled if amount view not shown?
    ' loop thru rsQB - make sure all QB info is filled in
  
    ' ********************************************
    ' *** Update setttings for all lines ***
    rsQB.MoveFirst
    Do
        UpdateGlobal
        rsQB.MoveNext
    Loop Until rsQB.EOF
  
    ' =====================================================================
    
    ' start session and open connection
    If QBOpen(Me, Me.lblMsg1) = False Then GoBack
    
    Me.lblMsg1 = "QBOpen Complete ..."
    
    ' ================================================================
    
    Me.MousePointer = vbHourglass
    
    Me.fg.Visible = False
    Me.fgDist.Visible = False
    
    ' Create the message set request object for the specific version messages.
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
  
    ' net pay update
    rsQB.Filter = adFilterNone
    SQLString = "Select = True AND TypeCode = " & PREquate.GlobalTypeQBPayNetPay
    rsQB.Filter = SQLString
    If rsQB.RecordCount > 0 Then
        qbFlag = True
        rsQB.MoveFirst
        Do

            Me.lblMsg1 = "Updating Net Pay to QuickBooks: " & rsQB!Desc
            Me.Refresh
                    
            boo = QBAccount.GetByID(Me.cmbQBChk.ItemData(Me.cmbQBChk.ListIndex))

            ' loop thru the dates selected
            If BatchNumbr > 0 Then
                
                If PRBatch.GetByID(BatchNumbr) = False Then
                    MsgBox "PR Batch Not Found: " & BatchNumbr, vbExclamation
                    GoBack
                End If
                    
                DirDepTotal = 0
                
                QBCheckAddRq "US", 5, 0, PRBatch.BatchID, _
                             QBAccount.QBID, "", _
                             rsQB!QBExpenseAcct, "", _
                             rsQB!QBPayTo, ""
            
                If DirDepTotal > 0 Then
                    ' batchid = 0 triggers dir deposit amount update
                    QBCheckAddRq "US", 5, 0, 0, _
                                 QBAccount.QBID, "", _
                                 rsQB!QBExpenseAcct, "", _
                                 rsQB!QBPayTo, ""
                
                End If
            
            Else
                
                If OptDate = "CHECK DATE" Then
                    SQLString = "SELECT * FROM PRBatch WHERE CheckDate >= " & CLng(StartDate) & _
                                " AND CheckDate <= " & CLng(EndDate) & _
                                " ORDER BY CheckDate"
                Else
                    SQLString = "SELECT * FROM PRBatch WHERE PEDate >= " & CLng(StartDate) & _
                                " AND PEDate <= " & CLng(EndDate) & _
                                " ORDER BY CheckDate"
                End If
                If PRBatch.GetBySQL(SQLString) = False Then
                    MsgBox "No batches found in the date range", vbInformation
                    GoBack
                End If
                Do
                    
                    DirDepTotal = 0
                    
                    QBCheckAddRq "US", 5, 0, PRBatch.BatchID, _
                                 QBAccount.QBID, "", _
                                 rsQB!QBExpenseAcct, "", _
                                 rsQB!QBPayTo, ""
                    
                    If DirDepTotal > 0 Then
                        ' batchid = 0 triggers dir deposit amount update
                        QBCheckAddRq "US", 5, 0, 0, _
                                     QBAccount.QBID, "", _
                                     rsQB!QBExpenseAcct, "", _
                                     rsQB!QBPayTo, ""
                    
                    End If
                    
                    If PRBatch.GetNext = False Then Exit Do
                Loop
            End If

            rsQB.MoveNext
        Loop Until rsQB.EOF
    End If
    
    Me.lblMsg1 = "Net Pay Update Complete ..."
  
    ' update the payables
    ' don't include direct deposit
    rsQB.Filter = adFilterNone
    SQLString = "Select = True AND DuePeriod <> 0 AND TypeCode <> " & PREquate.GlobalTypeQBPayNetPay & _
                " AND Desc <> 'Direct Deposit'"
    rsQB.Filter = SQLString
    If rsQB.RecordCount > 0 Then
        qbFlag = True
        rsQB.MoveFirst
        Do

            Me.lblMsg1 = "Updating Payables to QuickBooks: " & rsQB!Desc
            Me.Refresh

            If rsQB!Desc = "Direct Deposit" Then
                
                ' direct deposit - amount to the check reg
                boo = QBAccount.GetByID(Me.cmbQBChk.ItemData(Me.cmbQBChk.ListIndex))
                QBCheckAddRq "US", 5, 0, 0, _
                             QBAccount.QBID, "", _
                             rsQB!QBExpenseAcct, "", _
                             rsQB!QBPayTo, ""
                
            Else
                ' payables
                QBBuildBillAddRq requestMsgSet, "US"
            End If

            rsQB.MoveNext
        Loop Until rsQB.EOF

        Me.lblMsg1 = "Now processing Payable requests ......"
        Me.Refresh

        ' ??? do requests separate - parse results .... ???
        ' Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
        ' ParseBillAddRs responseMsgSet, "US"

    End If

    Me.lblMsg1 = "Payables Update Complete ..."
    
    ' update the journal entries
    ' update the payables
    rsQB.Filter = adFilterNone
    SQLString = "Select = True AND DuePeriod = 0 AND TypeCode <> " & PREquate.GlobalTypeQBPayNetPay
    rsQB.Filter = SQLString
    If rsQB.RecordCount > 0 Then
        qbFlag = True
        rsQB.MoveFirst
        Do

            Me.lblMsg1 = "Updating Journal Entries to QuickBooks: " & rsQB!Desc
            Me.Refresh

            If rsQB!Amount > 0 Then
                QBBuildJournalAddRq requestMsgSet, "US"
            Else
                negQBBuildJournalAddRq requestMsgSet, "US"
            End If
            
            QBCount = QBCount + 1

            rsQB.MoveNext

        Loop Until rsQB.EOF

        Me.lblMsg1 = "Now processing Payable requests ......"
        Me.Refresh

        ' parse results .....

    End If
    
    Me.lblMsg1 = "Jnl Update Complete ..."
    
    rsQB.Filter = adFilterNone
    
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    ' Perform the request and obtain a response from QuickBooks.
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
  
'    ' ****** parse jnl entry results ***********
'    If ResponseSet Is Nothing Then
'        MsgBox "ResponseSet = nothing "
'        End
'    End If
'
    Set ResponseList = responseMsgSet.ResponseList
'    If (ResponseList Is Nothing) Then
'        MsgBox "ResponseList = nothing "
'        End
'    End If
    
    For i = 0 To ResponseList.Count - 1
        
        Me.lblMsg1 = "Response Process " & i
    
        Set Response = ResponseList.GetAt(i)
 
        ' Check the status returned for the response.
        If (Response.StatusCode = 0) Then
 
'            ' Check to make sure the response is of the type we are expecting.
'            If (Not Response.Detail Is Nothing) Then
'                Dim ResponseType As Integer
'                ResponseType = Response.Type.GetValue
'                Dim j As Integer
'                ' Check for JournalEntryAddRs.
'                If (ResponseType = rtJournalEntryAddRs) Then
''                    Dim journalEntryRet As IJournalEntryRet
''                    Set journalEntryRet = response.Detail
''                    ParseJournalEntryRet journalEntryRet, country
'                End If
'
'            End If
        Else
            If Response.StatusSeverity <> "Warn" Then
                MsgBox Response.StatusCode & vbCr & Response.StatusMessage & vbCr & Response.StatusSeverity & _
                       vbCr & Response.Type.GetAsString & vbCr & i
            End If
        End If
    
    Next i
  
    ' ****** parse jnl entry results ***********
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
    '  update the PRHist.QBUpdateFlag
    If rsPRHist.RecordCount > 0 Then
        rsPRHist.MoveFirst
        Do
            If PRHist.GetByID(rsPRHist!HistID) = False Then
                MsgBox "HistID record nf: " & PRHist!HistID, vbExclamation
                GoBack
            End If
            PRHist.QBUpdateFlag = 1
            PRHist.Save (Equate.RecPut)
            rsPRHist.MoveNext
        Loop Until rsPRHist.EOF
    End If
  
    MsgBox Format(QBCount, "###,##0") & _
           " Tax Payments have been updated to QuickBooks", vbInformation
  
    GoBack
  
    ' ParseCheckAddRs responseMsgSet, country
  
    Exit Sub
  
Errs:
    MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error"
  
    ' SampleCodeForm.ErrorMsg.Text = Err.Description
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
    Me.MousePointer = vbArrow

End Sub
Private Sub negQBBuildJournalAddRq(requestMsgSet As IMsgSetRequest, country As String)
    
Dim LineCount As Long
    
    Dim QBAmount As Currency
    
    Set JnlAddReq = requestMsgSet.AppendJournalEntryAddRq
    
    JnlAddReq.TxnDate.SetValue PRBatch.CheckDate
    JnlAddReq.RefNumber.SetValue "PR"
    JnlAddReq.Memo.SetValue Mid(rsQB!Desc, 1, 30)
    ' JnlAddReq.IsAdjustment.SetValue True
    
'    ' get the QB Job
'    If JCJob.GetByID(976) = False Then
'        MsgBox "Job NF"
'        End
'    End If
'
'    ' get the DR account - Misc Exp
'    If QBAccount.GetByID(1272) = False Then
'        MsgBox "Acct NF"
'        End
'    End If
    
    ' debit the expense by job if necessary
    rsQBDist.Filter = adFilterNone
    rsQBDist.Filter = "RecID = " & rsQB!RecID
    If rsQBDist.RecordCount = 0 Then
        k = 1
    Else
        k = rsQBDist.RecordCount
    End If
    
    For j = 1 To k
    
        Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
        orJournalLine1.JournalCreditLine.TxnLineID.SetValue j
        ' orJournalLine1.JournalcreditLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
        
        orJournalLine1.JournalCreditLine.AccountRef.ListID.SetValue Trim(rsQB!QBExpenseAcct)
        orJournalLine1.JournalCreditLine.Memo.SetValue rsQB!Desc
        
        If rsQBDist.RecordCount > 0 Then
            
            boo = JCJob.GetByID(rsQBDist!JobID)
            If JCJob.QBID = "ORIG" Then
                X = JCJob.QBParentID
            Else
                X = JCJob.QBID
            End If
            
            QBAmount = rsQBDist!Amount * (-1)
            orJournalLine1.JournalCreditLine.Amount.SetValue QBAmount
            
            If rsQBDist!JobID = 999999999 Then
            Else
                orJournalLine1.JournalCreditLine.EntityRef.ListID.SetValue Trim(X)
            End If
        
        Else
            
            QBAmount = rsQB!Amount * (-1)
            orJournalLine1.JournalCreditLine.Amount.SetValue QBAmount
            ' orJournalLine1.JournalcreditLine.EntityRef.ListID.SetValue Trim(JCJob.QBParentID)
        End If
    
        If rsQBDist.RecordCount > 0 Then
            rsQBDist.MoveNext
        End If
    
    Next j
    
    ' debit info
    Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
    orJournalLine1.JournalDebitLine.TxnLineID.SetValue k + 1
    ' orJournalLine1.JournaldebitLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
    
    orJournalLine1.JournalDebitLine.AccountRef.ListID.SetValue Trim(rsQB!QBLiabilityAcct)
    
    QBAmount = rsQB!Amount * (-1)
    orJournalLine1.JournalDebitLine.Amount.SetValue QBAmount
    
    orJournalLine1.JournalDebitLine.Memo.SetValue rsQB!Desc
    ' orJournalLine1.JournaldebitLine.EntityRef.ListID.SetValue Trim(JCJob.QBID)

End Sub

Private Sub QBBuildJournalAddRq(requestMsgSet As IMsgSetRequest, country As String)
    
Dim LineCount As Long
    
    Set JnlAddReq = requestMsgSet.AppendJournalEntryAddRq
    
    JnlAddReq.TxnDate.SetValue PRBatch.CheckDate
    JnlAddReq.RefNumber.SetValue "PR"
    JnlAddReq.Memo.SetValue Mid(rsQB!Desc, 1, 30)
    ' JnlAddReq.IsAdjustment.SetValue True
    
'    ' get the QB Job
'    If JCJob.GetByID(976) = False Then
'        MsgBox "Job NF"
'        End
'    End If
'
'    ' get the DR account - Misc Exp
'    If QBAccount.GetByID(1272) = False Then
'        MsgBox "Acct NF"
'        End
'    End If
    
    ' debit the expense by job if necessary
    rsQBDist.Filter = adFilterNone
    rsQBDist.Filter = "RecID = " & rsQB!RecID
    If rsQBDist.RecordCount = 0 Then
        k = 1
    Else
        k = rsQBDist.RecordCount
    End If
    
    For j = 1 To k
    
        Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
        orJournalLine1.JournalDebitLine.TxnLineID.SetValue j
        ' orJournalLine1.JournalDebitLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
        
        orJournalLine1.JournalDebitLine.AccountRef.ListID.SetValue Trim(rsQB!QBExpenseAcct)
        orJournalLine1.JournalDebitLine.Memo.SetValue rsQB!Desc
        
        If rsQBDist.RecordCount > 0 Then
            
            boo = JCJob.GetByID(rsQBDist!JobID)
            If JCJob.QBID = "ORIG" Then
                X = JCJob.QBParentID
            Else
                X = JCJob.QBID
            End If
            
            orJournalLine1.JournalDebitLine.Amount.SetValue rsQBDist!Amount
            
            If rsQBDist!JobID = 999999999 Then
            Else
                orJournalLine1.JournalDebitLine.EntityRef.ListID.SetValue Trim(X)
            End If
        
        Else
            orJournalLine1.JournalDebitLine.Amount.SetValue rsQB!Amount
            ' orJournalLine1.JournalDebitLine.EntityRef.ListID.SetValue Trim(JCJob.QBParentID)
        End If
    
        If rsQBDist.RecordCount > 0 Then
            rsQBDist.MoveNext
        End If
    
    Next j
    
    ' credit info
    Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
    orJournalLine1.JournalCreditLine.TxnLineID.SetValue k + 1
    ' orJournalLine1.JournalCreditLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
    
    orJournalLine1.JournalCreditLine.AccountRef.ListID.SetValue Trim(rsQB!QBLiabilityAcct)
    
    orJournalLine1.JournalCreditLine.Amount.SetValue rsQB!Amount
    
    orJournalLine1.JournalCreditLine.Memo.SetValue rsQB!Desc
    ' orJournalLine1.JournalCreditLine.EntityRef.ListID.SetValue Trim(JCJob.QBID)

End Sub

Public Sub QBBuildBillAddRq(requestMsgSet As IMsgSetRequest, country As String)
 
    If (requestMsgSet Is Nothing) Then
        Exit Sub
    End If
 
    'Add the request to the message set request object.
    Set billAdd = requestMsgSet.AppendBillAddRq
 
    'Set the elements of IBillAdd.
 
    ' Set the FullName value.
    ' billAdd.VendorRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    billAdd.VendorRef.ListID.SetValue rsQB!QBPayTo
 
    ' Set the FullName value.
    ' billAdd.APAccountRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    boo = QBAccount.GetByID(Me.cmbQBAP.ItemData(Me.cmbQBAP.ListIndex))
    billAdd.APAccountRef.ListID.SetValue QBAccount.QBID
 
    ' Set the value of the IBillAdd.TxnDate element.
    billAdd.TxnDate.SetValue PRBatch.CheckDate
 
    ' Set the value of the IBillAdd.DueDate element.
    billAdd.DueDate.SetValue rsQB!DueDate
 
    ' Set the value of the IBillAdd.RefNumber element.
    ' 20 char max
    billAdd.RefNumber.SetValue "PR" & Format(PRBatch.CheckDate, "yyyymmdd")
 
    ' Set the FullName value.
    ' billAdd.TermsRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    ' billAdd.TermsRef.ListID.SetValue "ab"
 
    ' Set the value of the IBillAdd.Memo element.
    billAdd.Memo.SetValue rsQB!Desc
 
'    If (country = "US") Then
'        ' Set the value of the IBillAdd.LinkToTxnIDList element.
'        billAdd.LinkToTxnIDList.Add "val"
'    End If
    
    ' split per job ????
    rsQBDist.Filter = adFilterNone
    rsQBDist.Filter = "RecID = " & rsQB!RecID
    
    If rsQBDist.RecordCount = 0 Then
        k = 1
    Else
        k = rsQBDist.RecordCount
        rsQBDist.MoveFirst
    End If
    
    'Add multiple elements to the list. In this case we will add 5 elements.
    For j = 1 To k
        
        ' Append an element to the list and save the element in expenseLineAdd1 so we can set its values.
        Set expenseLineAdd1 = billAdd.ExpenseLineAddList.Append
 
        ' Set the FullName value.
        ' expenseLineAdd1.AccountRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        expenseLineAdd1.AccountRef.ListID.SetValue rsQB!QBExpenseAcct
 
        ' Set the value of the IExpenseLineAdd.Amount element.
        If rsQBDist.RecordCount = 0 Then
            expenseLineAdd1.Amount.SetValue rsQB!Amount
        Else
            expenseLineAdd1.Amount.SetValue rsQBDist!Amount
        End If
 
        ' Set the value of the IExpenseLineAdd.Memo element.
        If rsQBDist.RecordCount = 0 Then
            expenseLineAdd1.Memo.SetValue rsQB!Desc
        Else
            expenseLineAdd1.Memo.SetValue rsQB!Desc ' ??? use job name ???
        End If
 
        ' Set the FullName value.
        ' expenseLineAdd1.CustomerRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        If rsQBDist.RecordCount > 0 Then
            If rsQBDist!JobID = 999999999 Then
            Else
                boo = JCJob.GetByID(rsQBDist!JobID)
                If rsQBDist.RecordCount > 0 Then
                    If JCJob.QBID = "ORIG" Then
                        X = JCJob.QBParentID
                    Else
                        X = JCJob.QBID
                    End If
                    expenseLineAdd1.CustomerRef.ListID.SetValue X
                End If
            End If
        End If
        
        ' Set the FullName value.
        ' expenseLineAdd1.ClassRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        ' expenseLineAdd1.ClassRef.ListID.SetValue "ab"
 
        ' Set the value of the IExpenseLineAdd.BillableStatus element.
        ' expenseLineAdd1.BillableStatus.SetValue bsBillable
 
        'If Not (country = "US") Then
        '    ' Set the FullName value.
        '    expenseLineAdd1.TaxCodeRef.FullName.SetValue "ab"
        '
        '    ' Set the ListID value.
        '    expenseLineAdd1.TaxCodeRef.ListID.SetValue "ab"
        '
        'End If
        
        ' Set the value of the IExpenseLineAdd.defMacro element.
        QBCount = QBCount + 1
        expenseLineAdd1.defMacro.SetValue "Exp:" & QBCount
 
        If rsQBDist.RecordCount > 0 Then rsQBDist.MoveNext
    
    Next j
 
'    'Add multiple elements to the list. In this case we will add 5 elements.
'    Dim orItemLineAdd2 As IORItemLineAdd
'    Dim k As Integer
'    For k = 0 To 4
'
'        ' Append an element to the list and save the element in orItemLineAdd2 so we can set its values.
'        Set orItemLineAdd2 = billAdd.ORItemLineAddList.Append
'
'        ' Only can set one of the OR elements.
'        ' We will portray this restriction by using an If/Then/Else.
'        Dim orItemLineAddORElement3 As String
'        orItemLineAddORElement3 = "ItemLineAdd"
'        If (orItemLineAddORElement3 = "ItemLineAdd") Then
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.ItemRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.ItemRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemLineAdd.Desc element.
'            orItemLineAdd2.ItemLineAdd.Desc.SetValue "val"
'
'            ' Set the value of the IItemLineAdd.Quantity element.
'            orItemLineAdd2.ItemLineAdd.Quantity.SetValue 2#
'
'            ' Set the value of the IItemLineAdd.Cost element.
'            orItemLineAdd2.ItemLineAdd.Cost.SetValue 2#
'
'            ' Set the value of the IItemLineAdd.Amount element.
'            orItemLineAdd2.ItemLineAdd.Amount.SetValue 2#
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.CustomerRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.CustomerRef.ListID.SetValue "ab"
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.ClassRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.ClassRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemLineAdd.BillableStatus element.
'            orItemLineAdd2.ItemLineAdd.BillableStatus.SetValue bsBillable
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.ListID.SetValue "ab"
'
'            If Not (country = "US") Then
'                ' Set the FullName value.
'                orItemLineAdd2.ItemLineAdd.TaxCodeRef.FullName.SetValue "ab"
'
'                ' Set the ListID value.
'                orItemLineAdd2.ItemLineAdd.TaxCodeRef.ListID.SetValue "ab"
'
'            End If
'            If (country = "US") Then
'                ' Set the value of the ILinkToTxn.TxnID element.
'                orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnID.SetValue "val"
'
'                ' Set the value of the ILinkToTxn.TxnLineID element.
'                orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnLineID.SetValue "val"
'
'            End If
'        ElseIf (orItemLineAddORElement3 = "ItemGroupLineAdd") Then
'            ' Set the FullName value.
'            orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemGroupLineAdd.Desc element.
'            orItemLineAdd2.ItemGroupLineAdd.Desc.SetValue "val"
'
'            ' Set the value of the IItemGroupLineAdd.Quantity element.
'            orItemLineAdd2.ItemGroupLineAdd.Quantity.SetValue 2#
'
'        End If
'
'    Next k
'
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.Tax1Total element.
'        billAdd.Tax1Total.SetValue 2#
'
'    End If
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.Tax2Total element.
'        billAdd.Tax2Total.SetValue 2#
'
'    End If
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.ExchangeRate element.
'        billAdd.ExchangeRate.SetValue 2.5
'
'    End If
'    If (country = "UK") Then
'        ' Set the value of the IBillAdd.AmountIncludesVAT element.
'        billAdd.AmountIncludesVAT.SetValue True
'
'    End If
'    If (country = "US") Then
'        ' Set the value of the IBillAdd.IncludeRetElementList element.
'        billAdd.IncludeRetElementList.Add "val"
'
'    End If
    
    ' Set the value of the IBillAdd.defMacro element.
    QBCount = QBCount + 1
    billAdd.defMacro.SetValue "Bill:" & QBCount
 
End Sub

Public Sub QBParseBillAddRs(responseMsgSet As IMsgSetResponse, country As String)
 
    If (responseMsgSet Is Nothing) Then
        Exit Sub
    End If
 
    Set ResponseList = responseMsgSet.ResponseList
    If (ResponseList Is Nothing) Then
        Exit Sub
    End If
 
    ' Go through all of the responses in the list.
    Dim i As Integer
    For i = 0 To ResponseList.Count - 1
        Set Response = ResponseList.GetAt(i)
 
 ' MsgBox Response.StatusCode & vbCr & Response.StatusMessage & vbCr & Response.StatusSeverity
 
        ' Check the status returned for the response.
        If (Response.StatusCode = 0) Then
 
            ' Check to make sure the response is of the type we are expecting.
            If (Not Response.Detail Is Nothing) Then
                Dim ResponseType As Integer
                ResponseType = Response.Type.GetValue
                Dim j As Integer
                ' Check for BillAddRs.
                If (ResponseType = rtBillAddRs) Then
                    Dim billRet As IBillRet
                    Set billRet = Response.Detail
                    ' ParseBillRet billRet, country
                End If
            End If
        End If
    Next i
End Sub

Private Sub QBCheckAddRq(ByVal country As String, _
                        ByVal MajorVersion As Integer, _
                        ByVal MinorVersion As Integer, _
                        ByVal BatchID As Long, _
                        ByVal QBIDChk As String, _
                        ByVal QBChkName As String, _
                        ByVal QBIDExp As String, _
                        ByVal QBExpName As String, _
                        ByVal QBIDPay As String, _
                        ByVal QBPayName As String)
  
    If BatchID <> 0 Then
                  
        SQLString = "SELECT * FROM PRHist WHERE BatchID = " & BatchID & _
                    " ORDER BY CheckNumber"
        If PRHist.GetBySQL(SQLString) = False Then
            MsgBox "No Payroll data to export!", vbExclamation
            GoBack
        End If
    
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Do
          
            ' update flag
            If PRHist.QBUpdateFlag = 0 Or frmDateRange.chkQBOverride = 1 Then
          
                If PREmployee.GetByID(PRHist.EmployeeID) = False Then
                    MsgBox "Employee NF: " & PRHist.EmployeeID, vbExclamation
                    GoBack
                End If
                Me.lblMsg1 = "Building Check Add Request: " & PREmployee.LFName
                Me.Refresh
              
                If PRHist.Net > 0 Then
              
                    QBBuildCheckAddRq requestMsgSet, country, _
                                    QBIDChk, QBChkName, _
                                    QBIDExp, QBExpName, _
                                    QBIDPay, QBPayName, _
                                    BatchID
          
                End If
          
                DirDepTotal = DirDepTotal + PRHist.DirectDeposit
                          
            End If
          
            If PRHist.GetNext = False Then Exit Do
        Loop
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
    Else
                    
        ' direct deposit update to the check register
        ' batchid=0 triggers dir dep update
        If DirDepTotal > 0 Then
            QBBuildCheckAddRq requestMsgSet, country, _
                            QBIDChk, QBChkName, _
                            QBIDExp, QBExpName, _
                            QBIDPay, QBPayName, 0
        End If
    
    End If

End Sub

Private Sub QBBuildCheckAddRq(ByVal requestMsgSet As IMsgSetRequest, _
                           ByVal country As String, _
                           ByVal QBIDChk As String, _
                           ByVal QBChkName As String, _
                           ByVal QBIDExp As String, _
                           ByVal QBExpName As String, _
                           ByVal QBIDPay As String, _
                           ByVal QBPayName As String, _
                           ByVal BtchID As Long)
  
  If (requestMsgSet Is Nothing) Then
    Exit Sub
  End If
  
    QBCount = QBCount + 1
  
  'Add the request to the message set request object.
  Dim checkAdd As ICheckAdd
  
  Set checkAdd = requestMsgSet.AppendCheckAddRq
  
  'Set the elements of ICheckAdd.
  
  ' Set the FullName value.
  ' checkAdd.AccountRef.FullName.SetValue frmQBCheckUpdate.cmbChecking
  
  ' Set the ListID value.
  checkAdd.AccountRef.ListID.SetValue QBIDChk
  
  ' Set the FullName value.
  checkAdd.PayeeEntityRef.FullName.SetValue QBPayName
  
  ' Set the ListID value.
  checkAdd.PayeeEntityRef.ListID.SetValue QBIDPay
  
  ' Set the value of the ICheckAdd.RefNumber element.
  If BtchID <> 0 Then
     checkAdd.RefNumber.SetValue PRHist.CheckNumber
  Else
     checkAdd.RefNumber.SetEmpty
  End If
  
  ' Set the value of the ICheckAdd.TxnDate element.
  If BtchID <> 0 Then
     checkAdd.TxnDate.SetValue PRBatch.CheckDate
  Else
     checkAdd.TxnDate.SetValue rsQB!DueDate
  End If
  
  ' Set the value of the ICheckAdd.Memo element.
  If BtchID <> 0 Then
        If Me.chkNoName = 0 Then
            checkAdd.Memo.SetValue PREmployee.LFName
        Else
            checkAdd.Memo.SetValue "Emp#: " & PREmployee.EmployeeNumber
        End If
  Else
        checkAdd.Memo.SetValue "Direct Deposit"
  End If
  
  ' Set the value of the IAddress.Addr1 element.
  checkAdd.Address.Addr1.SetValue ""
  
  ' Set the value of the IAddress.Addr2 element.
  checkAdd.Address.Addr2.SetValue ""
  
  ' Set the value of the IAddress.Addr3 element.
  checkAdd.Address.Addr3.SetValue ""
  
  ' Set the value of the IAddress.Addr4 element.
  checkAdd.Address.Addr4.SetValue ""
  
  ' Set the value of the IAddress.City element.
  checkAdd.Address.City.SetValue ""
  
  If (country = "US") Then
    ' Set the value of the IAddress.State element.
    checkAdd.Address.State.SetValue ""
  
  End If
  If (country = "UK") Then
    ' Set the value of the IAddress.County element.
    checkAdd.Address.County.SetValue ""
  
  End If
  If (country = "CA") Then
    ' Set the value of the IAddress.Province element.
    checkAdd.Address.Province.SetValue ""
  
  End If
  ' Set the value of the IAddress.PostalCode element.
  checkAdd.Address.PostalCode.SetValue ""
  
  ' Set the value of the IAddress.Country element.
  checkAdd.Address.country.SetValue "l"
  
  ' Set the value of the ICheckAdd.IsToBePrinted element.
  checkAdd.IsToBePrinted.SetValue False
  
  'Add multiple elements to the list. In this case we will add 5 elements.
  Dim expenseLineAdd1 As IExpenseLineAdd
    
  ' Append an element to the list and save the element in expenseLineAdd1 so we can set its values.
  Set expenseLineAdd1 = checkAdd.ExpenseLineAddList.Append
  
  ' Set the FullName value.
  expenseLineAdd1.AccountRef.FullName.SetValue QBExpName
  
  ' Set the ListID value.
  expenseLineAdd1.AccountRef.ListID.SetValue QBIDExp
  
  ' Set the value of the IExpenseLineAdd.Amount element.
  If BtchID <> 0 Then
      expenseLineAdd1.Amount.SetValue PRHist.Net
  Else
      expenseLineAdd1.Amount.SetValue DirDepTotal
  End If
  
  ' Set the value of the IExpenseLineAdd.Memo element.
  If BtchID <> 0 Then
     expenseLineAdd1.Memo.SetValue PREmployee.LFName
  Else
     expenseLineAdd1.Memo.SetValue "Direct Deposit"
  End If
  
  ' ****************************************
  ' * Cust Ref
  ' Set the FullName value.
  ' expenseLineAdd1.CustomerRef.FullName.SetValue "ab"
  
  ' Set the ListID value.
  ' expenseLineAdd1.CustomerRef.ListID.SetValue "ab"
  
  ' ****************************************
  
  ' ****************************************
  ' * Class Ref
  ' Set the FullName value.
  ' expenseLineAdd1.ClassRef.FullName.SetValue "ab"
  
  ' Set the ListID value.
  ' expenseLineAdd1.ClassRef.ListID.SetValue "ab"
  
  ' Set the value of the IExpenseLineAdd.BillableStatus element.
  ' expenseLineAdd1.BillableStatus.SetValue bsBillable
  
'    If Not (country = "US") Then
'      ' Set the FullName value.
'      expenseLineAdd1.TaxCodeRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      expenseLineAdd1.TaxCodeRef.ListID.SetValue "ab"
'
'    End If
'    ' Set the value of the IExpenseLineAdd.defMacro element.
'    expenseLineAdd1.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
  
'  'Add multiple elements to the list. In this case we will add 5 elements.
'  Dim orItemLineAdd2 As IORItemLineAdd
'  Dim k As Integer
'  For k = 0 To 4
'    ' Append an element to the list and save the element in orItemLineAdd2 so we can set its values.
'    Set orItemLineAdd2 = checkAdd.ORItemLineAddList.Append
'
'    ' Only can set one of the OR elements.
'    ' We will portray this restriction by using an If/Then/Else.
'    Dim orItemLineAddORElement3 As String
'    orItemLineAddORElement3 = "ItemLineAdd"
'    If (orItemLineAddORElement3 = "ItemLineAdd") Then
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.ItemRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.ItemRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemLineAdd.Desc element.
'      orItemLineAdd2.ItemLineAdd.Desc.SetValue "val"
'
'      ' Set the value of the IItemLineAdd.Quantity element.
'      orItemLineAdd2.ItemLineAdd.Quantity.SetValue 2#
'
'      ' Set the value of the IItemLineAdd.Cost element.
'      orItemLineAdd2.ItemLineAdd.Cost.SetValue 2#
'
'      ' Set the value of the IItemLineAdd.Amount element.
'      orItemLineAdd2.ItemLineAdd.Amount.SetValue 2#
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.CustomerRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.CustomerRef.ListID.SetValue "ab"
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.ClassRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.ClassRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemLineAdd.BillableStatus element.
'      orItemLineAdd2.ItemLineAdd.BillableStatus.SetValue bsBillable
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.ListID.SetValue "ab"
'
'      If Not (country = "US") Then
'        ' Set the FullName value.
'        orItemLineAdd2.ItemLineAdd.TaxCodeRef.FullName.SetValue "ab"
'
'        ' Set the ListID value.
'        orItemLineAdd2.ItemLineAdd.TaxCodeRef.ListID.SetValue "ab"
'
'      End If
'      If (country = "US") Then
'        ' Set the value of the ILinkToTxn.TxnID element.
'        orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnID.SetValue "val"
'
'        ' Set the value of the ILinkToTxn.TxnLineID element.
'        orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnLineID.SetValue "val"
'
'      End If
'    ElseIf (orItemLineAddORElement3 = "ItemGroupLineAdd") Then
'      ' Set the FullName value.
'      orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemGroupLineAdd.Desc element.
'      orItemLineAdd2.ItemGroupLineAdd.Desc.SetValue "val"
'
'      ' Set the value of the IItemGroupLineAdd.Quantity element.
'      orItemLineAdd2.ItemGroupLineAdd.Quantity.SetValue 2#
'
'    End If
'
'  Next k
'
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.Tax1Total element.
'    checkAdd.Tax1Total.SetValue 2#
'
'  End If
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.Tax2Total element.
'    checkAdd.Tax2Total.SetValue 2#
'
'  End If
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.ExchangeRate element.
'    checkAdd.ExchangeRate.SetValue 2.5
'
'  End If
'  If (country = "UK") Then
'    ' Set the value of the ICheckAdd.AmountIncludesVAT element.
'    checkAdd.AmountIncludesVAT.SetValue True
'
'  End If
'  If (country = "US") Then
'    ' Set the value of the ICheckAdd.IncludeRetElementList element.
'    checkAdd.IncludeRetElementList.Add "val"
'
'  End If
'  ' Set the value of the ICheckAdd.defMacro element.
'  checkAdd.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
  
End Sub


