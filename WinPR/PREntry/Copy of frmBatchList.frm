VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBatchList 
   Caption         =   "Payroll Data Entry Batch List"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQBTaxPay 
      Caption         =   "Update ALL Amounts to QB"
      Height          =   615
      Left            =   9600
      TabIndex        =   18
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton cmdWBJReport 
      Caption         =   "Wage by Job Report"
      Height          =   615
      Left            =   9600
      TabIndex        =   17
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdQBInvUpdate 
      Caption         =   "Update Invoicing to QB"
      Height          =   615
      Left            =   9600
      TabIndex        =   16
      Top             =   6060
      Width           =   1695
   End
   Begin VB.CommandButton cmdQBUpdate 
      Caption         =   "Update Net Pay to QB"
      Height          =   615
      Left            =   9600
      TabIndex        =   15
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDirDep 
      Caption         =   "Dir Deposit"
      Height          =   495
      Left            =   9600
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdEntryForm 
      Caption         =   "Entry Form"
      Height          =   495
      Left            =   9600
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton chkChkReg 
      Caption         =   "Check Reg"
      Height          =   495
      Left            =   9600
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdDepositList 
      Caption         =   "Deposit List"
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheckPrint 
      Caption         =   "Check Print"
      Height          =   495
      Left            =   9600
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdEntry 
      Caption         =   "&ENTRY"
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3855
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   8175
      _cx             =   14420
      _cy             =   6800
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
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Reports:"
      Height          =   255
      Left            =   10080
      TabIndex        =   9
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Modify the current batch"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Delete the batch"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Create a new batch"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   450
      TabIndex        =   4
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "frmBatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Dim X As String
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending

Dim dbFileName As String
Dim dbFields(5) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte
Public BatchID As Long

Dim SQLStr As String

Private Sub Form_Load()
    
    Me.lblCompanyName = Trim(PRCompany.Name)
    
    ' no employees !!! ???
    SQLString = "SELECT * FROM PREmployee"
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No employees found!", vbExclamation, "PR Data Entry"
        GoBack
    End If
    
    ' delete all blank records
    rsInit "DELETE * FROM PRBatch WHERE IsNull(CreateDate)", cn, rs
    
    ' set the constants for the file
    dbFileName = "PRBatch"
    dbFields(0) = "BatchID"
    dbFields(1) = "CreateDate"
    dbFields(2) = "PEDate"
    dbFields(3) = "CheckDate"
    dbFields(4) = "RecCount"
    dbFields(5) = "YearMonth"
    dbSortCol = 0
    dbSortDesc = True
    
    GetSQLString
    
    rsInit GetSQLString, cn, rs
    SetGrid rs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 1000
    fg.ColWidth(1) = 1300
    fg.ColWidth(2) = 1300
    fg.ColWidth(3) = 1300
    fg.ColWidth(4) = 1500
    fg.ColWidth(5) = 1500
    
    fg.ColFormat(1) = "mm/dd/yyyy"
    fg.ColFormat(2) = "mm/dd/yyyy"
    fg.ColFormat(3) = "mm/dd/yyyy"
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = dbFields(0) & "+"
    fg.Cell(flexcpFontBold, 0, 0) = True
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    
    Unload frmSplash
    
    ' show the update to QB Inv button?
    Me.cmdQBInvUpdate.Visible = False
    Me.cmdWBJReport.Visible = False
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeScreenDefault & _
                " AND UserID = " & PRCompany.CompanyID & _
                " AND Description = 'TimeSheet'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        If PRGlobal.Byte1 = "1" Then
            Me.cmdQBInvUpdate.Visible = True
            Me.cmdWBJReport.Visible = True
        End If
    End If
    
    ' QB Tax Pay - .jb ONLY
    If UCase(User.Logon) = "JIM" Then Me.cmdQBTaxPay.Visible = True
    If UCase(User.Logon) = "DH" Then Me.cmdQBTaxPay.Visible = True
    
    Me.KeyPreview = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF7: CheckDateSweep
    End Select

End Sub

Private Sub cmdEntry_Click()
    
Dim OK As Boolean
    
    frmNewBatch.BatchID = fg.TextMatrix(fg.Row, 0)
    frmNewBatch.Show vbModal
    
    If Response Then
        
        ' time sheet(s) to use if originated as a Job Dist batch
        If PRBatch.GetByID(fg.TextMatrix(fg.Row, 0)) = False Then
            MsgBox "Batch Error: " & fg.TextMatrix(fg.Row, 0), vbExclamation
            GoBack
        End If
        
        If PRBatch.JobDist = 1 Then
            If TableExists("PRTimeSheet", cn) = True Then
                SQLString = "SELECT * FROM PRTimeSheet"
                If PRTimeSheet.GetBySQL(SQLString) = True Then
                    frmSelTimeSheets.Init
                    frmSelTimeSheets.fg.Enabled = False     ' can't change TimeSheets
                    frmSelTimeSheets.lblMsg1 = "Time Sheet Records already assigned to this batch!" & vbCr & _
                                               "No changes allowed!"
                    frmSelTimeSheets.fg.CellFontItalic = True
                    frmSelTimeSheets.cmdExit.Enabled = False
                    frmSelTimeSheets.Show vbModal           ' weeks selected
                End If
            End If
        Else
            frmSelTimeSheets.lblMsg1 = ""
            frmSelTimeSheets.cmdExit.Enabled = True
            frmSelTimeSheets.UseDist = False
        End If
        
        frmEntryTS.StartCheckNumber = frmNewBatch.tdbIntStartCheck
        frmEntryTS.BatchID = fg.TextMatrix(fg.Row, 0)
        frmEntryTS.Show vbModal
        UpdateGrid
    
    End If
    
    frmNewBatch.rsItem.Close
    Unload frmNewBatch
    
End Sub

' To Do:

Private Sub cmdAdd_Click()
    
Dim PRBilling As Boolean
    
    ' PR billing? - skip Time Sheet prompts
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                " AND Description = 'PayrollBilling' " & _
                " AND Var1 = 'Yes' " & _
                " AND Var2 = '" & PRCompany.GLCompanyID & "'"
    PRBilling = PRGlobal.GetBySQL(SQLString)
    
    ' job refresh
    If TableExists("JCJob", cn) = True Then
        ' don't bother if no jobs in list
        ' init from JC menu if a new install
        SQLString = "SELECT * from JCJob"
        If JCJob.GetBySQL(SQLString) = True Then
            frmJCGetQBData.Show vbModal
        End If
    End If
    
    ' AddAdo rs, fg
    frmNewBatch.BatchID = 0 ' from new
    frmNewBatch.Show vbModal
    RefreshGrid
    
    If Response Then
        
        ' time sheet(s) to use
        If PRBilling = False Then
            If TableExists("PRTimeSheet", cn) = True Then
                SQLString = "SELECT * FROM PRTimeSheet"
                If PRTimeSheet.GetBySQL(SQLString) = True Then
                    frmSelTimeSheets.lblMsg1 = ""
                    frmSelTimeSheets.Init
                    If frmSelTimeSheets.UseDist = True Then
                        frmSelTimeSheets.fg.Enabled = True
                        frmSelTimeSheets.lblMsg1 = ""
                        frmSelTimeSheets.cmdExit.Enabled = True
                        frmSelTimeSheets.Show vbModal
                    End If
                End If
            End If
        End If
        
        ' mark the batch if any time sheet records selected
        If frmSelTimeSheets.UseDist = True Then
            If PRBatch.GetByID(rs!BatchID) = False Then
                MsgBox "Batch Error: " & rs!BatchID, vbExclamation
                GoBack
            End If
            PRBatch.JobDist = 1
            PRBatch.Save (Equate.RecPut)
        End If
        
        frmEntryTS.StartCheckNumber = frmNewBatch.tdbIntStartCheck
        frmEntryTS.BatchID = rs!BatchID
        frmEntryTS.Show vbModal
        UpdateGrid
    
    End If
    
    frmNewBatch.rsItem.Close
    Unload frmNewBatch

End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrID As Long
    
    CurrID = fg.TextMatrix(fg.Row, 0)
        
    rs.Close
    rsInit GetSQLString, cn, rs
    Set fg.DataSource = rs.DataSource
       
    rw = fg.FindRow(CurrID, 0, 0)
       
    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus
    
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 
    If Col = 0 Then     ' validates for number - must enter a value that does not already exist
       
       If fg.EditText = "" Or fg.EditText = "0" Then
          MsgBox "Enter a value!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
       rw = fg.FindRow(fg.EditText, 0, 0)
       If rw <> -1 Then
          MsgBox "Number already exists!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
    End If

End Sub


Private Sub cmdExit_Click()
    
    If jbFlag = False Then BackName = "\Balint\GLMenu.exe"
    GoBack

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
    
    If fg.Rows = 1 Then Exit Sub
    
    ' what if no records left ????
        
    DelConfirm = MsgBox("Batch #: " & fg.TextMatrix(fg.Row, 0) & vbCr & "PE Date: " & fg.TextMatrix(fg.Row, 2), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")
    
    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If
    
    ' delete the PRBatch, PRHistory and PRDist records
    SQLString = "DELETE * FROM PRHist WHERE PRHist.BatchID = " & fg.TextMatrix(fg.Row, 0)
    cn.Execute SQLString
    
    SQLString = "DELETE * FROM PRDist WHERE PRDist.BatchID = " & fg.TextMatrix(fg.Row, 0)
    cn.Execute SQLString
    
    SQLString = "DELETE * FROM PRItemHist WHERE PRItemHist.BatchID = " & fg.TextMatrix(fg.Row, 0)
    cn.Execute SQLString
    
    SQLString = "DELETE * FROM PRBatch WHERE PRBatch.BatchID = " & fg.TextMatrix(fg.Row, 0)
    cn.Execute SQLString
    
    ' clear out time sheet selections
    If TableExists("PRTimeSheet", cn) = True Then
        
        ' clear PRBatch designation in PRGlobal
        SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePRBatchWE & _
                    " AND UserID = " & PRCompany.CompanyID & _
                    " AND Description = '" & fg.TextMatrix(fg.Row, 0) & "'"
        cnDes.Execute SQLString
                
        SQLString = "SELECT * FROM PRTimeSheet WHERE BatchID = " & fg.TextMatrix(fg.Row, 0)
        If PRTimeSheet.GetBySQL(SQLString) = True Then
            Do
                PRTimeSheet.BatchID = 0
                PRTimeSheet.HistID = 0
                PRTimeSheet.Save (Equate.RecPut)
                If PRTimeSheet.GetNext = False Then Exit Do
            Loop
        End If
    
    End If
    
    ' DelAdo rs, fg, fg.TextMatrix(fg.Row, 0)
    ' DelAdo rs, fg
    
    RefreshGrid

End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    ' clicking on a column header sorts based on that column
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then

       ' toggle the sort order
       If fg.MouseCol = dbSortCol Then
          If dbSortDesc = False Then
             dbSortDesc = True
          Else
             dbSortDesc = False
          End If
       Else
          ' switch the column
          fg.Cell(flexcpFontBold, 0, fg.MouseCol) = True
          fg.Cell(flexcpFontBold, 0, dbSortCol) = False
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol)
          dbSortCol = fg.MouseCol
       End If
       
       If dbSortDesc Then
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "-"
       Else
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "+"
       End If
    
       rs.Close
       
       rsInit GetSQLString, cn, rs
       Set fg.DataSource = rs.DataSource
       
       fg.ShowCell 1, 0

    End If
    
End Sub

Private Function GetSQLString() As String
    
Dim aa As Integer
    
' set the SQL string
'    x = "SELECT [Number],[Description] " & _
'        "FROM GLDescriptions ORDER BY [Number] DESC"

    GetSQLString = "SELECT "
    
    For aa = 0 To UBound(dbFields, 1)
        GetSQLString = GetSQLString & " [" & dbFields(aa) & "]"
        If aa <> UBound(dbFields, 1) Then GetSQLString = GetSQLString & ","
        GetSQLString = GetSQLString & " "
    Next aa
    
    GetSQLString = GetSQLString & "FROM " & dbFileName & " ORDER BY [" & dbFields(dbSortCol) & "]"
    If dbSortDesc Then
       GetSQLString = GetSQLString & " DESC"
    End If

End Function


Private Sub RefreshGrid()
    
    rs.Close
    rsInit GetSQLString, cn, rs
    SetGrid rs, fg
    
    rw = fg.Row
    If rw = fg.Rows Then rw = fg.Rows - 1
    fg.Select rw, 0
    fg.ShowCell rw, 0

End Sub

Private Sub UpdateGrid()
        
    BatchID = fg.TextMatrix(fg.Row, 0)
    If Not PRBatch.GetByID(BatchID) Then
        MsgBox "Batch Error: " & BatchID
        End
    End If
    rs!RecCount = PRBatch.RecCount
    rs.Update

End Sub

Private Sub ReportRun(ByVal EXEName As String, ByVal ProgName As String)
Dim ShellString As String
Dim TID As Double
         
    ShellString = Mid(App.Path, 1, 3) & "Balint\" & Trim(EXEName) & ".exe" & _
         " SysFile=\Balint\Data\GLSystem.mdb" & _
         " UserID=" & User.ID & _
         " BackName=\Balint\PREntry.exe" & _
         " Batch=" & fg.TextMatrix(fg.Row, 0) & _
         " ProgName=" & ProgName
     
     If dbPwd <> "" Then
        ShellString = ShellString & " dbPwd=" & dbPwd
     End If
             
    TID = Shell(ShellString, vbMaximizedFocus)
    
    Unload Me
    End

End Sub

Private Sub cmdCheckPrint_Click()
    ReportRun "PRGReps", "CheckPrint"
End Sub

Private Sub cmdDepositList_Click()
    ReportRun "PRReport", "Deposit"
End Sub

Private Sub chkChkReg_Click()
    ReportRun "PRReport", "CheckReg"
End Sub

Private Sub cmdEntryForm_Click()
    ReportRun "PRReport", "EntryForm"
End Sub
Private Sub cmdDirDep_Click()
    ReportRun "PRReport", "DirDep"
End Sub
Private Sub cmdWBJReport_Click()
    ReportRun "PRReport", "WageByJob"
End Sub

Private Sub cmdQBUpdate_Click()
    BatchID = fg.TextMatrix(fg.Row, 0)
    frmQBCheckUpdate.Show vbModal
End Sub

Private Sub cmdQBInvUpdate_Click()
    BatchID = fg.TextMatrix(fg.Row, 0)
    frmQBInvUpdate.Show vbModal
End Sub
Private Sub cmdQBTaxPay_Click()
    BatchID = fg.TextMatrix(fg.Row, 0)
    
    ' forces TaxPay to ask for date range
    BatchID = 0
    
    ReportRun "PRQBFunc", "TaxPay"
End Sub


Private Sub CheckDateSweep()

Dim ChkDays, Ct, Rct As Long
Dim LastBatch As Long

    If MsgBox("OK to run Check Date Fix?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    frmProgress.Show
    frmProgress.lblMsg1 = "Now Running Check Date Fix ..."
    frmProgress.Refresh
    
    ChkDays = PRCompany.CheckDays
    LastBatch = 0
    
    SQLString = "SELECT * FROM PRDist WHERE CheckDate = 0"
    If PRDist.GetBySQL(SQLString) = True Then
        Rct = PRDist.Records
        Ct = 0
        Do
            Ct = Ct + 1
            If Ct Mod 100 = 1 Then
                frmProgress.lblMsg2 = "Dist record: " & Ct & " Of: " & Rct
                frmProgress.Refresh
            End If
            If LastBatch = 0 Or PRDist.BatchID <> LastBatch Then
                If PRBatch.GetByID(PRDist.BatchID) = False Then
                    MsgBox "PR Batch not found: " & PRDist.BatchID, vbExclamation
                    GoBack
                End If
            End If
            LastBatch = PRDist.BatchID
            PRDist.CheckDate = PRBatch.CheckDate
            PRDist.Save (Equate.RecPut)
            If PRDist.GetNext = False Then Exit Do
        Loop
    End If

    LastBatch = 0
    
    SQLString = "SELECT * FROM PRItemHist WHERE CheckDate = 0"
    If PRItemHist.GetBySQL(SQLString) = True Then
        Rct = PRItemHist.Records
        Ct = 0
        Do
            Ct = Ct + 1
            If Ct Mod 100 = 1 Then
                frmProgress.lblMsg2 = "Item Hist record: " & Ct & " Of: " & Rct
                frmProgress.Refresh
            End If
            If LastBatch = 0 Or PRItemHist.BatchID <> LastBatch Then
                If PRBatch.GetByID(PRItemHist.BatchID) = False Then
                    MsgBox "PR Batch not found: " & PRItemHist.BatchID, vbExclamation
                    GoBack
                End If
            End If
            LastBatch = PRItemHist.BatchID
            PRItemHist.CheckDate = PRBatch.CheckDate
            PRItemHist.Save (Equate.RecPut)
            If PRItemHist.GetNext = False Then Exit Do
        Loop
    End If

    MsgBox "Check Date correction sweep is complete ...", vbInformation
    GoBack

End Sub

