VERSION 5.00
Begin VB.Form frmWageByJob 
   Caption         =   "Wage By Job Report"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
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
   ScaleHeight     =   4665
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6248
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   3128
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox chkSepPage 
      Caption         =   "&Separate page per job"
      Height          =   375
      Left            =   3068
      TabIndex        =   1
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      Height          =   1335
      Left            =   2648
      TabIndex        =   5
      Text            =   "PLEASE SELECT A DATE RANGE !!!"
      Top             =   1200
      Width           =   6615
   End
   Begin VB.CommandButton cmdDaterange 
      Caption         =   "&DATE RANGE"
      Height          =   735
      Left            =   1328
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmWageByJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsTS As New ADODB.Recordset

Dim i, j, k As Long
Dim X, Y, Z As String
Dim P2, p1 As Currency

Dim RegTot(4), OvtTot(4), OthTot(4) As Currency

Dim BillTot, BillTotG As Currency

Dim boo As Boolean

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name

    ' BatchID assigned? - use it
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
        Me.txtDisplay.Text = "Batch #: " & PRBatch.BatchID & _
                             " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy") & _
                             " Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
        RangeType = PREquate.RangeTypeBatch
        BatchNumbr = PRBatchID

    End If

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdDateRange_Click()
    
    frmDateRange.lblProgram = "DATE RANGE"
    frmDateRange.Show vbModal
   
    If frmDateRange.optCheckDate = True Then
        OptDate = "CHECK DATE"
    ElseIf frmDateRange.optPEDate = True Then
        OptDate = "P/E DATE"
    End If
    
    If BatchNumbr > 0 Then
        txtDisplay = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        PEDate = PRBatch.PEDate
        PRBatchID = BatchNumbr
        CheckDate = PRBatch.CheckDate
        OptDate = " "
    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
        
    End If

    Me.Refresh

End Sub

Private Sub cmdOK_Click()

Dim Flg As Boolean
Dim LastJobID As Long
Dim FirstFlag As Boolean

    ' alternate format for PR Billing clients
    If PRBilling = True Then
        BillingReport
        GoBack
    End If
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "JobID", adDouble
    rs.Fields.Append "EmpName", adVarChar, 20, adFldIsNullable
    rs.Fields.Append "EmpID", adDouble
    rs.Fields.Append "DptName", adVarChar, 15, adFldIsNullable
    rs.Fields.Append "DptID", adDouble
    rs.Fields.Append "PEDate", adDate
    rs.Fields.Append "RegAmt", adCurrency
    rs.Fields.Append "RegHrs", adCurrency
    rs.Fields.Append "OvtAmt", adCurrency
    rs.Fields.Append "OvtHrs", adCurrency
    rs.Fields.Append "OthAmt", adCurrency
    rs.Fields.Append "OthHrs", adCurrency
    rs.Fields.Append "TotAmt", adCurrency
    rs.Fields.Append "TotHrs", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    rsJob.CursorLocation = adUseClient
    rsJob.Fields.Append "JobID", adDouble
    rsJob.Fields.Append "JobName", adVarChar, 100, adFldIsNullable
    rsJob.Open , , adOpenDynamic, adLockOptimistic
    
    ' add the company name for unassigned jobs
    rsJob.AddNew
    rsJob!JobID = 0
    rsJob!JobName = PRCompany.Name
    rsJob.Update
    
    If BatchNumbr > 0 Then
        SQLString = "SELECT * FROM PRDist WHERE BatchID = " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRDist WHERE CheckDate >= " & CLng(StartDate) & _
                        " AND CheckDate <= " & CLng(EndDate)
        Else
            SQLString = "SELECT * FROM PRDist WHERE PEDate >= " & CLng(StartDate) & _
                        " AND PEDate <= " & CLng(EndDate)
        End If
    End If
    
    If PRDist.GetBySQL(SQLString) = False Then
        MsgBox "No PR data found!", vbExclamation
        GoBack
    End If

    frmProgress.Caption = "Wage report by job for: " & PRCompany.Name
    frmProgress.lblMsg1 = Me.txtDisplay
    frmProgress.Show
    frmProgress.Refresh

    ' gather data entered in TimeSheet but not PRDE
    On Error Resume Next
    rsTS.Close
    On Error GoTo 0
    
    rsTS.CursorLocation = adUseClient
    rsTS.Fields.Append "WEDate", adDate
    rsTS.Fields.Append "EmployeeID", adDouble
    rsTS.Fields.Append "JobID", adDouble
    rsTS.Fields.Append "TotalHours", adCurrency
    rsTS.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePRBatchWE
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            i = PRGlobal.Description
            If BatchNumbr <> 0 Then
                If i <> BatchNumbr Then GoTo NxtGlobal
            Else
                If PRBatch.GetByID(i) Then
                    If OptDate = "CHECK DATE" Then
                        If PRBatch.CheckDate < StartDate Then GoTo NxtGlobal
                        If PRBatch.CheckDate > EndDate Then GoTo NxtGlobal
                    Else
                        If PRBatch.PEDate < StartDate Then GoTo NxtGlobal
                        If PRBatch.PEDate > EndDate Then GoTo NxtGlobal
                    End If
                End If
            End If
            For i = 1 To 10
                X = ""
                If i = 1 Then X = PRGlobal.Var1
                If i = 2 Then X = PRGlobal.Var2
                If i = 3 Then X = PRGlobal.Var3
                If i = 4 Then X = PRGlobal.Var4
                If i = 5 Then X = PRGlobal.Var5
                If i = 6 Then X = PRGlobal.Var6
                If i = 7 Then X = PRGlobal.Var7
                If i = 8 Then X = PRGlobal.Var8
                If i = 9 Then X = PRGlobal.Var9
                If i = 10 Then X = PRGlobal.Var10
                If X = "" Then Exit Do
                SQLString = "SELECT * FROM PRTimeSheet WHERE " & _
                             "(BatchID = 0 or IsNull(BatchID)) " & _
                             " AND TotalHours <> 0"
                If PRTimeSheet.GetBySQL(SQLString) = True Then
                    Do
                        rsTS.AddNew
                        rsTS!WEDate = PRTimeSheet.WEDate
                        rsTS!EmployeeID = PRTimeSheet.EmployeeID
                        rsTS!JobID = PRTimeSheet.JobID
                        rsTS!TotalHours = PRTimeSheet.TotalHours
                        rsTS.Update
                        If PRTimeSheet.GetNext = False Then Exit Do
                    Loop
                End If
            Next i
            
NxtGlobal:
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If
    
    j = PRDist.Records
    
    Do
        
        ' lump job=none and unassigned together
        If PRDist.JobID = 999999 Then PRDist.JobID = 0
        
        i = i + 1
        If i Mod 10 = 1 Then
            frmProgress.lblMsg2 = "On Record: " & Format(i, "#,###,##0") & " Of: " & Format(j, "#,###,##0")
            frmProgress.Refresh
        End If
        
        If rs.RecordCount = 0 Then
            Flg = False
        Else
            rs.MoveFirst
            Do
                Flg = True
                If rs!JobID <> PRDist.JobID Then Flg = False
                If rs!EmpID <> PRDist.EmployeeID Then Flg = False
                If rs!DptID <> PRDist.DepartmentID Then Flg = False
                If rs!PEDate <> PRDist.PEDate Then Flg = False
                If Flg = True Then Exit Do
                rs.MoveNext
            Loop Until rs.EOF
        End If
        If Flg = False Then
            rs.AddNew
            
            ' separate record set for Job Name
            rsJob.Find "JobID = " & PRDist.JobID, 0, adSearchForward, 1
            If rsJob.EOF Then
                rsJob.AddNew
                rsJob!JobID = PRDist.JobID
                If PRDist.JobID = 0 Then
                    rsJob!JobName = Mid(PRCompany.Name, 1, 40)
                Else
                    If JCJob.GetByID(PRDist.JobID) = False Then
                        MsgBox "Job Not Found: " & PRDist.JobID, vbExclamation
                        GoBack
                    End If
                    rsJob!JobName = Mid(JCJob.FullName, 1, 100)
                End If
                rsJob.Update
            End If
            
            rs!JobID = PRDist.JobID
            
            rs!EmpID = PRDist.EmployeeID
            If PREmployee.GetByID(PRDist.EmployeeID) = False Then
                MsgBox "Employee not found: " & PRDist.EmployeeID, vbExclamation
                GoBack
            End If
            rs!EmpName = Mid(PREmployee.LFName, 1, 20)
            
            rs!DptID = PRDist.DepartmentID
            If PRDist.DepartmentID <> 0 Then
                If PRDepartment.GetByID(PRDist.DepartmentID) = False Then
                    MsgBox "Deparment not found: " & PRDist.DepartmentID, vbExclamation
                    GoBack
                End If
                rs!DptName = Mid(PRDepartment.Name, 1, 15)
            Else
                rs!DptName = "Misc"
            End If
            
            rs!PEDate = PRDist.PEDate
            rs.Update
        End If
        If PRDist.ItemType = PREquate.ItemTypeRegPay Then
            rs!RegAmt = rs!RegAmt + PRDist.Amount
            rs!RegHrs = rs!RegHrs + PRDist.Hours
        ElseIf PRDist.ItemType = PREquate.ItemTypeOvtPay Then
            rs!OvtAmt = rs!OvtAmt + PRDist.Amount
            rs!OvtHrs = rs!OvtHrs + PRDist.Hours
        ElseIf PRDist.ItemType = PREquate.ItemTypeOE Then
            rs!OthAmt = rs!OthAmt + PRDist.Amount
            rs!OthHrs = rs!OthHrs + PRDist.Hours
        End If
        rs!TotAmt = rs!RegAmt + rs!OvtAmt + rs!OthAmt
        rs!TotHrs = rs!RegHrs + rs!OvtHrs + rs!OthHrs
        rs.Update
        
        If PRDist.GetNext = False Then Exit Do
    
    Loop
    
    frmProgress.lblMsg2 = "Now sorting data ...."
    frmProgress.Refresh
    
    If rs.RecordCount = 0 Then
        MsgBox "No data found ...", vbExclamation
        GoBack
    End If
    
    rs.Sort = "JobID, EmpName, DptName, PEDate"
    
    i = 0
    j = rs.RecordCount
    LastJobID = 999999
    FirstFlag = True
    
    PrtInit ("Land")
    SetFont 8, Equate.LandScape
    Columns = Columns - 15
    WBJHeader
    
    ' report TimeSheet records not entered in data entry
    If rsTS.RecordCount > 0 Then
        
        Prvw.vsp.Font.Bold = True
        PrintValue(1) = " ":                                         FormatString(1) = "a5"
        PrintValue(2) = "*** TimeSheets not entered in Payroll ***": FormatString(2) = "a50"
        PrintValue(3) = " ":                                         FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        rsTS.Sort = "WEDate, JobID, EmployeeID"
        rsTS.MoveFirst
        Do
            
            boo = JCJob.GetByID(rsTS!JobID)
            boo = PREmployee.GetByID(rsTS!EmployeeID)
            
            PrintValue(1) = PREmployee.LFName:                  FormatString(1) = "a20"
            PrintValue(2) = JCJob.Name:                         FormatString(2) = "a20"
            PrintValue(3) = " Week Ended:":                     FormatString(3) = "a13"
            PrintValue(4) = Format(rsTS!WEDate, " mm/dd/yy "):  FormatString(4) = "a10"
            PrintValue(5) = "Hours: ":                          FormatString(5) = "a9"
            PrintValue(6) = rsTS!TotalHours:                    FormatString(6) = "d7"
            PrintValue(7) = " ":                                FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            If Ln > MaxLines Then
                FormFeed
                WBJHeader
            End If
            
            rsTS.MoveNext
            
        Loop Until rsTS.EOF
    
        Ln = Ln + 1
    
    End If
    
    rs.MoveFirst
    Do
        
        ' break in job
        If LastJobID <> 999999 And rs!JobID <> LastJobID Then
            WBJSecFooter
            If Me.chkSepPage = 1 Then
                FormFeed
                WBJHeader
            Else
                Ln = Ln + 1
            End If
            FirstFlag = True
        End If
        LastJobID = rs!JobID
        
        ' form feed
        If Ln >= MaxLines Then
            FormFeed
            WBJHeader
            FirstFlag = True
            WBJSecHeader
            FirstFlag = False
        End If
        
        ' section header
        If FirstFlag = True Then WBJSecHeader
        FirstFlag = False
        
        PrintValue(1) = rs!EmpName:                         FormatString(1) = "a22"
        PrintValue(2) = rs!DptName:                         FormatString(2) = "a17"
        PrintValue(3) = Format(rs!PEDate, " mm/dd/yyyy"):   FormatString(3) = "a12"
        PrintValue(4) = rs!RegHrs:                          FormatString(4) = "d10"
        PrintValue(5) = rs!OvtHrs:                          FormatString(5) = "d10"
        PrintValue(6) = rs!OthHrs:                          FormatString(6) = "d10"
        PrintValue(7) = rs!TotHrs:                          FormatString(7) = "d10"
        PrintValue(8) = rs!RegAmt:                          FormatString(8) = "d13"
        PrintValue(9) = rs!OvtAmt:                          FormatString(9) = "d13"
        PrintValue(10) = rs!OthAmt:                         FormatString(10) = "d13"
        PrintValue(11) = rs!TotAmt:                         FormatString(11) = "d13"
        PrintValue(12) = " ":                               FormatString(12) = "~"
        FormatPrint
        Ln = Ln + 1

        ' accum totals
        For j = 1 To 2
            RegTot(j) = RegTot(j) + rs!RegHrs
            RegTot(j + 2) = RegTot(j + 2) + rs!RegAmt
            OvtTot(j) = OvtTot(j) + rs!OvtHrs
            OvtTot(j + 2) = OvtTot(j + 2) + rs!OvtAmt
            OthTot(j) = OthTot(j) + rs!OthHrs
            OthTot(j + 2) = OthTot(j + 2) + rs!OthAmt
        Next j

        rs.MoveNext
    
    Loop Until rs.EOF

    If Ln > MaxLines - 5 Then
        FormFeed
        WBJHeader
        FirstFlag = True
        WBJSecHeader
    End If

    WBJSecFooter
    If Me.chkSepPage = 1 Then
        FormFeed
        WBJHeader
    Else
        Ln = Ln + 1
    End If
    WBJSecFooter True
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    Unload Me

End Sub

Private Sub WBJHeader()

    PageHeader "Wage by Job report", Trim(Me.txtDisplay)
    Ln = Ln + 1
    
    PrintValue(1) = " ":                    FormatString(1) = "a65"
    PrintValue(2) = "H  O  U  R  S":        FormatString(2) = "a13"
    PrintValue(3) = " ":                    FormatString(3) = "a35"
    PrintValue(4) = "W  A  G  E  S":        FormatString(4) = "a13"
    PrintValue(5) = " ":                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                    FormatString(1) = "a51"
    PrintValue(2) = String(39, "-"):        FormatString(2) = "a39"
    PrintValue(3) = " ":                    FormatString(3) = "a2"
    PrintValue(4) = String(50, "-"):        FormatString(4) = "a50"
    PrintValue(5) = " ":                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "Employee":             FormatString(1) = "a22"
    PrintValue(2) = "Department":           FormatString(2) = "a17"
    PrintValue(3) = "   PE Date":           FormatString(3) = "a12"
    PrintValue(4) = "Reg Hours ":           FormatString(4) = "r10"
    PrintValue(5) = "Ovt Hours ":           FormatString(5) = "r10"
    PrintValue(6) = "Oth Hours ":           FormatString(6) = "r10"
    PrintValue(7) = "Tot Hours ":           FormatString(7) = "r10"
    PrintValue(8) = "Reg Wage ":            FormatString(8) = "r13"
    PrintValue(9) = "Ovt Wage ":            FormatString(9) = "r13"
    PrintValue(10) = "Other Wages ":        FormatString(10) = "r13"
    PrintValue(11) = "Tot Wages ":          FormatString(11) = "r13"
    PrintValue(12) = " ":                   FormatString(12) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(Columns - 7, "="): FormatString(1) = "a" & Columns - 7
    PrintValue(2) = " ":                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
End Sub

Private Sub WBJSecHeader()
            
    Prvw.vsp.Font.Bold = True
            
    ' get the job record
    rsJob.Find "JobID = " & rs!JobID, 0, adSearchForward, 1
    If rsJob.EOF Then
        MsgBox "Job error: " & rs!JobID, vbExclamation
        GoBack
    End If
    PrintValue(1) = Space(6) & rsJob!JobName:       FormatString(1) = "a110"
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
    Prvw.vsp.Font.Bold = False

End Sub

Private Sub WBJSecFooter(Optional Final As Boolean)

Dim Tot, TotHrs As Currency

    Prvw.vsp.Font.Bold = True
    
    If Final Then
        X = "FINAL TOTAL"
        k = 2
    Else
        X = "TOTAL " & Mid(rsJob!JobName, 1, 45)
        k = 1
    End If
    
    TotHrs = RegTot(k) + OvtTot(k) + OthTot(k)
    Tot = RegTot(k + 2) + OvtTot(k + 2) + OthTot(k + 2)
    
    PrintValue(1) = X:              FormatString(1) = "a51"
    PrintValue(2) = RegTot(k):      FormatString(2) = "d10"
    PrintValue(3) = OvtTot(k):      FormatString(3) = "d10"
    PrintValue(4) = OthTot(k):      FormatString(4) = "d10"
    PrintValue(5) = TotHrs:         FormatString(5) = "d10"
    PrintValue(6) = RegTot(k + 2):  FormatString(6) = "d13"
    PrintValue(7) = OvtTot(k + 2):  FormatString(7) = "d13"
    PrintValue(8) = OthTot(k + 2):  FormatString(8) = "d13"
    PrintValue(9) = Tot:            FormatString(9) = "d13"
    PrintValue(10) = " ":           FormatString(10) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = "Total Hours:": FormatString(1) = "a51"
    PrintValue(2) = RegTot(k) + OvtTot(k) + OthTot(k)
    FormatString(2) = "d10"
    PrintValue(3) = " ":            FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
    
    RegTot(1) = 0
    RegTot(3) = 0
    OvtTot(1) = 0
    OvtTot(3) = 0
    OthTot(1) = 0
    OthTot(3) = 0

    Prvw.vsp.Font.Bold = False

End Sub

Private Sub BillingReport()

Dim ct, RecCt As Long

    rs.CursorLocation = adUseClient
    rs.Fields.Append "JobID", adDouble
    rs.Fields.Append "EmpName", adVarChar, 20, adFldIsNullable
    rs.Fields.Append "EmpID", adDouble
    rs.Fields.Append "PEDate", adDate
    rs.Fields.Append "RegAmt", adCurrency
    rs.Fields.Append "RegHrs", adCurrency
    rs.Fields.Append "OvtAmt", adCurrency
    rs.Fields.Append "OvtHrs", adCurrency
    rs.Fields.Append "OthAmt", adCurrency
    rs.Fields.Append "OthHrs", adCurrency
    rs.Fields.Append "BillingRate", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    rsJob.CursorLocation = adUseClient
    rsJob.Fields.Append "JobID", adDouble
    rsJob.Fields.Append "JobName", adVarChar, 100, adFldIsNullable
    rsJob.Open , , adOpenDynamic, adLockOptimistic
    
    If BatchNumbr > 0 Then
        SQLString = "SELECT * FROM PRDist WHERE BatchID = " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRDist WHERE CheckDate >= " & CLng(StartDate) & _
                        " AND CheckDate <= " & CLng(EndDate)
        Else
            SQLString = "SELECT * FROM PRDist WHERE PEDate >= " & CLng(StartDate) & _
                        " AND PEDate <= " & CLng(EndDate)
        End If
    End If
    
    If PRDist.GetBySQL(SQLString) = False Then
        MsgBox "No PR data found!", vbExclamation
        GoBack
    End If

    frmProgress.Caption = "Wage report by job for: " & PRCompany.Name
    frmProgress.lblMsg1 = Me.txtDisplay
    frmProgress.Show
    frmProgress.Refresh

    ct = 0
    RecCt = PRDist.Records
    
    Do
        
        ct = ct + 1
        If ct Mod 10 = 1 Then
            frmProgress.lblMsg2 = Format(ct, "#,###,##0") & " Of: " & Format(RecCt, "#,###,##0")
            frmProgress.Refresh
        End If
        
        ' maintain job record set
        rsJob.Find "JobID = " & PRDist.JobID, 0, adSearchForward, 1
        If rsJob.EOF = True Then
            
            If PRDist.JobID = 0 Then
                rsJob.AddNew
                rsJob!JobName = PRCompany.Name
                rsJob!JobID = 0
                rsJob.Update
            Else
                If JCJob.GetByID(PRDist.JobID) = False Then
                    MsgBox "Job Not Found: " & PRDist.JobID, vbExclamation
                    GoBack
                End If
                
                rsJob.AddNew
                rsJob!JobName = JCJob.FullName
                rsJob!JobID = PRDist.JobID
                rsJob.Update
        
            End If
        
        End If

        If PREmployee.GetByID(PRDist.EmployeeID) = False Then
            MsgBox "Employee Not Found: " & PRDist.EmployeeID, vbExclamation
            GoBack
        End If

        rs.AddNew
        rs!JobID = PRDist.JobID
        rs!EmpName = Mid(PREmployee.LFName, 1, 20)
        rs!EmpID = PRDist.EmployeeID
        rs!PEDate = PRDist.PEDate
            
        rs!RegAmt = 0
        rs!RegHrs = 0
        rs!OvtAmt = 0
        rs!OvtHrs = 0
        rs!OthAmt = 0
        rs!OthHrs = 0
        
        If PRDist.DistType = PREquate.DistTypeReg Then
            rs!RegAmt = PRDist.Amount
            rs!RegHrs = PRDist.Hours
        ElseIf PRDist.DistType = PREquate.DistTypeOT Then
            rs!OvtAmt = PRDist.Amount
            rs!OvtHrs = PRDist.Hours
        Else
            rs!OthAmt = PRDist.Amount
            rs!OthHrs = PRDist.Hours
        End If
    
        rs!BillingRate = PRDist.BillingRate
    
        rs.Update
        
        If PRDist.GetNext = False Then Exit Do
    
    Loop

    If rsJob.RecordCount = 0 Then
        MsgBox "No Data Found", vbExclamation
        GoBack
    End If

    ' sort
    rsJob.Sort = "JobName"
    rs.Sort = "JobID, EmpName, PEDate"

    ' print
    frmProgress.lblMsg2 = "Now Printing ..."
    frmProgress.Refresh

    PrtInit ("Port")
    SetFont 8, Equate.Portrait
    Columns = Columns

    rsJob.MoveFirst
    Do
        
        If Ln = 0 Or Ln > MaxLines Then
            BBJHeader
        End If
        
        PrintValue(1) = " ":                    FormatString(1) = "a6"
        PrintValue(2) = Trim(rsJob!JobName):    FormatString(2) = "a70"
        PrintValue(3) = " ":                    FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
    
        rs.MoveFirst
        Do
            If rs!JobID <> rsJob!JobID Then GoTo NxtRS
            
            ' print detail line
            p1 = SuperRound(rs!RegHrs + rs!OvtHrs + rs!OthHrs, rs!BillingRate)
            P2 = rs!RegAmt + rs!OvtAmt + rs!OthAmt
            PrintValue(1) = rs!EmpName:                         FormatString(1) = "a18"
            PrintValue(2) = Format(rs!PEDate, " mm/dd/yy "):    FormatString(2) = "a10"
            PrintValue(3) = rs!BillingRate:                     FormatString(3) = "d12"
            PrintValue(4) = rs!RegHrs:                          FormatString(4) = "d8"
            PrintValue(5) = rs!OvtHrs:                          FormatString(5) = "d8"
            PrintValue(6) = rs!OthHrs:                          FormatString(6) = "d8"
            PrintValue(7) = p1:                                 FormatString(7) = "d10"
            PrintValue(8) = " ":                                FormatString(8) = "a1"
            PrintValue(9) = rs!RegAmt:                          FormatString(9) = "d10"
            PrintValue(10) = rs!OvtAmt:                         FormatString(10) = "d10"
            PrintValue(11) = rs!OthAmt:                         FormatString(11) = "d10"
            PrintValue(12) = P2:                                FormatString(12) = "d10"
            PrintValue(13) = " ":                               FormatString(13) = "~"
            FormatPrint
            Ln = Ln + 1

            ' update the totals
            For j = 1 To 2
                RegTot(j) = RegTot(j) + rs!RegHrs
                RegTot(j + 2) = RegTot(j + 2) + rs!RegAmt
                OvtTot(j) = OvtTot(j) + rs!OvtHrs
                OvtTot(j + 2) = OvtTot(j + 2) + rs!OvtAmt
                OthTot(j) = OthTot(j) + rs!OthHrs
                OthTot(j + 2) = OthTot(j + 2) + rs!OthAmt
            Next j

            BillTot = BillTot + p1
            BillTotG = BillTotG + p1

NxtRS:
            rs.MoveNext
        Loop Until rs.EOF

        ' job totals
        BBJFooter False
        
        rsJob.MoveNext
        
    Loop Until rsJob.EOF
    
    ' grand totals
    BBJFooter True

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    Unload Me

End Sub

Private Sub BBJFooter(ByVal Final As Boolean)

Dim Tot, TotHrs As Currency

    Prvw.vsp.Font.Bold = True
    
    If Final Then
        X = "FINAL TOTAL"
        k = 2
    Else
        X = "TOTAL " & Mid(rsJob!JobName, 1, 70)
        k = 1
    End If
    
    Tot = RegTot(k + 2) + OvtTot(k + 2) + OthTot(k + 2)
    
    If Final = True Then
        p1 = BillTotG
    Else
        p1 = BillTot
    End If
    
    PrintValue(1) = X:              FormatString(1) = "a40"
    PrintValue(2) = RegTot(k):      FormatString(2) = "d8"
    PrintValue(3) = OvtTot(k):      FormatString(3) = "d8"
    PrintValue(4) = OthTot(k):      FormatString(4) = "d8"
    PrintValue(5) = p1:             FormatString(5) = "d10"
    PrintValue(6) = " ":            FormatString(6) = "a1"
    PrintValue(7) = RegTot(k + 2):  FormatString(7) = "d10"
    PrintValue(8) = OvtTot(k + 2):  FormatString(8) = "d10"
    PrintValue(9) = OthTot(k + 2):  FormatString(9) = "d10"
    PrintValue(10) = Tot:           FormatString(10) = "d10"
    PrintValue(11) = " ":           FormatString(11) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "Total Hours:": FormatString(1) = "a40"
    PrintValue(2) = RegTot(k) + OvtTot(k) + OthTot(k)
    FormatString(2) = "d8"
    PrintValue(3) = " ":            FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2

    RegTot(1) = 0
    RegTot(3) = 0
    OvtTot(1) = 0
    OvtTot(3) = 0
    OthTot(1) = 0
    OthTot(3) = 0

    BillTot = 0

    Prvw.vsp.Font.Bold = False

End Sub

Private Sub BBJHeader()
    
    If Ln > 0 Then FormFeed
    
    PageHeader "Wage by Job report", Trim(Me.txtDisplay)
    Ln = Ln + 1
    
    PrintValue(1) = " ":                    FormatString(1) = "a40"
    PrintValue(2) = "H  O  U  R  S":        FormatString(2) = "a17"
    PrintValue(3) = " ":                    FormatString(3) = "a33"
    PrintValue(4) = "W  A  G  E  S":        FormatString(4) = "a13"
    PrintValue(5) = " ":                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                    FormatString(1) = "a30"
    PrintValue(2) = String(33, "-"):        FormatString(2) = "a33"
    PrintValue(3) = " ":                    FormatString(3) = "a13"
    PrintValue(4) = String(38, "-"):        FormatString(4) = "a38"
    PrintValue(5) = " ":                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "Employee":             FormatString(1) = "a18"
    PrintValue(2) = "  PE Date ":           FormatString(2) = "a10"
    PrintValue(3) = "Bill Rate ":           FormatString(3) = "r12"
    PrintValue(4) = "Reg Hrs ":             FormatString(4) = "r8"
    PrintValue(5) = "Ovt Hrs ":             FormatString(5) = "r8"
    PrintValue(6) = "Oth Hrs ":             FormatString(6) = "r8"
    PrintValue(7) = "$ Billed ":            FormatString(7) = "r10"
    PrintValue(8) = " ":                    FormatString(8) = "a1"
    PrintValue(9) = "Reg Wage ":            FormatString(9) = "r10"
    PrintValue(10) = "Ovt Wage ":           FormatString(10) = "r10"
    PrintValue(11) = "Oth Wage ":           FormatString(11) = "r10"
    PrintValue(12) = "Tot Wage ":           FormatString(12) = "r10"
    PrintValue(13) = " ":                   FormatString(13) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(Columns - 7, "="): FormatString(1) = "a" & Columns - 7
    PrintValue(2) = " ":                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub
