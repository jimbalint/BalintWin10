VERSION 5.00
Begin VB.Form frmErnDed 
   Caption         =   "Earnings and Deduction Summary"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
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
   ScaleHeight     =   9585
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold Print"
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   7560
      Width           =   3615
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Include All Options"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   7200
      Width           =   3615
   End
   Begin VB.CheckBox chkPrtCompany 
      Caption         =   "Print Company Totals"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   6600
      Width           =   3615
   End
   Begin VB.CheckBox chkPrtDept 
      Caption         =   "Print Department Totals"
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   6120
      Width           =   3615
   End
   Begin VB.CheckBox chkPrtEmployee 
      Caption         =   "Print Employee Detail"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CheckBox chkTaxInfo 
      Caption         =   "Print Employee Tax Information"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CheckBox chkDeductions 
      Caption         =   "Include Deduction Detail"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   7080
      Width           =   3615
   End
   Begin VB.CheckBox chkEarnings 
      Caption         =   "Include Other Earnings Detail"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   6600
      Width           =   3615
   End
   Begin VB.CheckBox chkHours 
      Caption         =   "Include Other Hours Detail"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6120
      Width           =   3615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Frame frsSortBy 
      Caption         =   "   Sort Order   "
      Height          =   1095
      Left            =   2708
      TabIndex        =   15
      Top             =   4200
      Width           =   3615
      Begin VB.OptionButton optSortName 
         Caption         =   "Name"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optSortNumber 
         Caption         =   "Number"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSelectEmp 
      Caption         =   "SELECT EMPLOYEES"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox cmbQuarter 
      Height          =   390
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox cmbYear 
      Height          =   390
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblEmployees 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Employees To Print:"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Quarter to Report:"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Year to Report:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "frmErnDed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rsTitle As New ADODB.Recordset
Dim rsOEDED As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim BOYDate, BOQDate, EOQDate As Date
Dim I, J, K, JB As Long
Dim X, Y, Z As String
Dim PRTL As Byte
Dim QY, PRTL1, PRTL2 As String
Dim p1, P2, P3 As Currency
Dim boo As Boolean
Dim QtrFlag As Boolean
Dim LastEmpID As Long
Dim DirectDeposit As Boolean
Dim ERQID, ERYID, EEQID, EEYID As Long
Dim BlankFlag As Boolean
Dim OECount, HRCount, DEDCount, DLines As Byte
Dim GlobalID As Long
Dim OEFlag As Boolean
Dim DEDFlag As Boolean
Dim OEDLineCount As Byte

Private Sub Form_Load()

    ' *********** TO DO *****************
    ' no items test
    '    no OE OR no DED test
    '    >>>>> disable check boxes <<<<<<
    ' if no QTD data but has YTD data
    ' include dir dep / SD tax ....
    ' hours option
    '    no hours - skip reg hrs line also
    '    hrs=yes / oe=no
    ' *********** TO DO *****************
    
    Me.lblCompanyName = PRCompany.Name
    Me.lblEmployees = "ALL EMPLOYEES"

    ' get the years in history
    rs.CursorLocation = adUseClient
    rs.Fields.Append "PRYear", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRBatch ORDER BY CheckDate DESC"
    If PRBatch.GetBySQL(SQLString) = False Then
        MsgBox "No Payroll Data Found!", vbExclamation
        GoBack
    End If
    
    Do
        rs.Find "PRYear = " & Int(PRBatch.YearMonth / 100), 0, adSearchForward, 1
        If rs.EOF Then
            rs.AddNew
            rs!PRYear = Int(PRBatch.YearMonth / 100)
            rs.Update
        End If
        If PRBatch.GetNext = False Then Exit Do
    Loop
    
    rs.MoveFirst
    Do
        Me.cmbYear.AddItem rs!PRYear
        rs.MoveNext
    Loop Until rs.EOF
    cmbYear.ListIndex = 0

    Me.cmbQuarter.AddItem "1"
    Me.cmbQuarter.AddItem "2"
    Me.cmbQuarter.AddItem "3"
    Me.cmbQuarter.AddItem "4"
    
    If cmbYrQtrSet(Me.cmbYear, Me.cmbQuarter) = False Then GoBack
    
    ' cmbQuarter.ListIndex = 0

    Me.optSortNumber = True

    frmEmpSelect.AllEmployees = True

    ' screen defaults
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeScreenDefault & _
                " AND Description = 'ErnDed'" & _
                " AND UserID = " & User.ID & _
                " AND Var1 = '" & PRCompany.CompanyID & "'"
    If PRGlobal.GetBySQL(SQLString) = True Then
    
        Me.chkPrtEmployee = PRGlobal.Byte1
        Me.chkPrtDept = PRGlobal.Byte2
        Me.chkPrtCompany = PRGlobal.Byte3
        Me.chkTaxInfo = PRGlobal.Byte4
        Me.chkHours = PRGlobal.Byte5
        Me.chkEarnings = PRGlobal.Byte6
        Me.chkDeductions = PRGlobal.Byte7
        Me.chkBold = PRGlobal.Byte8

    Else
        
        Me.chkDeductions = 1
        Me.chkEarnings = 1
        Me.chkHours = 1
        Me.chkPrtCompany = 1
        Me.chkPrtDept = 1
        Me.chkPrtEmployee = 1
        Me.chkTaxInfo = 1
    
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeScreenDefault
        PRGlobal.Description = "ErnDed"
        PRGlobal.UserID = User.ID
        PRGlobal.Var1 = PRCompany.CompanyID
        PRGlobal.Byte1 = 1
        PRGlobal.Byte2 = 1
        PRGlobal.Byte3 = 1
        PRGlobal.Byte4 = 1
        PRGlobal.Byte5 = 1
        PRGlobal.Byte6 = 1
        PRGlobal.Byte7 = 1
        PRGlobal.Byte8 = 0
        PRGlobal.Save (Equate.RecAdd)
    
    End If
    GlobalID = PRGlobal.GlobalID

    Me.KeyPreview = True

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdSelectEmp_Click()
    frmEmpSelect.Show vbModal
    If frmEmpSelect.AllEmployees = False Then
        Me.lblEmployees = frmEmpSelect.SelCount & " Employees Selected"
    End If
End Sub

Private Sub cmdOK_Click()

    If frmEmpSelect.SelCount = 0 And frmEmpSelect.AllEmployees = False Then
        MsgBox "No employees selected!", vbExclamation
        Exit Sub
    End If

    ' save defaults
    If PRGlobal.GetByID(GlobalID) Then
        PRGlobal.Byte1 = Me.chkPrtEmployee
        PRGlobal.Byte2 = Me.chkPrtDept
        PRGlobal.Byte3 = Me.chkPrtCompany
        PRGlobal.Byte4 = Me.chkTaxInfo
        PRGlobal.Byte5 = Me.chkHours
        PRGlobal.Byte6 = Me.chkEarnings
        PRGlobal.Byte7 = Me.chkDeductions
        PRGlobal.Byte8 = Me.chkBold
        PRGlobal.Save (Equate.RecPut)
    End If
        
    OEFlag = False
    DEDFlag = False
    If Me.chkHours Or Me.chkEarnings Then OEFlag = True
    If Me.chkDeductions Then DEDFlag = True
    
    ' direct deposit for this company?
    SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeDirDepDed
    DirectDeposit = PRItem.GetBySQL(SQLString)
    
    ProcessData

' frmProgress.Hide
' SetGrid rsOEDED, fg
' Exit Sub

    ' recordset of titles
    On Error Resume Next
    rsTitle.Close
    On Error GoTo 0
    rsTitle.CursorLocation = adUseClient
    rsTitle.Fields.Append "ItemType", adInteger
    rsTitle.Fields.Append "ItemID", adDouble
    rsTitle.Fields.Append "Title", adVarChar, 15, adFldIsNullable
    rsTitle.Fields.Append "Hours", adBoolean
    rsTitle.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0" & _
                " AND (ItemType = " & PREquate.ItemTypeOE & _
                " OR ItemType = " & PREquate.ItemTypeDED & _
                " OR ItemType = " & PREquate.ItemTypeSDTax & ")" & _
                " ORDER BY ItemType, ItemID"
    If PRItem.GetBySQL(SQLString) = True Then
        Do
            
            ' items to skip
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                If Me.chkHours Then
                    HRCount = HRCount + 1
                    AddTitle True
                End If
                If Me.chkEarnings Then
                    OECount = OECount + 1
                    AddTitle
                End If
            Else
                If Me.chkDeductions Then
                    DEDCount = DEDCount + 1
                    AddTitle
                End If
            End If
            If PRItem.GetNext = False Then Exit Do
        Loop
    
    End If
    
    ' add deduction item for direct deposit???
    If DirectDeposit Then
        rsTitle.AddNew
        rsTitle!ItemType = PREquate.ItemTypeDED
        rsTitle!ItemID = 999999
        rsTitle!Title = "DIR DEP"
        rsTitle!Hours = False
        rsTitle.Update
        DEDCount = DEDCount + 1
    End If
    
    PrtInit ("Port")
    SetFont 8, Equate.Portrait
    Columns = Columns - 7
    
    ' how many rows per header / entry ?
    DLines = 4       ' Q&Y Pay/Tax
    If Me.chkHours Then
        DLines = DLines + 2     ' reg/ovt/oth hrs
    End If
    DLines = DLines + 2 + (2 * Calc6(HRCount)) + (2 * Calc6(OECount)) + (2 * Calc6(DEDCount))
    
    ' include tax info ???
    If Me.chkTaxInfo Then DLines = DLines + 2
    
    If Me.chkPrtEmployee Then PrintReport PREquate.RecTypeEmployee
    
    If Me.chkPrtDept And rsDept.RecordCount > 0 Then
        If Ln <> 0 Then FormFeed
        PrintReport PREquate.RecTypeDepartment
    End If

    If Me.chkPrtCompany Then
        If Ln <> 0 Then FormFeed
        PrintReport PREquate.RecTypeEmployer
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

    GoBack

End Sub

Private Sub PrintReport(ByVal RecType As Byte)
    
Dim LastID As Long
    
    frmProgress.lblMsg2 = "Now Printing Report ..."
    frmProgress.Refresh
    
    PRTotal.TFilter ("RecType = " & RecType)
    If RecType = PREquate.RecTypeDepartment Then
        PRTotal.TSortByString ("RecType, IDNumber, PeriodType")
    ElseIf Me.optSortName = True Then
        PRTotal.TSortByString ("RecType, Name, PeriodType")
    Else
        PRTotal.TSortByString ("RecType, IDNumber, PeriodType")
    End If
    
    PRTotal.FindFirst
    
    LastID = 0
    
    EDHeader
    
    If Me.chkBold Then
        Prvw.vsp.Font.Bold = True
    End If
    
    Do
    
        frmProgress.lblMsg2 = "Printing: " & PRTotal.Name
        frmProgress.Refresh
    
        If PRTotal.PeriodType = PREquate.PeriodTypeQuarter And PRTotal.Count = 0 Then
            ' is the YTD blank also?
            boo = PRTotal.GetNext
            If boo = False Then Exit Do
            If PRTotal.Count = 0 Then
                GoTo NxtPRTotal     ' no qtd and ytd - next emp/dept
            Else
                PRTotal.GetPrev     ' has ytd - process it
            End If
        End If
        
        OEDLineCount = 0
    
        ' page break? - not between Q/Y for same entry
        If LastID <> 0 And LastID <> PRTotal.RecID Then
            Ln = Ln + 1
            If Ln > MaxLines - DLines + 2 Then
                FormFeed
                EDHeader
            End If
        End If
        LastID = PRTotal.RecID
    
        ' print tax info?
        If Me.chkTaxInfo _
           And PRTotal.RecType = PREquate.RecTypeEmployee _
           And PRTotal.PeriodType = PREquate.PeriodTypeQuarter Then

            boo = PREmployee.GetByID(PRTotal.RecID)

            PrintValue(1) = " ":                FormatString(1) = "a20"

            PrintValue(2) = "SS#: " & Format(PRTotal.SSN, "000-00-0000")
            FormatString(2) = "a17"

            PrintValue(3) = "Pay/Yr: " & PREmployee.PaysPerYear:        FormatString(3) = "a12"

            If PREmployee.FWTBasis = PREquate.BasisExemptions Then
                X = "FWT#: " & PREmployee.FWTAmount
                If PREmployee.FWTMarried = 1 Then
                    X = X & " M"
                Else
                    X = X & " S"
                End If
            Else
                X = "FWT%: " & Format(PREmployee.FWTAmount, "##0.00 %")
            End If
            PrintValue(4) = X:                  FormatString(4) = "a13"

            If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                X = "SWT#: " & PREmployee.SWTAmount
                If PREmployee.SWTMarried = 1 Then
                    X = X & " M"
                Else
                    X = X & " S"
                End If
            Else
                X = "SWT%: " & Format(PREmployee.SWTAmount, "##0.00 %")
            End If
            PrintValue(5) = X:                  FormatString(5) = "a13"

            I = 6
            If PREmployee.DefaultCityID <> 0 Then
                boo = PRCity.GetByID(PREmployee.DefaultCityID)
                I = 7
                X = "CWT%: " & Format(PRCity.CityRate, "##0.00 ") & PRCity.ShortName
                PrintValue(6) = X:              FormatString(6) = "a25"
            End If

            PrintValue(I) = " ":                FormatString(I) = "~"
            FormatPrint
            Ln = Ln + 1

            X = "SALARIED: "
            If PREmployee.Salaried = 1 Then
                X = X & "Y"
            Else
                X = X & "N"
            End If
            PrintValue(1) = " ":                    FormatString(1) = "a20"
            PrintValue(2) = X:                      FormatString(2) = "a17"
            PrintValue(3) = "Rate:":                FormatString(3) = "a5"
            If PREmployee.Salaried = 1 Then
                PrintValue(4) = PREmployee.SalaryAmount
            Else
                PrintValue(4) = PREmployee.HourlyAmount
            End If
            FormatString(4) = "d10"

            X = "Inactive: "
            If PREmployee.Inactive = 1 Then
                X = X & "Y"
            Else
                X = X & "N"
            End If
            PrintValue(5) = X:                      FormatString(5) = "a12"
            PrintValue(6) = " ":                    FormatString(6) = "~"
            FormatPrint

            Ln = Ln + 1

        End If
    
        If Me.chkHours Then
            JB = 3
        Else
            JB = 2
        End If
        For PRTL = 1 To JB
        
            If PRTL = 1 And PRTotal.PeriodType = PREquate.PeriodTypeQuarter Then
                If PRTotal.IDNumber <> 0 Then
                    X = PRTotal.IDNumber
                Else
                    X = ""
                End If
            Else
                X = AlphaNum(PRTL)
            End If
            PRTL1 = Space(5 - Len(X)) & X & " "
            
            If PRTotal.PeriodType = PREquate.PeriodTypeQuarter Then
                QY = "QTD "
            Else
                QY = "YTD "
            End If
        
            If PRTL = 1 Then
                If PRTotal.PeriodType = PREquate.PeriodTypeQuarter Then
                    If PRTotal.RecType = PREquate.RecTypeEmployer Then
                        PRTL2 = PRCompany.Name
                    Else
                        PRTL2 = PRTotal.Name
                    End If
                Else
                    PRTL2 = QY & "PAY"
                End If
            ElseIf PRTL = 2 Then
                PRTL2 = QY & "REG TAXES"
            Else
                PRTL2 = QY & "HOURS"
            End If
        
            PrintValue(1) = PRTL1:                  FormatString(1) = "a6"
            PrintValue(2) = PRTL2:                  FormatString(2) = "a20"
            
            K = 2
            For I = 1 To 6
                
                BlankFlag = False
                
                If PRTL = 1 Then
                    If I = 1 Then p1 = PRTotal.RegAmount
                    If I = 2 Then p1 = PRTotal.OTAmount
                    If I = 3 Then p1 = PRTotal.OEAmount
                    If I = 4 Then p1 = PRTotal.Gross
                    If I = 5 Then
                        p1 = PRTotal.SSTax + PRTotal.MedTax + PRTotal.FWTTax + PRTotal.StateTax + PRTotal.CityTax
                    End If
                    If I = 6 Then p1 = PRTotal.Net
                ElseIf PRTL = 2 Then
                    If I = 1 Then p1 = PRTotal.SSTax
                    If I = 2 Then p1 = PRTotal.MedTax
                    If I = 3 Then p1 = PRTotal.FWTTax
                    If I = 4 Then p1 = PRTotal.StateTax
                    If I = 5 Then p1 = PRTotal.CityTax
                    If I = 6 Then
                        p1 = PRTotal.SSTax + PRTotal.MedTax + PRTotal.FWTTax + PRTotal.StateTax + PRTotal.CityTax
                    End If
                Else
                    If I = 1 Then p1 = PRTotal.RegHours
                    If I = 2 Then p1 = PRTotal.OTHours
                    If I = 3 Then p1 = PRTotal.OEHours
                    If I = 4 Then BlankFlag = True
                    If I = 5 Then BlankFlag = True
                    If I = 6 Then
                        p1 = PRTotal.RegHours + PRTotal.OTHours + PRTotal.OEHours
                    End If
                End If
                
                If BlankFlag = False Then
                    PrintValue(I + K) = p1:             FormatString(I + K) = "d14"
                Else
                    PrintValue(I + K) = " ":            FormatString(I + K) = "a14"
                End If
            
            Next I
    
            PrintValue(9) = " ":        FormatString(9) = "~"
            FormatPrint
            Ln = Ln + 1
        
            ' OE / DED print
            If PRTL = JB Then
                If OEFlag Then
                    If Me.chkHours Then
                        PrintOEDED PREquate.ItemTypeOE, True ' hours
                    End If
                    If Me.chkEarnings Then
                        PrintOEDED PREquate.ItemTypeOE
                    End If
                End If
                If DEDFlag Then PrintOEDED PREquate.ItemTypeDED
                Ln = Ln + 1
            End If
        
        Next PRTL
        
NxtPRTotal:
        If PRTotal.GetNext = False Then Exit Do
   
    Loop

End Sub

Private Sub EDHeader()

Dim HeadCount, HeadLine As Byte

    Prvw.vsp.Font.Bold = True

    PageHeader "Earnings and Deduction Summary", _
               "Year: " & Me.cmbYear & " " & _
               "Quarter Ended: " & Me.cmbQuarter

    Ln = Ln + 1
    
    PrintValue(1) = Space(3) & "a) ":           FormatString(1) = "a6"
    PrintValue(2) = "Pay":                      FormatString(2) = "a20"
    For I = 1 To 6
        Select Case I
            Case 1:           X = "REGULAR PAY "
            Case 2:           X = "OVERTIME PAY "
            Case 3:           X = "TOTAL OE "
            Case 4:           X = "GROSS PAY "
            Case 5:           X = "TOTAL TAXES "
            Case 6:           X = "NET PAY "
        End Select
        PrintValue(I + 2) = X:          FormatString(I + 2) = "r14"
    Next I
    PrintValue(9) = " ":        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = Space(3) & "b) ":           FormatString(1) = "a6"
    PrintValue(2) = "Reg Taxes":                FormatString(2) = "a20"
    For I = 1 To 6
        Select Case I
            Case 1:           X = "SS TAX "
            Case 2:           X = "MED TAX "
            Case 3:           X = "FWT TAX "
            Case 4:           X = "SWT TAX "
            Case 5:           X = "CWT TAX "
            Case 6:           X = "TOT REG TAX "
        End Select
        PrintValue(I + 2) = X:          FormatString(I + 2) = "r14"
    Next I
    PrintValue(9) = " ":        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 1
    
    HeadLine = 2
    
    ' hours?
    If Me.chkHours Then
        PrintValue(1) = Space(3) & "c) ":           FormatString(1) = "a6"
        PrintValue(2) = "Hours":                    FormatString(2) = "a20"
        For I = 1 To 6
            Select Case I
                Case 1:           X = "REG HOURS "
                Case 2:           X = "OT HOURS "
                Case 3:           X = "OE HOURS "
                Case 4:           X = " "
                Case 5:           X = " "
                Case 6:           X = "TOT HOURS "
            End Select
            PrintValue(I + 2) = X:          FormatString(I + 2) = "r14"
        Next I
        PrintValue(9) = " ":        FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 1
        HeadLine = HeadLine + 1
    End If

    ' no items for this company
    If rsTitle.RecordCount = 0 Then
        Ln = Ln + 1
        Prvw.vsp.Font.Bold = False
        Exit Sub
    End If
    
    For I = 1 To 3      ' 1=Hours / 2=OE / 3=DED

        If I = 1 Then J = HRCount
        If I = 2 Then J = OECount
        If I = 3 Then J = DEDCount
        If J = 0 Then GoTo NxtI
        
        HeadCount = 0
        HeadLine = HeadLine + 1
        rsTitle.MoveFirst
        Do
            
            If I = 1 Then boo = rsTitle!ItemType = PREquate.ItemTypeOE And rsTitle!Hours = True
            If I = 2 Then boo = rsTitle!ItemType = PREquate.ItemTypeOE And rsTitle!Hours = False
            If I = 3 Then boo = rsTitle!ItemType = PREquate.ItemTypeDED
            If boo = True Then
            
                HeadCount = HeadCount + 1
                If HeadCount = 7 Then
                    PrintValue(HeadCount + 3) = " ":    FormatString(HeadCount + 3) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    HeadCount = 1
                    HeadLine = HeadLine + 1
                End If
                If HeadCount = 1 Then
                    X = AlphaNum(HeadLine) & " "
                    PrintValue(1) = X:                  FormatString(1) = "r6"
                    If I = 1 Then X = "OTHER HOURS"
                    If I = 2 Then X = "OTHER EARN"
                    If I = 3 Then X = "DEDUCTIONS"
                    PrintValue(2) = X:                  FormatString(2) = "a20"
                End If
                PrintValue(HeadCount + 2) = rsTitle!Title & " ":  FormatString(HeadCount + 2) = "r14"
            
            End If
            rsTitle.MoveNext
        Loop Until rsTitle.EOF
        
        If HeadCount <= 6 Then
            PrintValue(HeadCount + 3) = " ":    FormatString(HeadCount + 3) = "~"
            FormatPrint
            Ln = Ln + 1
            ' HeadLine = HeadLine + 1
        End If
        
NxtI:
    Next I

    Ln = Ln + 1
    
    Prvw.vsp.Font.Bold = False
    
End Sub

Private Sub PrintOEDED(ByVal ItemType, _
                       Optional Hrs As Boolean)
    
Dim FieldCount As Byte
    
    If rsTitle.RecordCount = 0 Then Exit Sub
    
    FieldCount = 0
    rsTitle.MoveFirst
    
    Do
        
        If rsTitle!ItemType <> ItemType Then GoTo NxtRSTitle
        If rsTitle!Hours <> Hrs Then GoTo NxtRSTitle
            
        FieldCount = FieldCount + 1
        
        If FieldCount Mod 6 = 1 Then
            
            If FieldCount <> 1 Then
                PrintValue(FieldCount + 3) = " ":       FormatString(FieldCount + 3) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
            
            OEDLineCount = OEDLineCount + 1
            FieldCount = 1
            
            X = AlphaNum(PRTL + OEDLineCount)
            X = Space(5 - Len(X)) & X
            PrintValue(1) = X:          FormatString(1) = "a6"
            
            If PRTotal.PeriodType = PREquate.PeriodTypeQuarter Then
                X = "QTD "
            Else
                X = "YTD "
            End If
            If Hrs = True Then
                X = X & "HR"
            ElseIf ItemType = PREquate.ItemTypeOE Then
                X = X & "ERN"
            Else
                X = X & "DED"
            End If
            PrintValue(2) = X:          FormatString(2) = "a20"
        
        End If
        
        ' try to find the value in rsOEDED
        SQLString = "RecType = " & PRTotal.RecType & _
                    " AND RecID = " & PRTotal.RecID & _
                    " AND PeriodType = " & PRTotal.PeriodType & _
                    " AND ItemID = " & rsTitle!ItemID
        rsOEDED.Filter = SQLString
 
        If rsOEDED.RecordCount > 0 Then
            If Hrs = True Then
                p1 = rsOEDED!Hours
            Else
                p1 = rsOEDED!Amount
            End If
        Else
            p1 = 0
        End If
        rsOEDED.Filter = adFilterNone
            
        PrintValue(FieldCount + 2) = p1:        FormatString(FieldCount + 2) = "d14"

NxtRSTitle:
        rsTitle.MoveNext
    
    Loop Until rsTitle.EOF
            
    PrintValue(FieldCount + 3) = " ":       FormatString(FieldCount + 3) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub AddTitle(Optional Hrs As Boolean)
            
    rsTitle.AddNew
    
    If PRItem.ItemType = PREquate.ItemTypeSDTax Then
        rsTitle!ItemType = PREquate.ItemTypeDED
    Else
        rsTitle!ItemType = PRItem.ItemType
    End If
    rsTitle!ItemID = PRItem.ItemID
    If Hrs = True Then
        rsTitle!Title = Mid(PRItem.Abbreviation, 1, 12) & " HR"
    Else
        rsTitle!Title = Mid(PRItem.Abbreviation, 1, 15)
    End If
    rsTitle!Hours = Hrs
    rsTitle.Update

End Sub

Private Function Calc6(ByVal Count As Byte) As Byte

    If IsNull(Count) Then
        Calc6 = 0
    ElseIf Count = 0 Then
        Calc6 = 0
    ElseIf Count Mod 6 = 0 Then
        Calc6 = Int(Count / 6)
    Else
        Calc6 = Int(Count / 6) + 1
    End If

End Function

Private Sub ProcessData()

    ' standard fields in PRTotal Class
    '    field added for period type - Year or Quarter
    ' TempRS for OE/DED titles / separate RS for amounts

    frmProgress.Caption = "Earnings and Deduction Summary Report"
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.lblMsg2 = "Initializing Report ..."
    frmProgress.Show

    ' set up the dates
    BOYDate = DateSerial(Me.cmbYear, 1, 1)
    I = (Me.cmbQuarter.ListIndex + 1) * 3 - 2
    BOQDate = DateSerial(Me.cmbYear, I, 1)
    EOQDate = DateSerial(Me.cmbYear, I + 3, 1) - 1

    PRTotal.CreateRS
    
    On Error Resume Next
    rsOEDED.Close
    On Error GoTo 0
    rsOEDED.CursorLocation = adUseClient
    rsOEDED.Fields.Append "RecType", adInteger      ' Employee / Dept / Total
    rsOEDED.Fields.Append "RecID", adDouble
    rsOEDED.Fields.Append "PeriodType", adInteger
    rsOEDED.Fields.Append "ItemID", adDouble
    rsOEDED.Fields.Append "Hours", adCurrency
    rsOEDED.Fields.Append "Amount", adCurrency
    rsOEDED.Open , , adOpenDynamic, adLockOptimistic

    ' create total records for Dept and Company
    On Error Resume Next
    rsDept.Close
    On Error GoTo 0
    rsDept.CursorLocation = adUseClient
    rsDept.Fields.Append "DeptID", adDouble
    rsDept.Fields.Append "QID", adDouble
    rsDept.Fields.Append "YID", adDouble
    rsDept.Open , , adOpenDynamic, adLockOptimistic
    
    PRTotal.Clear
    PRTotal.RecType = PREquate.RecTypeEmployer
    PRTotal.RecID = 0
    PRTotal.PeriodType = PREquate.PeriodTypeQuarter
    PRTotal.Save (Equate.RecAdd)
    ERQID = PRTotal.PRTotalID
    
    PRTotal.PeriodType = PREquate.PeriodTypeYear
    PRTotal.Save (Equate.RecAdd)
    ERYID = PRTotal.PRTotalID
    
    SQLString = "SELECT * FROM PRDepartment"
    If PRDepartment.GetBySQL(SQLString) Then
        Do
            PRTotal.Clear
            PRTotal.RecType = PREquate.RecTypeDepartment
            PRTotal.RecID = PRDepartment.DepartmentID
            PRTotal.Name = PRDepartment.Name
            PRTotal.IDNumber = PRDepartment.DepartmentNumber
            PRTotal.PeriodType = PREquate.PeriodTypeQuarter
            PRTotal.Save (Equate.RecAdd)
            
            rsDept.AddNew
            rsDept!DeptID = PRDepartment.DepartmentID
            rsDept!QID = PRTotal.PRTotalID
            
            PRTotal.PeriodType = PREquate.PeriodTypeYear
            PRTotal.Save (Equate.RecAdd)
                    
            rsDept!YID = PRTotal.PRTotalID
            rsDept.Update
            
            If PRDepartment.GetNext = False Then Exit Do
        Loop
    End If
    
    ' if not all employees selected - separate SQL for each selected
    '    else - one sql for all data
    With frmEmpSelect
        
        If .AllEmployees = True Or .SelCount > 5 Then
            SQLString = "SELECT * FROM PRHist WHERE CheckDate >= " & CLng(BOYDate) & _
                        " AND CheckDate <= " & CLng(EOQDate) & _
                        " ORDER BY EmployeeID"
        Else
            SQLString = "SELECT * FROM PRHist WHERE CheckDate >= " & CLng(BOYDate) & _
                        " AND CheckDate <= " & CLng(EOQDate) & " AND ("
            
            I = 0
            .rsEmp.Filter = "Select = true"
            .rsEmp.MoveFirst
            Do
                SQLString = SQLString & " EmployeeID = " & .rsEmp!EmployeeID
                I = I + 1
                If I = .rsEmp.RecordCount Then
                    SQLString = SQLString & ") "
                Else
                    SQLString = SQLString & " OR "
                End If
                .rsEmp.MoveNext
            Loop Until .rsEmp.EOF
        
            SQLString = SQLString & "ORDER BY EmployeeID"
        
        End If
    
    End With
    
    If PRHist.GetBySQL(SQLString) = False Then
        MsgBox "No PR History data found for the range given!", vbExclamation
        Exit Sub
    End If

    LastEmpID = 0
    
    Do
    
        '  change in employee
        If LastEmpID = 0 Or LastEmpID <> PRHist.EmployeeID Then
            
            If PREmployee.GetByID(PRHist.EmployeeID) Then
                frmProgress.lblMsg2 = "Processing: " & PREmployee.LFName
                frmProgress.Refresh
            End If
        
            ' create employee PRTotal records
            PRTotal.Clear
            PRTotal.RecType = PREquate.RecTypeEmployee
            PRTotal.RecID = PREmployee.EmployeeID
            PRTotal.EmployeeID = PREmployee.EmployeeID
            PRTotal.IDNumber = PREmployee.EmployeeNumber
            PRTotal.Name = PREmployee.LFName
            PRTotal.SSN = PREmployee.SSN
            PRTotal.PeriodType = PREquate.PeriodTypeQuarter
            PRTotal.Save (Equate.RecAdd)
            EEQID = PRTotal.PRTotalID
            
            PRTotal.PeriodType = PREquate.PeriodTypeYear
            PRTotal.Save (Equate.RecAdd)
            EEYID = PRTotal.PRTotalID
        
        End If
        LastEmpID = PRHist.EmployeeID

        ' process the PRHist record - EE & ER totals
        If PRHist.CheckDate >= BOQDate And PRHist.CheckDate <= EOQDate Then
            QtrFlag = True
        Else
            QtrFlag = False
        End If
            
        If QtrFlag = True Then
            ProcessHist (EEQID)
            ProcessHist (ERQID)
        End If
        ProcessHist (EEYID)
        ProcessHist (ERYID)
        
        ' process dept totals
        rsDept.Find "DeptID = " & PRHist.DepartmentID, 0, adSearchForward, 1
        If rsDept.EOF = False Then
            If QtrFlag = True Then
                ProcessHist (rsDept!QID)
            End If
            ProcessHist (rsDept!YID)
        End If
            
        If OEFlag = True Then ProcessOE
        If DEDFlag = True Then ProcessDED
        If DirectDeposit = True Then ProcessDirDep
            
        If PRHist.GetNext = False Then Exit Do
    
    Loop

End Sub
Private Sub ProcessHist(ByVal PRTID As Long)

    boo = PRTotal.GetByID(PRTID)

    PRTotal.RegAmount = PRTotal.RegAmount + PRHist.RegAmount
    PRTotal.OTAmount = PRTotal.OTAmount + PRHist.OTAmount
    PRTotal.OEAmount = PRTotal.OEAmount + PRHist.OEAmount
    PRTotal.Gross = PRTotal.Gross + PRHist.Gross
    PRTotal.SSTax = PRTotal.SSTax + PRHist.SSTax
    PRTotal.MedTax = PRTotal.MedTax + PRHist.MedTax
    PRTotal.FWTTax = PRTotal.FWTTax + PRHist.FWTTax
    PRTotal.StateTax = PRTotal.StateTax + PRHist.SWTTax
    PRTotal.CityTax = PRTotal.CityTax + PRHist.CWTTax
    PRTotal.Net = PRTotal.Net + PRHist.Net
    PRTotal.DirectDeposit = PRTotal.DirectDeposit + PRHist.DirectDeposit
    
    PRTotal.RegHours = PRTotal.RegHours + PRHist.RegHours
    PRTotal.OTHours = PRTotal.OTHours + PRHist.OTHours
    PRTotal.OEHours = PRTotal.OEHours + PRHist.OEHours
    
    PRTotal.Count = PRTotal.Count + 1
    
    PRTotal.Save (Equate.RecPut)

End Sub
Private Sub ProcessOE()
        
    ' other earnings for PRDist
    SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
    If PRDist.GetBySQL(SQLString) = False Then Exit Sub
        
    Do
        If QtrFlag = True Then
            
            ProcessItem PREquate.RecTypeEmployee, _
                        PREmployee.EmployeeID, _
                        PREquate.PeriodTypeQuarter, _
                        PRDist.EmployerItemID, _
                        PRDist.Amount, _
                        PRDist.Hours
                        
            ProcessItem PREquate.RecTypeEmployer, _
                        0, _
                        PREquate.PeriodTypeQuarter, _
                        PRDist.EmployerItemID, _
                        PRDist.Amount, _
                        PRDist.Hours
                        
        End If
                        
        ProcessItem PREquate.RecTypeEmployee, _
                    PREmployee.EmployeeID, _
                    PREquate.PeriodTypeYear, _
                    PRDist.EmployerItemID, _
                    PRDist.Amount, _
                    PRDist.Hours
                    
        ProcessItem PREquate.RecTypeEmployer, _
                    0, _
                    PREquate.PeriodTypeYear, _
                    PRDist.EmployerItemID, _
                    PRDist.Amount, _
                    PRDist.Hours
                        
        rsDept.Find "DeptID = " & PRHist.DepartmentID, 0, adSearchForward, 1
        If rsDept.EOF = False Then
        
            If QtrFlag = True Then
                
                ProcessItem PREquate.RecTypeDepartment, _
                            rsDept!DeptID, _
                            PREquate.PeriodTypeQuarter, _
                            PRDist.EmployerItemID, _
                            PRDist.Amount, _
                            PRDist.Hours
                            
            End If
                                
            ProcessItem PREquate.RecTypeDepartment, _
                        rsDept!DeptID, _
                        PREquate.PeriodTypeYear, _
                        PRDist.EmployerItemID, _
                        PRDist.Amount, _
                        PRDist.Hours
                        
        End If
    
        If PRDist.GetNext = False Then Exit Do

    Loop
    
End Sub

Private Sub ProcessDED()
    
    ' deductions from PRItemHist
    SQLString = "SELECT * FROM PRItemHist WHERE HistID = " & PRHist.HistID
    If PRItemHist.GetBySQL(SQLString) = False Then Exit Sub
        
    Do
            
        If QtrFlag = True Then
            
            ProcessItem PREquate.RecTypeEmployee, _
                        PREmployee.EmployeeID, _
                        PREquate.PeriodTypeQuarter, _
                        PRItemHist.EmployerItemID, _
                        PRItemHist.Amount, 0
                        
            ProcessItem PREquate.RecTypeEmployer, _
                        0, _
                        PREquate.PeriodTypeQuarter, _
                        PRItemHist.EmployerItemID, _
                        PRItemHist.Amount, 0
                        
        End If
                        
        ProcessItem PREquate.RecTypeEmployee, _
                    PREmployee.EmployeeID, _
                    PREquate.PeriodTypeYear, _
                    PRItemHist.EmployerItemID, _
                    PRItemHist.Amount, 0
                    
        ProcessItem PREquate.RecTypeEmployer, _
                    0, _
                    PREquate.PeriodTypeYear, _
                    PRItemHist.EmployerItemID, _
                    PRItemHist.Amount, 0

        rsDept.Find "DeptID = " & PRHist.DepartmentID, 0, adSearchForward, 1
        If rsDept.EOF = False Then
        
            If QtrFlag = True Then
                
                ProcessItem PREquate.RecTypeDepartment, _
                            rsDept!DeptID, _
                            PREquate.PeriodTypeQuarter, _
                            PRItemHist.EmployerItemID, _
                            PRItemHist.Amount, 0
                            
            End If
                            
            ProcessItem PREquate.RecTypeDepartment, _
                        rsDept!DeptID, _
                        PREquate.PeriodTypeYear, _
                        PRItemHist.EmployerItemID, _
                        PRItemHist.Amount, 0
                        
        End If
    
        If PRItemHist.GetNext = False Then Exit Do

    Loop
    
End Sub
Private Sub ProcessDirDep()
    
    If QtrFlag = True Then
        
        ProcessItem PREquate.RecTypeEmployee, _
                    PREmployee.EmployeeID, _
                    PREquate.PeriodTypeQuarter, _
                    999999, _
                    PRHist.DirectDeposit, 0
                    
        ProcessItem PREquate.RecTypeEmployer, _
                    0, _
                    PREquate.PeriodTypeQuarter, _
                    999999, _
                    PRHist.DirectDeposit, 0
                    
    End If
                    
    ProcessItem PREquate.RecTypeEmployee, _
                PREmployee.EmployeeID, _
                PREquate.PeriodTypeYear, _
                999999, _
                PRHist.DirectDeposit, 0
                
    ProcessItem PREquate.RecTypeEmployer, _
                0, _
                PREquate.PeriodTypeYear, _
                999999, _
                PRHist.DirectDeposit, 0

    rsDept.Find "DeptID = " & PRHist.DepartmentID, 0, adSearchForward, 1
    If rsDept.EOF = False Then
    
        If QtrFlag = True Then
            
            ProcessItem PREquate.RecTypeDepartment, _
                        rsDept!DeptID, _
                        PREquate.PeriodTypeQuarter, _
                        999999, _
                        PRHist.DirectDeposit, 0
                        
        End If
                        
        ProcessItem PREquate.RecTypeDepartment, _
                    rsDept!DeptID, _
                    PREquate.PeriodTypeYear, _
                    999999, _
                    PRHist.DirectDeposit, 0
                    
    End If

End Sub


Private Sub ProcessItem(ByVal RecType As Byte, _
                        ByVal RecID As Long, _
                        ByVal PeriodType As Byte, _
                        ByVal ItemID As Long, _
                        ByVal Amount As Currency, _
                        ByVal ItemHours As Currency)
                        
Dim Flg As Boolean

    Flg = False
    If rsOEDED.RecordCount = 0 Then
    Else
        rsOEDED.MoveFirst
        Do
            If rsOEDED!RecType = RecType And _
               rsOEDED!RecID = RecID And _
               rsOEDED!PeriodType = PeriodType And _
               rsOEDED!ItemID = ItemID Then
                    Flg = True
                    Exit Do
            End If
            rsOEDED.MoveNext
        Loop Until rsOEDED.EOF
    End If
    
    If Flg = False Then
        rsOEDED.AddNew
        rsOEDED!RecType = RecType
        rsOEDED!RecID = RecID
        rsOEDED!PeriodType = PeriodType
        rsOEDED!ItemID = ItemID
        rsOEDED!Amount = 0
        rsOEDED!Hours = 0
    End If
    
    rsOEDED!Amount = rsOEDED!Amount + Amount
    rsOEDED!Hours = rsOEDED!Hours + ItemHours
    
    rsOEDED.Update

End Sub

Private Function AlphaNum(ByVal aNum As Byte) As String
    Select Case aNum
        Case 1:       AlphaNum = "a)"
        Case 2:       AlphaNum = "b)"
        Case 3:       AlphaNum = "c)"
        Case 4:       AlphaNum = "d)"
        Case 5:       AlphaNum = "e)"
        Case 6:       AlphaNum = "f)"
        Case 7:       AlphaNum = "g)"
        Case 8:       AlphaNum = "h)"
        Case 9:       AlphaNum = "i)"
        Case 10:       AlphaNum = "j)"
        Case 11:       AlphaNum = "k)"
        Case 12:       AlphaNum = "l)"
        Case 13:       AlphaNum = "m)"
        Case 14:       AlphaNum = "n)"
        Case 15:       AlphaNum = "o)"
        Case 16:       AlphaNum = "p)"
        Case 17:       AlphaNum = "q)"
        Case 18:       AlphaNum = "r)"
        Case 19:       AlphaNum = "s)"
        Case 20:       AlphaNum = "t)"
        Case 21:       AlphaNum = "u)"
        Case 22:       AlphaNum = "v)"
        Case 23:       AlphaNum = "w)"
        Case 24:       AlphaNum = "x)"
        Case 25:       AlphaNum = "y)"
        Case 26:       AlphaNum = "z)"
        Case Else:     AlphaNum = " )"
    End Select
End Function

Private Sub cmdSelectAll_Click()
    Me.chkDeductions = 1
    Me.chkEarnings = 1
    Me.chkHours = 1
    Me.chkPrtCompany = 1
    Me.chkPrtDept = 1
    Me.chkPrtEmployee = 1
    Me.chkTaxInfo = 1
    Me.Refresh
End Sub

