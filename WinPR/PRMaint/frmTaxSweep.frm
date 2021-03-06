VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmTaxSweep 
   Caption         =   "Taxable Wage Sweep"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate tdbStartDate 
      Height          =   375
      Left            =   653
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calendar        =   "frmTaxSweep.frx":0000
      Caption         =   "frmTaxSweep.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTaxSweep.frx":0176
      Keys            =   "frmTaxSweep.frx":0194
      Spin            =   "frmTaxSweep.frx":01F2
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
      Text            =   "07/25/2009"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40019
      CenturyMode     =   0
   End
   Begin VB.CheckBox chkAllCheckDates 
      Caption         =   "ALL Check Dates:"
      Height          =   375
      Left            =   3233
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   5453
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   2333
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   390
      Left            =   4553
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox chkAllYears 
      Caption         =   "Run for ALL Years"
      Height          =   495
      Left            =   3233
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin TDBDate6Ctl.TDBDate tdbEndDate 
      Height          =   375
      Left            =   4613
      TabIndex        =   4
      Top             =   3240
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calendar        =   "frmTaxSweep.frx":021A
      Caption         =   "frmTaxSweep.frx":031A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTaxSweep.frx":038C
      Keys            =   "frmTaxSweep.frx":03AA
      Spin            =   "frmTaxSweep.frx":0408
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
      Text            =   "07/25/2009"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40019
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Select Year to Run:"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxSweep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Yr, LastYear As Long
Dim rsState As New ADODB.Recordset
Dim rsERItem As New ADODB.Recordset

Private Sub Form_Load()

    ' init recordset of state unem max
    rsState.CursorLocation = adUseClient
    rsState.Fields.Append "StateID", adDouble
    rsState.Fields.Append "UnEmpMax", adCurrency
    rsState.Fields.Append "UnEmpRmn", adCurrency
    rsState.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRState"
    If PRState.GetBySQL(SQLString) Then
        Do
            rsState.AddNew
            rsState!StateID = PRState.StateID
            rsState!UnEmpMax = PRState.UnEmpMax
            rsState!UnEmpRmn = 0
            rsState.Update
            If Not PRState.GetNext Then Exit Do
        Loop
    End If

    ' init the recordset of ER items
    rsERItem.CursorLocation = adUseClient
    rsERItem.Fields.Append "ItemID", adDouble
    rsERItem.Fields.Append "NoSSTax", adBoolean
    rsERItem.Fields.Append "NoMEDTax", adBoolean
    rsERItem.Fields.Append "NoFWTTax", adBoolean
    rsERItem.Fields.Append "NoSWTTax", adBoolean
    rsERItem.Fields.Append "NoCWTTax", adBoolean
    rsERItem.Fields.Append "NoSUNTax", adBoolean
    rsERItem.Fields.Append "NoFUNTax", adBoolean
    rsERItem.Open , , adOpenDynamic, adLockOptimistic

    ' get record set of Employer Items
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 AND " & _
                "(ItemType = " & PREquate.ItemTypeOE & " OR " & _
                "ItemType = " & PREquate.ItemTypeDED & ")"
    
    If PRItem.GetBySQL(SQLString) Then
        Do
            rsERItem.AddNew
            rsERItem!ItemID = PRItem.ItemID
            rsERItem!NoSSTax = PRItem.NoSSTax
            rsERItem!NoMedTax = PRItem.NoMedTax
            rsERItem!NoFWTTax = PRItem.NoFWTTax
            rsERItem!NoSWTTax = PRItem.NoSWTTax
            rsERItem!NoCWTTax = PRItem.NoCWTTax
            rsERItem!NoFUNTax = PRItem.NoFUNTax
            rsERItem!NoSUNTax = PRItem.NoSUNTax
            rsERItem.Update
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If

    ' form setups
    Me.chkAllYears = 1
    Me.lblCompanyName = PRCompany.Name
    Me.Label1.Enabled = False
    Me.cmbTaxYear.Enabled = False
    Me.tdbStartDate.Visible = False
    Me.tdbEndDate.Visible = False
    Me.chkAllCheckDates = 1

    ' get the years available
    SQLString = "SELECT * FROM PRHist ORDER BY YearMonth DESC"
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History Found!", vbCritical
        GoBack
    End If

    LastYear = 0
    Do
        Yr = Int(PRHist.YearMonth / 100)
        If LastYear <> Yr Then
            Me.cmbTaxYear.AddItem Yr
        End If
        LastYear = Yr
        If Not PRHist.GetNext Then Exit Do
    Loop
    Me.cmbTaxYear.ListIndex = 0

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub chkAllYears_Click()
    
    If Me.chkAllYears Then
        Me.Label1.Enabled = False
        Me.cmbTaxYear.Enabled = False
        Me.tdbStartDate.Visible = False
        Me.tdbEndDate.Visible = False
        Me.chkAllCheckDates = 1
    Else
        Me.Label1.Enabled = True
        Me.cmbTaxYear.Enabled = True
    End If

End Sub
Private Sub chkAllCheckDates_Click()
    
Dim Dt As Date
    
    If Me.chkAllCheckDates Then
        Me.tdbStartDate.Visible = False
        Me.tdbEndDate.Visible = False
    Else
        Me.tdbStartDate.Visible = True
        Dt = DateSerial(Me.cmbTaxYear.Text, 1, 1)
        tdbDateSet Me.tdbStartDate, Dt
        Me.tdbEndDate.Visible = True
        Dt = DateSerial(Me.cmbTaxYear.Text, 12, 31)
        tdbDateSet Me.tdbEndDate, Dt
        Me.chkAllYears = 0
    End If
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()
    
Dim i As Long
    
    If Me.chkAllYears = 0 Then
        TaxWageCalc Me.cmbTaxYear.Text
    Else
        For i = 0 To Me.cmbTaxYear.ListCount - 1
            Me.cmbTaxYear.ListIndex = i
            TaxWageCalc Me.cmbTaxYear.Text
        Next i
    End If
    
    MsgBox "Taxable Wage Calculation complete for " & Trim(PRCompany.Name), vbInformation
    
    GoBack

End Sub

Private Sub TaxWageCalc(ByVal TaxYear As Long)

Dim CWTDeduct, SSMax, FUNMax, SUNMax As Currency
Dim YTDSSRmn, YTDFUNRmn, YTDSUNRmn As Currency
Dim SSWage, MEDWage, FWTWage, SWTWage, CWTWage, FUNWage, SUNWage As Currency
Dim SSWageBase, FUNWageBase, SUNWageBase As Currency

Dim StartYM, EndYM As Long
Dim LastEE As Long
Dim p1, P2 As Currency

Dim WkcPct As Double
Dim WkcId As Long
Dim DistCount As Long


    ' ***********************************************************
    ' multi state pay not allowed on same check
    ' PRDist.StateWage is either the PRDist amount or zero
    ' no need to split up deduction amounts from state tax
    
    ' SD Export - get PRD city wage?
    '    different versions for different .tb files?
    '    if imported don't override here
    
    ' quarterly reports QC - run for each state
    ' is the UnEmp pct per state?  does FUN% or SUN% vary at all???
    
    ' DE testing
    '    one state per PRHist
    '    assign PRHist.StateID and PRHist.SUNWage
    
    ' test
    '     non-tax items
    '     multi state tests
    '     SS max test
    
    ' WKC??? EIC???
    
    ' ***********************************************************

    frmProgress.Caption = "Payroll Taxable Wage Set"
    frmProgress.lblMsg1 = Trim(PRCompany.Name) & " " & TaxYear
    frmProgress.Show

    ' get max wage amounts
    SSMax = PRGlobal.GetAmount(PREquate.GlobalTypeSSMax, TaxYear)
    FUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, TaxYear)
    SUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeSUNMax, TaxYear)

    StartYM = TaxYear * 100 + 1
    EndYM = TaxYear * 100 + 12

    ' sweep to fix PRItemHist.EmployerItemID not assigned
    ' for Escott - 03/02/2010
    SQLString = "SELECT * FROM PRItemHist WHERE EmployerItemID = 0 " & _
                " AND YearMonth >= " & StartYM & _
                " AND YearMonth <= " & EndYM
    If PRItemHist.GetBySQL(SQLString) Then
        Do
            If PRItem.GetByID(PRItemHist.ItemID) Then
                PRItemHist.EmployerItemID = PRItem.EmployerItemID
                PRItemHist.Save (Equate.RecPut)
            End If
            If PRItemHist.GetNext = False Then Exit Do
        Loop
    End If

    If Me.chkAllCheckDates Then
        SQLString = "SELECT * FROM PRHist WHERE " & _
                    "YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                    " ORDER BY EmployeeID, CheckDate, HistID"
    Else
        SQLString = "SELECT * FROM PRHist WHERE " & _
                    "YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                    " AND CheckDate < " & CLng(Me.tdbEndDate) & _
                    " ORDER BY EmployeeID, CheckDate, HistID"
    End If
    
    If Not PRHist.GetBySQL(SQLString) Then Exit Sub
        
    Recs = PRHist.Records
        
    Do
        
        Ct = Ct + 1
        If Ct Mod 100 = 1 Then
            frmProgress.lblMsg2 = "On Record: " & Format(Ct, "##,###,##0") & " Of: " & _
                                  Format(Recs, "##,###,##0")
            frmProgress.Refresh
        End If
        
        ' clear totals
        If LastEE = 0 Or PRHist.EmployeeID <> LastEE Then
            
            YTDSSRmn = SSMax
            YTDFUNRmn = FUNMax
            YTDSUNRmn = SUNMax
        
            ' clear the state totals
            rsState.MoveFirst
            Do
                rsState!UnEmpRmn = rsState!UnEmpMax
                rsState.Update
                rsState.MoveNext
            Loop Until rsState.EOF
        
            ' get the employee record
            If Not PREmployee.GetByID(PRHist.EmployeeID) Then
                MsgBox "Employee record NF: " & PRHist.EmployeeID & vbCr & PRHist.HistID, vbCritical
                End
            End If
        
            ' get Wkc Pct
            WkcId = 0
            If PREmployee.WkcUseDept = 1 Then
                If PRDepartment.GetByID(PREmployee.DepartmentID) Then
                    WkcId = PRDepartment.WkcCat
                End If
            Else
                WkcId = PREmployee.WkcCat
            End If
            If WkcId = 0 Then
                WkcPct = 0
            Else
                If PRGlobal.GetByID(WkcId) Then
                    WkcPct = PRGlobal.Percent
                End If
            End If
            
            frmProgress.lblMsg2 = PREmployee.LFName
            frmProgress.Show
        
        End If
        LastEE = PRHist.EmployeeID
        
        ' just accum YTD if running for a check date range
        ' and check date is before start date
        If Me.chkAllCheckDates = 0 Then
            If PRHist.CheckDate < Me.tdbStartDate Then
                
                YTDSSRmn = YTDSSRmn - PRHist.SSWage
                YTDFUNRmn = YTDFUNRmn - PRHist.FUNWage
                
                ' track SUN by State
                SQLString = "StateID = " & PRHist.StateID
                rsState.Find SQLString, 0, adSearchForward, 1
                If rsState.EOF Then
                    MsgBox "StateID NF in PRHist: " & PRHist.StateID & vbCr & PRHist.HistID, vbCritical
                    End
                End If
                rsState!UnEmpRmn = rsState!UnEmpRmn - PRHist.SUNWage
                rsState.Update
            
                ' skip to next PRHist record
                GoTo NextPRHist
            
            End If
        End If
        
        SSWage = PRHist.Gross
        MEDWage = PRHist.Gross
        FWTWage = PRHist.Gross
        SWTWage = PRHist.Gross
        CWTWage = PRHist.Gross
        FUNWage = PRHist.Gross
        SUNWage = PRHist.Gross
 
        ' loop prdist
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
        If PRDist.GetBySQL(SQLString) Then
            Do
                
                ' get the Employee Item
                ' 2019-04-23 - don't get for Reg/OT
                If PRDist.DistType <> PREquate.DistTypeReg And PRDist.DistType <> PREquate.DistTypeOT Then
                    If Not PRItem.GetByID(PRDist.ItemID) Then
                        MsgBox "PRItem NF: " & PRDist.ItemID & " " & PRDist.DistID, vbCritical
                        GoBack
                    End If
                End If
                
                PRDist.CityWage = PRDist.Amount
 
                ' use the employer record? for other earnings
                If PRDist.DistType <> PREquate.DistTypeReg And PRDist.DistType <> PREquate.DistTypeOT Then
                    If PRItem.UseEmployer = 1 Then
                    
                        SQLString = "ItemID = " & PRDist.EmployerItemID
                        rsERItem.Find SQLString, 0, adSearchForward, 1
                        If rsERItem.EOF Then
                            MsgBox "Item ID NF for Employer: " & PRDist.EmployerItemID & " " & _
                                   PRDist.DistID, vbCritical
                            End
                        End If
                        If rsERItem!NoSSTax Then SSWage = SSWage - PRDist.Amount
                        If rsERItem!NoMedTax Then MEDWage = MEDWage - PRDist.Amount
                        If rsERItem!NoFWTTax Then FWTWage = FWTWage - PRDist.Amount
                        If rsERItem!NoSWTTax Then SWTWage = SWTWage - PRDist.Amount
                        
                        If rsERItem!NoCWTTax Then
                            CWTWage = CWTWage - PRDist.Amount
                            PRDist.CityWage = 0
                        End If
                        
                        If rsERItem!NoFUNTax Then FUNWage = FUNWage - PRDist.Amount
                        If rsERItem!NoSUNTax Then SUNWage = SUNWage - PRDist.Amount
                    
                    Else
                        
                        If PRItem.NoSSTax Then SSWage = SSWage - PRDist.Amount
                        If PRItem.NoMedTax Then MEDWage = MEDWage - PRDist.Amount
                        If PRItem.NoFWTTax Then FWTWage = FWTWage - PRDist.Amount
                        If PRItem.NoSWTTax Then SWTWage = SWTWage - PRDist.Amount
                        
                        If PRItem.NoCWTTax Then
                            CWTWage = CWTWage - PRDist.Amount
                            PRDist.CityWage = 0
                        End If
                        
                        If PRItem.NoFUNTax Then FUNWage = FUNWage - PRDist.Amount
                        If PRItem.NoSUNTax Then SUNWage = SUNWage - PRDist.Amount
                    
                    End If
                End If
                                    
                ' *** 11/12/09 - keep taxable wage for EE marked as no tax ***
                ' employee marked as non-taxable
                ' If PREmployee.NoCityTax Then PRDist.CityWage = 0
                
                PRDist.Save (Equate.RecPut)
                If Not PRDist.GetNext Then Exit Do
            
            Loop
        
        End If

        ' loop the deductions from PRItemHist
        CWTDeduct = 0
        SQLString = "SELECT * FROM PRItemHist WHERE HistID = " & PRHist.HistID & _
                    " AND (PRItemHist.ItemType = " & PREquate.ItemTypeOE & _
                    " OR PRItemHist.ItemType = " & PREquate.ItemTypeDED & ")"
                    
        If PRItemHist.GetBySQL(SQLString) Then
            Do
                
                ' get the Employee Item
                If Not PRItem.GetByID(PRItemHist.ItemID) Then
                    If PREmployee.GetByID(PRItemHist.EmployeeID) Then
                    End If
                    MsgBox "PRItem NF: " & PRItemHist.ItemID & " " & PRItemHist.ItemHistID & vbCr & _
                           PRItem.Title & vbCr & _
                           PREmployee.LFName, vbExclamation
                    GoBack
                End If
                
                ' use the employer record?
                If PRItem.UseEmployer Then
                    SQLString = "ItemID = " & PRItemHist.EmployerItemID
                    rsERItem.Find SQLString, 0, adSearchForward, 1
                    If rsERItem.EOF Then
                        If PREmployee.GetByID(PRItemHist.EmployeeID) Then
                        End If
                        MsgBox "Item ID NF for Employer: " & PRItemHist.EmployerItemID & " " & _
                               PRItemHist.ItemHistID & vbCr & _
                               PREmployee.LFName, vbExclamation
                        GoBack
                    End If
                    
                    If rsERItem!NoSSTax Then SSWage = SSWage - PRItemHist.Amount
                    If rsERItem!NoMedTax Then MEDWage = MEDWage - PRItemHist.Amount
                    If rsERItem!NoFWTTax Then FWTWage = FWTWage - PRItemHist.Amount
                    If rsERItem!NoSWTTax Then SWTWage = SWTWage - PRItemHist.Amount
                    If rsERItem!NoCWTTax Then CWTDeduct = CWTDeduct + PRItemHist.Amount
                    If rsERItem!NoCWTTax Then CWTWage = CWTWage - PRItemHist.Amount
                    If rsERItem!NoFUNTax Then FUNWage = FUNWage - PRItemHist.Amount
                    If rsERItem!NoSUNTax Then SUNWage = SUNWage - PRItemHist.Amount
                Else
                    If PRItem.NoSSTax Then SSWage = SSWage - PRItemHist.Amount
                    If PRItem.NoMedTax Then MEDWage = MEDWage - PRItemHist.Amount
                    If PRItem.NoFWTTax Then FWTWage = FWTWage - PRItemHist.Amount
                    If PRItem.NoSWTTax Then SWTWage = SWTWage - PRItemHist.Amount
                    If PRItem.NoCWTTax Then CWTDeduct = CWTDeduct + PRItemHist.Amount
                    If PRItem.NoCWTTax Then CWTWage = CWTWage - PRItemHist.Amount
                    If PRItem.NoFUNTax Then FUNWage = FUNWage - PRItemHist.Amount
                    If PRItem.NoSUNTax Then SUNWage = SUNWage - PRItemHist.Amount
                End If
                
                If Not PRItemHist.GetNext Then Exit Do
            
            Loop
        
        End If
        
        ' *** 11/12/09 - keep taxable wage for EE marked as no tax ***
        ' *** 02/25/10 - Unemp - ER taxes - zero out the wage
        ' employee marked as non-taxable
'        If PREmployee.NoSSTax Then SSWage = 0
'        If PREmployee.NoMedTax Then MEDWage = 0
'        If PREmployee.NoFedTax Then FWTWage = 0
'        If PREmployee.NoStateTax Then SWTWage = 0
'        If PREmployee.NoCityTax Then CWTWage = 0
        
        If PREmployee.NoFedUnemp Then FUNWage = 0
        If PREmployee.NoStateUnemp Then SUNWage = 0
        If PREmployee.NoSSTax Then SSWage = 0
        If PREmployee.NoMedTax Then MEDWage = 0
        
        ' compare to yearly max
        PRHist.SSWageBase = SSWage
        If SSWage < YTDSSRmn Then
            YTDSSRmn = YTDSSRmn - SSWage
        Else
            SSWage = YTDSSRmn
            YTDSSRmn = 0
        End If
        
        PRHist.FUNWageBase = FUNWage
        If FUNWage < YTDFUNRmn Then
            YTDFUNRmn = YTDFUNRmn - FUNWage
        Else
            FUNWage = YTDFUNRmn
            YTDFUNRmn = 0
        End If
        
        PRHist.SSWage = SSWage
        PRHist.MEDWage = MEDWage
        PRHist.FWTWage = FWTWage
        PRHist.SWTWage = SWTWage
        PRHist.CWTWage = CWTWage
        PRHist.FUNWage = FUNWage
        PRHist.SUNWage = SUNWage

        ' get the YTD unem wage for the state
        SQLString = "StateID = " & PRHist.StateID
        rsState.Find SQLString, 0, adSearchForward, 1
        If rsState.EOF Then
            MsgBox "State NF ?:" & PRHist.StateID & " " & PRHist.HistID, vbCritical
            End
        End If
        
        PRHist.SUNWageBase = SUNWage
        If SUNWage < rsState!UnEmpRmn Then
            rsState!UnEmpRmn = rsState!UnEmpRmn - SUNWage
        Else
            SUNWage = rsState!UnEmpRmn
            rsState!UnEmpRmn = 0
        End If
        rsState.Update
        
        PRHist.SUNWage = SUNWage
 
        ' split up deductions not subject to CWT
        ' by proportion to the city wage
        If CWTDeduct <> 0 Then
            
'            ' loop back thru PRDist
'            ' PRDist.CityWage is either the PRDist.Amount or Zero
'            P2 = 0
'            SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
'            If PRDist.GetBySQL(SQLString) Then
'                Do
'                    If PRDist.CityWage <> 0 Then
'                        p1 = Round(PRDist.CityWage / CWTWage * CWTDeduct, 2)
'                        If p1 + P2 > CWTDeduct Then p1 = CWTDeduct - P2
'                        P2 = P2 + p1
'                        PRDist.CityWage = PRDist.CityWage - p1
'                        PRDist.Save (Equate.RecPut)
'                    End If
'                    If Not PRDist.GetNext Then Exit Do
'                Loop
'            End If
        
            SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
            If PRDist.GetBySQL(SQLString) Then
                DistCount = 0
                p1 = 0
                P2 = CWTDeduct
                Do
                    DistCount = DistCount + 1
                    If DistCount = PRDist.Records Then  ' last/only record - take remaining
                        PRDist.CityWage = PRDist.Amount - P2
                    Else
                        ' p1 = Round(PRDist.CityWage / CWTWage * CWTDeduct, 2)
                        p1 = Round(PRDist.Amount / PRHist.Gross * CWTDeduct, 2)
                        PRDist.CityWage = PRDist.Amount - p1
                        P2 = P2 - p1
                    End If
                    PRDist.Save (Equate.RecPut)
                    If PRDist.GetNext = False Then Exit Do
                Loop
            End If
        End If
                    
        ' Wkc Comp
        PRHist.WkcAmount = Round(WkcPct / 100 * PRHist.Gross, 2)
        
        PRHist.Save (Equate.RecPut)
        
NextPRHist:
        If Not PRHist.GetNext Then Exit Do
    
    Loop
    
End Sub
