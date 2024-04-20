Attribute VB_Name = "modReports"
Option Explicit
Dim rsCUR As New ADODB.Recordset
Dim rsQTD As New ADODB.Recordset
Dim rsYTD As New ADODB.Recordset
Dim rsMON As New ADODB.Recordset
Dim rsQTR As New ADODB.Recordset
Dim trs As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsB As New ADODB.Recordset
Dim RJAmount, WrittenAmount, BankNumber, BankAddress As String
Dim MaxAmt As Currency
Dim ReportTitle As String
Dim FmtString, TelFmtString As String
Dim Y As String
Dim Z As String
Dim w As String
Dim x As String
Dim LandSw As Byte
Dim I As Long
Dim J As Long
Dim QTR1Mo As Date
Dim QTR2Mo As Date
Dim TGross As Currency
Dim TSSWage As Currency
Dim TMedWage As Currency
Dim TFWTWage As Currency
Dim TSWTWAGE As Currency
Dim TCWTWage As Currency
Dim TSUNWage As Currency
Dim TFUNWage As Currency

Dim GTGross As Currency
Dim GTSSWage As Currency
Dim GTMedWage As Currency
Dim GTFWTWage As Currency
Dim GTSWTWAGE As Currency
Dim GTCWTWage As Currency
Dim GTSUNWage As Currency
Dim GTFUNWage As Currency
Dim LabelRows, ColumnCount, LRow, LabelCount As Long

Dim SkipFlag As Boolean
Dim ItemCount As Long
Dim LastEmpName As String
Dim LastEmpNo As Long
Dim LineCount As Long
Dim LastChkDate As Date
Dim dtlCount As Long
Public EndFlag As Boolean
Public LastEmpID As Long
Public ErnTotals(3, 9) As Currency
Public GrTotals(3, 9) As Currency

Public CheckNum As Long
Public ILineNo As Long
'Dim CustName As String
Public PRCheck As New cPRCheck
Dim TxtWidth As Double
Dim Msg1 As String

Dim StartString, ItemTitle As String
Dim ChkPrefix As String

Public Sub Form941Pt4Pt5(ByRef frm As Form)
    
    PRGlobal.Var1 = ""
    PRGlobal.Var2 = ""
    PRGlobal.Var3 = ""
    PRGlobal.Var4 = ""
    PRGlobal.Var5 = ""
    PRGlobal.Var6 = ""
    PRGlobal.Var7 = ""
    
    ' Part 4 - Third Party Designee - Per User
    If frm.Part4ID <> 0 Then
        If PRGlobal.GetByID(frm.Part4ID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part4
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
           
    PRGlobal.Var1 = frm.Part4Name
    PRGlobal.Var2 = frm.Part4Phone
    PRGlobal.Var3 = frm.Part4Pin
    PRGlobal.Save (Equate.RecPut)
    
    If frm.Part4CheckYes Then
        PosPrint 3500, 8200, PRGlobal.Var1
        PosPrint 3500, 8650, PRGlobal.Var2
        PosPrint 10000, 8650, PRGlobal.Var3
    End If
           
    ' Part 5 - Company Signature - Per Company
    If frm.Part5ID <> 0 Then
        If PRGlobal.GetByID(frm.Part5ID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part5
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If

    PRGlobal.Var1 = frm.Part5NameTitle

    PRGlobal.Save (Equate.RecPut)

    PosPrint 3000, 10600, PRGlobal.Var1
    
    If IsNull(frm.Part5Date) = False Then
        PosPrint 3000, 11100, frm.Part5Date
    End If
    PosPrint 5100, 11100, frm.Part5Phone
        
    'Paid Preparer - Per User
    If frm.PaidPrepID <> 0 Then
        If PRGlobal.GetByID(frm.PaidPrepID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941PaidPrep
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
               
    PRGlobal.Var1 = frm.PrepFirm
    PRGlobal.Var2 = frm.PrepAddr1
    PRGlobal.Var3 = frm.PrepAddr2
    PRGlobal.Var4 = frm.PrepPhone
    PRGlobal.Var5 = frm.PrepEIN
    PRGlobal.Var6 = frm.PrepZip
    PRGlobal.Var7 = frm.PrepSSN
    If frm.PrepCheck Then
        PRGlobal.Var8 = "1"
    Else
        PRGlobal.Var8 = "0"
    End If
    PRGlobal.Var9 = frm.cmbPrepName.text

    PRGlobal.Save (Equate.RecPut)
    
    PosPrint 3000, 12040, frm.cmbPrepName
    PosPrint 9000, 12500, PRGlobal.Var4
    PosPrint 3000, 12940, PRGlobal.Var1
    PosPrint 3000, 13480, PRGlobal.Var2
    PosPrint 3000, 13920, PRGlobal.Var3
    
    PosPrint 9000, 12940, PRGlobal.Var5
    PosPrint 9000, 13480, PRGlobal.Var6
    PosPrint 9000, 13920, PRGlobal.Var7
End Sub

'=======================================   FORM 941     ======================================

Public Sub Form941APrint()
Dim VertSpace, VertPosn As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
 
    CurrYear = Year(Now())
    Ln = 0
    SetEquates
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    VertSpace = 492
    FmtString = "##,###,##0.00"
    TelFmtString = "###-###-####"
    
    PosPrint 3200, 1020, PRCompany.FederalID
    PosPrint 2500, 1490, PRCompany.Name
    PosPrint 1500, 2200, PRCompany.Address1
    If PRCompany.Address2 <> "" Then
        PosPrint 1500, 2400, PRCompany.Address2
    End If
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        PosPrint 1500, 2675, PRCompany.City & ", " & PRState.StateAbbrev & "  " & PRCompany.ZipCode
    End If

    If frm941Entry.cmbQtr = 1 Then
        PosPrint 8400, 1020, "X"
    End If   'col  'line

    If frm941Entry.cmbQtr = 2 Then
        PosPrint 8400, 1500, "X"
    End If
    If frm941Entry.cmbQtr = 3 Then
        PosPrint 8400, 1970, "X"
    End If
    If frm941Entry.cmbQtr = 4 Then
        PosPrint 8400, 2490, "X"     '  over   down
    End If

    
    PosPrint 10200, 3410, PadRight(Format(frm941Entry.Line1, "##,##0"), 6)
    
    PosPrint 9400, 3890, PadRight(Format(frm941Entry.Line2, FmtString), 13)
    PosPrint 9400, 4370, PadRight(Format(frm941Entry.Line3, FmtString), 13)
    PosPrint 8800, 4900, frm941Entry.AlphaCheckLine4

    frm941Entry.Line5aa = frm941Entry.Line5a * 0.124
    frm941Entry.Line5bb = frm941Entry.Line5b * 0.124
    frm941Entry.Line5cc = frm941Entry.Line5c * 0.029

    PosPrint 4000, 5580, PadRight(Format(frm941Entry.Line5a, FmtString), 13)
    PosPrint 6900, 5580, PadRight(Format(frm941Entry.Line5aa, FmtString), 13)
    PosPrint 4000, 6050, PadRight(Format(frm941Entry.Line5b, FmtString), 13)
    PosPrint 6900, 6050, PadRight(Format(frm941Entry.Line5bb, FmtString), 13)
    PosPrint 4000, 6550, PadRight(Format(frm941Entry.Line5c, FmtString), 13)
    PosPrint 6900, 6550, PadRight(Format(frm941Entry.Line5cc, FmtString), 13)
    PosPrint 9400, 7030, PadRight(Format(frm941Entry.Line5d, FmtString), 13)
    PosPrint 9400, 7510, PadRight(Format(frm941Entry.Line6, FmtString), 13)

    PosPrint 6900, 8140, PadRight(Format(frm941Entry.Line7a, FmtString), 13)
    PosPrint 6900, 8635, PadRight(Format(frm941Entry.Line7b, FmtString), 13)
    PosPrint 6900, 9100, PadRight(Format(frm941Entry.Line7c, FmtString), 13)
    PosPrint 9400, 9525, PadRight(Format(frm941Entry.Line7d, FmtString), 13)
    PosPrint 9400, 10030, PadRight(Format(frm941Entry.Line8, FmtString), 13)
    PosPrint 9400, 10530, PadRight(Format(frm941Entry.Line9, FmtString), 13)
    PosPrint 9400, 11010, PadRight(Format(frm941Entry.Line10, FmtString), 13)
    PosPrint 9400, 11490, PadRight(Format(frm941Entry.Line10Total, FmtString), 13)
    PosPrint 6900, 11950, PadRight(Format(frm941Entry.Line11, FmtString), 13)
    PosPrint 6900, 12450, PadRight(Format(frm941Entry.Line12a, FmtString), 13)
    PosPrint 4650, 12920, PadRight(Format(frm941Entry.Line12b, FmtString), 13)

    PosPrint 9400, 13400, PadRight(Format(frm941Entry.Line13, FmtString), 13)
    PosPrint 9400, 13880, PadRight(Format(frm941Entry.Line14, FmtString), 13)
    PosPrint 6900, 14330, PadRight(Format(frm941Entry.Line15, FmtString), 13)
    PosPrint 9600, 14340, frm941Entry.AlphaCheckLine15a
    PosPrint 9600, 14520, frm941Entry.AlphaCheckLine15b     '  over   down
    FormFeed

'   #######################  FORM 941 - PAGE 2  ####################################
'
    PosPrint 900, 1120, PRCompany.Name
    PosPrint 8490, 1120, PRCompany.FederalID
    PosPrint 950, 2100, frm941Entry.Line16
    PosPrint 1830, 2550, frm941Entry.AlphaCheckLine17a
    PosPrint 1830, 3100, frm941Entry.AlphaCheckLine17b
    
    If frm941Entry.Line17Check2 = 1 Then
        PosPrint 5000, 3800, PadRight(Format(frm941Entry.Line17Mo1, FmtString), 13)
        PosPrint 5000, 4250, PadRight(Format(frm941Entry.Line17Mo2, FmtString), 13)
        PosPrint 5000, 4710, PadRight(Format(frm941Entry.Line17Mo3, FmtString), 13)
        PosPrint 5000, 5230, PadRight(Format(frm941Entry.Line17Total, FmtString), 13)
    End If
    
    PosPrint 1820, 5540, frm941Entry.AlphaCheckLine17c
    PosPrint 9200, 6500, frm941Entry.AlphaCheckLine18
    If frm941Entry.Line18Check = 1 Then
        PosPrint 3900, 6950, frm941Entry.Line18Date
    End If
    PosPrint 9200, 7220, frm941Entry.AlphaCheckLine19
    Form941Pt4Pt5 frm941Entry
    
    If frm941Entry.Part4CheckYes = 1 Then
        PosPrint 880, 8200, "X"
    Else
        PosPrint 900, 8900, "X"
    End If
    
    PosPrint 2660, 14450, frm941Entry.AlphaCheckPart5
    If IsNull(frm941Entry.PrepDate) = False Then
        PosPrint 8180, 14450, frm941Entry.PrepDate
    End If
    
End Sub

Public Sub Form941BPrint(ByVal VertPos As Long, ByRef fg As VSFlexGrid, ByVal BMoTax As Currency)

Dim VertSpace As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString As String

    SetEquates
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

'    SQLString = "SELECT * FROM PREmployee"
'    rsInit SQLString, cn, rs941

    VertSpace = 492
    FmtString = "##,###,##0.00"

    Col1X = 610
    Col2X = 2840
    Col3X = 5040
    Col4X = 7240

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(1, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(1, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(1, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(1, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(2, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(2, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(2, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(2, 7), FmtString), 13)
    PosPrint Col4X + 2300, VertPos, PadRight(Format(BMoTax, FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(3, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(3, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(3, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(3, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(4, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(4, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(4, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(4, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(5, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(5, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(5, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(5, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(6, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(6, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(6, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(6, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(7, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(7, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(7, 5), FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(fg.TextMatrix(7, 7), FmtString), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(Format(fg.TextMatrix(8, 1), FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(fg.TextMatrix(8, 3), FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(fg.TextMatrix(8, 5), FmtString), 13)

End Sub

Public Sub Form941BHdr(ByRef frm As Form, ByVal TaxYear As String)
    
    FmtString = "##,###,##0.00"

    With frm

        CurrYear = Year(Now())
        If .cmbQtr = 1 Then
            PosPrint 8490, 900, "X"
        ElseIf .cmbQtr = 2 Then
            PosPrint 8490, 1150, "X"
        ElseIf .cmbQtr = 3 Then
            PosPrint 8490, 1430, "X"
        ElseIf .cmbQtr = 4 Then
            PosPrint 8490, 1660, "X"
        End If
    
        PosPrint 3380, 910, PRCompany.FederalID
        PosPrint 3380, 1160, PRCompany.Name
        PosPrint 3380, 1440, TaxYear
        PosPrint 9500, 14380, PadRight(Format(.BTotalTax, FmtString), 13)
    
    End With
    
End Sub

Public Sub CheckPrint(ByVal CheckType As Byte)

'Dim PRCheck As New cPRCheck
Dim ReportTitle As String
Dim LineNo As Long
Dim CkCt As Long
Dim CoName As String
Dim BlankStock As Boolean
Dim ChkIncr As Byte
Dim FWTString, SWTString As String

    ' is this print on blank stock?
    BlankStock = False
    ChkPrefix = ""
    If CheckType = PREquate.CheckTypeBlankStock Then
    
        ' get the check prefix
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePRCheckPrefix
        If PRGlobal.GetBySQL(SQLString) Then
            ChkPrefix = Trim(PRGlobal.Description)
        End If

        If CNPRCKOpen(frmCheckPrint.CheckFileName, "pobox45") = False Then
            MsgBox "Can not open blank check stock definition!", vbExclamation
            GoBack
        End If
        BlankStock = True
        
        ' new fields - 02/15/2010
        If AddField("PRCheck", "BankAccountAdd", "Char (10)", cnPRCK) Then
        End If
        If AddField("PRCheck", "AddressAdjust", "Long", cnPRCK) Then
        End If
        
        ' find the first record
        If Not PRCheck.GetByID(1) Then
            MsgBox "No blank stock data exists!", vbExclamation
            GoBack
        End If
        
    End If

    ' temp record set for YTD & Current Pay hours and amounts
    rsYTD.CursorLocation = adUseClient
    
    rsYTD.Fields.Append "Side", adVarChar, 2, adFldIsNullable
    rsYTD.Fields.Append "ID", adDouble
    rsYTD.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsYTD.Fields.Append "CurrAmt", adCurrency
    rsYTD.Fields.Append "YTDAmt", adCurrency
    rsYTD.Fields.Append "CurrHours", adCurrency
    rsYTD.Fields.Append "YTDHours", adCurrency
    
    rsYTD.Open , , adOpenDynamic, adLockOptimistic

    frmCheckPrint.Hide
    LineNo = 25
    ILineNo = 25
    CkCt = 0

    PrtInit ("Port")
    
    ReportTitle = "Print Checks "
            
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    ' print in reverse order
    If frmCheckPrint.chkRevOrder = 1 Then
        frmCheckPrint.trs.MoveLast
        CheckNum = frmCheckPrint.tdbnumEndNumber
    Else
        frmCheckPrint.trs.MoveFirst
        CheckNum = frmCheckPrint.tdbnumStartNumber
    End If
    
    Do
    
        Prvw.vsp.Font.Name = "ARIAL"
        If frmCheckPrint.trs!PrintCheck Then
            
            ' w/ net pay only
            If frmCheckPrint.cmbAmtFilter.ListIndex = 1 Then
                If frmCheckPrint.trs!CheckAmount = 0 Then GoTo NxtCheck
            End If
            
            ' zero net pay only
            If frmCheckPrint.cmbAmtFilter.ListIndex = 2 Then
                If frmCheckPrint.trs!CheckAmount <> 0 Then GoTo NxtCheck
            End If
            
            If CkCt <> 0 Then Prvw.vsp.NewPage
            CkCt = CkCt + 1
                
            If Not PREmployee.GetByID(frmCheckPrint.trs!EmployeeID) Then
                MsgBox "Employee Not Found !!!", vbExclamation, "Print Checks"
                Exit Sub
            End If
        
            ' display withholding status for FWT / SWT
            If Not PRW4.GetByEmployeeID(PREmployee.EmployeeID) Then
                PRW4.Clear
                PRW4.EmployeeID = PREmployee.EmployeeID
                PRW4.Save (Equate.RecAdd)
            End If
            FWTString = "FWT "
            If PRW4.FilingType = PREquate.PRW4Standard Then
                If PREmployee.FWTBasis = PREquate.BasisExemptions Then
                    If PREmployee.FWTMarried = 1 Then
                        FWTString = FWTString & " M" & PREmployee.FWTAmount
                    Else
                        FWTString = FWTString & " S" & PREmployee.FWTAmount
                    End If
                Else
                    FWTString = FWTString & PREmployee.FWTAmount & "%"
                End If
                If PREmployee.FWTExtraAmount <> 0 Then
                    FWTString = FWTString & "*"
                End If
            Else
                If PRW4.FilingType = PREquate.PRW4Single Then FWTString = FWTString & "W4-S"
                If PRW4.FilingType = PREquate.PRW4Married Then FWTString = FWTString & "W4-M"
                If PRW4.FilingType = PREquate.PRW4HOH Then FWTString = FWTString & "W4-H"
                If PRW4.TwoJobs <> 0 Then FWTString = FWTString & "- 2Job"
                If PRW4.ExtraWH <> 0 Then FWTString = FWTString & "*"
            End If
        
            SWTString = "SWT "
            If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                If PREmployee.SWTMarried = 1 Then
                    SWTString = SWTString & "M" & PREmployee.SWTAmount
                Else
                    SWTString = SWTString & "S" & PREmployee.SWTAmount
                End If
            Else
                SWTString = SWTString & PREmployee.SWTAmount & "%"
            End If
            If PREmployee.SWTExtraAmount <> 0 Then
                SWTString = SWTString & "*"
            End If
        
            frmProgress.lblMsg2 = "Now Printing Check for: " & PREmployee.FLName & _
                                  " # " & frmCheckPrint.trs!CheckNumber
            frmProgress.Refresh
        
            ' gather the amounts for the stub
            SQLString = "SELECT * FROM PRHist WHERE PRHist.EmployeeID = " & frmCheckPrint.trs!EmployeeID & _
                        " AND PRHist.CheckDate <= " & CLng(Int(PRBatch.CheckDate)) & _
                        " AND PRHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
                        " AND PRHist.HistID <= " & frmCheckPrint.trs!HistID & _
                        " ORDER BY PRHist.HistID"
            
            If Not PRHist.GetBySQL(SQLString) Then
                ' ???
            End If
            
            Do
                
                ' problem w/ check date comparison in SQL statement
                If PRHist.CheckDate > PRBatch.CheckDate Then GoTo SkipHist
                
                ' right side stub info from PRHist
                CheckPrintYTDUpdate "R", 1, PRHist.RegHours, 0, PRHist.HistID, "REG HRS", PRHist.BatchID
                CheckPrintYTDUpdate "R", 2, PRHist.OTHours, 0, PRHist.HistID, "OVT HRS", PRHist.BatchID
                CheckPrintYTDUpdate "R", 3, PRHist.OEHours, 0, PRHist.HistID, "OTH HRS", PRHist.BatchID
                CheckPrintYTDUpdate "R", 4, PRHist.RegAmount, 0, PRHist.HistID, "REG PAY", PRHist.BatchID
                CheckPrintYTDUpdate "R", 5, PRHist.OTAmount, 0, PRHist.HistID, "OVT PAY", PRHist.BatchID
                CheckPrintYTDUpdate "R", 6, PRHist.OEAmount, 0, PRHist.HistID, "OTH PAY", PRHist.BatchID
                CheckPrintYTDUpdate "R", 7, PRHist.Gross, 0, PRHist.HistID, "GROSS PAY", PRHist.BatchID
                CheckPrintYTDUpdate "R", 8, PRHist.Deductions, 0, PRHist.HistID, "DEDUCTIONS", PRHist.BatchID
                CheckPrintYTDUpdate "R", 9, PRHist.SSTax, 0, PRHist.HistID, "SS  TAX", PRHist.BatchID
                CheckPrintYTDUpdate "R", 10, PRHist.MedTax, 0, PRHist.HistID, "MED TAX", PRHist.BatchID
                CheckPrintYTDUpdate "R", 11, PRHist.FWTTax, 0, PRHist.HistID, FWTString, PRHist.BatchID
                CheckPrintYTDUpdate "R", 12, PRHist.SWTTax, 0, PRHist.HistID, SWTString, PRHist.BatchID
                CheckPrintYTDUpdate "R", 13, PRHist.CWTTax, 0, PRHist.HistID, "CWT TAX", PRHist.BatchID
                CheckPrintYTDUpdate "R", 15, PRHist.Net + PRHist.DirectDeposit, 0, PRHist.HistID, "NET PAY", PRHist.BatchID
                CheckPrintYTDUpdate "R", 16, PRHist.Net, 0, PRHist.HistID, "CK AMT", PRHist.BatchID
                If PRHist.DirectDeposit <> 0 Then
                    CheckPrintYTDUpdate "R", 17, PRHist.DirectDeposit, 0, PRHist.HistID, "DIR DEP", PRHist.BatchID
                End If
                
                ' left side - OE then DED
                
                ' other earning detail
                SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID & _
                            " AND PRDist.ItemType <> " & PREquate.ItemTypeRegPay & _
                            " AND PRDist.ItemType <> " & PREquate.ItemTypeOvtPay
                
                If PRDist.GetBySQL(SQLString) Then
                    Do
                        CheckPrintYTDUpdate "L1", PRDist.EmployerItemID, PRDist.Amount, PRDist.Hours, PRHist.HistID, PRDist.EmployerItemID, PRDist.BatchID
                        If Not PRDist.GetNext Then Exit Do
                    Loop
                End If
                
                ' deduction detail
                SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID
                
                If PRItemHist.GetBySQL(SQLString) Then
                    Do
                        If PRItemHist.ItemType = PREquate.ItemTypeDirDepDed Then
                            ' get bank info from the employee PRItem records
                            SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & PRItemHist.ItemID
                            If Not PRItem.GetBySQL(SQLString) Then
                                CheckPrintYTDUpdate "L3", PRItemHist.ItemID, PRItemHist.Amount, PRItemHist.Hours, PRHist.HistID, "Dir Dep", PRItemHist.BatchID
                            Else
                                If PRItem.DirDepType = PREquate.DirDepTypeChecking Then
                                    x = PRItem.DirDepBank & " CK"
                                Else
                                    x = PRItem.DirDepBank & " SV"
                                End If
                                CheckPrintYTDUpdate "L3", PRItemHist.ItemID, PRItemHist.Amount, PRItemHist.Hours, PRHist.HistID, x, PRItemHist.BatchID
                            End If
                        Else
                            CheckPrintYTDUpdate "L2", PRItemHist.EmployerItemID, PRItemHist.Amount, PRItemHist.Hours, PRHist.HistID, PRItemHist.EmployerItemID, PRItemHist.BatchID
                        End If
                        If Not PRItemHist.GetNext Then Exit Do
                    Loop
                End If
                
SkipHist:
                If Not PRHist.GetNext Then Exit Do
                    
            Loop
            
            rsYTD.Sort = "Side, ID"
            If Not PRHist.GetByID(frmCheckPrint.trs!HistID) Then
            End If
            If frmCheckPrint.chkSveChk = 1 Then
                PRHist.CheckNumber = CheckNum
                PRHist.Save (Equate.RecPut)
            End If
 
            frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.LFName)
            frmProgress.Show
            
            If frmCheckPrint.chkBillStub = 0 Then
            
                If LCase(ChkPrefix) = "sct" Then
                    ' scott molders - type "A" w/ signature
                    CheckPrintPrePrintA
                ElseIf CheckType = PREquate.CheckTypeBlankStock Then
                    CheckPrintBlankForm
                ElseIf CheckType = PREquate.CheckTypePrePrintedA Then
                    CheckPrintPrePrintA
                ElseIf CheckType = PREquate.CheckTypePrePrintedB Then
                    CheckPrintPrePrintB
                Else
                    CheckPrintPrePrintC
                End If
                                    
                '''''''''''''''''''''''''' TOP STUB SECTION   '''''''''''''''''''''''''
                ' bring up the stub some ....
                If CheckType = PREquate.CheckTypePrePrintedB Then VertNudge = VertNudge - 2
                If CheckType = PREquate.CheckTypePrePrintedC Then VertNudge = VertNudge + 3
                CheckStub 1
                If CheckType = PREquate.CheckTypePrePrintedB Then VertNudge = VertNudge + 2
                If CheckType = PREquate.CheckTypePrePrintedC Then VertNudge = VertNudge - 3
    
                '''''''''''''''''''''''''' BOTTOM STUB SECTION   '''''''''''''''''''''''''
                If frmCheckPrint.chkBottomPanel = 1 Then
                    CheckStub 2
                End If
 
            Else
                CheckPrintBill      ' check panel on bottom / billing panel on top
            End If

'''''''''''''''''''''''''''''''''''  Clear Recordset   '''''''''''''''''''''''''''''''''
                               
NxtCheck:
        
            ' clear the temp record set
            If rsYTD.RecordCount > 0 Then
                rsYTD.MoveFirst
                Do
                    rsYTD.Delete
                    rsYTD.MoveNext
                    If rsYTD.EOF Then Exit Do
                Loop
            End If
   
            ChkIncr = 1
        
        Else            ' check skipped
            
            ChkIncr = 0
   
        End If
   
        ' Making sure the recordset is empty
        If rsYTD.RecordCount > 0 Then
            rsYTD.MoveFirst
            Do Until rsYTD.RecordCount = 0
                rsYTD.Delete
                rsYTD.MoveNext
            Loop
        End If
        
        If frmCheckPrint.chkRevOrder = 1 Then
            frmCheckPrint.trs.MovePrevious
            If frmCheckPrint.trs.BOF Then Exit Do
            CheckNum = CheckNum - ChkIncr
        Else
            frmCheckPrint.trs.MoveNext
            If frmCheckPrint.trs.EOF Then Exit Do
            CheckNum = CheckNum + ChkIncr
        End If
    
    Loop

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub CheckPrintBill()

Dim BotAdd, BotLine As Long
Dim LeftLineNo, RightLineNo As Long
Dim P1 As Currency
Dim RateType As String
Dim TotalHours, TotalAmount As Currency
Dim AIW As String
    
    ' ********************************************
    ' billing check print logic
    ' top panel 4"          pay info
    ' middle panel 3.5"     billing info
    ' bottom panel 3.5"     check
    ' ********************************************

    BotAdd = 0
    BotLine = 0

    Prvw.vsp.FontBold = True:                       SetFont 13, Equate.Portrait
    Prvw.vsp.Font.Name = "ARIAL"
    
    ' ****
    ' flush junk characters ????
    PosPrint 10, 10 + BotAdd, " "
    ' ****
    
    SetFont 10, Equate.Portrait                                                        '  FONT IS 10
    Prvw.vsp.Font.Name = "Courier new"
    
    ' *******************************************************************************
    ' *** top panel - pay info stub
    
    prt 1, 55, "CHK DATE: ":        prt 1, 65, Format(PRBatch.CheckDate, "mm/dd/yyyy ")
    prt 1, 77, "CHK #: ":           prt 1, 85, CheckNum
    
    PosPrint 180, 400, "- - - - - - - - - -  CURRENT PD - - YR TO DATE":
    PosPrint 5500, 400, "- - - - - - - - -  CURRENT PD  - - - YR TO DATE"
    
    If PREmployee.UseAltName = 0 Then
        prt 21, 44, Trim(PREmployee.FLName):  prt 21, 76, "EMP. ID: " & PREmployee.EmployeeNumber
    Else
        prt 21, 44, Trim(PREmployee.AltName): prt 21, 76, "EMP. ID: " & PREmployee.EmployeeNumber
    End If
    
    PRDepartment.GetByID (PREmployee.DepartmentID)
    prt 22, 47, "DPT: " & PRDepartment.DepartmentNumber
    If frmCheckPrint.chkNoRate = 0 Then
        prt 22, 58, "RATE: " & Format(PRHist.RegRate, "##,###.#0")
    End If
    prt 22, 74, "PE DATE: " & Format(PRHist.PEDate, "mm/dd/yy")
    
    LeftLineNo = 4:  RightLineNo = 4
        
    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' loop thru the YTD temp record set for stub amounts
    rsYTD.MoveFirst
    Do
        ' left side - OE / DED
        If rsYTD!Side = "L1" Or rsYTD!Side = "L2" Or rsYTD!Side = "L3" Then
            If rsYTD!YTDHours <> 0 Then
                prt LeftLineNo, 1, Trim(rsYTD!Title) & " HRS":  prt LeftLineNo, 18, CurrFormat(rsYTD!CurrHours)
                prt LeftLineNo, 33, CurrFormat(rsYTD!YTDHours): LeftLineNo = LeftLineNo + 1
                End If
            prt LeftLineNo, 1, Trim(rsYTD!Title):               prt LeftLineNo, 18, CurrFormat(rsYTD!CurrAmt)
            prt LeftLineNo, 33, CurrFormat(rsYTD!YTDAmt):       LeftLineNo = LeftLineNo + 1
        Else        ' right side - summaries
            prt RightLineNo, 47, rsYTD!Title:                   prt RightLineNo, 63, CurrFormat(rsYTD!CurrAmt)
            prt RightLineNo, 78, CurrFormat(rsYTD!YTDAmt):      RightLineNo = RightLineNo + 1
        End If
        
        rsYTD.MoveNext
        If rsYTD.EOF Then Exit Do
    Loop
    
    prt LeftLineNo + 1, 1, frmCheckPrint.tdbtextMsg

    ' *******************************************************************************
    ' *** middle panel - billing info stub

    ' billing info stub header
    Prvw.Font.Bold = True
        
    Ln = 27
    PrintValue(1) = PRCompany.Name:                         FormatString(1) = "a30"
    PrintValue(2) = "PE Date: ":                            FormatString(2) = "a9"
    PrintValue(3) = Format(PRHist.PEDate, " mm/dd/yy "):    FormatString(3) = "a10"
    PrintValue(4) = "Check Date: ":                         FormatString(4) = "a12"
    PrintValue(5) = Format(PRHist.CheckDate, " mm/dd/yy "): FormatString(5) = "a10"
    PrintValue(6) = " ":                                    FormatString(6) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = "Wk Ending":                            FormatString(1) = "a10"
    PrintValue(2) = "Customer":                             FormatString(2) = "a48"
    PrintValue(3) = "Ern Type":                             FormatString(3) = "a10"
    PrintValue(4) = "Hours ":                               FormatString(4) = "r9"
    PrintValue(5) = "Rate ":                                FormatString(5) = "r9"
    PrintValue(6) = "Amount ":                              FormatString(6) = "r9"
    PrintValue(7) = " ":                                    FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1
    
    Prvw.Font.Bold = False
    
    TotalHours = 0
    TotalAmount = 0
    
    With frmCheckPrint.rsTS
        .Filter = "EmployeeID = " & PREmployee.EmployeeID
        If .RecordCount > 0 Then
            .MoveFirst
            Do
                    
                x = !JobID
                If JCJob.GetByID(!JobID) Then
                    x = JCJob.Name
                End If
                
                ' get the rate - from PRDist
                If !ItemID = 99991 Then
                    SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID & _
                                " AND DistType = " & PREquate.DistTypeReg
                    RateType = "Regular"
                ElseIf !ItemID = 99992 Then
                    SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID & _
                                " AND DistType = " & PREquate.DistTypeOT
                    RateType = "Overtime"
                Else
                    SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID & _
                                " AND DistType = " & PREquate.DistTypeItem & _
                                " AND EmployerItemID = " & !ItemID
                End If
                
                If PRDist.GetBySQL(SQLString) = False Then
                    ' ????
                End If
                
                If !ItemID < 99991 Then
                    If PRItem.GetByID(PRDist.EmployerItemID) Then
                        RateType = PRItem.Abbreviation
                    Else
                        RateType = "Other"
                    End If
                End If
                
                P1 = SuperRound(!Hours, PRDist.Rate)
                
                PrintValue(1) = Format(PRHist.PEDate, " mm/dd/yy "):        FormatString(1) = "a10"
                PrintValue(2) = JCJob.FullName:                             FormatString(2) = "a48"
                PrintValue(3) = RateType:                                   FormatString(3) = "a10"
                PrintValue(4) = !Hours:                                     FormatString(4) = "d9"
                
                ' 2011-12-17 - take rate and amount from recordset already made
                ' PrintValue(5) = PRDist.Rate:                                FormatString(5) = "d9"
                PrintValue(5) = !Rate:                                      FormatString(5) = "d9"
                
                ' PrintValue(6) = P1:                                         FormatString(6) = "d9"
                PrintValue(6) = !Amount:                                    FormatString(6) = "d9"
                
                PrintValue(7) = " ":                                        FormatString(7) = "~"
                FormatPrint
                Ln = Ln + 1
                
                TotalHours = TotalHours + !Hours
                TotalAmount = TotalAmount + !Amount
                
                ' ??? multi stub check ???
                
                .MoveNext
            
            Loop Until .EOF
        End If
        
        .Filter = adFilterNone
    
        ' print totals
        PrintValue(1) = " ":                    FormatString(1) = "a10"
        PrintValue(2) = "T O T A L:":           FormatString(2) = "a48"
        PrintValue(3) = " ":                    FormatString(3) = "a10"
        PrintValue(4) = TotalHours:             FormatString(4) = "d9"
        PrintValue(5) = " ":                    FormatString(5) = "a9"
        PrintValue(6) = TotalAmount:            FormatString(6) = "d9"
        PrintValue(7) = " ":                    FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1
    
    End With

    ' *******************************************************************************
    ' *** bottom panel - check

    Ln = 54
    
    If PRHist.Net <= 0 Then         '   Net is <= Zero
        AIW = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
    Else
        AIW = AmountInWords(PRHist.Net, True)
    End If
    
    PrintValue(1) = AIW:                FormatString(1) = "r87"
    PrintValue(2) = " ":                FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 4
    
    PrintValue(1) = " ":                                    FormatString(1) = "a40"
    PrintValue(2) = PRHist.CheckNumber:                     FormatString(2) = "n9"
    PrintValue(3) = " ":                                    FormatString(3) = "a13"
    PrintValue(4) = Format(PRHist.CheckDate, "mm/dd/yy"):   FormatString(4) = "a8"
    PrintValue(5) = " ":                                    FormatString(5) = "a7"
    PrintValue(6) = CheckAmount(PRHist.Net):                FormatString(6) = "a16"
    PrintValue(7) = " ":                                    FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
 
    For I = 1 To 4
        If I = 1 Then
            If PREmployee.AltName = "" Then
                AIW = PREmployee.FLName
            Else
                AIW = PREmployee.AltName
            End If
        End If
        If I = 2 Then AIW = PREmployee.Address1
        If I = 3 Then AIW = PREmployee.Address2
        If I = 4 Then AIW = PREmployee.CSZ
        If AIW <> "" Then
            PrintValue(1) = " ":            FormatString(1) = "a10"
            PrintValue(2) = AIW:            FormatString(2) = "a40"
            PrintValue(3) = " ":            FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 1
        End If
    Next I
 
 
'    If PRHist.Net <= 0 Then         '   Net is <= Zero
'
'        Prt 8, 51, PRHist.CheckNumber
'        Prt 8, 60, Format(PRHist.CheckDate, "mm/dd/yyyy")
'        Prt 8, 70, "NO DOLLARS AND NO CENTS"
'        RJAmount = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
'        Prt 12, 7, RJAmount
'
'        PosPrint 7100, 3460, "   *** V O I D   V O I D  V O I D ***   "
'    Else
'
'        PosPrint 5700, 1850, PRHist.CheckNumber
'        PosPrint 7700, 1850, Format(PRHist.CheckDate, "mm/dd/yyyy")
'        PosPrint 9800, 1850, CheckAmount(PRHist.Net)
'
'        WrittenAmount = AmountInWords(PRHist.Net, False)
'        Prvw.vsp.CurrentX = 10700 - Prvw.vsp.TextWidth(Trim(WrittenAmount)) + (Nudge * HorzNudge)
'        Prvw.vsp.CurrentY = 2310 + (Nudge * VertNudge)
'        Prvw.vsp.Text = Trim(WrittenAmount)
'
'    End If
'
'    SetFont 10, Equate.Portrait                 '  FONT IS 10
'
'    If PREmployee.UseAltName = 0 Then
'        Prt 14, 10, Trim(PREmployee.FLName)
'    Else
'        Prt 14, 10, Trim(PREmployee.AltName)
'    End If
'
'    ILineNo = 14
'
'    If Trim(PREmployee.Address1) <> "" Then
'        ILineNo = ILineNo + 1
'        Prt ILineNo, 10, Trim(PREmployee.Address1)
'    End If
'
'    If Trim(PREmployee.Address2) <> "" Then
'        ILineNo = ILineNo + 1
'        Prt ILineNo, 10, Trim(PREmployee.Address2)
'    End If
'
'    If Trim(PREmployee.City) <> "" Then
'        ILineNo = ILineNo + 1
'        Prt ILineNo, 10, PREmployee.CSZ
'    End If
'
'    PosPrint 900, 4160, PREmployee.CheckComment
'

End Sub

Private Sub CheckStub(ByVal TopBot As Byte)

Dim BotAdd, BotLine As Long
Dim LeftLineNo, RightLineNo As Long

    If TopBot = 1 Then
        BotAdd = 0
        BotLine = 0
    Else
        BotAdd = 4950
        BotLine = 22        ' lines to add
    End If

    Prvw.vsp.FontBold = True:                       SetFont 13, Equate.Portrait
    Prvw.vsp.Font.Name = "ARIAL"
    
    ' ****
    ' flush junk characters ????
    PosPrint 10, 10 + BotAdd, " "
    ' ****
    
    PosPrint 400, 5090 + BotAdd, Trim(PRCheck.CustomerName)
    SetFont 10, Equate.Portrait                                                        '  FONT IS 10
    Prvw.vsp.Font.Name = "Courier new"
    
    prt 23 + BotLine, 55, "CHK DATE: ":    prt 23 + BotLine, 65, Format(PRBatch.CheckDate, "mm/dd/yyyy ")
    prt 23 + BotLine, 77, "CHK #: ":       prt 23 + BotLine, 85, CheckNum
    
    PosPrint 180, 5410 + BotAdd, "- - - - - - - - - -  CURRENT PD - - YR TO DATE":
    PosPrint 5800, 5410 + BotAdd, "- - - - - - - - -  CURRENT PD  - - - YR TO DATE"
    
    ' Prt 24 + BotLine, 47, "- - - - - - - - -  CURRENT PD  - - - YR TO DATE"
    
    SetFont 10, Equate.Portrait
    
    If PREmployee.UseAltName = 0 Then
        prt 42 + BotLine, 47, Trim(PREmployee.FLName):  prt 42 + BotLine, 76, "EMP. ID: " & PREmployee.EmployeeNumber
    Else
        prt 42 + BotLine, 47, Trim(PREmployee.AltName): prt 42 + BotLine, 76, "EMP. ID: " & PREmployee.EmployeeNumber
    End If
    PRDepartment.GetByID (PREmployee.DepartmentID)
    prt 43 + BotLine, 47, "DPT: " & PRDepartment.DepartmentNumber
    If frmCheckPrint.chkNoRate = 0 Then
        prt 43 + BotLine, 58, "RATE: " & Format(PRHist.RegRate, "##,###.#0")
    End If
    prt 43 + BotLine, 76, "PE DATE: " & Format(PRHist.PEDate, "mm/dd/yy")
    
    LeftLineNo = 25 + BotLine:  RightLineNo = 25 + BotLine
        
    ' +++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' loop thru the YTD temp record set for stub amounts
    rsYTD.MoveFirst
    SetFont 10, Equate.Portrait                     '   FONT IS 10
    Do
        ' left side - OE / DED
        If rsYTD!Side = "L1" Or rsYTD!Side = "L2" Or rsYTD!Side = "L3" Then
            If rsYTD!YTDHours <> 0 Then
                prt LeftLineNo, 1, Trim(rsYTD!Title) & " HRS":  prt LeftLineNo, 18, CurrFormat(rsYTD!CurrHours)
                prt LeftLineNo, 33, CurrFormat(rsYTD!YTDHours): LeftLineNo = LeftLineNo + 1
                End If
            prt LeftLineNo, 1, Trim(rsYTD!Title):               prt LeftLineNo, 18, CurrFormat(rsYTD!CurrAmt)
            prt LeftLineNo, 33, CurrFormat(rsYTD!YTDAmt):       LeftLineNo = LeftLineNo + 1
        Else        ' right side - summaries
            prt RightLineNo, 47, rsYTD!Title:                   prt RightLineNo, 63, CurrFormat(rsYTD!CurrAmt)
            prt RightLineNo, 80, CurrFormat(rsYTD!YTDAmt):      RightLineNo = RightLineNo + 1
        End If
        
        rsYTD.MoveNext
        If rsYTD.EOF Then Exit Do
    Loop
    
    prt LeftLineNo + 1, 1, frmCheckPrint.tdbtextMsg

End Sub


'          side id     amount     hrs    histid        title     ybatchid
'CheckPrintYtdUpdate "R", 10, PRHist.MedTax, 0, PRHist.HistID, "MED TAX", PRHist.BatchID

Private Sub CheckPrintYTDUpdate(ByVal Side As String, _
                      ByVal ID As Long, _
                      ByVal Amount As Currency, _
                      ByVal Hours As Currency, _
                      ByVal HistID As Long, _
                      ByVal Title As String, _
                      ByVal yBatchID As Long)

' update the temp record set for check printing

Dim YTDFlag As Boolean

    YTDFlag = False
    If rsYTD.RecordCount > 0 Then
        rsYTD.MoveFirst
        Do
            If rsYTD!Side = Side And rsYTD!ID = ID Then
                YTDFlag = True
                Exit Do
            Else
                '
            End If
            rsYTD.MoveNext
            If rsYTD.EOF Then Exit Do
        Loop
    End If
    
    ' add a new record
    If YTDFlag = False Then
        rsYTD.AddNew
        ' title "hard coded" for right side - use employer item id for right side
        If IsNumeric(Title) Then
            If Not PRItem.GetByID(CLng(Title)) Then
                rsYTD!Title = Left("Other " & Title, 30)
            Else
                rsYTD!Title = Left(PRItem.Title, 30)
            End If
        Else
            rsYTD!Title = Mid(Title, 1, 30)
        End If
        rsYTD!ID = ID
        rsYTD!Side = Side
    End If

    rsYTD!YTDAmt = rsYTD!YTDAmt + Amount
    rsYTD!YTDHours = rsYTD!YTDHours + Hours
    
    ' xxx If yBatchID = frmCheckPrint.BatchID Then
    If HistID = frmCheckPrint.trs!HistID Then
        rsYTD!CurrAmt = rsYTD!CurrAmt + Amount
        rsYTD!CurrHours = rsYTD!CurrHours + Hours
    End If

    rsYTD.Update

End Sub

Private Sub CheckPrintPrePrintA()        ' #########     PRE-PRINTED FORM   #############
Dim NoAsterisks As Long
    Prvw.vsp.Font.Name = "ARIAL":       Prvw.vsp.FontBold = False
    SetFont 10, Equate.Portrait          '  FONT IS 10
    
    If PREmployee.UseAltName = 0 Then
        prt 9, 13, Trim(PREmployee.FLName)
    Else
        prt 9, 13, Trim(PREmployee.AltName)
    End If
 ''''''''''''''''''''''''''''''''''''''    PAY SECTION    ''''''''''''''''''''''''''''''''''''''''
 
    If PRHist.Net <= 0 Then         '   Net is <= Zero
        
        prt 8, 51, PRHist.CheckNumber
        prt 8, 60, Format(PRHist.CheckDate, "mm/dd/yyyy")
        prt 8, 70, "NO DOLLARS AND NO CENTS"
        RJAmount = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
        prt 12, 7, RJAmount

        PosPrint 7100, 3460, "   *** V O I D   V O I D  V O I D ***   "
    Else
     
        PosPrint 5700, 1850, PRHist.CheckNumber
        PosPrint 7700, 1850, Format(PRHist.CheckDate, "mm/dd/yyyy")
        PosPrint 9800, 1850, CheckAmount(PRHist.Net)

        WrittenAmount = AmountInWords(PRHist.Net, False)
        Prvw.vsp.CurrentX = 10700 - Prvw.vsp.TextWidth(Trim(WrittenAmount)) + (Nudge * HorzNudge)
        Prvw.vsp.CurrentY = 2310 + (Nudge * VertNudge)
        Prvw.vsp.text = Trim(WrittenAmount)

    End If
    
    SetFont 10, Equate.Portrait                 '  FONT IS 10
    
    If PREmployee.UseAltName = 0 Then
        prt 14, 10, Trim(PREmployee.FLName)
    Else
        prt 14, 10, Trim(PREmployee.AltName)
    End If
    
    ILineNo = 14
    
    If Trim(PREmployee.Address1) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address1)
    End If
    
    If Trim(PREmployee.Address2) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address2)
    End If
    
    If Trim(PREmployee.City) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, PREmployee.CSZ
    End If
    
    PosPrint 900, 4160, PREmployee.CheckComment
                    
    ' scott molders - type "a" w/ signature
    If PRHist.Net > 0 And LCase(ChkPrefix) = "sct" Then
        If Trim(PRCheck.SignImage1) <> "" Then
            CheckPrintSignPrint PRCheck.Sign1Left + (Nudge * HorzNudge), PRCheck.Sign1Top + (Nudge * VertNudge) - 150, PRCheck.Sign1Width, PRCheck.Sign1Height, PRCheck.SignImage1
        End If
    End If
                    
End Sub

Private Sub CheckPrintPrePrintB()        ' #########     PRE-PRINTED FORM   #############
Dim NoAsterisks As Long
    Prvw.vsp.Font.Name = "ARIAL":       Prvw.vsp.FontBold = False
    SetFont 10, Equate.Portrait          '  FONT IS 10
    
    prt 6, 82, Format(PRHist.CheckDate, "mm/dd/yyyy")
    
    If PREmployee.UseAltName = 0 Then
        prt 9, 13, Trim(PREmployee.FLName)
    Else
        prt 9, 13, Trim(PREmployee.AltName)
    End If
 ''''''''''''''''''''''''''''''''''''''    PAY SECTION    ''''''''''''''''''''''''''''''''''''''''
 
    If PRHist.Net <= 0 Then         '   Net is <= Zero
        
        prt 8, 70, "NO DOLLARS AND NO CENTS"
        RJAmount = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
        prt 12, 7, RJAmount

        PosPrint 7100, 3460, "   *** V O I D   V O I D  V O I D ***   "
    Else
        
        PosPrint 9750, 2030, CheckAmount(PRHist.Net)

        WrittenAmount = AmountInWords(PRHist.Net, False)
        Prvw.vsp.CurrentX = 10200 - Prvw.vsp.TextWidth(Trim(WrittenAmount)) + (Nudge * HorzNudge)
        Prvw.vsp.CurrentY = 2485 + (Nudge * VertNudge)
        Prvw.vsp.text = Trim(WrittenAmount)

    End If
    
    SetFont 10, Equate.Portrait                 '  FONT IS 10
    
    If PREmployee.UseAltName = 0 Then
        prt 14, 10, Trim(PREmployee.FLName)
    Else
        prt 14, 10, Trim(PREmployee.AltName)
    End If
    
    ILineNo = 14
    
    If Trim(PREmployee.Address1) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address1)
    End If
    
    If Trim(PREmployee.Address2) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address2)
    End If
    
    If Trim(PREmployee.City) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, PREmployee.CSZ
    End If
    
    PosPrint 900, 4160, PREmployee.CheckComment
                    
End Sub
Private Sub CheckPrintPrePrintC()        ' #########     PRE-PRINTED FORM   #############
Dim NoAsterisks As Long
    
    Prvw.vsp.Font.Name = "ARIAL":       Prvw.vsp.FontBold = False
    SetFont 10, Equate.Portrait          '  FONT IS 10
    
    prt 6, 82, Format(PRHist.CheckDate, "mm/dd/yyyy")
    
    If PREmployee.UseAltName = 0 Then
        prt 9, 13, Trim(PREmployee.FLName)
    Else
        prt 9, 13, Trim(PREmployee.AltName)
    End If
 ''''''''''''''''''''''''''''''''''''''    PAY SECTION    ''''''''''''''''''''''''''''''''''''''''
 
    If PRHist.Net <= 0 Then         '   Net is <= Zero
        
        prt 9, 80, "XXXXXXXXXXXX"
        RJAmount = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
        prt 11, 7, RJAmount

        PosPrint 7700, 3600, "   *** V O I D   V O I D  V O I D ***   "
    Else
        
        PosPrint 9750, 2030, CheckAmount(PRHist.Net)

        WrittenAmount = AmountInWords(PRHist.Net, False)
        Prvw.vsp.CurrentX = 10200 - Prvw.vsp.TextWidth(Trim(WrittenAmount)) + (Nudge * HorzNudge)
        Prvw.vsp.CurrentY = 2485 + (Nudge * VertNudge)
        Prvw.vsp.text = Trim(WrittenAmount)

    End If
    
    SetFont 10, Equate.Portrait                 '  FONT IS 10
    
    If PREmployee.UseAltName = 0 Then
        prt 14, 10, Trim(PREmployee.FLName)
    Else
        prt 14, 10, Trim(PREmployee.AltName)
    End If
    
    ILineNo = 14
    
    If Trim(PREmployee.Address1) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address1)
    End If
    
    If Trim(PREmployee.Address2) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address2)
    End If
    
    If Trim(PREmployee.City) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, PREmployee.CSZ
    End If
    
    PosPrint 900, 4160, PREmployee.CheckComment
                    
End Sub

Private Sub CheckPrintBlankForm()
Dim CheckBankName As String
Dim BoldVertAdd, NormlVertAdd, VertAdd As Long
Dim CkNum As String
Dim VertPosn, HorzPosn As Integer
Dim BankAcctString As String
    
    Prvw.vsp.Font.Name = "ARIAL"
    Prvw.vsp.FontBold = True      ' BOLD IS ON
    
    SetFont 13, Equate.Portrait             '  FONT IS 13
    CkNum = Space(9 - Len(Trim(CStr(CheckNum)))) & CheckNum
    PosPrint 10000, 300, CkNum
    
    ILineNo = 1
    
'''''''''''''''''''''''''''''''''''  Bank Address Section  '''''''''''''''''''''''''''''''''

    If PRHist.Net > 0 Then

        HorzPosn = 70

        '   FONT IS 8 and BOLD is ON
        SetFont 8, Equate.Portrait:                         Prvw.vsp.FontBold = True
        prt ILineNo, HorzPosn, Trim(PRCheck.Bank1)
        
        If Trim(PRCheck.Bank2) <> "" Then
            ILineNo = ILineNo + 1
            prt ILineNo, HorzPosn, Trim(PRCheck.Bank2)
        End If
        
        If Trim(PRCheck.Bank3) <> "" Then
            ILineNo = ILineNo + 1
            prt ILineNo, HorzPosn, Trim(PRCheck.Bank3)
        End If
        
        If Trim(PRCheck.Bank4) <> "" Then
            ILineNo = ILineNo + 1
            prt ILineNo, HorzPosn, Trim(PRCheck.Bank4)
        End If
        
        If Trim(PRCheck.BankFraction) <> "" Then
            ILineNo = ILineNo + 1
            prt ILineNo, HorzPosn, Trim(PRCheck.BankFraction)
        End If

    End If

'''''''''''''''''''''''''''''''''''  Company Address Section  '''''''''''''''''''''''''''''''''

    '  Turn off BOLD feature and reduce font size for Company Address Section
                                                    '  FONT IS 10
    HorzPosn = 800
    
    Prvw.vsp.FontBold = True
    SetFont 13, Equate.Portrait             '  FONT IS 13
    PosPrint HorzPosn, 300, Trim(PRCheck.CustomerName):   Prvw.vsp.FontBold = False  '  BOLD IS OFF
    
    Prvw.vsp.FontBold = False:                      SetFont 10, Equate.Portrait
    
    BoldVertAdd = 300
    NormlVertAdd = 250
    ILineNo = 580
    
    If Trim(PRCheck.Address1) <> "" Then
        If PRCheck.Addr1Bold = 1 Then
            Prvw.vsp.FontBold = True
            SetFont 13, Equate.Portrait             '  FONT IS 10
            VertAdd = BoldVertAdd
        Else
            Prvw.vsp.FontBold = False
            SetFont 8, Equate.Portrait              '  FONT IS 8
            VertAdd = NormlVertAdd
        End If
        PosPrint HorzPosn, ILineNo, Trim(PRCheck.Address1)
    End If

    ILineNo = ILineNo + VertAdd
    
    If Trim(PRCheck.Address2) <> "" Then
        If PRCheck.Addr2Bold = 1 Then
            Prvw.vsp.FontBold = True
            SetFont 13, Equate.Portrait              '  FONT IS 8
            VertAdd = BoldVertAdd
        Else
            Prvw.vsp.FontBold = False
            SetFont 8, Equate.Portrait                      '  FONT IS 8
            VertAdd = NormlVertAdd
        End If
        PosPrint HorzPosn, ILineNo, Trim(PRCheck.Address2)
    End If
    
    ILineNo = ILineNo + VertAdd
    VertAdd = NormlVertAdd
    
    SetFont 8, Equate.Portrait                      '  FONT IS 8
    
    If Trim(PRCheck.Address3) <> "" Then
        If PRCheck.Addr3Bold = 1 Then
            Prvw.vsp.FontBold = True
        Else
            Prvw.vsp.FontBold = False
        End If
        PosPrint HorzPosn, ILineNo, Trim(PRCheck.Address3)
    End If
    
    ILineNo = ILineNo + VertAdd
    
    If Trim(PRCheck.Address4) <> "" Then
        If PRCheck.Addr4Bold = 1 Then
            Prvw.vsp.FontBold = True
        Else
            Prvw.vsp.FontBold = False
        End If
        PosPrint HorzPosn, ILineNo, Trim(PRCheck.Address4)
    
    End If
    
''''''''''''''''''''''''''''''''''''''    Employee Address Section  ''''''''''''''''''''''''''''''''''''''''
    SetFont 10, Equate.Portrait
    Prvw.vsp.FontBold = False
    Prvw.vsp.Font.Name = "ARIAL":
    
    ' Prt 8, 1, "PAY"
    Ln = 8
    PrintValue(1) = "PAY":      FormatString(1) = "a5"
    PrintValue(2) = "":         FormatString(2) = "~"
    FormatPrint

    SetFont 8, Equate.Portrait                      '  FONT IS 8
    prt 13, 1, "TO THE":             prt 14, 1, "ORDER"
    prt 15, 3, "OF"
    
    SetFont 10, Equate.Portrait                 '  FONT IS 10

    If PREmployee.UseAltName = 1 Then
        prt 13, 10, Trim(PREmployee.AltName)
    Else
        prt 13, 10, Trim(PREmployee.FLName)
    End If
    
    ILineNo = 13
    
    If Trim(PREmployee.Address1) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address1)
    End If
    
    If Trim(PREmployee.Address2) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.Address2)
    End If
        
    If Trim(PREmployee.City) <> "" Then
        ILineNo = ILineNo + 1
        prt ILineNo, 10, Trim(PREmployee.CSZ)
    End If
    
    prt 17, 10, PREmployee.CheckComment
            
''''''''''''''''''''''''''''''''''''''    PAY SECTION    ''''''''''''''''''''''''''''''''''''''''
            
    If PRHist.Net <= 0 Then         '   Net is Zero
        prt 8, 70, "NO DOLLARS AND NO CENTS"
        RJAmount = "*******    VOID ** VOID **  VOID ** NOTICE OF DEPOSIT ONLY - SEE STUB FOR DEPOSIT INFORMATION    *******"
        prt 10, 7, RJAmount

        PosPrint 7100, 3460, "   *** V O I D   V O I D  V O I D ***   "
    Else
        
        WrittenAmount = AmountInWords(PRHist.Net, True)
                
        ' RJAmount = String(77 - Len(Trim(WrittenAmount)), " ") & Trim(WrittenAmount)
        ' Prt 8, 37, RJAmount
        
'        Ln = 8
'        FormatString(1) = "r77":    PrintValue(1) = Trim(WrittenAmount)
'        FormatString(2) = "~":      PrintValue(2) = ""
'        FormatPrint
        
        Prvw.vsp.CurrentX = 11350 - Prvw.vsp.TextWidth(Trim(WrittenAmount)) + (Nudge * HorzNudge)
        Prvw.vsp.CurrentY = 1900 + (Nudge * VertNudge)
        Prvw.vsp.text = Trim(WrittenAmount)
        
        prt 10, 55, CheckNum
        prt 10, 67, Format(PRBatch.CheckDate, "mm/dd/yyyy")
        prt 10, 82, CheckAmount(PRHist.Net)
    
    End If

     ''''''''''''' ''''''''''''''''''    SIGNATURE SECTION   ''''''''''''''''''''''''''''''''''''
    
    If PRHist.Net > 0 Then
    
        If Trim(PRCheck.SignImage1) <> "" Then
            CheckPrintSignPrint PRCheck.Sign1Left + (Nudge * HorzNudge), PRCheck.Sign1Top + (Nudge * VertNudge) - 150, PRCheck.Sign1Width, PRCheck.Sign1Height, PRCheck.SignImage1
        End If
        ILineNo = 14

    End If
    
    If PRCheck.TwoSignLines = 1 Then
        PosPrint 6600, 3050, String(43, "_")
        PosPrint 6600, 3450, String(43, "_"):        ILineNo = ILineNo + 1
    Else
        ILineNo = ILineNo + 2
        PosPrint 6600, 3450, String(43, "_"):        ILineNo = ILineNo + 1
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''  MICR ENCODING SECTION  ''''''''''''''''''''''''''''''

    If PRHist.Net > 0 Then

        VertPosn = 4380

        Prvw.vsp.Font.Name = "MICR Encoding"
        Prvw.vsp.Font.Size = 18
        
        ' check number
        PosPrint 1590, VertPosn, "C" & Format(PRHist.CheckNumber, "000000000") & "C"
        
        ' ABA Number
        PosPrint 3745, VertPosn, "A" & Trim(PRCheck.BankABA) & "A"
        
        ' Account Number
        BankAcctString = Trim(PRCheck.BankAccount) & "C"
        If PRCheck.BankAccountAdd <> "" Then
            BankAcctString = Trim(BankAcctString) & PRCheck.BankAccountAdd
        End If
        
        If PRCheck.AccountSpace = 2 Then
            PosPrint 6070, VertPosn, Trim(BankAcctString)
        Else
            PosPrint 5890, VertPosn, Trim(BankAcctString)
        End If
                
    End If
                
End Sub

Private Sub CheckPrintSignPrint(ByVal SignLeft As Long, _
                      ByVal SignTop As Long, _
                      ByVal SignWidth As Long, _
                      ByVal SignHeight As Long, _
                      ByVal FileName As String)
        
Dim prFileName As String
        
    If BalintFolder = "" Then
        prFileName = "\Balint\Data\" & FileName
    Else
        prFileName = Replace(BalintFolder, "^", " ") & "\Data\" & FileName
    End If
    
    
    Prvw.Picture1.Picture = LoadPicture(Trim(prFileName))
    
    Prvw.vsp.DrawPicture Prvw.Picture1, SignLeft, SignTop, SignWidth, SignHeight, 10

End Sub

Public Sub TaxableWageRpt(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal StartDate As Long, _
                         ByVal EndDate As Long, _
                         ByVal OptDate As String)
                         
Dim EmpName As String
Dim TotAmt As Currency
Dim sqlstring1 As String
Dim LastFLName As String
Dim LastPEDate As Date
Dim LastCheckDate As Date
Dim ct As Long
   
    frmTaxWage.Hide
    ReportTitle = "PAYROLL TAXABLE WAGE REPORT"
    PrtInit ("Land")
    LandSw = 1
    Columns = 145
    SetFont 8, Equate.LandScape
    PRTotal.CreateRS

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "EmpNo", adDouble:                            rs.Fields.Append "DeptID", adDouble
    rs.Fields.Append "DeptNo", adDouble
    rs.Fields.Append "EmpID", adDouble:                            rs.Fields.Append "FLName", adChar, 30, adFldMayBeNull
    rs.Fields.Append "PEDate", adDate:                             rs.Fields.Append "CheckDate", adDate
    rs.Fields.Append "Gross", adCurrency:                          rs.Fields.Append "SSWage", adCurrency
    rs.Fields.Append "MEDWage", adCurrency:                        rs.Fields.Append "FWTWage", adCurrency
    rs.Fields.Append "SWTWage", adCurrency:                        rs.Fields.Append "CWTWage", adCurrency
    rs.Fields.Append "FUNWage", adCurrency:                        rs.Fields.Append "SUNWage", adCurrency

    rs.Open , , adOpenDynamic, adLockOptimistic

    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PR Batch Not Found: " & PRBatchID, vbExclamation
            GoBack
        End If
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & PRBatch.PEDate
    End If
    
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
        OptDate = "P/E DATE"

    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRHist WHERE PRHist.CheckDate >= " & CLng(StartDate) & _
                        " AND PRHist.CheckDate <= " & CLng(EndDate)
            Msg1 = "CHECK DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        ElseIf OptDate = "P/E DATE" Then
             SQLString = "SELECT * FROM PRHist WHERE PRHist.PEDate >= " & CLng(StartDate) & _
                        " AND PRHist.PEDate <= " & CLng(EndDate)
            Msg1 = "PERIOD ENDING DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If
    
    If frmTaxWage.optEmpNo Then
        ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NUMBER"
    Else
        ReportTitle = Trim(ReportTitle) & " BY CHECK DATE"
    End If
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbExclamation, "Tax Wage Report"
        GoBack
    End If

    Do
        ' employee select filter?  Populate Recordset
        If frmEmpSelect.AllEmployees = False Then

            SQLString = "EmployeeID = " & PRHist.EmployeeID
            frmEmpSelect.rsEmp.Find SQLString, 0, adSearchForward, 1
            If frmEmpSelect.rsEmp.EOF Then
                GoTo NextHist
            End If
            If frmEmpSelect.rsEmp!Select = False Then GoTo NextHist
        End If

        rs.AddNew
        rs!EmpID = PRHist.EmployeeID
        If PREmployee.GetByID(PRHist.EmployeeID) Then
            rs!EmpNo = PREmployee.EmployeeNumber
            rs!FLName = PREmployee.FLName
        End If
        rs!PEDate = PRHist.PEDate
        rs!CheckDate = PRHist.CheckDate
        rs!DeptID = PRHist.DepartmentID
        rs!Gross = PRHist.Gross
        rs!SSWage = PRHist.SSWage
        rs!MEDWage = PRHist.MEDWage
        rs!FWTWage = PRHist.FWTWage
        rs!SWTWage = PRHist.SWTWage
        rs!CWTWage = PRHist.CWTWage
        rs!FUNWage = PRHist.FUNWage
        rs!SUNWage = PRHist.SUNWage
        rs!DeptNo = PREmployee.DepartmentID
        rs.Update
NextHist:
        If Not PRHist.GetNext Then Exit Do
    Loop
    
    If frmTaxWage.optEmpNo Then
        rs.Sort = "EmpNo"
    Else
        rs.Sort = "CheckDate"
    End If
    
    rs.MoveFirst
    
    LastEmpNo = 0
    LastFLName = " "
    LastCheckDate = 0
    LastPEDate = 0
    
    Do Until rs.EOF
        If Ln = 0 Or Ln >= MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = Ln + 1
            TaxableWageHeader
        End If
                
        ' Check to see if subtotal should be printed
        If frmTaxWage.optEmpNo Then
            If LastEmpNo <> 0 And LastEmpNo <> rs!EmpNo Then
                If ct > 1 Then
                    TaxableWageSub ("TOTAL " & LastEmpNo & " - " & LastFLName)
                    ct = 0
                Else
                    TGross = 0
                    TSSWage = 0
                    TMedWage = 0
                    TFWTWage = 0
                    TSWTWAGE = 0
                    TCWTWage = 0
                    TSUNWage = 0
                    TFUNWage = 0
                    ct = 0
                End If
                Ln = Ln + 1
            End If
        Else
            If LastCheckDate <> 0 And LastCheckDate <> rs!CheckDate Then
                If ct > 1 Then
                    TaxableWageSub ("TOTAL - CHECK DATE: " & LastCheckDate)
                    ct = 0
                Else
                    TGross = 0
                    TSSWage = 0
                    TMedWage = 0
                    TFWTWage = 0
                    TSWTWAGE = 0
                    TCWTWage = 0
                    TSUNWage = 0
                    TFUNWage = 0
                    ct = 0
                End If
                Ln = Ln + 1
            End If
        End If
        ct = ct + 1
        
        ' Print detail line
        PrintValue(1) = rs!EmpNo:                               FormatString(1) = "a6"
        PrintValue(2) = rs!FLName:                              FormatString(2) = "a40"
        PrintValue(3) = Format(rs!PEDate, "mm/dd/yy"):          FormatString(3) = "a10"
        PrintValue(4) = Format(rs!CheckDate, "mm/dd/yy"):       FormatString(4) = "a11"
        PrintValue(5) = rs!Gross:                               FormatString(5) = "d14"
        PrintValue(6) = rs!SSWage:                              FormatString(6) = "d14"
        PrintValue(7) = rs!MEDWage:                             FormatString(7) = "d14"
        PrintValue(8) = rs!FWTWage:                             FormatString(8) = "d14"
        PrintValue(9) = rs!SWTWage:                             FormatString(9) = "d14"
        PrintValue(10) = " ":                                   FormatString(10) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                                    FormatString(1) = "a67"
        PrintValue(2) = rs!CWTWage:                             FormatString(2) = "d14"
        PrintValue(3) = rs!FUNWage:                             FormatString(3) = "d14"
        PrintValue(4) = rs!SUNWage:                             FormatString(4) = "d14"
        PrintValue(5) = " ":                                    FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        ' Accumulate totals
        TGross = TGross + rs!Gross
        TSSWage = TSSWage + rs!SSWage
        TMedWage = TMedWage + rs!MEDWage
        TFWTWage = TFWTWage + rs!FWTWage
        TSWTWAGE = TSWTWAGE + rs!SWTWage
        TCWTWage = TCWTWage + rs!CWTWage
        TSUNWage = TSUNWage + rs!SUNWage
        TFUNWage = TFUNWage + rs!FUNWage
        
        GTGross = GTGross + rs!Gross
        GTSSWage = GTSSWage + rs!SSWage
        GTMedWage = GTMedWage + rs!MEDWage
        GTFWTWage = GTFWTWage + rs!FWTWage
        GTSWTWAGE = GTSWTWAGE + rs!SWTWage
        GTCWTWage = GTCWTWage + rs!CWTWage
        GTSUNWage = GTSUNWage + rs!SUNWage
        GTFUNWage = GTFUNWage + rs!FUNWage
        
        ' Save Variables for comparison or printing
        LastEmpNo = rs!EmpNo
        LastFLName = rs!FLName
        LastPEDate = rs!PEDate
        LastCheckDate = rs!CheckDate
        
        ' write/update total records in prtotal
        If Not PRTotal.tFind(PREquate.GLTypeDept, rs!DeptID) Then
            PRTotal.Clear
            PRTotal.RecType = PREquate.GLTypeDept
            PRTotal.RecID = rs!DeptID
            PRTotal.Save (Equate.RecAdd)
        End If
        
        PRTotal.Gross = PRTotal.Gross + rs!Gross
        PRTotal.SSWage = PRTotal.SSWage + rs!SSWage
        PRTotal.MEDWage = PRTotal.MEDWage + rs!MEDWage
        PRTotal.FWTWage = PRTotal.FWTWage + rs!FWTWage
        PRTotal.StateWage = PRTotal.StateWage + rs!SWTWage
        PRTotal.CityWage = PRTotal.CityWage + rs!CWTWage
        PRTotal.FUNWage = PRTotal.FUNWage + rs!FUNWage
        PRTotal.SUNWage = PRTotal.SUNWage + rs!SUNWage
                    
        PRTotal.Save (Equate.RecPut)

        rs.MoveNext
        If rs.EOF Then
            Exit Do
        End If
    Loop

    ' Print Last Subtotal and Grand Totals
    If ct > 1 Then
        If frmTaxWage.optEmpNo Then
            TaxableWageSub ("TOTAL - " & LastEmpNo & " - " & LastFLName)
        Else
            TaxableWageSub ("TOTAL - CHECK DATE: " & LastCheckDate)
        End If
    End If
    
    ' Print Department Totals?
    If frmTaxWage.chkDeptTotals Then
        Ln = Ln + 1
        TaxableWageDeptTotals
    End If
    
    ' Print Grand Totals
    Ln = Ln + 1
    PrintValue(1) = "GRAND TOTALS: ":                       FormatString(1) = "a67"
    PrintValue(2) = GTGross:                                FormatString(2) = "d14"
    PrintValue(3) = GTSSWage:                               FormatString(3) = "d14"
    PrintValue(4) = GTMedWage:                              FormatString(4) = "d14"
    PrintValue(5) = GTFWTWage:                              FormatString(5) = "d14"
    PrintValue(6) = GTSWTWAGE:                              FormatString(6) = "d14"
    PrintValue(7) = " ":                                    FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                                    FormatString(1) = "a67"
    PrintValue(2) = GTCWTWage:                              FormatString(2) = "d14"
    PrintValue(3) = GTFUNWage:                              FormatString(3) = "d14"
    PrintValue(4) = GTSUNWage:                              FormatString(4) = "d14"
    PrintValue(5) = " ":                                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub TaxableWageSub(ByVal SubVar As String)
    If Ln = 0 Or Ln > MaxLines Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        Ln = 5
        TaxableWageHeader
    End If
    PrintValue(1) = SubVar:                                 FormatString(1) = "a67"
    PrintValue(2) = TGross:                                 FormatString(2) = "d14"
    PrintValue(3) = TSSWage:                                FormatString(3) = "d14"
    PrintValue(4) = TMedWage:                               FormatString(4) = "d14"
    PrintValue(5) = TFWTWage:                               FormatString(5) = "d14"
    PrintValue(6) = TSWTWAGE:                               FormatString(6) = "d14"
    PrintValue(7) = " ":                                    FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                                    FormatString(1) = "a67"
    PrintValue(2) = TCWTWage:                               FormatString(2) = "d14"
    PrintValue(3) = TFUNWage:                               FormatString(3) = "d14"
    PrintValue(4) = TSUNWage:                               FormatString(4) = "d14"
    PrintValue(5) = " ":                                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    TGross = 0
    TSSWage = 0
    TMedWage = 0
    TFWTWage = 0
    TSWTWAGE = 0
    TCWTWage = 0
    TSUNWage = 0
    TFUNWage = 0
End Sub

Public Sub TaxableWageHeader()

    PrintValue(1) = "Emp #":            FormatString(1) = "a6"
    PrintValue(2) = "Employee Name":    FormatString(2) = "a40"
    PrintValue(3) = "PE Date":          FormatString(3) = "a10"
    PrintValue(4) = "CK Date":          FormatString(4) = "a16"
    PrintValue(5) = "Gross":            FormatString(5) = "a15"
    PrintValue(6) = "SS Wage":          FormatString(6) = "a13"
    PrintValue(7) = "MED Wage":         FormatString(7) = "a14"
    PrintValue(8) = "FWT Wage":         FormatString(8) = "a14"
    PrintValue(9) = "SWT Wage":         FormatString(9) = "a13"
    PrintValue(10) = " ":               FormatString(10) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":               FormatString(1) = "a72"
    PrintValue(2) = "CWT Wage":        FormatString(2) = "a14"
    PrintValue(3) = "FUN Wage":        FormatString(3) = "a14"
    PrintValue(4) = "SUN Wage":        FormatString(4) = "a13"
    PrintValue(5) = " ":               FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(140, "-"):   FormatString(1) = "a140"
    PrintValue(2) = " ":                FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
End Sub

Private Sub TaxableWageDeptTotals()

    PRTotal.TSortByString "RecId"
    If PRTotal.FindFirst = False Then
    End If

    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = 5
            TaxableWageHeader
        End If
  
        If PRDepartment.GetByID(PRTotal.RecID) Then
            PrintValue(1) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & PRDepartment.Name:  FormatString(1) = "a67"
            PrintValue(2) = PRTotal.Gross:                                  FormatString(2) = "d14"
            PrintValue(3) = PRTotal.SSWage:                                 FormatString(3) = "d14"
            PrintValue(4) = PRTotal.MEDWage:                                FormatString(4) = "d14"
            PrintValue(5) = PRTotal.FWTWage:                                FormatString(5) = "d14"
            PrintValue(6) = PRTotal.StateWage:                              FormatString(6) = "d14"
            PrintValue(7) = " ":                                            FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = " ":                                            FormatString(1) = "a67"
            PrintValue(2) = PRTotal.CityWage:                               FormatString(2) = "d14"
            PrintValue(3) = PRTotal.FUNWage:                                FormatString(3) = "d14"
            PrintValue(4) = PRTotal.SUNWage:                                FormatString(4) = "d14"
            PrintValue(5) = " ":                                            FormatString(5) = "~"
            FormatPrint
            Ln = Ln + 1
        End If

        If Not PRTotal.GetNext Then Exit Do
    Loop
                
End Sub

Public Sub PRBatchList(ByVal StartDate As Long, ByVal EndDate As Long)
Dim yr, yrmo, mo, Qtr As String
Dim Gross, TGross As Currency
Dim Recs, StartChkNo, EndChkNo, TRecs As Long


    frmBatchList.Hide
    ReportTitle = "PAYROLL BATCH LIST"
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    SetEquates
     
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If frmBatchList.chkAllYears Then
        SQLString = "SELECT * FROM PRBatch ORDER BY BatchID"     ' ALL YEARS
        Msg1 = "ALL HISTORY"
    Else
        If StartDate <> 0 And EndDate <> 0 Then
            SQLString = "SELECT * FROM PRBatch WHERE PRBatch.CheckDate >= " & StartDate & " And PRBatch.CheckDate <= " & EndDate
            Msg1 = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & _
                " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If

    If Not PRBatch.GetBySQL(SQLString) Then
        MsgBox "No Payroll Batch History Found for that Check Date Range", vbExclamation, "Payroll Batch List"
        GoBack
    End If
            
    Do
    
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = 5
            
            PrintValue(1) = " ":                    FormatString(1) = "a9"
            PrintValue(2) = "Create":               FormatString(2) = "a55"
            PrintValue(3) = "Total":                FormatString(3) = "a12"
            PrintValue(4) = "Start":                FormatString(4) = "a13"
            PrintValue(5) = "End":                  FormatString(5) = "a6"
            PrintValue(6) = " ":                    FormatString(6) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "Batch #":              FormatString(1) = "a10"
            PrintValue(2) = "Date":                 FormatString(2) = "a8"
            PrintValue(3) = "PE Date":              FormatString(3) = "a10"
            PrintValue(4) = "Chk Date":             FormatString(4) = "a10"
            PrintValue(5) = "Quarter":              FormatString(5) = "a9"
            PrintValue(6) = "# Recs":               FormatString(6) = "a17"
            PrintValue(7) = "Gross":                FormatString(7) = "a11"
            PrintValue(8) = "Check #":              FormatString(8) = "a12"
            PrintValue(9) = "Check #":              FormatString(9) = "a11"
            PrintValue(10) = " ":                   FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
    
            PrintValue(1) = String(95, "-"):   FormatString(1) = "a95"
            PrintValue(2) = " ":                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1

        End If
        
        yrmo = PRBatch.YearMonth
        mo = Mid(yrmo, 5, 2)
        yr = Mid(yrmo, 1, 4)
        
        If mo = "01" Then
            Qtr = yr & "-" & 1
        ElseIf mo = "02" Then
            Qtr = yr & "-" & 1
        ElseIf mo = "03" Then
            Qtr = yr & "-" & 1
        ElseIf mo = "04" Then
            Qtr = yr & "-" & 2
        ElseIf mo = "05" Then
            Qtr = yr & "-" & 2
        ElseIf mo = "06" Then
            Qtr = yr & "-" & 2
        ElseIf mo = "07" Then
            Qtr = yr & "-" & 3
        ElseIf mo = "08" Then
            Qtr = yr & "-" & 3
        ElseIf mo = "09" Then
            Qtr = yr & "-" & 3
        ElseIf mo = "10" Then
            Qtr = yr & "-" & 4
        ElseIf mo = "11" Then
            Qtr = yr & "-" & 4
        ElseIf mo = "12" Then
            Qtr = yr & "-" & 4
        Else
            Qtr = "??"
        End If
        
        ' loop thru PRHist and gather additional info

        SQLString = "SELECT * FROM PRHist WHERE BatchId = " & PRBatch.BatchID & _
                    " ORDER BY PRHist.CheckNumber"
    
        If Not PRHist.GetBySQL(SQLString) Then
            MsgBox "No Payroll History Found for that Batch", vbExclamation, "Payroll Batch List"
            GoBack
        End If
        
        Recs = 0
        Gross = 0
        Do
            Recs = Recs + 1
            TRecs = TRecs + 1
            Gross = Gross + PRHist.Gross
            TGross = TGross + PRHist.Gross
            If Recs = 1 Then
                StartChkNo = PRHist.CheckNumber
            End If
            EndChkNo = PRHist.CheckNumber
            If Not PRHist.GetNext Then Exit Do
        Loop
        
        PrintValue(1) = PRBatch.BatchID:                        FormatString(1) = "a8"
        PrintValue(2) = Format(PRBatch.CreateDate, "mm/dd/yy"): FormatString(2) = "a10"
        PrintValue(3) = Format(PRBatch.PEDate, "mm/dd/yy"):     FormatString(3) = "a10"
        PrintValue(4) = Format(PRBatch.CheckDate, "mm/dd/yy"):  FormatString(4) = "a10"
        PrintValue(5) = Qtr:                                    FormatString(5) = "a9"
        PrintValue(6) = Recs:                                   FormatString(6) = "n6"
        PrintValue(7) = " ":                                    FormatString(7) = "a3"
        PrintValue(8) = Gross:                                  FormatString(8) = "d14"
        PrintValue(9) = StartChkNo:                             FormatString(9) = "n12"
        PrintValue(10) = EndChkNo:                              FormatString(10) = "n12"
        PrintValue(11) = " ":                                   FormatString(11) = "~"
        FormatPrint
        Ln = Ln + 1
        
        If Not PRBatch.GetNext Then Exit Do
   Loop
            
    ' Print Total Line
    Ln = Ln + 1
    PrintValue(1) = "Grand Total ":                             FormatString(1) = "a47"
    PrintValue(2) = TRecs:                                      FormatString(2) = "n6"
    PrintValue(3) = " ":                                        FormatString(3) = "a3"
    PrintValue(4) = TGross:                                     FormatString(4) = "d14"
    PrintValue(5) = " ":                                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
        
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
        
End Sub
'Private Sub PageHeader(ByVal ReportName As String, _
'                       ByVal Msg1 As String, _
'                       ByVal Msg2 As String, _
'                       ByVal msg3 As String)
'
'    Ln = 0
'    Pg = Pg + 1
'
'    ' 29 characters for fixed left and right portion of first header line
'    '    1             8       1   8                    10         1
'    ' first line - system date & time / company name / page #
'    x = PRCompany.Name
'    Y = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
'    z = "Page: " & Format(Pg, "####")
'
'    If Len(x) > Columns - 17 Then
'       x = Mid(Trim(PRCompany.Name), 1, Columns - 27)
'    End If
'
'    If LandSw = 1 Then
'        Columns = 145               '  originally columns = 165
'        i = ((Columns - Len(x)) / 2) - 29
'        w = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss") & _
'            Space(i) & x
'        i = Columns - Len(w) - 30
'        w = w & Space(i) & "Page: " & Format(Pg, "###0")
'    Else
'        i = ((Columns - Len(x)) / 2) - 19
'        w = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss") & _
'            Space(i) & x
'        i = Columns - Len(w) - 10
'        w = w & Space(i) & "Page: " & Format(Pg, "###0")
'
'    End If
'
'    PrtCenter Ln, w
'    Ln = Ln + 1
'
'    If ReportName <> "" Then
'        PrtCenter 0, ReportName
'        Ln = Ln + 1
'    End If
'
'    If Msg1 <> "" Then
'       PrtCenter Ln, Msg1
'       Ln = Ln + 1
'    End If
'
'    If Msg2 <> "" Then
'       PrtCenter Ln, Msg2
'       Ln = Ln + 1
'    End If
'
'    If msg3 <> "" Then
'       PrtCenter Ln, msg3
'       Ln = Ln + 1
'    End If
'
'    Ln = Ln + 1
'
'End Sub


Public Sub ItemDetail(ByVal RangeType As Byte, _
                      ByVal BatchNumbr As Long, _
                      ByVal PEDate As Long, _
                      ByVal CheckDt As Long, _
                      ByVal StartDate As Long, _
                      ByVal EndDate As Long, _
                      ByVal OptDate As String)

Dim Item1ID, Item2ID, Item3ID, Item4ID, Item5ID, ctr, dtlCount As Long
Dim ItsHours1, ItsHours2, ItsHours3, ItsHours4, ItsHours5 As Boolean
Dim SGross, SItem1, SItem2, SItem3, SItem4, SItem5, SMax, SRemaining As Currency
Dim TGross, TItem1, TItem2, TItem3, TItem4, TItem5, AmtRemaining, TMax, TRemaining As Currency

    ' special grand totals if remaining displayed
Dim RmnMax, RmnRemain As Currency
Dim RmnMaxTL, RmnRemainTL As Currency
Dim ItemFlag As Boolean
    
    RmnMax = 0
    RmnRemain = 0
    
    SetEquates
    PrtInit ("Land")
    LandSw = 1
    SetFont 8, Equate.LandScape
    Columns = Columns - 5
    
    ReportTitle = "PAYROLL ITEM DETAIL LISTING"
    If frmItemDetail.optChkDate = True Then
        Msg2 = "ORDER BY CHECK DATE"
    Else
        Msg2 = "ORDER BY EMPLOYEE NUMBER"
    End If
    
    frmItemDetail.rsItem.MoveFirst
    ItemCount = 0
    Do
        If frmItemDetail.rsItem!Selected = True Then
            ItemCount = ItemCount + 1
            If ItemCount = 1 Then
                Item1ID = frmItemDetail.rsItem!ItemID
                If frmItemDetail.rsItem!IsItHours Then
                    ItsHours1 = True
                End If
            End If
            If ItemCount = 2 Then
                Item2ID = frmItemDetail.rsItem!ItemID
                If frmItemDetail.rsItem!IsItHours Then
                    ItsHours2 = True
                End If
            End If
            If ItemCount = 3 Then
                Item3ID = frmItemDetail.rsItem!ItemID
                If frmItemDetail.rsItem!IsItHours Then
                    ItsHours3 = True
                End If
            End If
            If ItemCount = 4 Then
                Item4ID = frmItemDetail.rsItem!ItemID
                If frmItemDetail.rsItem!IsItHours Then
                    ItsHours4 = True
                End If
            End If
            If ItemCount = 5 Then
                Item5ID = frmItemDetail.rsItem!ItemID
                If frmItemDetail.rsItem!IsItHours Then
                    ItsHours5 = True
                End If
            End If
            
        End If
        frmItemDetail.rsItem.MoveNext
        If frmItemDetail.rsItem.EOF Then Exit Do
    Loop
    
    trs.CursorLocation = adUseClient
    trs.Fields.Append "EmpName", adChar, 30, adFldMayBeNull
    trs.Fields.Append "EmpID", adDouble:                            trs.Fields.Append "EmpNo", adDouble:
    trs.Fields.Append "HistID", adDouble:                           trs.Fields.Append "PEDate", adDouble:
    trs.Fields.Append "ChkDate", adDate:                            trs.Fields.Append "Gross", adCurrency:
    trs.Fields.Append "Item1Amount", adCurrency:                    trs.Fields.Append "Item2Amount", adCurrency
    trs.Fields.Append "Item3Amount", adCurrency:                    trs.Fields.Append "Item4Amount", adCurrency
    trs.Fields.Append "Item5Amount", adCurrency:                    trs.Fields.Append "YTDGross", adCurrency
    trs.Fields.Append "MaxAmt", adCurrency
      
    trs.Open , , adOpenDynamic, adLockOptimistic

    frmItemDetail.Hide

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
            
    ' =================== OE from PRDist =========================================
    
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.BatchID = " & BatchNumbr & _
        " AND (ItemType = " & PREquate.ItemTypeDED & " OR ItemType = " & PREquate.ItemTypeOE & ")"
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRDist WHERE PRDist.CheckDate >= " & (StartDate) & " AND " & _
                                     " PRDist.CheckDate <= " & (EndDate) & _
                                     " AND (ItemType = " & PREquate.ItemTypeDED & _
                                     " OR ItemType = " & PREquate.ItemTypeOE & ")"
            Msg1 = "CHECK DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
                                    
        Else
            SQLString = "SELECT * FROM PRDist WHERE PRDist.PEDate >= " & (StartDate) & " AND " & _
                                    " PRDist.PEDate <= " & (EndDate) & _
                                     " AND (ItemType = " & PREquate.ItemTypeDED & _
                                     " OR ItemType = " & PREquate.ItemTypeOE & ")"
            Msg1 = "P/E DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If

    If PRDist.GetBySQL(SQLString) Then

        Recs = PRDist.Records
        ct = 0

        Do
            
            ct = ct + 1
            If ct = 1 Or ct Mod 20 = 0 Then
                frmProgress.lblMsg2 = "On Dist Record: " & Format(ct, "##,###,##0") & _
                                   " Of: " & Format(Recs, "##,###,##0")
                frmProgress.Refresh
            End If
            
            ' has this employee been selected?
            frmItemDetail.rsEmp.Find "EmpID = " & PRDist.EmployeeID, 0, adSearchForward, 1
    
            If frmItemDetail.rsEmp.EOF = False Then
                If frmItemDetail.rsEmp!Selected = True Then
                    SQLString = "HistID = " & PRDist.HistID
                    trs.Find SQLString, 0, adSearchForward, 1
                    If trs.EOF Then
                        trs.AddNew
                        trs!EmpID = PRDist.EmployeeID
                        If Not PREmployee.GetByID(PRDist.EmployeeID) Then
                            MsgBox "Employee not found in Employee Master File!!!", vbExclamation, "Item Detail Report"
                            GoBack
                        End If
                        
                        trs!EmpNo = PREmployee.EmployeeNumber
                        trs!EmpName = Mid(PREmployee.LFName, 1, 30)
                        trs!HistID = PRDist.HistID
                        trs!ChkDate = PRDist.CheckDate
                        trs!PEDate = PRDist.PEDate
                        trs!Item1Amount = 0
                        trs!Item2Amount = 0
                        trs!Item3Amount = 0
                        trs!Item4Amount = 0
                        trs!Item5Amount = 0
                        trs!MaxAmt = 0
                                            
                        AmtRemaining = 0
                        ctr = ctr + 1
                        
                        If Not PRHist.GetByID(PRDist.HistID) Then
                            trs!Gross = 0
                            trs!YTDGross = trs!YTDGross + PRHist.Gross
                        Else
                            trs!Gross = PRHist.Gross
                            trs!YTDGross = trs!YTDGross + PRHist.Gross
                        End If
        
                        ItemCount = 0
                        frmItemDetail.rsItem.MoveFirst
                    
                    End If
                            
                End If
                
                ItemFlag = False
                
                If PRDist.EmployerItemID = Item1ID Then
                    If ItsHours1 = True Then
                        trs!Item1Amount = trs!Item1Amount + PRDist.Hours
                    Else
                        trs!Item1Amount = trs!Item1Amount + PRDist.Amount
                    End If
                    ' get the employee item for max amt
                    If PRItem.GetByID(PRDist.ItemID) Then
                        trs!MaxAmt = PRItem.MaxAmount
                    End If
                    ItemFlag = True
                End If
                If PRDist.EmployerItemID = Item2ID Then
                    If ItsHours2 = True Then
                        trs!Item2Amount = trs!Item2Amount + PRDist.Hours
                    Else
                        trs!Item2Amount = trs!Item2Amount + PRDist.Amount
                    End If
                    ItemFlag = True
                End If
                If PRDist.EmployerItemID = Item3ID Then
                    If ItsHours3 Then
                        trs!Item3Amount = trs!Item3Amount + PRDist.Hours
                    Else
                        trs!Item3Amount = trs!Item3Amount + PRDist.Amount
                    End If
                    ItemFlag = True
                End If
                If PRDist.EmployerItemID = Item4ID Then
                    If ItsHours4 Then
                        trs!Item4Amount = trs!Item4Amount + PRDist.Hours
                    Else
                        trs!Item4Amount = trs!Item4Amount + PRDist.Amount
                    End If
                    ItemFlag = True
                End If
                If PRDist.EmployerItemID = Item5ID Then
                    If ItsHours5 Then
                        trs!Item5Amount = trs!Item5Amount + PRDist.Hours
                    Else
                        trs!Item5Amount = trs!Item5Amount + PRDist.Amount
                    End If
                    ItemFlag = True
                End If

                If ItemFlag Then trs.Update
                
            End If
            If Not PRDist.GetNext Then Exit Do
        Loop
    
    End If
            
    ' =================== Deductions from PRItemHist =============================
            
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.BatchID = " & BatchNumbr & _
        " AND (ItemType = " & PREquate.ItemTypeDED & " OR ItemType = " & PREquate.ItemTypeOE & _
        " OR ItemType = " & PREquate.ItemTypeSDTax & ")"
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.CheckDate >= " & (StartDate) & " AND " & _
                                     " PRItemHist.CheckDate <= " & (EndDate) & _
                                     " AND (ItemType = " & PREquate.ItemTypeDED & _
                                     " OR ItemType = " & PREquate.ItemTypeOE & _
                                     " OR ItemType = " & PREquate.ItemTypeSDTax & ")"
            Msg1 = "CHECK DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
                                    
        Else
            SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.PEDate >= " & (StartDate) & " AND " & _
                                    " PRItemHist.PEDate <= " & (EndDate) & _
                                     " AND (ItemType = " & PREquate.ItemTypeDED & _
                                     " OR ItemType = " & PREquate.ItemTypeOE & _
                                     " OR ItemType = " & PREquate.ItemTypeSDTax & ")"
            Msg1 = "P/E DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If

    ct = 0

    If PRItemHist.GetBySQL(SQLString) Then
            
        Recs = PRItemHist.Records
            
        Do
            
            ct = ct + 1
            If ct = 1 Or ct Mod 20 = 0 Then
                frmProgress.lblMsg2 = "On Deduction Record: " & Format(ct, "##,###,##0") & _
                                   " Of: " & Format(Recs, "##,###,##0")
                frmProgress.Refresh
            End If
    
            ' has this employee been selected?
            frmItemDetail.rsEmp.Find "EmpID = " & PRItemHist.EmployeeID, 0, adSearchForward, 1

            If Not frmItemDetail.rsEmp.EOF And frmItemDetail.rsEmp!Selected = True Then
                SQLString = "HistID = " & PRItemHist.HistID

                trs.Find SQLString, 0, adSearchForward, 1
                If trs.EOF Then
                    trs.AddNew
                    trs!EmpID = PRItemHist.EmployeeID
                    If Not PREmployee.GetByID(PRItemHist.EmployeeID) Then
                        MsgBox "Employee not found in Employee Master File!!!", vbExclamation, "Item Detail Report"
                        GoBack
                    End If
                    trs!EmpNo = PREmployee.EmployeeNumber
                    trs!EmpName = Mid(PREmployee.LFName, 1, 30)
                    trs!HistID = PRItemHist.HistID
                    trs!ChkDate = PRItemHist.CheckDate
                    trs!PEDate = PRItemHist.PEDate
                    trs!Item1Amount = 0
                    trs!Item2Amount = 0
                    trs!Item3Amount = 0
                    trs!Item4Amount = 0
                    trs!Item5Amount = 0
                    trs!MaxAmt = 0
                    
                    AmtRemaining = 0
   
                    ctr = ctr + 1
                    If Not PRHist.GetByID(PRItemHist.HistID) Then
                        trs!Gross = 0
                        trs!YTDGross = trs!YTDGross + PRHist.Gross
                    Else
                        trs!Gross = PRHist.Gross
                        trs!YTDGross = trs!YTDGross + PRHist.Gross
                    End If
                    
                    ItemCount = 0
                    ' If ctr > 11 Then Exit Do
                    frmItemDetail.rsItem.MoveFirst
                    
                End If

                If PRItemHist.EmployerItemID = Item1ID Then
                    If ItsHours1 Then
                        trs!Item1Amount = trs!Item1Amount + PRDist.Hours
                    Else
                        trs!Item1Amount = trs!Item1Amount + PRItemHist.Amount
                    End If
                    ' get the employee item for max amt
                    If PRItem.GetByID(PRItemHist.ItemID) Then
                        trs!MaxAmt = PRItem.MaxAmount
                    End If
                End If
                If PRItemHist.EmployerItemID = Item2ID Then
                    If ItsHours2 Then
                        trs!Item2Amount = trs!Item2Amount + PRDist.Hours
                    Else
                        trs!Item2Amount = trs!Item2Amount + PRItemHist.Amount
                    End If
                End If
                
                If PRItemHist.EmployerItemID = Item3ID Then
                    If ItsHours3 Then
                        trs!Item3Amount = trs!Item3Amount + PRDist.Hours
                    Else
                        trs!Item3Amount = trs!Item3Amount + PRItemHist.Amount
                    End If
                End If
                If PRItemHist.EmployerItemID = Item4ID Then
                    If ItsHours4 Then
                        trs!Item4Amount = trs!Item4Amount + PRDist.Hours
                    Else
                        trs!Item4Amount = trs!Item4Amount + PRItemHist.Amount
                    End If
                End If
                If PRItemHist.EmployerItemID = Item5ID Then
                    If ItsHours5 Then
                        trs!Item5Amount = trs!Item5Amount + PRDist.Hours
                    Else
                        trs!Item5Amount = trs!Item5Amount + PRItemHist.Amount
                    End If
                End If

                trs.Update
                
            End If
            If Not PRItemHist.GetNext Then Exit Do
        Loop

    End If

    If trs.RecordCount = 0 Then
        MsgBox "No PR Data Found!", vbExclamation
        GoBack
    End If

    ''''''''''''''''''''      PRINT REPORT DETAIL     ''''''''''''''''''''''''''''''''''''''
    
    If frmItemDetail.optChkDate = True Then
        trs.Sort = "ChkDate, EmpNo"
    ElseIf frmItemDetail.optEmpNo = True Then
        trs.Sort = "EmpNo, ChkDate"
    Else
        trs.Sort = "EmpName, ChkDate"
    End If
    
    LastEmpID = 0
    
    If trs.RecordCount = 0 Then
        MsgBox "There are no Employees that fit your Criteria Selection", vbExclamation, "Item Detail Report"
        GoBack
    End If
    
    trs.MoveFirst

    LastEmpNo = 0
    LastChkDate = 0
    dtlCount = 0
    Do
        ' dtlCount = dtlCount + 1
        If Ln = 0 Or Ln > MaxLines - LineCount Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, Msg2, ""
            ItemDetailHeader
        End If
        If PREmployee.GetByID(trs!EmpID) Then
        End If
        
        SkipFlag = True
        ' Set skip flag to false (to process) if item amount <> 0
        If trs!Item1Amount <> 0 Then SkipFlag = False
        If trs!Item2Amount <> 0 Then SkipFlag = False
        If trs!Item3Amount <> 0 Then SkipFlag = False
        If trs!Item4Amount <> 0 Then SkipFlag = False
        If trs!Item5Amount <> 0 Then SkipFlag = False

        If SkipFlag = False Then
            If frmItemDetail.optChkDate Then
                ' Print SubTotal Line if Option Check Date is selected
                If LastChkDate <> 0 And LastChkDate <> trs!ChkDate Then
                    If dtlCount > 1 Then
                        PrintValue(1) = "     Check Date: ":            FormatString(1) = "a17"
                        PrintValue(2) = LastChkDate:                    FormatString(2) = "a46"
                                                
                        If frmItemDetail.chkNoGross = 0 Then
                            PrintValue(3) = SGross:                         FormatString(3) = "d14"
                        Else
                            PrintValue(3) = "":                             FormatString(3) = "a14"
                        End If
                        PrintValue(4) = SItem1:                         FormatString(4) = "d14"
                        If frmItemDetail.chkShowRemain Then
                            PrintValue(5) = SMax:                       FormatString(5) = "d14"
                            PrintValue(6) = SRemaining:                 FormatString(6) = "d14"
                            PrintValue(7) = " ":                        FormatString(7) = "~"
                        Else
                            PrintValue(5) = SItem2:                     FormatString(5) = "d14"
                            PrintValue(6) = SItem3:                     FormatString(6) = "d14"
                            PrintValue(7) = SItem4:                     FormatString(7) = "d14"
                            PrintValue(8) = SItem5:                     FormatString(8) = "d14"
                            PrintValue(9) = " ":                        FormatString(9) = "~"
                        End If
                        FormatPrint
                        Ln = Ln + 1
                    End If
                    Ln = Ln + 1
                    SGross = 0
                    SItem1 = 0
                    SItem2 = 0
                    SItem3 = 0
                    SItem4 = 0
                    SItem5 = 0
                    SMax = 0
                    SRemaining = 0
                    dtlCount = 0
                End If
            Else        ' works if by EE number or name
                ' ' Print SubTotal Line if Option Employee Number OR NAME is selected
                If LastEmpNo <> 0 And LastEmpNo <> trs!EmpNo Then
                    If dtlCount > 1 Then
                        PrintValue(1) = "     Employee: ":                          FormatString(1) = "a17"
                        PrintValue(2) = LastEmpNo & " - " & LastEmpName: FormatString(2) = "a46"
                        If frmItemDetail.chkNoGross = 0 Then
                            PrintValue(3) = SGross:                       FormatString(3) = "d14"
                        Else
                            PrintValue(3) = "":                           FormatString(3) = "a14"
                        End If
                        PrintValue(4) = SItem1:                       FormatString(4) = "d14"
                        If frmItemDetail.chkShowRemain Then
                            ' PrintValue(5) = SMax:                       FormatString(5) = "d14"
                            ' PrintValue(6) = SRemaining:                 FormatString(6) = "d14"
                            PrintValue(5) = " ":                        FormatString(5) = "~"
                        Else
                            PrintValue(5) = SItem2:                     FormatString(5) = "d14"
                            PrintValue(6) = SItem3:                     FormatString(6) = "d14"
                            PrintValue(7) = SItem4:                     FormatString(7) = "d14"
                            PrintValue(8) = SItem5:                     FormatString(8) = "d14"
                            PrintValue(9) = " ":                        FormatString(9) = "~"
                        End If
                        FormatPrint
                        Ln = Ln + 1
                    End If
                    
                    Ln = Ln + 1
                    SGross = 0
                    SItem1 = 0
                    SItem2 = 0
                    SItem3 = 0
                    SItem4 = 0
                    SItem5 = 0
                    SMax = 0
                    SRemaining = 0
                    dtlCount = 0
                        
                    ' update special grand total if reporting remaining
                    RmnMaxTL = RmnMaxTL + RmnMax
                    RmnRemainTL = RmnRemainTL + RmnRemain
                    
                End If
                    
            End If
    
            '  update Subtotal and Grand totals
            SGross = SGross + trs!Gross
            TGross = TGross + trs!Gross
            SItem1 = SItem1 + trs!Item1Amount
            TItem1 = TItem1 + trs!Item1Amount
            SItem2 = SItem2 + trs!Item2Amount
            TItem2 = TItem2 + trs!Item2Amount
            SItem3 = SItem3 + trs!Item3Amount
            TItem3 = TItem3 + trs!Item3Amount
            SItem4 = SItem4 + trs!Item4Amount
            TItem4 = TItem4 + trs!Item4Amount
            SItem5 = SItem5 + trs!Item5Amount
            TItem5 = TItem5 + trs!Item5Amount
            SMax = SMax + trs!MaxAmt
            TMax = TMax + trs!MaxAmt
            SRemaining = SRemaining + AmtRemaining
            TRemaining = TRemaining + AmtRemaining
            
            If frmItemDetail.chkTotalsOnly = 0 Then
                
                ' get the item
                PrintValue(1) = trs!EmpNo:                           FormatString(1) = "a7"
                PrintValue(2) = trs!EmpName:                         FormatString(2) = "a32"
                PrintValue(3) = Format(trs!PEDate, "mm/dd/yyyy"):    FormatString(3) = "a12"
                PrintValue(4) = Format(trs!ChkDate, "mm/dd/yyyy"):   FormatString(4) = "a12"
                
                If frmItemDetail.chkNoGross = 0 Then
                    PrintValue(5) = trs!Gross:                           FormatString(5) = "d14"
                Else
                    PrintValue(5) = "":                                  FormatString(5) = "a14"
                End If
                PrintValue(6) = trs!Item1Amount:                     FormatString(6) = "d14"
                If frmItemDetail.chkShowRemain Then
                    PrintValue(7) = trs!MaxAmt:                      FormatString(7) = "d14"
                    AmtRemaining = trs!MaxAmt - SItem1
                    PrintValue(8) = AmtRemaining:                    FormatString(8) = "d14"
                    PrintValue(9) = " ":                             FormatString(9) = "~"
                Else
                    PrintValue(7) = trs!Item2Amount:                 FormatString(7) = "d14"
                    PrintValue(8) = trs!Item3Amount:                 FormatString(8) = "d14"
                    PrintValue(9) = trs!Item4Amount:                 FormatString(9) = "d14"
                    PrintValue(10) = trs!Item5Amount:                FormatString(10) = "d14"
                    PrintValue(11) = " ":                            FormatString(11) = "~"
                End If
                                
                FormatPrint
                Ln = Ln + 1
            
            End If
            
            dtlCount = dtlCount + 1
            
            LastEmpNo = PREmployee.EmployeeNumber
            LastChkDate = trs!ChkDate
            LastEmpName = trs!EmpName
        
            RmnMax = trs!MaxAmt
            RmnRemain = AmtRemaining
        
        End If
        
        trs.MoveNext
        If trs.EOF Then Exit Do
    
    Loop
    
    ' Print SubTotal Line
    If dtlCount > 1 Then
        If frmItemDetail.optChkDate Then
            PrintValue(1) = "     Check Date: ":                FormatString(1) = "a17"
            PrintValue(2) = LastChkDate:                        FormatString(2) = "a46"
            If frmItemDetail.chkNoGross = 0 Then
                PrintValue(3) = SGross:                             FormatString(3) = "d14"
            Else
                PrintValue(3) = "":                                 FormatString(3) = "a14"
            End If
            
            PrintValue(4) = SItem1:                             FormatString(4) = "d14"
            If frmItemDetail.chkShowRemain Then
    '            PrintValue(5) = SMax:                       FormatString(5) = "d14"
    '            PrintValue(6) = SRemaining:                 FormatString(6) = "d14"
                PrintValue(5) = " ":                        FormatString(5) = "~"
            Else
                PrintValue(5) = SItem2:                     FormatString(5) = "d14"
                PrintValue(6) = SItem3:                     FormatString(6) = "d14"
                PrintValue(7) = SItem4:                     FormatString(7) = "d14"
                PrintValue(8) = SItem5:                     FormatString(8) = "d14"
                PrintValue(9) = " ":                        FormatString(9) = "~"
            End If
            FormatPrint
            Ln = Ln + 2
        Else
            PrintValue(1) = "     Employee: ":                  FormatString(1) = "a17"
            PrintValue(2) = LastEmpNo & " - " & LastEmpName:    FormatString(2) = "a46"
            
            If frmItemDetail.chkNoGross = 0 Then
                PrintValue(3) = SGross:                             FormatString(3) = "d14"
            Else
                PrintValue(3) = "":                             FormatString(3) = "a14"
            End If
            
            PrintValue(4) = SItem1:                             FormatString(4) = "d14"
            If frmItemDetail.chkShowRemain Then
    '            PrintValue(5) = SMax:                           FormatString(5) = "d14"
    '            PrintValue(6) = SRemaining:                     FormatString(6) = "d14"
                PrintValue(5) = " ":                            FormatString(5) = "~"
            Else
                PrintValue(5) = SItem2:                         FormatString(5) = "d14"
                PrintValue(6) = SItem3:                         FormatString(6) = "d14"
                PrintValue(7) = SItem4:                         FormatString(7) = "d14"
                PrintValue(8) = SItem5:                         FormatString(8) = "d14"
                PrintValue(9) = " ":                            FormatString(9) = "~"
            End If
            FormatPrint
            Ln = Ln + 2
                    
            ' update special grand total if reporting remaining
            RmnMaxTL = RmnMaxTL + RmnMax
            RmnRemainTL = RmnRemainTL + RmnRemain
        
        End If
    Else
        Ln = Ln + 1
    End If
    
    ' Print Grand Total Line
    If frmItemDetail.chkShowRemain = 1 Then
        PrintValue(1) = "GRAND TOTAL ":                         FormatString(1) = "a63"
        
        If frmItemDetail.chkNoGross = 0 Then
            PrintValue(2) = TGross:                                 FormatString(2) = "d14"
        Else
            PrintValue(2) = "":                                 FormatString(2) = "a14"
        End If
        PrintValue(3) = TItem1:                                 FormatString(3) = "d14"
        PrintValue(4) = RmnMaxTL:                               FormatString(4) = "d14"
        PrintValue(5) = RmnRemainTL:                            FormatString(5) = "d14"
        PrintValue(6) = " ":                                    FormatString(6) = "~"
        FormatPrint
    Else
        PrintValue(1) = "GRAND TOTAL ":                         FormatString(1) = "a63"
        
        If frmItemDetail.chkNoGross = 0 Then
            PrintValue(2) = TGross:                                 FormatString(2) = "d14"
        Else
            PrintValue(2) = "":                                 FormatString(2) = "a14"
        End If
        
        PrintValue(3) = TItem1:                                 FormatString(3) = "d14"
        PrintValue(4) = TItem2:                                 FormatString(4) = "d14"
        PrintValue(5) = TItem3:                                 FormatString(5) = "d14"
        PrintValue(6) = TItem4:                                 FormatString(6) = "d14"
        PrintValue(7) = TItem5:                                 FormatString(7) = "d14"
        PrintValue(8) = " ":                                    FormatString(8) = "~"
        FormatPrint
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

    
Public Sub ItemDetailHeader()
Dim ColNumber As Long
    
    PrintValue(1) = "EMP NO.":                      FormatString(1) = "a7"
    PrintValue(2) = "EMPLOYEE NAME":                FormatString(2) = "a32"
    PrintValue(3) = "P/E DATE":                     FormatString(3) = "a12"
    PrintValue(4) = "CHK DATE":                     FormatString(4) = "a12"
    
    If frmItemDetail.chkNoGross = 0 Then
        PrintValue(5) = "GROSS":                        FormatString(5) = "r13"
    Else
        PrintValue(5) = "":                             FormatString(5) = "r13"
    End If
    
    ColNumber = 6
         
    frmItemDetail.rsItem.MoveFirst
    Do
 
        If frmItemDetail.rsItem!Selected = True Then
            ' get the item
            If PRItem.GetByID(frmItemDetail.rsItem!ItemID) Then
                
                ' Check to see if the item is HOURS
                If frmItemDetail.rsItem!IsItHours Then
                    PrintValue(ColNumber) = PRItem.Abbreviation & " HR"        '#################    HOURS   #################
                Else
                    PrintValue(ColNumber) = PRItem.Abbreviation
                End If
            
            ElseIf ColNumber = 11 Then
                Exit Do
            End If
            
            FormatString(ColNumber) = "r14"
            ColNumber = ColNumber + 1
            
        End If
        frmItemDetail.rsItem.MoveNext
        
    Loop Until frmItemDetail.rsItem.EOF
    
    ' Show Maximum and Remaining Dollars if box is checked
    If frmItemDetail.chkShowRemain Then
        PrintValue(ColNumber) = "MAXIMUM":                  FormatString(ColNumber) = "r14"
        ColNumber = ColNumber + 1
        PrintValue(ColNumber) = "REMAINING":                FormatString(ColNumber) = "r14"
        ColNumber = ColNumber + 1
    End If
    
    PrintValue(ColNumber) = " ":                    FormatString(ColNumber) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(147, "="):               FormatString(1) = "a147"
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
End Sub

Public Sub AnnivRemain(ByVal DateRecall As Byte, _
                       ByRef rsEmp As ADODB.Recordset, _
                       ByRef rsItem As ADODB.Recordset)

Dim RmnMax, RmnRemain, TotalHours As Currency
Dim GrandMax, GrandRemain, GrandHours As Currency

Dim boo As Boolean
Dim ItmID As Long
    
Dim D2, HireDate, StartDate, EndDate As Date
Dim YearCount As Byte
    
    RmnMax = 0
    RmnRemain = 0
    
    GrandMax = 0
    GrandRemain = 0
    GrandHours = 0
    
    PrtInit ("Port")
    LandSw = 1
    SetFont 8, Equate.Portrait
    Columns = Columns - 5
    
    ReportTitle = "HOURS REMAINING REPORT - BASED ON ANNIVERSARY DATE'"
    
    ' filter for what was selected
    rsItem.Filter = "Selected = True"
    rsItem.MoveFirst
    ItmID = rsItem!ItemID
    If PRItem.GetByID(ItmID) Then
        ItemTitle = PRItem.Abbreviation
    End If
    ItemTitle = Trim(ItemTitle) & " HR"
    
    rsEmp.Filter = "Selected = True"
    rsEmp.Sort = "EmpNo"
    rsEmp.MoveFirst
    
    frmProgress.Caption = "Hours Remaining - Anniv Rollover Report"
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.Show
    
    
    Do
        
        boo = PREmployee.GetByID(rsEmp!EmpID)
        
        If Ln = 0 Or Ln > MaxLines - 1 Then AnnivHeader
            
        ' is the hire date filled in ?
        If PREmployee.DateHired = 0 Then
            SQLString = "SELECT * FROM PRDist WHERE EmployeeID = " & rsEmp!EmpID & _
                        " AND EmployerItemID = " & ItmID
            If PRDist.GetBySQL(SQLString) = True Then
                ' print message on report if any data exists for the item for the employee
                PrintValue(1) = "   **** Hire Date Not Entered: " & PREmployee.EmployeeNumber & _
                               " " & PREmployee.LFName
                FormatString(1) = "a70"
                PrintValue(2) = " ":        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 2
            End If
            GoTo NxtrsEMP
        End If
    
        ' get the Employee Item
        SQLString = "SELECT * FROM PRItem WHERE EmployerItemID = " & rsItem!ItemID & _
                    " AND EmployeeID = " & rsEmp!EmpID
        If PRItem.GetBySQL(SQLString) = False Then GoTo NxtrsEMP
        
        RmnMax = PRItem.MaxAmount
        RmnRemain = PRItem.MaxAmount
        TotalHours = 0
    
        ' setup the dates
        ' *** option for years back ***
        YearCount = 0
        EndDate = DateSerial(Year(Now()) - YearCount, Month(Now()), Day(Now()))
        
        StartDate = PREmployee.DateHired
        StartString = "Hired: "
        If frmItemDetail.chkRecall And PREmployee.DateLastRecall <> 0 Then
            StartDate = PREmployee.DateLastRecall
            StartString = "Recalled: "
        End If
        HireDate = StartDate
        D2 = StartDate
        
        Do
            D2 = DateSerial(Year(D2) + 1, Month(D2), Day(D2))
            If D2 > EndDate Then Exit Do
            StartDate = D2
        Loop
        
        SQLString = "SELECT * FROM PRDist WHERE EmployeeID = " & rsEmp!EmpID & _
                    " AND PEDate >= " & CLng(StartDate) & _
                    " AND PEDate <= " & CLng(EndDate) & _
                    " AND EmployerItemID = " & rsItem!ItemID & _
                    " ORDER BY PEDate"
        
        If PRDist.GetBySQL(SQLString) = True Then
        
            frmProgress.lblMsg2 = Trim(PREmployee.LFName) & " " & PRDist.Records & " records to process"
            frmProgress.Refresh
        
            Do
                
                RmnRemain = RmnRemain - PRDist.Hours
                TotalHours = TotalHours + PRDist.Hours
                
                PrintValue(1) = CStr(PREmployee.EmployeeNumber):            FormatString(1) = "a9"
                PrintValue(2) = PREmployee.LFName:                          FormatString(2) = "a30"
                PrintValue(3) = Format(PRDist.PEDate, " mm/dd/yy "):        FormatString(3) = "a17"
                PrintValue(4) = PRDist.Hours:                               FormatString(4) = "d12"
                PrintValue(5) = PRItem.MaxAmount:                           FormatString(5) = "d12"
                PrintValue(6) = RmnRemain:                                  FormatString(6) = "d12"
                PrintValue(7) = " ":                                        FormatString(7) = "~"
                FormatPrint
                Ln = Ln + 1
                
                If PRDist.GetNext = False Then Exit Do
                
            Loop
        
        End If
        
        ' employee subtotal
        PrintValue(1) = "     Employee: ":              FormatString(1) = "a15"
        PrintValue(2) = PREmployee.EmployeeNumber & " - " & PREmployee.LFName
        FormatString(2) = "a41"
        PrintValue(3) = TotalHours:                     FormatString(3) = "d12"
        PrintValue(4) = PRItem.MaxAmount:               FormatString(4) = "d12"
        PrintValue(5) = RmnRemain:                      FormatString(5) = "d12"
        PrintValue(6) = " ":                            FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = Space(5) & StartString:             FormatString(1) = "a20"
        PrintValue(2) = Format(HireDate, " mm/dd/yy "):     FormatString(2) = "a10"
        PrintValue(3) = "Date Range: ":                     FormatString(3) = "a12"
        PrintValue(4) = Format(StartDate, " mm/dd/yy "):    FormatString(4) = "a10"
        PrintValue(5) = Format(EndDate, " mm/dd/yy "):      FormatString(5) = "a10"
        PrintValue(6) = " ":                                FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 2
        
        GrandHours = GrandHours + TotalHours
        GrandMax = GrandMax + PRItem.MaxAmount
        GrandRemain = GrandRemain + RmnRemain
                
NxtrsEMP:
        rsEmp.MoveNext
    
    Loop Until rsEmp.EOF

    ' grand totals
    If Ln > MaxLines - 3 Then
        AnnivHeader
    End If
    Ln = Ln + 1
    PrintValue(1) = "     TOTALS: ":                FormatString(1) = "a56"
    PrintValue(2) = GrandHours:                     FormatString(2) = "d12"
    PrintValue(3) = GrandMax:                       FormatString(3) = "d12"
    PrintValue(4) = GrandRemain:                    FormatString(4) = "d12"
    PrintValue(5) = " ":                            FormatString(5) = "~"
    FormatPrint

    frmProgress.Hide

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub AnnivHeader()
            
    If Ln <> 0 Then FormFeed
    PageHeader ReportTitle
    Ln = Ln + 1
    
    PrintValue(1) = "Employee Number / Name":           FormatString(1) = "a39"
    PrintValue(2) = "PdEnd Date":                       FormatString(2) = "a10"
    PrintValue(3) = ItemTitle & " ":                    FormatString(3) = "r19"
    PrintValue(4) = "MAXIMUM ":                         FormatString(4) = "r12"
    PrintValue(5) = "REMAINING ":                       FormatString(5) = "r12"
    PrintValue(6) = " ":                                FormatString(6) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(Columns - 5, "="):           FormatString(1) = "a" & Columns - 5
    PrintValue(2) = " ":                                FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub


Public Sub EEList(ByVal ReportType As String)
Dim ReportTitle As String
Dim FDept As String
Dim LabelColumns As Integer
Dim ActInact As String
    
    SetEquates
    frmLists.Hide
    Msg2 = "Date: " & Format(Now, "mm/dd/yyyy")
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    ' page set up based on the report type
    Select Case ReportType

        Case "NumberName"
            PrtInit ("Port")
            ReportTitle = "EMPLOYEE NUMBER/NAME LIST"
            SetFont 10, Equate.Portrait
        Case "DetailList"
            PrtInit ("Land")
            LandSw = 1
            ReportTitle = "EMPLOYEE DETAIL LISTING"
            SetFont 8, Equate.LandScape
        Case "EmployeeRateList"
            PrtInit ("Port")
            ReportTitle = "EMPLOYEE RATE LISTING"
            SetFont 10, Equate.Portrait
        Case "SSNFormat"
            PrtInit ("Port")
            ReportTitle = "SOCIAL SECURITY NUMBER FORMAT"
            SetFont 10, Equate.Portrait
        Case "RateTaxList"
            PrtInit ("Land")
            LandSw = 1
            ReportTitle = "RATE TAX LISTING"
            SetFont 8, Equate.LandScape
        Case "TimeCardLabels"
            PrtInit ("Port")
            ReportTitle = "labels "
            SetFont 8, Equate.Portrait
        Case "MailingLabels"
            PrtInit ("Port")
            ReportTitle = "labels "
            SetFont 8, Equate.Portrait
    End Select
                                                                
    ' set up SQL statement based upon order requested

    SQLString = "Select * from PREmployee"

    If frmLists.optNumber Then
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NO."
        End If
        SQLString = Trim(SQLString) & " ORDER BY EmployeeNumber"
    ElseIf frmLists.optName Then
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NAME"
        End If
        SQLString = Trim(SQLString) & " ORDER BY LastName, FirstName"
    ElseIf frmLists.optZipCode Then
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE ZIP CODE"
        End If
        SQLString = Trim(SQLString) & " ORDER BY ZipCode"
    End If
    
    If frmLists.optSal = True Then
        Msg1 = "All Salaried Employees"
    ElseIf frmLists.optHrly = True Then
        Msg1 = "All Hourly Employees"
    End If

    If frmLists.optActive = True Then
        If Trim(Msg1) = "" Then
            Msg1 = "All Active Employees"
        Else
            Msg1 = Trim(Msg1) & " & Active Employees"
        End If
    ElseIf frmLists.optInactive = True Then
        If Trim(Msg1) = "" Then
            Msg1 = "All InActive Employees"
        Else
            Msg1 = Trim(Msg1) & " & InActive Employees"
        End If
    End If
        
    If frmLists.optAllA = True And frmLists.optAllS = True Then
        Msg1 = "All Employees"
    End If
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Found !!!", vbExclamation, "Employee Lists and Labels"
        Exit Sub
    End If
    Do
        If Ln = 0 Or Ln > MaxLines Then
         
            If Ln Then FormFeed
            If ReportTitle = "labels " Then
            Else
                PageHeader ReportTitle, Msg1, Msg2, ""
            End If
            Ln = Ln + 2
                     
            Select Case ReportType
                
                Case "NumberName"
                    PrintValue(1) = "EMPLOYEE# ":                   FormatString(1) = "a10"
                    PrintValue(2) = " DEPT# ":                      FormatString(2) = "a7"
                    PrintValue(3) = "DEPT NAME:":                   FormatString(3) = "a30"
                    PrintValue(4) = "EMPLOYEE NAME":                FormatString(4) = "a40"
                    PrintValue(5) = " ":                            FormatString(5) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = String(90, "-"):                FormatString(1) = "a90"
                    PrintValue(2) = " ":                            FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                Case "DetailList"
                    SetFont 8, Equate.LandScape
                    PrintValue(1) = "EMPLOYEE NO.":                 FormatString(1) = "a11"  ' First Heading
                    PrintValue(2) = " ":                            FormatString(2) = "a2"
                    PrintValue(3) = "DEPARTMENT":                   FormatString(3) = "a10"
                    PrintValue(4) = " ":                            FormatString(4) = "a9"
                    PrintValue(5) = "EMPLOYEE NAME":                FormatString(5) = "a35"
                    PrintValue(6) = " ":                            FormatString(6) = "a1"
                    PrintValue(7) = "ADDRESS":                      FormatString(7) = "a30"
                    PrintValue(8) = " ":                            FormatString(8) = "a4"
                    PrintValue(9) = "CITY":                         FormatString(9) = "a20"
                    PrintValue(10) = " ":                           FormatString(10) = "a1"
                    PrintValue(11) = "ST":                          FormatString(11) = "a2"
                    PrintValue(12) = " ":                           FormatString(12) = "a2"
                    PrintValue(13) = "ZIP":                         FormatString(13) = "a6"
                                                                                
                 '*** Print SS Number?
                    If frmLists.chkSSN Then
                        PrintValue(14) = "SSN ":              FormatString(14) = "a9"
                        PrintValue(15) = " ":                       FormatString(15) = "~"
                    Else
                        PrintValue(14) = " ":                       FormatString(14) = "~"
                    End If
         
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = " ":                            FormatString(1) = "a8"  ' Print 2nd heading
                    PrintValue(2) = "DATE LAST":                    FormatString(2) = "a9"  ' PAID
                    PrintValue(3) = " ":                            FormatString(3) = "a3"
                    PrintValue(4) = "DATE":                         FormatString(4) = "a4"
                    PrintValue(5) = " ":                            FormatString(5) = "a8"
                    PrintValue(6) = "DATE LAST":                    FormatString(6) = "a9"  ' RAISE
                    PrintValue(7) = " ":                            FormatString(7) = "a3"
                    PrintValue(8) = "DATE LAST":                    FormatString(8) = "a9"  ' REVIEW
                    PrintValue(9) = " ":                            FormatString(9) = "a3"
                    PrintValue(10) = "DATE LAST":                   FormatString(10) = "a9" ' LAYOFF
                    PrintValue(11) = " ":                           FormatString(11) = "a3"
                    PrintValue(12) = "DATE LAST":                   FormatString(12) = "a9" ' RECALL
                    PrintValue(13) = " ":                           FormatString(13) = "a2"
                    PrintValue(14) = "DATE":                        FormatString(14) = "a4" ' TERMINATED"
                    PrintValue(15) = " ":                           FormatString(15) = "a6"
                    PrintValue(16) = "TERM":                        FormatString(16) = "a7" ' REASON
                    PrintValue(17) = " ":                           FormatString(17) = "a1"
                    PrintValue(18) = "DATE OF":                     FormatString(18) = "a7" ' BIRTH
                    PrintValue(19) = " ":                           FormatString(19) = "a3"
                    PrintValue(20) = "RACE":                        FormatString(20) = "a4" ' CODE
                    PrintValue(21) = " ":                           FormatString(21) = "a3"
                    PrintValue(22) = "MARITAL":                     FormatString(22) = "a7" ' STATUS
                    PrintValue(23) = " ":                           FormatString(23) = "a3"
                    PrintValue(24) = "EDU":                         FormatString(24) = "a3" ' LEVEL
                    PrintValue(25) = " ":                           FormatString(25) = "a4"
                    PrintValue(26) = "SHIFT":                       FormatString(26) = "a5" ' CODE
                    PrintValue(27) = " ":                           FormatString(27) = "~"
                    FormatPrint
                    Ln = Ln + 1
                                        
                    PrintValue(1) = " ":                            FormatString(1) = "a10" '  Print 3rd Heading
                    PrintValue(2) = "PAID":                         FormatString(2) = "a4"
                    PrintValue(3) = " ":                            FormatString(3) = "a6"
                    PrintValue(4) = "HIRED":                        FormatString(4) = "a5"
                    PrintValue(5) = " ":                            FormatString(5) = "a9"
                    PrintValue(6) = "RAISE":                        FormatString(6) = "a5"
                    PrintValue(7) = " ":                            FormatString(7) = "a6"
                    PrintValue(8) = "REVIEW":                       FormatString(8) = "a6"
                    PrintValue(9) = " ":                            FormatString(9) = "a6"
                    PrintValue(10) = "LAYOFF":                      FormatString(10) = "a6"
                    PrintValue(11) = " ":                           FormatString(11) = "a6"
                    PrintValue(12) = "RECALL":                      FormatString(12) = "a6"
                    PrintValue(13) = " ":                           FormatString(13) = "a4"
                    PrintValue(14) = "TERM":                        FormatString(14) = "a9"
                    PrintValue(15) = " ":                           FormatString(15) = "a0"
                    PrintValue(16) = "REASON":                      FormatString(16) = "a9"
                    PrintValue(17) = " ":                           FormatString(17) = "a1"
                    PrintValue(18) = "BIRTH":                       FormatString(18) = "a5"
                    PrintValue(19) = " ":                           FormatString(19) = "a4"
                    PrintValue(20) = "CODE":                        FormatString(20) = "a4"
                    PrintValue(21) = " ":                           FormatString(21) = "a3"
                    PrintValue(22) = "STATUS":                      FormatString(22) = "a6"
                    PrintValue(23) = " ":                           FormatString(23) = "a3"
                    PrintValue(24) = "LEVEL":                       FormatString(24) = "a5"
                    PrintValue(25) = " ":                           FormatString(25) = "a4"
                    PrintValue(26) = "CODE":                        FormatString(26) = "a4"
                    PrintValue(27) = " ":                           FormatString(27) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = String(144, "="):               FormatString(1) = "a144"
                    PrintValue(2) = " ":                            FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                Case "EmployeeRateList"
                    PrintValue(1) = " ":                            FormatString(1) = "a80"
                    PrintValue(2) = "HOURLY/":                      FormatString(2) = "a7"
                    PrintValue(3) = " ":                            FormatString(3) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "EMPL #":                       FormatString(1) = "a9"
                    PrintValue(2) = "EMPLOYEE NAME":                FormatString(2) = "a31"
                    PrintValue(3) = "DEPARTMENT":                   FormatString(3) = "a30"
                    PrintValue(4) = "RATE":                         FormatString(4) = "a10"
                    PrintValue(5) = "SALARY":                       FormatString(5) = "a6"
                    PrintValue(6) = " ":                            FormatString(6) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = String(91, "="):                FormatString(1) = "a91"
                    PrintValue(2) = " ":                            FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                Case "SSNFormat"
                    PrintValue(1) = "EMPLOYEE NO.":                 FormatString(1) = "a12"
                    PrintValue(2) = " ":                            FormatString(2) = "a3"
                    PrintValue(3) = "EMPLOYEE NAME":                FormatString(3) = "a40"
                    PrintValue(4) = " ":                            FormatString(4) = "a3"
                    PrintValue(5) = "SS NUMBER ":                   FormatString(5) = "a11"
                    PrintValue(6) = " ":                            FormatString(6) = "a3"
                    PrintValue(7) = "BIRTH DATE":                   FormatString(7) = "a10"
                    PrintValue(8) = " ":                            FormatString(8) = "a3"
                    PrintValue(9) = "GENDER ":                      FormatString(9) = "a6"
                    PrintValue(10) = " ":                           FormatString(10) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = String(91, "="):                FormatString(1) = "a91"
                    PrintValue(2) = " ":                            FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1

               Case "RateTaxList"
                    SetFont 8, Equate.LandScape
                    PrintValue(1) = " ":                            FormatString(1) = "a61" ' Header 1
                    PrintValue(2) = "NO":                           FormatString(2) = "a4"  ' SS Tax
                    PrintValue(3) = "NO":                           FormatString(3) = "a4"  ' MED Tax
                    PrintValue(4) = "NO":                           FormatString(4) = "a4"  ' FED Tax
                    PrintValue(5) = "NO":                           FormatString(5) = "a4"  ' ST Tax
                    PrintValue(6) = "NO":                           FormatString(6) = "a5"  ' Cty Tax
                    PrintValue(7) = "NO":                           FormatString(7) = "a7"  ' FED Unemp
                    PrintValue(8) = "NO":                           FormatString(8) = "a4"  ' State Unemp
                    PrintValue(9) = " ":                            FormatString(9) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = " ":                            FormatString(1) = "a52" ' Header 2
                    PrintValue(2) = "HOURLY/":                      FormatString(2) = "a8"
                    PrintValue(3) = " ":                            FormatString(3) = "a1"
                    PrintValue(4) = "SS":                           FormatString(4) = "a4"  ' TAX
                    PrintValue(5) = "MED":                          FormatString(5) = "a4"  ' TAX
                    PrintValue(6) = "FED":                          FormatString(6) = "a4"  ' TAX
                    PrintValue(7) = "ST":                           FormatString(7) = "a4"  ' TAX
                    PrintValue(8) = "CTY":                          FormatString(8) = "a5"  ' TAX
                    PrintValue(9) = "FED":                          FormatString(9) = "a5"  ' UNEMP
                    PrintValue(10) = "STATE":                       FormatString(10) = "a6" ' UNEMP
                    PrintValue(11) = "FWT":                         FormatString(11) = "a4" ' MAR
                    PrintValue(12) = "FWT":                         FormatString(12) = "a8" ' EXMP
                    PrintValue(13) = "FWT":                         FormatString(13) = "a5" ' PCT
                    PrintValue(14) = " ":                           FormatString(14) = "a3"
                    PrintValue(15) = "FWT":                         FormatString(15) = "a7" ' XAMT
                    PrintValue(16) = "SWT":                         FormatString(16) = "a4" ' MAR
                    PrintValue(17) = "SWT":                         FormatString(17) = "a8" ' EXMP
                    PrintValue(18) = "SWT":                         FormatString(18) = "a4" ' PCT
                    PrintValue(19) = " ":                           FormatString(19) = "a5"
                    PrintValue(20) = "SWT":                         FormatString(20) = "a4" ' XAMT
                    PrintValue(21) = " ":                           FormatString(21) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "EMP #":                        FormatString(1) = "a5"  ' Header 3
                    PrintValue(2) = " ":                            FormatString(2) = "a2"
                    PrintValue(3) = "NAME":                         FormatString(3) = "a18"
                    PrintValue(4) = " ":                            FormatString(4) = "a3"
                    PrintValue(5) = "DEPARTMENT":                   FormatString(5) = "a10"
                    PrintValue(6) = " ":                            FormatString(6) = "a9"
                    PrintValue(7) = "RATE":                         FormatString(7) = "a5"
                    PrintValue(8) = " ":                            FormatString(8) = "a0"
                    PrintValue(9) = "SALARY":                       FormatString(9) = "a8"
                    PrintValue(10) = " ":                           FormatString(10) = "a1"
                    PrintValue(11) = "TAX":                         FormatString(11) = "a4" ' SS
                    PrintValue(12) = "TAX":                         FormatString(12) = "a4" ' MED
                    PrintValue(13) = "TAX":                         FormatString(13) = "a4" ' FED
                    PrintValue(14) = "TAX":                         FormatString(14) = "a4" ' ST
                    PrintValue(15) = "TAX":                         FormatString(15) = "a4" ' CTY
                    PrintValue(16) = "UNEMP":                       FormatString(16) = "a6" ' FED
                    PrintValue(17) = "UNEMP":                       FormatString(17) = "a6" ' ST
                    PrintValue(18) = "MAR":                         FormatString(18) = "a4" ' FWT
                    PrintValue(19) = "EXMP":                        FormatString(19) = "a8" ' FWT
                    PrintValue(20) = "PCT":                         FormatString(20) = "a5" ' FWT
                    PrintValue(21) = " ":                           FormatString(21) = "a3"
                    PrintValue(22) = "XAMT":                        FormatString(22) = "a7" ' FWT
                    PrintValue(23) = "MAR":                         FormatString(23) = "a4" ' SWT
                    PrintValue(24) = "EXMP":                        FormatString(24) = "a8" ' SWT
                    PrintValue(25) = "PCT":                         FormatString(25) = "a4" ' SWT
                    PrintValue(26) = " ":                           FormatString(26) = "a4"
                    PrintValue(27) = "XAMT":                        FormatString(27) = "a5" ' SWT
                    PrintValue(28) = " ":                           FormatString(28) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = String(130, "="):               FormatString(1) = "a130"
                    PrintValue(2) = " ":                            FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
               
            End Select
         
         End If

        ' **** Active/Inactive filter
         If frmLists.optActive = True Then          ' All Active Employees
            If PREmployee.Inactive = 1 Then GoTo Cycle1
         ElseIf frmLists.optInactive = True Then
            If PREmployee.Inactive = 0 Then GoTo Cycle1
         End If
    
        ' **** Salaried/Hourly filter
         If frmLists.optSal = True Then
            If PREmployee.Salaried = 0 Then GoTo Cycle1
         ElseIf frmLists.optHrly = True Then
            If PREmployee.Salaried = 1 Then GoTo Cycle1
         End If
 
        If PREmployee.Inactive = True Then
            ActInact = "Inactive: Y"
        Else
            ActInact = "Inactive: N"
        End If

        ' **** department filter
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            PRDepartment.Name = "Not Assigned"
            PRDepartment.DepartmentNumber = 0
        End If
        FDept = PRDepartment.DepartmentNumber

'=============================================================================================
'============================      PRINT REPORT DETAIL     ===================================
'=============================================================================================
        frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
        frmProgress.Show
        Select Case ReportType
            Case "NumberName"
    
                PrintValue(1) = PREmployee.EmployeeNumber:          FormatString(1) = "n10"
                PrintValue(2) = PRDepartment.DepartmentNumber:      FormatString(2) = "n7"
                PrintValue(3) = PRDepartment.Name:                  FormatString(3) = "a30"
                PrintValue(4) = PREmployee.LFName:                  FormatString(4) = "a40"
                PrintValue(5) = " ":                                FormatString(5) = "~"
                FormatPrint
                Ln = Ln + 1
                
            Case "DetailList"
                MaxLines = 47
                
                PrintValue(1) = PREmployee.EmployeeNumber:          FormatString(1) = "a9"
                PrintValue(2) = " ":                                FormatString(2) = "a4"
                PrintValue(3) = RTrim(PRDepartment.DepartmentNumber):   FormatString(3) = "n2"
                PrintValue(4) = " - ":                              FormatString(4) = "a4"
                PrintValue(5) = PRDepartment.Name:                  FormatString(5) = "a8"
                PrintValue(6) = " ":                                FormatString(6) = "a5"
                PrintValue(7) = PREmployee.LFName:                  FormatString(7) = "a35"
                PrintValue(8) = " ":                                FormatString(8) = "a1"
                PrintValue(9) = RTrim(PREmployee.Address1):         FormatString(9) = "a30"
                PrintValue(10) = " ":                               FormatString(10) = "a4"
                PrintValue(11) = RTrim(PREmployee.City):            FormatString(11) = "a20"
                PrintValue(12) = " ":                               FormatString(12) = "a1"
                PrintValue(13) = PREmployee.State:                  FormatString(13) = "a2"
                PrintValue(14) = " ":                               FormatString(14) = "a2"
                PrintValue(15) = PREmployee.ZipCode:                FormatString(15) = "a5"

                 '*** Print SS Number?
                If frmLists.chkSSN Then
                    PrintValue(16) = " ":                           FormatString(16) = "a1"
                    PrintValue(17) = PREmployee.SSString:           FormatString(17) = "a11"
                    PrintValue(18) = " ":                           FormatString(18) = "~"
                Else
                    PrintValue(16) = " ":                           FormatString(16) = "~"
                End If
                
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = "Gen Info: ":                       FormatString(1) = "a11" '  Print 2nd Detail Line
                PrintValue(2) = ActInact:                           FormatString(2) = "a11"
                PrintValue(3) = " ":                                FormatString(3) = "a2"
                If PREmployee.Salaried = 1 Then
                    PrintValue(4) = "Salaried: Y":                  FormatString(4) = "a12"
                    PrintValue(5) = " ":                            FormatString(5) = "a2"
                    PrintValue(6) = "Salary Amt: ":                 FormatString(6) = "a12"
                    PrintValue(7) = PREmployee.SalaryAmount:        FormatString(7) = "d10"
                Else
                    PrintValue(4) = "Salaried: N":                  FormatString(4) = "a12"
                    PrintValue(5) = " ":                            FormatString(5) = "a2"
                    PrintValue(6) = "Hourly Amt: ":                 FormatString(6) = "a12"
                    PrintValue(7) = PREmployee.HourlyAmount:        FormatString(7) = "d10"
                End If
                
                PrintValue(8) = " ":                                FormatString(8) = "a2"
                PrintValue(9) = "Pays Per Yr: ":                    FormatString(9) = "a13"
                PrintValue(10) = PREmployee.PaysPerYear:            FormatString(10) = "n2"
                PrintValue(11) = " ":                               FormatString(11) = "a2"
                                
                If PREmployee.DefaultCityID > 0 Then
                    If PRCity.GetBySQL("Select * from PRCity where PRCity.CityID = " & PREmployee.DefaultCityID) Then
                        PrintValue(12) = "City: ":                  FormatString(12) = "a6"
                        PrintValue(13) = Trim(PRCity.CityName):     FormatString(13) = "a15"
                        PrintValue(14) = " ":                       FormatString(14) = "a2"
                        PrintValue(15) = "Rate: ":                  FormatString(15) = "a6"
                        PrintValue(16) = PRCity.CityRate:           FormatString(16) = "d6"
                        PrintValue(17) = " ":                       FormatString(17) = "a2"
                        If PRCity.StateID > 0 Then
                            PrintValue(18) = "State: ":             FormatString(18) = "a7"
                            If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCity.StateID) Then
                                PrintValue(19) = PRState.StateAbbrev:   FormatString(19) = "a2"
                                If PREmployee.WkcCat <> 0 Then
                                    PrintValue(20) = "  W/C: " & PREmployee.WkcCat:  FormatString(20) = "a6"
                                End If
                            Else
                                If PREmployee.WkcCat <> 0 Then
                                    PrintValue(19) = "W/C: " & PREmployee.WkcCat:  FormatString(19) = "a6"
                                End If
                            End If
                        Else
                            If PREmployee.WkcCat <> 0 Then
                                PrintValue(18) = "W/C: " & PREmployee.WkcCat:  FormatString(18) = "a6"
                            End If
                        End If

                    End If
                Else
                    PrintValue(12) = " ":                               FormatString(12) = "~"
                    If PREmployee.WkcCat <> 0 Then
                        PrintValue(13) = "W/C: " & PREmployee.WkcCat:   FormatString(18) = "a6"
                    Else
                        PrintValue(13) = " ":                           FormatString(18) = "a1"
                    End If
                End If

                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = "Tax Base: ":                           FormatString(1) = "a11" '  Print 3rd Detail Line
                
                If PREmployee.FWTMarried = 1 Then
                    PrintValue(2) = "FWT Married: Y":                   FormatString(2) = "a14"
                Else
                    PrintValue(2) = "FWT Married: N":                   FormatString(2) = "a14"
                End If
                
                PrintValue(3) = " ":                                    FormatString(3) = "a2"
                If PREmployee.FWTBasis = PREquate.BasisExemptions Then
                    PrintValue(4) = "FWT Exemps: ":                     FormatString(4) = "a12"
                    PrintValue(5) = PREmployee.FWTAmount:               FormatString(5) = "n2"
                    PrintValue(6) = " ":                                FormatString(6) = "a2"
                ElseIf PREmployee.FWTBasis = PREquate.BasisPercent Then
                    PrintValue(4) = "FWT: ":                            FormatString(4) = "a5"
                    PrintValue(5) = PREmployee.FWTAmount:               FormatString(5) = "d10"
                    PrintValue(6) = "%  ":                              FormatString(6) = "a3"
                End If
                
                If PREmployee.FWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(7) = "FWT Extra: ":                      FormatString(7) = "a11"
                    PrintValue(8) = PREmployee.FWTExtraAmount:          FormatString(8) = "d10"
                    PrintValue(9) = "% ":                               FormatString(9) = "a3"
                ElseIf PREmployee.FWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(7) = "FWT Extra: ":                      FormatString(7) = "a11"
                    PrintValue(8) = "$ " & PREmployee.FWTExtraAmount:   FormatString(8) = "d10"
                    PrintValue(9) = " ":                                FormatString(9) = "a2"
                End If
                
                If PREmployee.SWTMarried = 1 Then
                    PrintValue(10) = "SWT Married: Y":                  FormatString(10) = "a14"
                Else
                    PrintValue(10) = "SWT Married: N":                  FormatString(10) = "a14"
                End If
                
                PrintValue(11) = " ":                                   FormatString(11) = "a2"
                If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                    PrintValue(12) = "SWT Exemps: ":                    FormatString(12) = "a12"
                    PrintValue(13) = PREmployee.SWTAmount:              FormatString(13) = "n2"
                    PrintValue(14) = " ":                               FormatString(14) = "a2"
                ElseIf PREmployee.SWTBasis = PREquate.BasisPercent Then
                    PrintValue(12) = "SWT: ":                           FormatString(12) = "a5"
                    PrintValue(13) = PREmployee.SWTAmount:              FormatString(13) = "d9"
                    PrintValue(14) = "%  ":                             FormatString(14) = "a3"
                End If
                
                If PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(15) = "SWT Extra: ":                     FormatString(15) = "a11"
                    PrintValue(16) = PREmployee.SWTExtraAmount:         FormatString(16) = "d6"
                    PrintValue(17) = "%":                               FormatString(17) = "a1"
                    PrintValue(18) = " ":                               FormatString(18) = "~"
                ElseIf PREmployee.SWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(15) = "SWT Extra: ":                     FormatString(15) = "a11"
                    PrintValue(16) = "$ " & PREmployee.SWTExtraAmount:  FormatString(16) = "d8"
                    PrintValue(17) = " ":                               FormatString(17) = "~"
                End If
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = "Tax Flags: ":                          FormatString(1) = "a11" '  Print 4th Detail Line
                If PREmployee.NoSSTax = 1 Then
                    PrintValue(2) = "No SS Tax? " & "Y":                FormatString(2) = "a12"
                Else
                    PrintValue(2) = "No SS Tax? " & "N":                FormatString(2) = "a12"
                End If
                
                PrintValue(3) = " ":                                    FormatString(3) = "a4"
                If PREmployee.NoMedTax = 1 Then
                    PrintValue(4) = "No Med Tax? " & "Y":               FormatString(4) = "a13"
                Else
                    PrintValue(4) = "No Med Tax? " & "N":               FormatString(4) = "a13"
                End If
                
                PrintValue(5) = " ":                                    FormatString(5) = "a4"
                If PREmployee.NoFedTax = 1 Then
                    PrintValue(6) = "No Fed Tax? " & "Y":               FormatString(6) = "a13"
                Else
                    PrintValue(6) = "No Fed Tax? " & "N":               FormatString(6) = "a13"
                End If
                
                PrintValue(7) = " ":                                    FormatString(7) = "a4"
                If PREmployee.NoStateTax = 1 Then
                    PrintValue(8) = "No State Tax? " & "Y":             FormatString(8) = "a15"
                Else
                    PrintValue(8) = "No State Tax? " & "N":             FormatString(8) = "a15"
                End If
                
                PrintValue(9) = " ":                                    FormatString(9) = "a4"
                If PREmployee.NoCityTax = 1 Then
                    PrintValue(10) = "No City Tax? " & "Y":             FormatString(10) = "a14"
                Else
                    PrintValue(10) = "No City Tax? " & "N":             FormatString(10) = "a14"
                End If
                
                PrintValue(11) = " ":                                   FormatString(11) = "a4"
                If PREmployee.NoFedUnemp = 1 Then
                    PrintValue(12) = "No Fed Unemp? " & "Y":            FormatString(12) = "a15"
                Else
                    PrintValue(12) = "No Fed Unemp? " & "N":            FormatString(12) = "a15"
                End If
                
                PrintValue(13) = " ":                                   FormatString(13) = "a4"
                If PREmployee.NoStateUnemp = 1 Then
                    PrintValue(14) = "No State Unemp? " & "Y":          FormatString(14) = "a17"
                Else
                    PrintValue(14) = "No State Unemp? " & "N":          FormatString(14) = "a17"
                End If
                
                PrintValue(15) = " ":                                   FormatString(15) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = " ":                                    FormatString(1) = "a8"  '  Print 5th Detail Line
                If PREmployee.DateLastPaid <> 0 Then
                    PrintValue(2) = Format(PREmployee.DateLastPaid, "mm/dd/yy"):   FormatString(2) = "a12"
                    PrintValue(3) = " ":                                FormatString(3) = "a0"
                Else
                    PrintValue(2) = " ":                                FormatString(2) = "a12"
                    PrintValue(3) = " ":                                FormatString(3) = "a0"
                End If
                If PREmployee.DateHired <> 0 Then
                    PrintValue(4) = Format(PREmployee.DateHired, "mm/dd/yy"): FormatString(4) = "a12"
                    PrintValue(5) = " ":                                FormatString(5) = "a0"
                Else
                    PrintValue(4) = " ":                                FormatString(4) = "a12"
                    PrintValue(5) = " ":                                FormatString(5) = "a0"
                End If
                If PREmployee.DateLastRaise <> 0 Then
                    PrintValue(6) = Format(PREmployee.DateLastRaise, "mm/dd/yy"): FormatString(6) = "a12"
                    PrintValue(7) = " ":                                FormatString(7) = "a0"
                Else
                    PrintValue(6) = " ":                                FormatString(6) = "a12"
                    PrintValue(7) = " ":                                FormatString(7) = "a0"
                End If
                If PREmployee.DateLastReview <> 0 Then
                    PrintValue(8) = Format(PREmployee.DateLastReview, "mm/dd/yy"):    FormatString(8) = "a12"
                    PrintValue(9) = " ":                                FormatString(9) = "a0"
                Else
                    PrintValue(8) = " ":                                FormatString(8) = "a12"
                    PrintValue(9) = " ":                                FormatString(9) = "a0"
                End If
                If PREmployee.DateLastLayoff <> 0 Then
                    PrintValue(10) = Format(PREmployee.DateLastLayoff, "mm/dd/yy"):    FormatString(10) = "a12"
                    PrintValue(11) = " ":                               FormatString(11) = "a0"
                Else
                    PrintValue(10) = " ":                               FormatString(10) = "a12"
                    PrintValue(11) = " ":                               FormatString(11) = "a0"
                End If
                If PREmployee.DateLastRecall <> 0 Then
                    PrintValue(12) = Format(PREmployee.DateLastRecall, "mm/dd/yy"):   FormatString(12) = "a11"
                    PrintValue(13) = " ":                               FormatString(13) = "a0"
                Else
                    PrintValue(12) = " ":                               FormatString(12) = "a11"
                    PrintValue(13) = " ":                               FormatString(13) = "a0"
                End If
                If PREmployee.DateTerminated <> 0 Then
                    PrintValue(14) = Format(PREmployee.DateTerminated, "mm/dd/yy") & " " & PREmployee.TermReason:   FormatString(14) = "a15"
                    PrintValue(15) = " ":                               FormatString(15) = "a3"
                Else
                    PrintValue(14) = " ":                               FormatString(14) = "a15"
                    PrintValue(15) = " ":                               FormatString(15) = "a3"
                End If

                If PREmployee.DateOfBirth <> 0 Then
                    PrintValue(16) = Format(PREmployee.DateOfBirth, "mm/dd/yy"):  FormatString(16) = "a13"
                    PrintValue(17) = " ":                               FormatString(17) = "a0"
                Else
                    PrintValue(16) = " ":                               FormatString(16) = "a12"
                    PrintValue(17) = " ":                               FormatString(17) = "a0"
                End If
                If Trim(PREmployee.RaceCode) <> 0 Then
                    PrintValue(18) = PREmployee.RaceCode:               FormatString(18) = "a6"
                    PrintValue(19) = " ":                               FormatString(19) = "a1"
                Else
                    PrintValue(18) = " ":                               FormatString(18) = "a0"
                    PrintValue(19) = " ":                               FormatString(19) = "a1"
                End If
                If Trim(PREmployee.MaritalStatus) <> "" Then
                    PrintValue(20) = PREmployee.MaritalStatus:          FormatString(20) = "a7"
                    PrintValue(21) = " ":                               FormatString(21) = "a1"
                Else
                    PrintValue(20) = " ":                               FormatString(20) = "a0"
                    PrintValue(21) = " ":                               FormatString(21) = "a1"
                End If
                If PREmployee.EducationLevel <> 0 Then
                    PrintValue(22) = PREmployee.EducationLevel:         FormatString(22) = "a6"
                    PrintValue(23) = " ":                               FormatString(23) = "a1"
                Else
                    PrintValue(22) = " ":                               FormatString(22) = "a0"
                    PrintValue(23) = " ":                               FormatString(23) = "a1"
                End If
                If PREmployee.ShiftCode <> 0 Then
                    PrintValue(24) = PREmployee.ShiftCode:              FormatString(24) = "a7"
                    PrintValue(25) = " ":                               FormatString(25) = "a1"
                Else
                    PrintValue(24) = " ":                               FormatString(24) = "a1"
                    PrintValue(25) = " ":                               FormatString(25) = "a1"
                End If

                PrintValue(26) = " ":                                   FormatString(26) = "~"
                FormatPrint
                Ln = Ln + 1
                 
                PrintValue(1) = String(140, "-"):                       FormatString(1) = "a140"
                PrintValue(2) = " ":                                    FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
             
             Case "EmployeeRateList"
                    
                 PrintValue(1) = PREmployee.EmployeeNumber:             FormatString(1) = "a9"
                 PrintValue(2) = PREmployee.LFName:                     FormatString(2) = "a31"
                 PrintValue(3) = RTrim(PRDepartment.DepartmentNumber) & " " & Trim(PRDepartment.Name):  FormatString(3) = "a25"
                 If PREmployee.Salaried = 1 Then
                     PrintValue(4) = PREmployee.SalaryAmount:           FormatString(4) = "d10"
                 Else
                     PrintValue(4) = PREmployee.HourlyAmount:           FormatString(4) = "d10"
                 End If
                 PrintValue(5) = " ":                                   FormatString(5) = "a5"
                 If PREmployee.Salaried = 1 Then
                     PrintValue(6) = "SALARY":                          FormatString(6) = "a6"
                 Else
                     PrintValue(6) = "HOURLY":                          FormatString(6) = "a6"
                 End If
                 PrintValue(7) = " ":                                   FormatString(7) = "~"
                 FormatPrint
                 Ln = Ln + 1
                 
             Case "SSNFormat"
                 PrintValue(1) = PREmployee.EmployeeNumber:             FormatString(1) = "a9"
                 PrintValue(2) = " ":                                   FormatString(2) = "a6"
                 PrintValue(3) = PREmployee.LFName:                     FormatString(3) = "a40"
                 PrintValue(4) = " ":                                   FormatString(4) = "a3"
                 PrintValue(5) = PREmployee.SSString:                   FormatString(5) = "a12"
                 PrintValue(6) = " ":                                   FormatString(6) = "a3"
                 If PREmployee.DateOfBirth <> 0 Then
                     PrintValue(7) = PREmployee.DateOfBirth:            FormatString(7) = "a10"
                 Else
                     PrintValue(7) = " ":                               FormatString(7) = "a10"
                 End If
                 PrintValue(8) = " ":                                   FormatString(8) = "a4"
                 PrintValue(9) = PREmployee.Sex:                        FormatString(9) = "a6"
                 PrintValue(10) = " ":                                  FormatString(10) = "~"
                 FormatPrint
                 Ln = Ln + 1
                 
             Case "RateTaxList"
                 PrintValue(1) = PREmployee.EmployeeNumber:             FormatString(1) = "n5"
                 PrintValue(2) = " ":                                   FormatString(2) = "a2"
                 PrintValue(3) = PREmployee.LFName:                     FormatString(3) = "a18"
                 PrintValue(4) = " ":                                   FormatString(4) = "a3"
                 PrintValue(5) = RTrim(PRDepartment.DepartmentNumber) & "-" & Trim(PRDepartment.Name):  FormatString(5) = "a10"
                 PrintValue(6) = " ":                                   FormatString(6) = "a4"
                 If PREmployee.Salaried = 1 Then
                     PrintValue(7) = PREmployee.SalaryAmount:           FormatString(7) = "d10"
                     PrintValue(8) = "SALARY":                          FormatString(8) = "a7"
                 Else
                     PrintValue(7) = PREmployee.HourlyAmount:           FormatString(7) = "d10"
                     PrintValue(8) = "HOURLY":                          FormatString(8) = "a7"
                 End If
                 PrintValue(9) = " ":                                   FormatString(9) = "a3"
                 If PREmployee.NoSSTax = 1 Then
                    PrintValue(10) = "Y":                               FormatString(10) = "a4"
                 Else
                    PrintValue(10) = "N":                               FormatString(10) = "a4"
                 End If
                 If PREmployee.NoMedTax = 1 Then
                    PrintValue(11) = "Y":                               FormatString(11) = "a4"
                 Else
                    PrintValue(11) = "N":                               FormatString(11) = "a4"
                 End If
                 If PREmployee.NoFedTax = 1 Then
                    PrintValue(12) = "Y":                               FormatString(12) = "a4"
                 Else
                    PrintValue(12) = "N":                               FormatString(12) = "a4"
                 End If
                 If PREmployee.NoStateTax = 1 Then
                    PrintValue(13) = "Y":                               FormatString(13) = "a4"
                 Else
                    PrintValue(13) = "N":                               FormatString(13) = "a4"
                 End If
                 PrintValue(14) = " ":                                  FormatString(14) = "a0"
                 If PREmployee.NoCityTax = 1 Then
                    PrintValue(15) = "Y":                               FormatString(15) = "a4"
                 Else
                    PrintValue(15) = "N":                               FormatString(15) = "a4"
                 End If
                 PrintValue(16) = " ":                                  FormatString(16) = "a1"
                 If PREmployee.NoFedUnemp = 1 Then
                    PrintValue(17) = "Y":                               FormatString(17) = "a4"
                 Else
                    PrintValue(17) = "N":                               FormatString(17) = "a4"
                 End If
                 PrintValue(18) = " ":                                  FormatString(18) = "a2"
                 If PREmployee.NoStateUnemp = 1 Then
                    PrintValue(19) = "Y":                               FormatString(19) = "a5"
                 Else
                    PrintValue(19) = "N":                               FormatString(19) = "a5"
                 End If
                 If PREmployee.FWTMarried = 1 Then
                    PrintValue(20) = "Y":                               FormatString(20) = "a3"
                 Else
                    PrintValue(20) = "N":                               FormatString(20) = "a3"
                 End If
                 
                 PrintValue(21) = PREmployee.FWTBasis:                  FormatString(21) = "n3"
                 PrintValue(22) = PREmployee.FWTAmount:                 FormatString(22) = "d9"
                 PrintValue(23) = Format(PREmployee.FWTExtraAmount):    FormatString(23) = "d9"
                 PrintValue(24) = " ":                                  FormatString(24) = "a3"
                 If PREmployee.SWTMarried = 1 Then
                    PrintValue(25) = "Y":                               FormatString(25) = "a1"
                 Else
                    PrintValue(25) = "N":                               FormatString(25) = "a1"
                 End If
                 PrintValue(26) = " ":                                  FormatString(26) = "a3"
                 PrintValue(27) = PREmployee.SWTBasis:                  FormatString(27) = "n2"
                 PrintValue(28) = " ":                                  FormatString(28) = "a0"
                 PrintValue(29) = PREmployee.SWTAmount:                 FormatString(29) = "d9"
                 PrintValue(30) = PREmployee.SWTExtraAmount:            FormatString(30) = "d9"
                 PrintValue(31) = " ":                                  FormatString(31) = "~"
                 FormatPrint
                 Ln = Ln + 1
                
                Case "TimeCardLabels"
                    
                    YUnits = 225 + 15
                    SetFont 9, Equate.Portrait
                    LabelColumns = 1
                    
                    If NoLabels = 0 Then
                       ' GetDeptInfo (PREmployee.DepartmentID)
                       LabelString(1, 1) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a38"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                    
                       LabelString(1, 1) = RTrim(PREmployee.FLName)
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a38"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = "PERIOD ENDING DATE: " & frmLists.tdbPEDate
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a38"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                                             
                       LabelString(1, 1) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name)
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a38"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 3
                       
                    ElseIf NoLabels = 1 Then
                       LabelColumns = 2
                       ColumnCount = ColumnCount + 1
                       'GetDeptInfo (PREmployee.DepartmentID)
                       Label2String(1, ColumnCount) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       Label2String(2, ColumnCount) = RTrim(PREmployee.FLName)
                       Label2String(3, ColumnCount) = "PERIOD ENDING DATE: " & frmLists.tdbPEDate
                       Label2String(4, ColumnCount) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name)
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 3
                          If LabelCount = 10 Then
                            FormFeed
                            Ln = Ln + 4
                            LabelCount = 0
                          End If
                          For LRow = 1 To 4
                              ColumnCount = 0
                              PrintValue(1) = Label2String(LRow, 1):    FormatString(1) = "a38"
                              PrintValue(2) = Label2String(LRow, 2):    FormatString(2) = "a38"
                              PrintValue(3) = Label2String(LRow, 3):    FormatString(3) = "a38"
                              PrintValue(4) = Label2String(LRow, 4):    FormatString(4) = "a38"
                              PrintValue(5) = " ":                      FormatString(5) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
                          LabelCount = LabelCount + 1
                       End If
                    ElseIf NoLabels = 2 Then
                       LabelColumns = 3
                       ColumnCount = ColumnCount + 1
                       'GetDeptInfo (PREmployee.DepartmentID)
                       Label2String(1, ColumnCount) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       Label2String(2, ColumnCount) = RTrim(PREmployee.FLName)
                       Label2String(3, ColumnCount) = "P/E DATE: " & frmLists.tdbPEDate
                       Label2String(4, ColumnCount) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name)
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 2
                          If LabelCount = 10 Then
                              FormFeed
                              Ln = Ln + 4
                              LabelCount = 0
                          End If

                          For LRow = 1 To 4
                              ColumnCount = 0
                              PrintValue(1) = Label2String(LRow, 1):    FormatString(1) = "a38"
                              PrintValue(2) = Label2String(LRow, 2):    FormatString(2) = "a38"
                              PrintValue(3) = Label2String(LRow, 3):    FormatString(3) = "a38"
                              PrintValue(4) = Label2String(LRow, 4):    FormatString(4) = "a38"
                              PrintValue(5) = " ":                      FormatString(5) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
                       LabelCount = LabelCount + 1
                       End If
                       
                    End If
                Case "MailingLabels"
                    
                    YUnits = 225 + 15
                    LabelColumns = 1
                    SetFont 8, Equate.Portrait
                    
                    If NoLabels = 0 Then
                       LabelString(1, 1) = PREmployee.FLName
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a41"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = PREmployee.Address1
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a41"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a41"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 3
                       
                    ElseIf NoLabels = 1 Then
                       LabelColumns = 2
                       ColumnCount = ColumnCount + 1
 
                       LabelString(1, ColumnCount) = PREmployee.FLName
                       LabelString(2, ColumnCount) = PREmployee.Address1
                       LabelString(3, ColumnCount) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                       
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 2
                          If LabelCount = 10 Then
                              FormFeed
                              Ln = Ln + 5
                              LabelCount = 0
                          End If
                          For LRow = 1 To 3
                              ColumnCount = 0

                              PrintValue(1) = LabelString(LRow, 1):     FormatString(1) = "a41"
                              PrintValue(2) = LabelString(LRow, 2):     FormatString(2) = "a41"
                              PrintValue(3) = LabelString(LRow, 3):     FormatString(3) = "a41"
                              PrintValue(4) = " ":                      FormatString(4) = "~"
                              FormatPrint

                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                              
                          Next LRow
                          LabelCount = LabelCount + 1
                       End If
                       
                   ElseIf NoLabels = 2 Then
                       LabelColumns = 3
                       ColumnCount = ColumnCount + 1
                       LabelString(1, ColumnCount) = PREmployee.FLName
                       LabelString(2, ColumnCount) = PREmployee.Address1
                       LabelString(3, ColumnCount) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode

                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 3
                          If LabelCount = 10 Then
                              FormFeed
                              Ln = Ln + 5
                              LabelCount = 0
                          End If
                          
                          For LRow = 1 To 3
                              ColumnCount = 0

                              PrintValue(1) = LabelString(LRow, 1):     FormatString(1) = "a41"
                              PrintValue(2) = LabelString(LRow, 2):     FormatString(2) = "a41"
                              PrintValue(3) = LabelString(LRow, 3):     FormatString(3) = "a41"
                              PrintValue(4) = " ":                      FormatString(4) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
                          LabelCount = LabelCount + 1
                       End If
                   End If

           End Select
                
Cycle1:
           If Not PREmployee.GetNext Then
               Exit Do
           End If
    Loop
    
Select Case ReportType
   Case "MailingLabels"
      If NoLabels = 2 Then
         Ln = Ln + 3
         If ColumnCount = 1 Then
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         ElseIf ColumnCount = 2 Then    '(Row, Col)
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a41"
               PrintValue(2) = LabelString(1, 2):                       FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a41"
               PrintValue(2) = LabelString(2, 2):                       FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a41"
               PrintValue(2) = LabelString(3, 2):                       FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
         End If
      
      ElseIf NoLabels = 1 Then
         If ColumnCount > 0 Then
            Ln = Ln + 2
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         End If
      End If
      
   Case "TimeCardLabels"
      If NoLabels = 2 Then
         Ln = Ln + 3
         If ColumnCount = 1 Then
            
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                                    
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         ElseIf ColumnCount = 2 Then
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a41"
               PrintValue(2) = Label2String(1, 2):                      FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a41"
               PrintValue(2) = Label2String(2, 2):                      FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a41"
               PrintValue(2) = Label2String(3, 2):                      FormatString(2) = "a41"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a41"
               PrintValue(2) = Label2String(4, 2):                      FormatString(2) = "a41"
               PrintValue(3) = " "
               FormatString(3) = "~"
               FormatPrint
         End If
      
      ElseIf NoLabels = 1 Then
         If ColumnCount > 0 Then
            Ln = Ln + 2
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a41"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         End If
      End If
 End Select
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub CheckDetailHdr()
    PrintValue(1) = "CHECK DATE":           FormatString(1) = "a14"
    PrintValue(2) = "P/E":                  FormatString(2) = "a8"
    PrintValue(3) = "CHK NOS.":             FormatString(3) = "a14"
    PrintValue(4) = "NO. CKS":              FormatString(4) = "a8"
    PrintValue(5) = "GROSS":                FormatString(5) = "r14"
    PrintValue(6) = "SS TAX":                 FormatString(6) = "r15"
    PrintValue(7) = "MED":                  FormatString(7) = "r15"
    PrintValue(8) = "FWT":                  FormatString(8) = "r15"
    PrintValue(9) = "SWT":                  FormatString(9) = "r15"
    PrintValue(10) = "CWT":                 FormatString(10) = "r14"
    PrintValue(11) = "EIC":                 FormatString(11) = "r14"
    PrintValue(12) = " ":                   FormatString(12) = "~"
    FormatPrint
    Ln = Ln + 1
    

    PrintValue(1) = "CITY TAX":             FormatString(1) = "r30"
    PrintValue(2) = "DED":                  FormatString(2) = "r13"
    PrintValue(3) = "NET":                  FormatString(3) = "r15"
    PrintValue(4) = "FWT PMT":              FormatString(4) = "r15"
    PrintValue(5) = "FUTA PMT":             FormatString(5) = "r15"
    PrintValue(6) = "SWT PMT":              FormatString(6) = "r15"
    PrintValue(7) = "CWT PMT":              FormatString(7) = "r15"
    PrintValue(8) = "UNEMP TAX":            FormatString(8) = "r14"
    PrintValue(9) = "TOT BANK CR":          FormatString(9) = "r14"
    PrintValue(10) = " ":                   FormatString(10) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = String(148, "="):       FormatString(1) = "a148"
    PrintValue(2) = " ":                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Public Sub CheckDetail(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal StartDate As Long, _
                         ByVal EndDate As Long, _
                         ByVal OptDate As String)

Dim BegChkNo As Long
Dim EndChkNo As Long
Dim LastChkDate As Date
Dim sqlstring1 As String
Dim GTotBankCr As Currency
    
    ReportTitle = "PAYROLL CHECK DETAIL REPORT"
    PrtInit ("Land")
    LandSw = 1
    Columns = 145
    SetFont 8, Equate.LandScape
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "CheckDate", adDate:               rs.Fields.Append "PEDate", adDate
    rs.Fields.Append "ChkNos", adDouble:                rs.Fields.Append "NoChks", adDouble
    rs.Fields.Append "Gross", adCurrency:               rs.Fields.Append "SS TAX", adCurrency
    rs.Fields.Append "MED", adCurrency:                 rs.Fields.Append "FWT", adCurrency
    rs.Fields.Append "SWT", adCurrency:                 rs.Fields.Append "CWT", adCurrency
    rs.Fields.Append "EIC", adCurrency:                 rs.Fields.Append "CityTax", adCurrency
    rs.Fields.Append "DED", adCurrency:                 rs.Fields.Append "NET", adCurrency
    rs.Fields.Append "FWTPMT", adCurrency:              rs.Fields.Append "FUTAPMT", adCurrency
    rs.Fields.Append "SWTPMT", adCurrency:              rs.Fields.Append "CWTPMT", adCurrency
    rs.Fields.Append "UNEMPTAX", adCurrency:            rs.Fields.Append "TotBankCr", adCurrency

    rs.Open , , adOpenDynamic, adLockOptimistic

    rsMON.CursorLocation = adUseClient
    rsMON.Fields.Append "CheckDate", adDate:               rsMON.Fields.Append "PEDate", adDate
    rsMON.Fields.Append "MonYr", adVarChar, 6, adFldIsNullable: rsQTR.Fields.Append "NoChks", adDouble
    rsMON.Fields.Append "Gross", adCurrency:               rsMON.Fields.Append "SS TAX", adCurrency
    rsMON.Fields.Append "MED", adCurrency:                 rsMON.Fields.Append "FWT", adCurrency
    rsMON.Fields.Append "SWT", adCurrency:                 rsMON.Fields.Append "CWT", adCurrency
    rsMON.Fields.Append "EIC", adCurrency:                 rsMON.Fields.Append "CityTax", adCurrency
    rsMON.Fields.Append "DED", adCurrency:                 rsMON.Fields.Append "NET", adCurrency
    rsMON.Fields.Append "FWTPMT", adCurrency:              rsMON.Fields.Append "FUTAPMT", adCurrency
    rsMON.Fields.Append "SWTPMT", adCurrency:              rsMON.Fields.Append "CWTPMT", adCurrency
    rsMON.Fields.Append "UNEMPTAX", adCurrency:            rsMON.Fields.Append "TotBankCr", adCurrency

    rsMON.Open , , adOpenDynamic, adLockOptimistic
    
    rsQTR.CursorLocation = adUseClient
    rsQTR.Fields.Append "CheckDate", adDate:               rsQTR.Fields.Append "PEDate", adDate
    rsQTR.Fields.Append "QTRYr", adVarChar, 6, adFldIsNullable: rsQTR.Fields.Append "NoChks", adDouble
    rsQTR.Fields.Append "Gross", adCurrency:               rsQTR.Fields.Append "SS TAX", adCurrency
    rsQTR.Fields.Append "MED", adCurrency:                 rsQTR.Fields.Append "FWT", adCurrency
    rsQTR.Fields.Append "SWT", adCurrency:                 rsQTR.Fields.Append "CWT", adCurrency
    rsQTR.Fields.Append "EIC", adCurrency:                 rsQTR.Fields.Append "CityTax", adCurrency
    rsQTR.Fields.Append "DED", adCurrency:                 rsQTR.Fields.Append "NET", adCurrency
    rsQTR.Fields.Append "FWTPMT", adCurrency:              rsQTR.Fields.Append "FUTAPMT", adCurrency
    rsQTR.Fields.Append "SWTPMT", adCurrency:              rsQTR.Fields.Append "CWTPMT", adCurrency
    rsQTR.Fields.Append "UNEMPTAX", adCurrency:            rsQTR.Fields.Append "TotBankCr", adCurrency

    rsQTR.Open , , adOpenDynamic, adLockOptimistic
    
    rsYTD.CursorLocation = adUseClient
    rsYTD.Fields.Append "CheckDate", adDate:            rsYTD.Fields.Append "PEDate", adDate
    rsYTD.Fields.Append "ChkNos", adDouble:             rsYTD.Fields.Append "NoChks", adDouble
    rsYTD.Fields.Append "Gross", adCurrency:            rsYTD.Fields.Append "SS TAX", adCurrency
    rsYTD.Fields.Append "MED", adCurrency:              rsYTD.Fields.Append "FWT", adCurrency
    rsYTD.Fields.Append "SWT", adCurrency:              rsYTD.Fields.Append "CWT", adCurrency
    rsYTD.Fields.Append "EIC", adCurrency:              rsYTD.Fields.Append "CityTax", adCurrency
    rsYTD.Fields.Append "DED", adCurrency:              rsYTD.Fields.Append "NET", adCurrency
    rsYTD.Fields.Append "FWTPMT", adCurrency:           rsYTD.Fields.Append "FUTAPMT", adCurrency
    rsYTD.Fields.Append "SWTPMT", adCurrency:           rsYTD.Fields.Append "CWTPMT", adCurrency
    rsYTD.Fields.Append "UNEMPTAX", adCurrency:         rsYTD.Fields.Append "TotBankCr", adCurrency
    
    rsYTD.Open , , adOpenDynamic, adLockOptimistic
    
    rsMON.AddNew
    rsMON!CheckDate = PRHist.CheckDate
    rsMON!PEDate = PRHist.PEDate
    rsMON!ChkNos = 0
    rsMON!NoChks = 0
    rsMON!Gross = 0
    rsMON!SSTax = 0
    rsMON!Med = 0
    rsMON!FWT = 0
    rsMON!SWT = 0
    rsMON!CWT = 0
    rsMON!EIC = 0
    rsMON!CityTax = 0
    rsMON!DED = 0
    rsMON!Net = 0
    rsMON!FWTPmt = 0
    rsMON!FUTAPmt = 0
    rsMON!SWTPmt = 0
    rsMON!CWTPmt = 0
    rsMON!UnempTax = 0
    rsMON!TotBankCr = 0
                
    rsQTR.AddNew
    rsQTR!CheckDate = PRHist.CheckDate
    rsQTR!PEDate = PRHist.PEDate
    rsQTR!ChkNos = 0
    rsQTR!NoChks = 0
    rsQTR!Gross = 0
    rsQTR!SSTax = 0
    rsQTR!Med = 0
    rsQTR!FWT = 0
    rsQTR!SWT = 0
    rsQTR!CWT = 0
    rsQTR!EIC = 0
    rsQTR!CityTax = 0
    rsQTR!DED = 0
    rsQTR!Net = 0
    rsQTR!FWTPmt = 0
    rsQTR!FUTAPmt = 0
    rsQTR!SWTPmt = 0
    rsQTR!CWTPmt = 0
    rsQTR!UnempTax = 0
    rsQTR!TotBankCr = 0
                                                                
    rsYTD.AddNew
    rsYTD!NoChks = 0
    rsYTD!Gross = 0
    rsYTD!SSTax = 0
    rsYTD!Med = 0
    rsYTD!FWT = 0
    rsYTD!SWT = 0
    rsYTD!CWT = 0
    rsYTD!EIC = 0
    rsYTD!CityTax = 0
    rsYTD!DED = 0
    rsYTD!Net = 0
    rsYTD!FWTPmt = 0
    rsYTD!FUTAPmt = 0
    rsYTD!SWTPmt = 0
    rsYTD!CWTPmt = 0
    rsYTD!UnempTax = 0
    rsYTD!TotBankCr = 0
            
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PR Batch Not Found: " & PRBatchID, vbExclamation
            GoBack
        End If
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & PRBatch.PEDate
    End If
    
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
        OptDate = "P/E DATE"
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRHist WHERE PRHist.CheckDate >= " & CLng(StartDate) & _
                        " AND PRHist.CheckDate <= " & CLng(EndDate)
            Msg1 = "CHECK DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        ElseIf OptDate = "P/E DATE" Then
             SQLString = "SELECT * FROM PRHist WHERE PRHist.PEDate >= " & CLng(StartDate) & _
                        " AND PRHist.PEDate <= " & CLng(EndDate)
            Msg1 = "PERIOD ENDING DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If
    
    SQLString = Trim(SQLString) & " Order by CheckDate, CheckNumber"
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbExclamation, "Tax Wage Report"
        GoBack
    End If
    LastChkDate = 0
 
    Do
'       ' employee select filter?  Populate Recordset
'        If frmEmpSelect.AllEmployees = False Then

            sqlstring1 = "CheckDate = " & Format(PRHist.CheckDate, "mm/dd/yyyy")
            rs.Find sqlstring1, 0, adSearchForward, 1
                    
            If rs.EOF Then
                rs.AddNew
                rs!CheckDate = PRHist.CheckDate
                rs!PEDate = PRHist.PEDate
                rs!ChkNos = 0
                rs!NoChks = 0
                rs!Gross = 0
                rs!SSTax = 0
                rs!Med = 0
                rs!FWT = 0
                rs!SWT = 0
                rs!CWT = 0
                rs!EIC = 0
                rs!CityTax = 0
                rs!DED = 0
                rs!Net = 0
                rs!FWTPmt = 0
                rs!FUTAPmt = 0
                rs!SWTPmt = 0
                rs!CWTPmt = 0
                rs!UnempTax = 0
                rs!TotBankCr = 0
                rs.MoveFirst
            End If
            
            rs!NoChks = rs!NoChks + 1
            
            If rs!NoChks = 1 Then
                BegChkNo = PRHist.CheckNumber
            End If
            
            If LastChkDate <> 0 And LastChkDate <> PRHist.CheckDate Then
                EndChkNo = 0
                rs!NoChks = 0
                rs!Gross = 0
                rs!SSTax = 0
                rs!Med = 0
                rs!FWT = 0
                rs!SWT = 0
                rs!CWT = 0
                rs!EIC = 0
                rs!CityTax = 0
                rs!DED = 0
                rs!Net = 0
                rs!FWTPmt = 0
                rs!FUTAPmt = 0
                rs!SWTPmt = 0
                rs!CWTPmt = 0
                rs!UnempTax = 0
                rs!TotBankCr = 0
            End If
            
            rs!PEDate = PRHist.PEDate
            rs!Gross = rs!Gross + PRHist.Gross
            rs!SSTax = rs!SSTax + PRHist.SSTax
            rs!Med = rs!Med + PRHist.MedTax
            rs!FWT = rs!FWT + PRHist.FWTTax
            rs!SWT = rs!SWT + PRHist.SWTTax
            rs!CWT = rs!CWT + PRHist.CWTTax
            rs!EIC = 0                                  ''''   CHANGE   ''''''''''''''''''''''''''''
            rs!CityTax = rs!CityTax + 0                 ''''   CHANGE   ''''''''''''''''''''''''''''
            rs!DED = rs!DED + PRHist.Deductions
            rs!Net = rs!Net + PRHist.Net
            rs!FWTPmt = rs!FWTPmt + 0
            rs!FUTAPmt = rs!FUTAPmt + 0                 ''''   CHANGE   ''''''''''''''''''''''''''''
            rs!SWTPmt = rs!SWTPmt + 0
            rs!CWTPmt = rs!CWTPmt + 0
            rs!UnempTax = rs!UnempTax + 0   ''''   CHANGE   ''''''''''''''''''''''''''''
            rs!TotBankCr = rs!TotBankCr + PRHist.Net    ''''   CHANGE   ''''''''''''''''''''''''''''
            
            rsYTD!NoChks = rsYTD!NoChks + 1
            rsYTD!Gross = rsYTD!Gross + PRHist.Gross
            rsYTD!SSTax = rsYTD!SSTax + PRHist.SSTax
            rsYTD!Med = rsYTD!Med + PRHist.MedTax
            rsYTD!FWT = rsYTD!FWT + PRHist.FWTTax
            rsYTD!SWT = rsYTD!SWT + PRHist.SWTTax
            rsYTD!CWT = rsYTD!CWT + PRHist.CWTTax
            rsYTD!EIC = rsYTD!EIC + 0                   ''''   CHANGE   ''''''''''''''''''''''''''''
            rsYTD!CityTax = rsYTD!CityTax + 0
            rsYTD!DED = rsYTD!DED + PRHist.Deductions
            rsYTD!Net = rsYTD!Net + PRHist.Net
            rsYTD!FWTPmt = rsYTD!FWTPmt + 0
            rsYTD!FUTAPmt = rsYTD!FUTAPmt + 0
            rsYTD!SWTPmt = rsYTD!SWTPmt + 0
            rsYTD!CWTPmt = rsYTD!CWTPmt + 0
            rsYTD!UnempTax = rsYTD!UnempTax + 0
            rsYTD!TotBankCr = rsYTD!TotBankCr + PRHist.Net
            rsYTD.Update
            
            EndChkNo = PRHist.CheckNumber
            LastChkDate = PRHist.CheckDate
            rs.Update

        If Not PRHist.GetNext Then Exit Do
    Loop

    rs.MoveFirst

    Do Until rs.EOF
        If Ln = 0 Or Ln > MaxLines - LineCount Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, Msg2, ""
            CheckDetailHdr
        End If

        PrintValue(1) = rs!CheckDate:               FormatString(1) = "a14"
        PrintValue(2) = Format(rs!PEDate, "mm/dd"): FormatString(2) = "a8"
        PrintValue(3) = BegChkNo & "-" & EndChkNo:  FormatString(3) = "a14"
        PrintValue(4) = rs!NoChks:                  FormatString(4) = "r7"
        PrintValue(5) = rs!Gross:                   FormatString(5) = "r15"
        PrintValue(6) = rs!SSTax:                   FormatString(6) = "r15"
        PrintValue(7) = rs!Med:                     FormatString(7) = "r15"
        PrintValue(8) = rs!FWT:                     FormatString(8) = "r15"
        PrintValue(9) = rs!SWT:                     FormatString(9) = "r15"
        PrintValue(10) = rs!CWT:                    FormatString(10) = "r14"
        PrintValue(11) = rs!EIC:                    FormatString(11) = "r14"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = rs!CityTax:                 FormatString(1) = "r30"
        PrintValue(2) = rs!DED:                     FormatString(2) = "r13"
        PrintValue(3) = rs!Net:                     FormatString(3) = "r15"
        PrintValue(4) = rs!FWTPmt:                  FormatString(4) = "r15"
        PrintValue(5) = rs!FUTAPmt:                 FormatString(5) = "r15"
        PrintValue(6) = rs!SWTPmt:                  FormatString(6) = "r15"
        PrintValue(7) = rs!CWTPmt:                  FormatString(7) = "r15"
        PrintValue(8) = rs!UnempTax:                FormatString(8) = "r14"
        PrintValue(9) = rs!TotBankCr:               FormatString(9) = "r14"
    
        FormatPrint
        Ln = Ln + 1
        rs.MoveNext
    Loop
    
    rsYTD.MoveFirst

    Do Until rsYTD.EOF
        If Ln = 0 Or Ln > MaxLines - LineCount Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, Msg2, ""
            CheckDetailHdr
        End If

        PrintValue(1) = rsYTD!NoChks:               FormatString(1) = "r43"
        PrintValue(2) = rsYTD!Gross:                FormatString(2) = "r15"
        PrintValue(3) = rsYTD!SSTax:                FormatString(3) = "r15"
        PrintValue(4) = rsYTD!Med:                  FormatString(4) = "r15"
        PrintValue(5) = rsYTD!FWT:                  FormatString(5) = "r15"
        PrintValue(6) = rsYTD!SWT:                  FormatString(6) = "r15"
        PrintValue(7) = rsYTD!CWT:                  FormatString(7) = "r14"
        PrintValue(8) = rsYTD!EIC:                  FormatString(8) = "r14"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = rsYTD!CityTax:              FormatString(1) = "r30"
        PrintValue(2) = rsYTD!DED:                  FormatString(2) = "r13"
        PrintValue(3) = rsYTD!Net:                  FormatString(3) = "r15"
        PrintValue(4) = rsYTD!FWTPmt:               FormatString(4) = "r15"
        PrintValue(5) = rsYTD!FUTAPmt:              FormatString(5) = "r15"
        PrintValue(6) = rsYTD!SWTPmt:               FormatString(6) = "r15"
        PrintValue(7) = rsYTD!CWTPmt:               FormatString(7) = "r15"
        PrintValue(8) = rsYTD!UnempTax:             FormatString(8) = "r14"
        PrintValue(9) = rsYTD!TotBankCr:            FormatString(9) = "r14"
    
        FormatPrint
        Ln = Ln + 1
        rsYTD.MoveNext
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub


Public Sub EarnSummary(ByVal RangeType As Byte, _
                       ByVal BatchNumbr As Long, _
                       ByVal PEDate As Long, _
                       ByVal StartDate As Long, _
                       ByVal EndDate As Long, _
                       ByVal OptDate As String)
Dim sqlstring1 As String
Dim SQLString2 As String
Dim ct As Long
Dim TestTotal As Currency
Dim cntr As Long
Dim AddrString As String
Dim PageFlag As Boolean
Dim PrintCount As Byte

    LastEmpID = 0
    ct = 0
    cntr = 0

    ReportTitle = "EARNINGS SUMMARY REPORT"
    PrtInit ("Land")
    LandSw = 1
    Columns = 145
    SetFont 8, Equate.LandScape
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    MaxLines = 46
    MaxLines = 44

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
        
    rs.CursorLocation = adUseClient
    rs.Fields.Append "EmpNo", adDouble:                     rs.Fields.Append "DeptID", adDouble
    rs.Fields.Append "EmpID", adDouble:                     rs.Fields.Append "FLName", adChar, 30, adFldMayBeNull
    rs.Fields.Append "Addr1", adChar, 30, adFldMayBeNull:   rs.Fields.Append "Addr2", adChar, 30, adFldMayBeNull
    rs.Fields.Append "City", adChar, 30, adFldMayBeNull:    rs.Fields.Append "State", adChar, 2, adFldMayBeNull
    rs.Fields.Append "Zip", adChar, 10, adFldMayBeNull:     rs.Fields.Append "SSN", adDouble:
    rs.Fields.Append "Date", adDate:                        rs.Fields.Append "ChkNo", adDouble:
    rs.Fields.Append "DeptNo", adDouble:                    rs.Fields.Append "Rate", adCurrency
    rs.Fields.Append "Gross", adCurrency:                   rs.Fields.Append "SSTAX", adCurrency
    rs.Fields.Append "MED", adCurrency:                     rs.Fields.Append "FWT", adCurrency
    rs.Fields.Append "SWT", adCurrency:                     rs.Fields.Append "CWT", adCurrency
    rs.Fields.Append "OtherTax", adCurrency:                rs.Fields.Append "TotDed", adCurrency
    rs.Fields.Append "Net", adCurrency:                     rs.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    rs.Fields.Append "FWTBasis", adSingle:                  rs.Fields.Append "FWTExtraBasis", adSingle
    rs.Fields.Append "FWTAmount", adCurrency:               rs.Fields.Append "FWTExtraAmount", adCurrency
    rs.Fields.Append "SWTBasis", adSingle:                  rs.Fields.Append "SWTExtraBasis", adSingle
    rs.Fields.Append "SWTAmount", adCurrency:               rs.Fields.Append "SWTExtraAmount", adCurrency
    rs.Fields.Append "InclCurr", adBoolean:                 rs.Fields.Append "InclQTD", adBoolean

    rs.Open , , adOpenDynamic, adLockOptimistic
    
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PR Batch Not Found: " & PRBatchID, vbExclamation
            GoBack
        End If
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & PRBatch.PEDate
    End If

    sqlstring1 = "SELECT * FROM PRHist WHERE PRHist.CheckDate >= " & CLng(frmEarnSumm.BOYDate) & _
                " AND PRHist.CheckDate <= " & CLng(frmEarnSumm.EOQDate)

    If RangeType = PREquate.RangeTypeBatch Then
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
        OptDate = "P/E DATE"
    Else
        If OptDate = "CHECK DATE" Then
            Msg1 = "CHECK DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        ElseIf OptDate = "P/E DATE" Then
            Msg1 = "PERIOD ENDING DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If
    
    Msg1 = Trim(Msg1) & "  \  QTD Totals for: " & Format(frmEarnSumm.BOQDate, "mm/dd/yy") & " To: " & _
           Format(frmEarnSumm.EOQDate, "mm/dd/yy")
    
    If Not PRHist.GetBySQL(sqlstring1) Then
        MsgBox "No Data Found !!!", vbExclamation, "Tax Wage Report"
        GoBack
    End If

    Recs = PRHist.Records
    
    Do
       
        ' employee select filter?  Populate Recordset
        ct = ct + 1
        If ct = 1 Or ct Mod 20 = 0 Then
            frmProgress.lblMsg1 = Trim(PRCompany.Name) & " Earnings Summary"
            frmProgress.lblMsg2 = "Processing History: " & Format(ct, "##,###,##0") & " Of: " & Format(Recs, "##,###,##0")
            frmProgress.Refresh
        End If

        If frmEmpSelect.AllEmployees = False Then
            SQLString2 = "EmployeeID = " & PRHist.EmployeeID
            frmEmpSelect.rsEmp.Find SQLString2, 0, adSearchForward, 1
            If frmEmpSelect.rsEmp.EOF Then
                GoTo NextHist
            End If
            If frmEmpSelect.rsEmp!Select = False Then GoTo NextHist
        End If

        rs.AddNew
        rs!EmpID = PRHist.EmployeeID
        If PREmployee.GetByID(PRHist.EmployeeID) Then
            rs!EmpNo = PREmployee.EmployeeNumber
            
            rs!FLName = Mid(PREmployee.FLName, 1, 30)
            rs!EmployeeName = Mid(PREmployee.LFName, 1, 80)
            rs!Addr1 = Mid(PREmployee.Address1, 1, 30)
            rs!Addr2 = Mid(PREmployee.Address2, 1, 30)
            rs!City = Mid(PREmployee.City, 1, 30)
            rs!State = Mid(PREmployee.State, 1, 2)
            rs!zip = Mid(PREmployee.ZipCode, 1, 10)
            rs!SSN = PREmployee.SSN
            rs!FWTBasis = PREmployee.FWTBasis
            rs!FWTExtraBasis = PREmployee.FWTExtraBasis
            rs!FWTAmount = PREmployee.FWTAmount
            rs!FWTExtraAmount = PREmployee.FWTExtraAmount
            rs!SWTBasis = PREmployee.SWTBasis
            rs!SWTExtraBasis = PREmployee.SWTExtraBasis
            rs!SWTAmount = PREmployee.SWTAmount
            rs!SWTExtraAmount = PREmployee.SWTExtraAmount
        
            rs!DeptNo = 0
            If PREmployee.DepartmentID <> 0 Then
                If PRDepartment.GetByID(PREmployee.DepartmentID) Then
                    rs!DeptNo = PRDepartment.DepartmentNumber
                End If
            End If
        
        End If

        rs!inclcurr = False
        rs!InclQTD = False
        If PRHist.CheckDate >= CLng(StartDate) And PRHist.CheckDate <= CLng(EndDate) Then
            rs!inclcurr = True
            rs!Rate = PRHist.RegRate
            rs!DeptID = PRHist.DepartmentID
            PRDepartment.GetByID (PRHist.DepartmentID)
            rs!DeptNo = PRDepartment.DepartmentNumber
        End If
        
        If PRHist.CheckDate >= CLng(frmEarnSumm.BOQDate) And PRHist.CheckDate <= CLng(frmEarnSumm.EOQDate) Then
            rs!InclQTD = True
        End If

        rs!Date = PRHist.CheckDate
        rs!ChkNo = PRHist.CheckNumber
        rs!Gross = PRHist.Gross
        rs!SSTax = PRHist.SSTax
        rs!Med = PRHist.MedTax
        rs!FWT = PRHist.FWTTax
        rs!SWT = PRHist.SWTTax
        rs!CWT = PRHist.CWTTax
        rs!OtherTax = 0
        rs!TotDed = PRHist.Deductions
        rs!Net = PRHist.Net + PRHist.DirectDeposit

        rs.Update

NextHist:
        If Not PRHist.GetNext Then Exit Do
    Loop

    If frmEarnSumm.optEmpNo Then
        rs.Sort = "EmpNo, Date"
    Else
        rs.Sort = "EmployeeName, Date"
    End If
    
    Ln = 0
    ct = 0

    rs.MoveFirst

    PrintCount = 255
    
    Do

        PageFlag = False
        If frmEarnSumm.chkExcludeDetail = 0 Then
            If Ln = 0 Or Ln >= MaxLines Then PageFlag = True
        ElseIf PrintCount = 255 Then
            PageFlag = True
            PrintCount = 1
        ElseIf LastEmpID <> rs!EmpID Then
            PrintCount = PrintCount + 1
            If PrintCount = 7 Then
                PageFlag = True
                PrintCount = 1
            End If
        End If
        
        If PageFlag = True Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = Ln + 1
            EarnSummaryHdr
        End If
        ct = ct + 1

        '  EMPLOYEE HEADER LINE
        If LastEmpID <> rs!EmpID Then
            PrintValue(1) = " ":                                                    FormatString(1) = "a5"
            PrintValue(2) = rs!EmpNo:                                               FormatString(2) = "a7"
            If frmEarnSumm.chkSSN = 1 Then
                PrintValue(3) = Format(rs!SSN, "###-##-####"):                      FormatString(3) = "a13"
            Else
                PrintValue(3) = " ":                                                FormatString(3) = "a3"
            End If
            PrintValue(4) = rs!EmployeeName:                                        FormatString(4) = "a31"
            PrintValue(5) = "DEPT: " & rs!DeptNo:                                   FormatString(5) = "a10"
            PrintValue(6) = "RATE: " & Format(rs!Rate, "###,###.#0"):               FormatString(6) = "a14"
            
            If rs!FWTBasis = PREquate.BasisExemptions Then
                PrintValue(7) = "FWT Exmps: " & Format(rs!FWTAmount, "##0"):        FormatString(7) = "a15"
            Else
                PrintValue(7) = "FWT Exmps: " & Format(0, "##0"):                   FormatString(7) = "a15"
            End If
            
            If rs!FWTBasis = PREquate.BasisPercent Then
                PrintValue(8) = "FWT %: " & Format(rs!FWTAmount, "##0.00") & " %":  FormatString(8) = "a20"
            Else
                PrintValue(8) = "FWT %: " & Format(0, "##0.00") & " %":             FormatString(8) = "a15"
            End If
            
            If rs!FWTExtraBasis = PREquate.BasisAmount Then
                PrintValue(9) = "FWT Extra $: " & Format(rs!FWTExtraAmount, "##0.00"):   FormatString(9) = "a20"
            Else
                PrintValue(9) = "FWT Extra $: " & Format(0, "##0.00"):              FormatString(9) = "a20"
            End If
            
            If rs!FWTExtraBasis = PREquate.BasisPercent Then
                PrintValue(10) = "FWT Extra %: " & Format(rs!FWTExtraAmount, "##0.00") & " %": FormatString(10) = "a20"
            Else
                PrintValue(10) = "FWT Extra %: " & Format(0, "##0.00") & " %":      FormatString(10) = "a20"
            End If
            FormatPrint
            Ln = Ln + 1
            
            
            PrintValue(1) = " ":                                                    FormatString(1) = "a70"
            If rs!FWTBasis = PREquate.BasisExemptions Then
                PrintValue(2) = "SWT Exmps: " & Format(rs!SWTAmount, "##0"):        FormatString(2) = "a15"
            Else
                PrintValue(2) = "SWT Exmps: " & Format(0, "##0"):                   FormatString(2) = "a15"
            End If
            If rs!SWTBasis = PREquate.BasisPercent Then
                PrintValue(3) = "SWT %: " & Format(rs!SWTAmount, "##0.00") & " %":  FormatString(3) = "a20"
            Else
                PrintValue(3) = "SWT %: " & Format(0, "##0.00") & " %":             FormatString(3) = "a15"
            End If
            
            If rs!SWTExtraBasis = PREquate.BasisAmount Then
                PrintValue(4) = "SWT Extra $: " & Format(rs!SWTExtraAmount, "##0.00"):   FormatString(4) = "a20"
            Else
                PrintValue(4) = "SWT Extra $: " & Format(0, "##0.00"):              FormatString(4) = "a20"
            End If
            
            If rs!SWTExtraBasis = PREquate.BasisPercent Then
                PrintValue(5) = "SWT Extra %: " & Format(rs!SWTExtraAmount, "##0.00") & " %": FormatString(5) = "a20"
            Else
                PrintValue(5) = "SWT Extra %: " & Format(0, "##0.00") & " %":      FormatString(5) = "a20"
            End If
            
            PrintValue(6) = " ":                                                   FormatString(6) = "~"

            FormatPrint
            Ln = Ln + 1
            
            If frmEarnSumm.chkAddr = 1 Then
                PrintValue(1) = " ":                                                FormatString(1) = "a15"
                AddrString = Trim(rs!Addr1) & "  " & Trim(rs!Addr2) & "  " & Trim(rs!City) & "  " & rs!State & "  " & Trim(rs!zip)
                PrintValue(2) = AddrString:                                         FormatString(2) = "a120"
                PrintValue(3) = " ":                                                FormatString(3) = "~"

                FormatPrint
                Ln = Ln + 1
            End If

        End If
 
        ' Print Employee detail line
        If rs!inclcurr Then
            
            If frmEarnSumm.chkExcludeDetail = 0 Then
                PrintValue(1) = Format(rs!Date, "yyyymmdd"):         FormatString(1) = "a9"
                PrintValue(2) = rs!ChkNo:                            FormatString(2) = "a7"
                PrintValue(3) = rs!DeptNo:                           FormatString(3) = "a6"
                PrintValue(4) = Format(rs!Rate, "###,###.#0"):       FormatString(4) = "d9"
                PrintValue(5) = Format(rs!Gross, "####,###.#0"):     FormatString(5) = "d13"
                PrintValue(6) = Format(rs!SSTax, "###,###.#0"):      FormatString(6) = "d13"
                PrintValue(7) = Format(rs!Med, "###,###.#0"):        FormatString(7) = "d13"
                PrintValue(8) = Format(rs!FWT, "###,###.#0"):        FormatString(8) = "d13"
                PrintValue(9) = Format(rs!SWT, "###,###.#0"):        FormatString(9) = "d13"
                PrintValue(10) = Format(rs!CWT, "###,###.#0"):       FormatString(10) = "d13"
                PrintValue(11) = Format(rs!OtherTax, "###,###.#0"):  FormatString(11) = "d13"
                PrintValue(12) = Format(rs!TotDed, "###,###.#0"):    FormatString(12) = "d13"
                PrintValue(13) = Format(rs!Net, "###,###.#0"):       FormatString(13) = "r12"
                PrintValue(14) = " ":                                FormatString(14) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
            
            ' Update Current Totals
            ErnTotals(1, 1) = ErnTotals(1, 1) + rs!Gross
            ErnTotals(1, 2) = ErnTotals(1, 2) + rs!SSTax
            ErnTotals(1, 3) = ErnTotals(1, 3) + rs!Med
            ErnTotals(1, 4) = ErnTotals(1, 4) + rs!FWT
            ErnTotals(1, 5) = ErnTotals(1, 5) + rs!SWT
            ErnTotals(1, 6) = ErnTotals(1, 6) + rs!CWT
            ErnTotals(1, 7) = ErnTotals(1, 7) + rs!OtherTax
            ErnTotals(1, 8) = ErnTotals(1, 8) + rs!TotDed
            ErnTotals(1, 9) = ErnTotals(1, 9) + rs!Net
        
        End If
        
        ' Update QTD Totals
        If rs!InclQTD = True Then
            ErnTotals(2, 1) = ErnTotals(2, 1) + rs!Gross
            ErnTotals(2, 2) = ErnTotals(2, 2) + rs!SSTax
            ErnTotals(2, 3) = ErnTotals(2, 3) + rs!Med
            ErnTotals(2, 4) = ErnTotals(2, 4) + rs!FWT
            ErnTotals(2, 5) = ErnTotals(2, 5) + rs!SWT
            ErnTotals(2, 6) = ErnTotals(2, 6) + rs!CWT
            ErnTotals(2, 7) = ErnTotals(2, 7) + rs!OtherTax
            ErnTotals(2, 8) = ErnTotals(2, 8) + rs!TotDed
            ErnTotals(2, 9) = ErnTotals(2, 9) + rs!Net
        End If


        ' Update YTD Totals

        ErnTotals(3, 1) = ErnTotals(3, 1) + rs!Gross
        ErnTotals(3, 2) = ErnTotals(3, 2) + rs!SSTax
        ErnTotals(3, 3) = ErnTotals(3, 3) + rs!Med
        ErnTotals(3, 4) = ErnTotals(3, 4) + rs!FWT
        ErnTotals(3, 5) = ErnTotals(3, 5) + rs!SWT
        ErnTotals(3, 6) = ErnTotals(3, 6) + rs!CWT
        ErnTotals(3, 7) = ErnTotals(3, 7) + rs!OtherTax
        ErnTotals(3, 8) = ErnTotals(3, 8) + rs!TotDed
        ErnTotals(3, 9) = ErnTotals(3, 9) + rs!Net
        
        LastEmpID = rs!EmpID
        rs.MoveNext
            
        If rs.EOF = False Then
            If rs!EmpID <> LastEmpID Then
                If Ln > MaxLines And frmEarnSumm.chkExcludeDetail = 0 Then
                    FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    Ln = Ln + 1
                    EarnSummaryHdr
                End If
                Ln = Ln + 1
                EarnSummaryPrtTotals
                Ln = Ln + 1
                If ct <> 1 And frmEarnSumm.chkPgEmp Then
                    FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    Ln = Ln + 1
                    EarnSummaryHdr
                End If
            End If
        End If
        
    Loop Until rs.EOF
    
    ' total for last employee
    Ln = Ln + 1
    If Ln >= MaxLines Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        Ln = Ln + 1
        EarnSummaryHdr
    End If
    EarnSummaryPrtTotals
    
    ' grand totals
    EndFlag = True
    EarnSummaryPrtTotals
    Ln = Ln + 1

    MonthlySummary

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
        
End Sub

Public Sub MonthlySummary()
    
    Dim strSQL As String
    Dim mos(12) As String
    mos(1) = "JAN"
    mos(2) = "FEB"
    mos(3) = "MAR"
    mos(4) = "APR"
    mos(5) = "MAY"
    mos(6) = "JUN"
    mos(7) = "JUL"
    mos(8) = "AUG"
    mos(9) = "SEP"
    mos(10) = "OCT"
    mos(11) = "NOV"
    mos(12) = "DEC"
    
    strSQL = ""
    strSQL = strSQL & "select year(CheckDate) as Yr, month(CheckDate) as Mo, sum(Gross) as Gr, sum(SSTax) as SST, sum(MedTax) as MED, sum(FWTTax) as FWT, sum(SWTTax) as SWT, sum(CWTTax) as CWT"
    strSQL = strSQL & " From PRHist"
    strSQL = strSQL & " WHERE PRHist.CheckDate >= " & CLng(frmEarnSumm.BOYDate)
    strSQL = strSQL & " AND PRHist.CheckDate <= " & CLng(frmEarnSumm.EOQDate)
    strSQL = strSQL & " group by year(CheckDate), month(CheckDate)"
    strSQL = strSQL & " order by year(CheckDate), month(CheckDate)"
    
    rsInit strSQL, cn, rs
    If rs.RecordCount = 0 Then Exit Sub
    
    FormFeed
    PageHeader ReportTitle, "", "Payroll Amounts by Year/Month", ""
    
    Ln = Ln + 5
    
    Dim ii As Integer
    Dim Totals(6) As Currency
    Dim Amts(6) As Currency
    For ii = 1 To 6
        Totals(ii) = 0
    Next ii
    
    Dim Titles(8) As String
    Titles(1) = "YEAR"
    Titles(2) = "MONTH"
    Titles(3) = "GROSS"
    Titles(4) = "SS TX"
    Titles(5) = "MED"
    Titles(6) = "FWT"
    Titles(7) = "SWT"
    Titles(8) = "CWT"
    For ii = 1 To 8
        PrintValue(ii) = Titles(ii)
        If ii <= 2 Then
            FormatString(ii) = "a10"
        Else
            FormatString(ii) = "r13"
        End If
    Next ii
    PrintValue(9) = " ":    FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(148, "="):       FormatString(1) = "a148"
    PrintValue(2) = " ":                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 2
    
    rs.MoveFirst
    Do While Not rs.EOF
        
        PrintValue(1) = rs("Yr"):       FormatString(1) = "a10"
        PrintValue(2) = mos(rs("Mo")):       FormatString(2) = "a10"
        
        Amts(1) = rs("Gr")
        Amts(2) = rs("SST")
        Amts(3) = rs("MED")
        Amts(4) = rs("FWT")
        Amts(5) = rs("SWT")
        Amts(6) = rs("CWT")
        For ii = 1 To 6
            Totals(ii) = Totals(ii) + Amts(ii)
            PrintValue(ii + 2) = Format(Amts(ii), "###,###.#0")
            FormatString(ii + 2) = "d13"
        Next ii
        PrintValue(9) = " "
        FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 2
        
        rs.MoveNext
    Loop

    Ln = Ln + 2
    PrintValue(1) = "TOTALS":       FormatString(1) = "a10"
    PrintValue(2) = "":             FormatString(2) = "a10"
    For ii = 1 To 6
        PrintValue(ii + 2) = Format(Totals(ii), "###,###.#0")
        FormatString(ii + 2) = "d13"
    Next ii
    PrintValue(9) = " "
    FormatString(9) = "~"
    FormatPrint

End Sub

Public Sub EarnSummaryHdr()
    PrintValue(1) = "CHK DATE":             FormatString(1) = "a9"
    PrintValue(2) = "CHK #":                FormatString(2) = "a6"
    PrintValue(3) = "DEPT":                 FormatString(3) = "a4"
    PrintValue(4) = "RATE":                 FormatString(4) = "r11"
    PrintValue(5) = "GROSS":                FormatString(5) = "r13"
    PrintValue(6) = "SS TX":                FormatString(6) = "r13"
    PrintValue(7) = "MED":                  FormatString(7) = "r13"
    PrintValue(8) = "FWT":                  FormatString(8) = "r13"
    PrintValue(9) = "SWT":                  FormatString(9) = "r13"
    PrintValue(10) = "CWT":                 FormatString(10) = "r13"
    PrintValue(11) = "OTH TAX":             FormatString(11) = "r13"
    PrintValue(12) = "TOT DED":             FormatString(12) = "r13"
    PrintValue(13) = "NET":                 FormatString(13) = "r13"
    PrintValue(14) = " ":                   FormatString(14) = "~"
    FormatPrint
    Ln = Ln + 1
 
    PrintValue(1) = String(148, "="):       FormatString(1) = "a148"
    PrintValue(2) = " ":                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub


Public Sub EarnSummaryPrtTotals()
        If EndFlag = True Then
            Ln = Ln + 1
            PrintValue(1) = "GRAND TOTALS":         FormatString(1) = "a31"
            PrintValue(2) = " ":                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
        End If
        
        For I = 1 To 3
            If I = 1 Then x = "CURR TOTAL"
            If I = 2 Then x = "QTD TOTAL"
            If I = 3 Then x = "YTD TOTAL"
            PrintValue(1) = x:                              FormatString(1) = "a31"
            For J = 1 To 9
                If EndFlag = True Then
                    PrintValue(J + 1) = GrTotals(I, J):     FormatString(J + 1) = "d13"
                Else
                    PrintValue(J + 1) = ErnTotals(I, J):    FormatString(J + 1) = "d13"
                    GrTotals(I, J) = GrTotals(I, J) + ErnTotals(I, J)
                    ErnTotals(I, J) = 0
                End If
            Next J
    
            PrintValue(11) = " ":                           FormatString(11) = "~"
            FormatPrint
            Ln = Ln + 1
        Next I
        
'
'        For i = 1 To 3
'            For j = 1 To 9
'                GrTotals(i, 9) = GrTotals(i, j) + ErnTotals(i, j)
'                ErnTotals(i, j) = 0
'            Next j
'        Next i

End Sub


