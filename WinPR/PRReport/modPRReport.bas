Attribute VB_Name = "modPRReport"
Option Explicit
Dim ChkEmp As Byte

Dim rs As New ADODB.Recordset
Dim rsYTD As New ADODB.Recordset
Dim trs As New ADODB.Recordset
Dim trsODT As New ADODB.Recordset
Dim trsODTTot As New ADODB.Recordset
Dim trsEntry As New ADODB.Recordset
Dim FormColor As String

Dim PrintFlag As Boolean
Dim LastLine As Byte
Dim xx As String
Dim ItemMax, ItemCount As Byte
Dim FFSW As Byte
Dim RJAmount As String

Dim LastEmpNo, LastEmpID, TempID As Long
Dim LastChkDate As Date
Dim LastEmpName As String
Dim SkipFlag As Boolean

Dim ODTTypeHist As Byte
Dim ODTTypeEE As Byte
Dim ODTTypeDpt As Byte
Dim ODTTypeER As Byte
    
Dim ODTLineHr As Byte
Dim ODTLineOE As Byte
Dim ODTLineDed As Byte

Dim I, J, K As Long
Dim ReportTitle, QtrEnding, w, X, Y, Z As String
Dim CurrDate As Date
Dim StartMonth, EndMonth, LandSw As Byte
Dim FindStr As String
Dim LabelRows, ColumnCount, LRow As Long

Dim BankAddress, BankNumber, WrittenAmount As String
Dim WageGross, TotWageGross, DWageGross, WageSS, TotWageSS, DWageSS As Currency
Dim TaxFed, TotTaxFed, DTaxFed, WageMed, TotWageMed, DWageMed As Currency
Dim TaxSS, TotTaxSS, DTaxSS As Currency
Dim TaxState, TotTaxState, DTaxState As Currency
Dim TaxCity, TotTaxCity, DTaxCity As Currency
Dim TotFICA, FinalFICA, TaxMed, TotTaxMed, DTaxMed As Currency
Dim WageFed, TotWageFed, DWageFed As Currency
Dim WageFic, TotWageFic, DWageFic As Currency
Dim TipsFic, TotTipsFic, DTipsFic As Currency
Dim TipsMed, TotTipsMed, DTipsMed As Currency
Dim WageState, TotWageState, DWageState As Currency
Dim WageCity, TotWageCity, DWageCity As Currency
Dim QTDWageGross, TotQTDWageGross, DQTDWageGross As Currency
Dim YTDWageGross, TotYTDWageGross, DYTDWageGross As Currency
Dim HistCount As Long
Dim DepSTUnempAmt, DepSTUnempMatch, DepSTUnempPct As Currency
Dim DepFedUnempAmt, DepFedUnempPct, DepFedUnempMatch As Currency

Dim OECount, DEDCount, LineCount As Integer
Dim ODTFlag As Boolean

Dim DirDepEE, DirDepTl, CheckTotal As Currency
Dim DirDepFlag As Boolean
Dim BtchHeader As String

Dim Hdr1, Hdr2, Hdr3, Hdr4 As String
Dim HourTotal, WageTotal, TaxTotal As Currency

' check recon report
Dim RecCount As Long
Dim LastType, PrintNum As Byte
Dim DepoAmt, TotAmt, CheckAmt As Currency
Dim NoRecords As Integer

' dir dep report
Dim OrderType As Byte
Dim LineCt, SubLineCt As Integer
Dim LastBatch As Long
Dim TotCredAmt, SubCredAmt, DebitTotal, DepositTotal As Currency
Dim SeqNo, ChkDigNo, ChkSve, RteNo As Long
Dim FedID, HashString As String
Dim BlockCt, WriteCt As Long
Dim TChannel As Variant
Dim Hash As Double

Dim DeptID, DeptNum As Long

' qtrly reports
Dim SSString, NameString, Frmt As String
Dim AmtLen As Long
Dim CustLen, LnCnt, CurrPg, NumPages, PadNumber, PadSpaces As Long
Dim qPadString, StateAbbrev As String

' ======= 941 Form =======
Dim TelFmtString, FmtString As String
' ======= 941 Form =======
Dim ddPEDate, ddCheckDate As Date

' ======= YE city tax report =======
Dim LastCityName As String
Dim LastCityNumber As Long
Dim SYTDGROSS As Currency
Dim SYTDTAX As Currency
Dim TYTDGross As Currency
Dim TYTDTAX As Currency
Dim StartYM As Long
Dim EndYM As Long
Dim CityName As String
Dim CityNumber As Long
Dim LastCityID As Long
' ======= YE city tax report =======

Dim RptTitle As String
Dim ct As Long

' ********* 1099 Processing ***************

Dim FormCount As Byte

Dim Misc1099_Box1 As Currency
Dim Misc1099_Box2 As Currency
Dim Misc1099_Box3 As Currency
Dim Misc1099_Box4 As Currency
Dim Misc1099_Box5 As Currency
Dim Misc1099_Box6 As Currency
Dim Misc1099_Box7 As Currency
Dim Misc1099_Box8 As Currency
Dim Misc1099_Box9 As String
Dim Misc1099_Box10 As Currency
Dim Misc1099_Box13 As Currency
Dim Misc1099_Box14 As Currency
Dim Misc1099_Box15a As Currency
Dim Misc1099_Box15b As Currency
Dim Misc1099_Box16a As Currency
Dim Misc1099_Box16b As Currency
Dim Misc1099_Box17a As String
Dim Misc1099_Box17b As String
Dim Misc1099_Box18a As Currency
Dim Misc1099_Box18b As Currency

Dim Misc1099(21) As String

Dim PayerName As String
Dim PayerAddr1 As String
Dim PayerAddr2 As String
Dim PayerAddr3 As String
Dim PayerAddr4 As String
Dim PayerID As String

Dim PayeeName As String
Dim PayeeAddr1 As String
Dim PayeeAddr2 As String
Dim PayeeAddr3 As String
Dim PayeeID As String

' OH BUC 2010
Dim ThisPage_Count As Byte
Dim ThisPage_Amount As Currency
Dim StartPageNum As Byte
Dim PageNum As Integer
' ********* 1099 Processing ***************

Public Sub PR1099(ByVal jTaxYear As String)

    With frm1099.rs
    
        If .RecordCount = 0 Then Exit Sub
        
        PrtInit ("Port")    ' "Port" = Portrait
        SetFont 10, Equate.Portrait
        
        ' init the payer info
        PayerName = PRCompany.Name
        PayerAddr1 = PRCompany.Address1
        
        If PRCompany.Address2 <> "" Then
            PayerAddr2 = PRCompany.Address2
            PayerAddr3 = PRCompany.CSZ
        Else
            PayerAddr2 = PRCompany.CSZ
            PayerAddr3 = ""
        End If
        PayerAddr4 = ""
        
        PayerID = PRCompany.FederalID
    
        FormCount = 0
    
        ' init the box string array
        For I = 1 To 21
            If I = 9 Then
                Misc1099(I) = ""
            Else
                Misc1099(I) = PadRight(Format(0, "##,###,##0.00"), 13)
            End If
        Next I
        
        .MoveFirst
        Do
        
            If !Select = True Then
            
                If PREmployee.GetByID(!EmployeeID) = False Then
                    MsgBox "Employee ID not found! " & !EmployeeID, vbExclamation
                    GoBack
                End If
                
                PayeeName = PREmployee.FLName
                
                PayeeAddr1 = PREmployee.Address1
                PayeeAddr2 = PREmployee.Address2
                'PayeeAddr1 = "addr1"
                'PayeeAddr2 = ""
                
                PayeeAddr3 = PREmployee.CSZ
                
                PayeeID = PREmployee.SSString
                
                Misc1099(1) = PadRight(Format(!Amount, "##,###,##0.00"), 13)
        
                Print1099MISC
            End If
            
            .MoveNext
        Loop Until .EOF
    
    End With
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub Print1099MISC()

    ' 2022-01-19 - 3 forms per page

    Dim fmtNm As String
    fmtNm = "a40"
    
    Dim FormSp, VertSp As Long
    Dim xPos, yPos As Long
    Dim FormDiff As Long
    
    ' name/id fields printed first
    ' then box fields
    FormDiff = 2500      ' space between forms
    VertSp = 200        ' line feed value
    
    FormCount = FormCount + 1
    yPos = 720 + ((FormCount - 1) * 5300)
    
    xPos = 1000
    
    ' print the payer info
    For I = 1 To 5
        If I = 1 Then X = PayerName
        If I = 2 Then X = PayerAddr1
        If I = 3 Then X = PayerAddr2
        If I = 4 Then X = PayerAddr3
        If I = 5 Then X = PayerAddr4
        PosPrint xPos, yPos + 40, X
        yPos = yPos + VertSp
    Next I
    
    ' print the ID numbers & Box 1 NEC
    yPos = yPos + 400
    PosPrint xPos, yPos, PayerID
    PosPrint xPos + 2600, yPos, PayeeID
    xPos = 5600
    PosPrint xPos, yPos, Misc1099(1)
    
    ' print the payee info
    yPos = yPos + 870
    xPos = 1000
    PosPrint xPos, yPos, PayeeName
    
    ' single address field
    yPos = yPos + VertSp * 2
    PosPrint 1000, yPos, Left(Trim(PayeeAddr1) & " " & Trim(PayeeAddr2) & " " & String(45, " "), 45)
    
    yPos = yPos + VertSp
    xPos = 5600
    PosPrint xPos, yPos, Misc1099(4)
    
    yPos = yPos + VertSp
    xPos = 1000
    PosPrint xPos, yPos, PayeeAddr3
    
    If FormCount = 3 Then
        Prvw.vsp.NewPage
        FormCount = 0
    End If

End Sub

Public Sub CityList()
     
    frmCityList.Hide
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    
    PrtInit ("Port")    ' "Port" = Portrait
    
    ' set up SQL statement based upon order requested
    ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NO."
    If frmCityList.optNumber Then
        ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NO."
        SQLString = "SELECT * FROM PRCITY ORDER BY CityNumber"
    Else
        ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NAME"
        SQLString = "SELECT * FROM PRCITY ORDER BY CityName"
    End If
    
    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If Not PRCity.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbExclamation, "Payroll Rate File Listing"
        Exit Sub
    End If
    Do
        If Ln = 0 Or Ln > MaxLines Then
         
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            SetFont 10, Equate.Portrait
            
            Ln = Ln + 2                ' Changed from Ln +1 to Ln + 2
            
            PrintValue(1) = "City Num":                             FormatString(1) = "a9"
            PrintValue(2) = " ":                                    FormatString(2) = "a3"
            PrintValue(3) = "City Name":                            FormatString(3) = "a30"
            PrintValue(4) = " ":                                    FormatString(4) = "a4"
            PrintValue(5) = "City Tax Rate":                        FormatString(5) = "a13"
            PrintValue(6) = " ":                                    FormatString(6) = "~"
            FormatPrint
            Ln = Ln + 1
             
            PrintValue(1) = String(94, "-"):                        FormatString(1) = "a94"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
        End If
    
        PrintValue(1) = PRCity.CityNumber:                          FormatString(1) = "n5"
        PrintValue(2) = " ":                                        FormatString(2) = "a7"
        PrintValue(3) = PRCity.CityName:                            FormatString(3) = "a30"
        PrintValue(4) = " ":                                        FormatString(4) = "a6"
        PrintValue(5) = PRCity.CityRate:                            FormatString(5) = "d8"
        PrintValue(6) = " ":                                        FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 1
        
        If Not PRCity.GetNext Then
            Exit Do
        End If
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub EEList(ByVal ReportType As String)
Dim ReportTitle As String
Dim FDept As String
Dim LabelColumns, MaxLabelRows As Integer
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
            SetFont 9, Equate.Portrait
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
    ElseIf frmLists.optDept Then
        If ReportTitle <> "labels" Then
            ReportTitle = Trim(ReportTitle) & " BY DEPT BY NAME"
        End If
        SQLString = Trim(SQLString) & " ORDER BY DepartmentID, LastName, FirstName"
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
    
        If Not PRW4.GetByEmployeeID(PREmployee.EmployeeID) Then
            PRW4.Clear
            PRW4.EmployeeID = PREmployee.EmployeeID
            PRW4.Save (Equate.RecAdd)
        End If
        
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
                    PrintValue(11) = "STATE":                       FormatString(11) = "a5"
                    PrintValue(12) = " ":                           FormatString(12) = "a5"
                    PrintValue(13) = "ZIP":                         FormatString(13) = "a8"
                                                                                
                 '*** Print SS Number?
                    If frmLists.chkSSN Then
                        PrintValue(14) = "SS NUMBER ":              FormatString(14) = "a9"
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
                    PrintValue(25) = " ":                           FormatString(25) = "a5"
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
                    
                    PrintValue(1) = String(140, "="):               FormatString(1) = "a140"
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
                    PrintValue(2) = "EMPLOYEE NAME":                FormatString(2) = "a37"
                    PrintValue(3) = "DEPT":                         FormatString(3) = "a7"
                    PrintValue(4) = "DEPT NAME":                    FormatString(4) = "a17"
                    PrintValue(5) = "RATE":                         FormatString(5) = "a10"
                    PrintValue(6) = "SALARY":                       FormatString(6) = "a6"
                    PrintValue(7) = " ":                            FormatString(7) = "~"
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
                PrintValue(3) = " " & PRDepartment.Name:            FormatString(3) = "a30"
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
                PrintValue(14) = " ":                               FormatString(14) = "a8"
                PrintValue(15) = PREmployee.ZipCode:                FormatString(15) = "a5"

                 '*** Print SS Number?
                If frmLists.chkSSN Then
                    PrintValue(16) = " ":                           FormatString(16) = "a3"
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
                
                Dim PrtNum As Integer
                If PRW4.FilingType = PREquate.PRW4Standard Then
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
                    
                    PrtNum = 9
                
                Else
                    
                    If PRW4.FilingType = PREquate.PRW4Single Then PrintValue(2) = "FWT W4 S"
                    If PRW4.FilingType = PREquate.PRW4Married Then PrintValue(2) = "FWT W4 M"
                    If PRW4.FilingType = PREquate.PRW4HOH Then PrintValue(2) = "FWT W4 HOH"
                    If PRW4.TwoJobs <> 0 Then PrintValue(2) = PrintValue(2) & "*2"
                    PrintValue(2) = PrintValue(2) & " "
                    FormatString(2) = "a14"
                    
                    PrtNum = 2
                    
                    If PRW4.Dependents <> 0 Then
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = "#Dep:" & PRW4.Dependents & " "
                        FormatString(PrtNum) = "a9"
                    End If
                        
                    If PRW4.DependentsOther <> 0 Then
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = "#Dep Other:" & PRW4.DependentsOther & " "
                        FormatString(PrtNum) = "a14"
                    End If
                    
                    If PRW4.OtherIncome <> 0 Then
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = "Other Inc:": FormatString(PrtNum) = "a10"
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = PRW4.OtherIncome: FormatString(PrtNum) = "d10"
                    End If
                    
                    If PRW4.Deductions <> 0 Then
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = "Deductions:": FormatString(PrtNum) = "a11"
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = PRW4.Deductions: FormatString(PrtNum) = "d10"
                    End If
                    
                    If PRW4.ExtraWH <> 0 Then
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = "Extra WH:": FormatString(PrtNum) = "a9"
                        PrtNum = PrtNum + 1
                        PrintValue(PrtNum) = PRW4.ExtraWH: FormatString(PrtNum) = "d8"
                    End If
                
                End If
                
                PrtNum = PrtNum + 1
                If PREmployee.SWTMarried = 1 Then
                    PrintValue(PrtNum) = "SWT Married: Y":                  FormatString(PrtNum) = "a14"
                Else
                    PrintValue(PrtNum) = "SWT Married: N":                  FormatString(PrtNum) = "a14"
                End If
                
                PrtNum = PrtNum + 1
                PrintValue(PrtNum) = " ":                                   FormatString(PrtNum) = "a2"
                If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "SWT Exemps: ":                    FormatString(PrtNum) = "a12"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = PREmployee.SWTAmount:              FormatString(PrtNum) = "n2"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = " ":                               FormatString(PrtNum) = "a2"
                ElseIf PREmployee.SWTBasis = PREquate.BasisPercent Then
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "SWT: ":                           FormatString(PrtNum) = "a5"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = PREmployee.SWTAmount:              FormatString(PrtNum) = "d9"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "%  ":                             FormatString(PrtNum) = "a3"
                End If
                
                If PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "SWT Extra: ":                     FormatString(PrtNum) = "a11"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = PREmployee.SWTExtraAmount:         FormatString(PrtNum) = "d6"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "%":                               FormatString(PrtNum) = "a1"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = " ":                               FormatString(PrtNum) = "~"
                ElseIf PREmployee.SWTExtraBasis = PREquate.BasisAmount Then
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "SWT Extra: ":                     FormatString(PrtNum) = "a11"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = "$ " & PREmployee.SWTExtraAmount:  FormatString(PrtNum) = "d8"
                    PrtNum = PrtNum + 1
                    PrintValue(PrtNum) = " ":                               FormatString(PrtNum) = "~"
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
                If Trim(PREmployee.RaceCode) <> "" Then
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
                 PrintValue(2) = PREmployee.LFName:                     FormatString(2) = "a37"
                 PrintValue(3) = RTrim(PRDepartment.DepartmentNumber):  FormatString(3) = "n4"
                 PrintValue(4) = " - ":                                 FormatString(4) = "a3"
                 PrintValue(5) = Trim(PRDepartment.Name):               FormatString(5) = "a8"
                 If PREmployee.Salaried = 1 Then
                     PrintValue(6) = PREmployee.SalaryAmount:           FormatString(6) = "d14"
                 Else
                     PrintValue(6) = PREmployee.HourlyAmount:           FormatString(6) = "d10"
                 End If
                 PrintValue(7) = " ":                                   FormatString(7) = "a5"
                 If PREmployee.Salaried = 1 Then
                     PrintValue(8) = "SALARY":                          FormatString(8) = "a6"
                 Else
                     PrintValue(8) = "HOURLY":                          FormatString(8) = "a6"
                 End If
                 PrintValue(9) = " ":                                   FormatString(9) = "~"
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
                 
                 If PRW4.FilingType = PREquate.PRW4Standard Then
                    If PREmployee.FWTMarried = 1 Then
                       PrintValue(20) = "Y":                               FormatString(20) = "a3"
                    Else
                       PrintValue(20) = "N":                               FormatString(20) = "a3"
                    End If
                 
                    PrintValue(21) = PREmployee.FWTBasis:                  FormatString(21) = "n3"
                    PrintValue(22) = PREmployee.FWTAmount:                 FormatString(22) = "d9"
                    PrintValue(23) = Format(PREmployee.FWTExtraAmount):    FormatString(23) = "d9"
                 Else
                    If PRW4.FilingType = PREquate.PRW4Married Then
                       PrintValue(20) = "*Y":                               FormatString(20) = "a3"
                    ElseIf PRW4.FilingType = PREquate.PRW4Single Then
                       PrintValue(20) = "*S":                               FormatString(20) = "a3"
                    Else
                       PrintValue(20) = "*H":                               FormatString(20) = "a3"
                    End If
                 
                    PrintValue(21) = PRW4.Dependents:                  FormatString(21) = "n3"
                    PrintValue(22) = "":                 FormatString(22) = "a9"
                    PrintValue(23) = Format(PRW4.ExtraWH):    FormatString(23) = "d9"
                 End If
                 
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
                    MaxLabelRows = 10
                    MaxLines = 60
                    LabelColumns = 1
                    If NoLabels = 0 Then
                       ' GetDeptInfo (PREmployee.DepartmentID)
                       LabelString(1, 1) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a37"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                    
                       LabelString(1, 1) = Left(RTrim(PREmployee.FLName), 37)
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a37"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = "PERIOD ENDING DATE: " & frmLists.tdbPEDate
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a37"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                                             
                       LabelString(1, 1) = Left("DEPT:" & PRDepartment.DepartmentNumber & "-" & RTrim(PRDepartment.Name), 37)
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a37"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 3
                       
                    ElseIf NoLabels = 1 Then
                       LabelColumns = 2
                       ColumnCount = ColumnCount + 1
                       'GetDeptInfo (PREmployee.DepartmentID)
                       Label2String(1, ColumnCount) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       Label2String(2, ColumnCount) = Left(RTrim(PREmployee.FLName), 37)
                       Label2String(3, ColumnCount) = "PERIOD ENDING DATE: " & frmLists.tdbPEDate
                       Label2String(4, ColumnCount) = Left("DEPT:" & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name), 37)
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 2
                          For LRow = 1 To 4
                              ColumnCount = 0
                              PrintValue(1) = Label2String(LRow, 1):    FormatString(1) = "a37"
                              PrintValue(2) = Label2String(LRow, 2):    FormatString(2) = "a37"
                              PrintValue(3) = Label2String(LRow, 3):    FormatString(3) = "a37"
                              PrintValue(4) = Label2String(LRow, 4):    FormatString(4) = "a37"
                              PrintValue(5) = " ":                      FormatString(5) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
                       End If
                    ElseIf NoLabels = 2 Then
                       LabelColumns = 3
                       ColumnCount = ColumnCount + 1
                       'GetDeptInfo (PREmployee.DepartmentID)
                       Label2String(1, ColumnCount) = "EMPLOYEE # : " & PREmployee.EmployeeNumber
                       Label2String(2, ColumnCount) = Left(RTrim(PREmployee.FLName), 37)
                       Label2String(3, ColumnCount) = "P/E DATE: " & frmLists.tdbPEDate
                       Label2String(4, ColumnCount) = Left("DEPT:" & PRDepartment.DepartmentNumber & "-" & RTrim(PRDepartment.Name), 37)
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 2
                          For LRow = 1 To 4
                              ColumnCount = 0
                              PrintValue(1) = Label2String(LRow, 1):    FormatString(1) = "a37"
                              PrintValue(2) = Label2String(LRow, 2):    FormatString(2) = "a37"
                              PrintValue(3) = Label2String(LRow, 3):    FormatString(3) = "a37"
                              PrintValue(4) = Label2String(LRow, 4):    FormatString(4) = "a37"
                              PrintValue(5) = " ":                      FormatString(5) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
                       End If
                    End If
                Case "MailingLabels"
                    MaxLabelRows = 10
                    MaxLines = 60
                    
                    LabelColumns = 1
                    If NoLabels = 0 Then
                       LabelString(1, 1) = PREmployee.FLName
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a35"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = PREmployee.Address1
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a35"
                       PrintValue(2) = " ":                             FormatString(2) = "~"
                       FormatPrint
                       Ln = Ln + 1
                       
                       LabelString(1, 1) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                       PrintValue(1) = LabelString(1, 1):               FormatString(1) = "a35"
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
                           
                          For LRow = 1 To 3
                              ColumnCount = 0
                              PrintValue(1) = LabelString(LRow, 1):     FormatString(1) = "a35"
                              PrintValue(2) = LabelString(LRow, 2):     FormatString(2) = "a35"
                              PrintValue(3) = LabelString(LRow, 3):     FormatString(3) = "a35"
                              PrintValue(4) = " ":                      FormatString(4) = "~"
                              FormatPrint
                              Ln = Ln + 2
                              LabelRows = LabelRows + 1
                          Next LRow
                       End If
                       
                   ElseIf NoLabels = 2 Then
                       LabelColumns = 3
                       ColumnCount = ColumnCount + 1
                       LabelString(1, ColumnCount) = PREmployee.FLName
                       LabelString(2, ColumnCount) = PREmployee.Address1
                       LabelString(3, ColumnCount) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                       
                       If ColumnCount = LabelColumns Then
                          Ln = Ln + 3
                           
                          For LRow = 1 To 3
                              ColumnCount = 0
                              PrintValue(1) = LabelString(LRow, 1):     FormatString(1) = "a35"
                              PrintValue(2) = LabelString(LRow, 2):     FormatString(2) = "a35"
                              PrintValue(3) = LabelString(LRow, 3):     FormatString(3) = "a35"
                              PrintValue(4) = " ":                      FormatString(4) = "~"
                              FormatPrint
                              Ln = Ln + 1
                              LabelRows = LabelRows + 1
                          Next LRow
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
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         ElseIf ColumnCount = 2 Then    '(Row, Col)
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a35"
               PrintValue(2) = LabelString(1, 2):                       FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a35"
               PrintValue(2) = LabelString(2, 2):                       FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a35"
               PrintValue(2) = LabelString(3, 2):                       FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
         End If
      
      ElseIf NoLabels = 1 Then
         If ColumnCount > 0 Then
            Ln = Ln + 2
               PrintValue(1) = LabelString(1, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = LabelString(2, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = LabelString(3, 1):                       FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         End If
      End If
   Case "TimeCardLabels"
      If NoLabels = 2 Then
         Ln = Ln + 3
         If ColumnCount = 1 Then
            
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                                    
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         ElseIf ColumnCount = 2 Then
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a35"
               PrintValue(2) = Label2String(1, 2):                      FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a35"
               PrintValue(2) = Label2String(2, 2):                      FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a35"
               PrintValue(2) = Label2String(3, 2):                      FormatString(2) = "a35"
               PrintValue(3) = " ":                                     FormatString(3) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a35"
               PrintValue(2) = Label2String(4, 2):                      FormatString(2) = "a35"
               PrintValue(3) = " "
               FormatString(3) = "~"
               FormatPrint
         End If
      
      ElseIf NoLabels = 1 Then
         If ColumnCount > 0 Then
            Ln = Ln + 2
               PrintValue(1) = Label2String(1, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
               
               PrintValue(1) = Label2String(2, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(3, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
               Ln = Ln + 1
                  
               PrintValue(1) = Label2String(4, 1):                      FormatString(1) = "a35"
               PrintValue(2) = " ":                                     FormatString(2) = "~"
               FormatPrint
         End If
      End If
 End Select
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub NewHireReport()
    
Dim HDate As Date
Dim FFlag As Boolean
    
    frmNewHire.Hide
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.lblMsg2 = ""
    frmProgress.Show
    
    PrtInit ("Port")    ' "Port" = Portrait
    
    ' set up SQL statement based upon order requested
    ReportTitle = ""
        
'    SQLString = "SELECT * FROM PREmployee WHERE (PREmployee.DateHired >= " & CLng(StartDate) & _
'                " AND PREmployee.DateHired <= " & CLng(EndDate) & ")" & _
'                " OR (PREmploy"
'
'                " ORDER BY PREmployee.EmployeeNumber"

    ' go thru all employees and then filter by Hire or Recall date
    SQLString = "SELECT * FROM PREmployee WHERE Inactive = 0 ORDER BY EmployeeNumber"

    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Found !!!", vbExclamation, "State of Ohio New Hire Reporting Form 7048"
        GoBack
    End If
            
    FFlag = True
            
    Do

        If PREmployee.DateLastRecall > PREmployee.DateHired Then
            HDate = PREmployee.DateLastRecall
        Else
            HDate = PREmployee.DateHired
        End If
        
        If frmNewHire.cmbState.text <> "All" Then
            If LCase(PREmployee.State) <> LCase(frmNewHire.cmbState.text) Then
                GoTo NewHireNext
            End If
        End If
        
        If HDate >= StartDate And HDate <= EndDate Then
        
            If FFlag = False Then FormFeed
            FFlag = False
            
            Ln = Ln + 4
            
            If frmNewHire.cmbState.text = "OH" Then
                PrintValue(1) = "STATE OF OHIO NEW HIRE REPORTING FORM 7048":       FormatString(1) = "a90"
            Else
                PrintValue(1) = "NEW HIRE REPORT":       FormatString(1) = "a90"
            End If
            
            PrintValue(2) = " ":                                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 4
            
            PrintValue(1) = "E M P L O Y E E   I N F O R M A T I O N":          FormatString(1) = "a50"
            PrintValue(2) = " ":                                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 4
            
            PrintValue(1) = "SOCIAL SECURITY NUMBER:":                          FormatString(1) = "a26"
            PrintValue(2) = PREmployee.SSString:                                FormatString(2) = "a11"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            PrintValue(1) = "NAME:":                                            FormatString(1) = "a26"
            PrintValue(2) = PREmployee.FLName:                                  FormatString(2) = "a30"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            If Trim(PREmployee.Address2) = "" Then
                PrintValue(1) = "ADDRESS:":                                     FormatString(1) = "a26"
            Else
                PrintValue(1) = "ADDRESS 1:":                                   FormatString(1) = "a26"
            End If
            PrintValue(2) = PREmployee.Address1:                                FormatString(2) = "a30"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            If Trim(PREmployee.Address2) <> "" Then
                PrintValue(1) = "ADDRESS 2:":                                   FormatString(1) = "a26"
                PrintValue(2) = PREmployee.Address2:                            FormatString(2) = "a30"
                PrintValue(3) = " ":                                            FormatString(3) = "~"
                FormatPrint
                Ln = Ln + 2
            End If
            
            PrintValue(1) = "CITY/STATE/ZIP:":                                  FormatString(1) = "a26"
            PrintValue(2) = Trim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
            FormatString(2) = "a90"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            PrintValue(1) = "EMPLOYEE DATE OF HIRE:":                           FormatString(1) = "a26"
            PrintValue(2) = Format(HDate, "mm/dd/yyyy"):         FormatString(2) = "a10"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            PrintValue(1) = "DATE OF BIRTH:":                                   FormatString(1) = "a26"
            If PREmployee.DateOfBirth = 0 Then
                PrintValue(2) = Format(PREmployee.DateOfBirth, "00/00/0000"):   FormatString(3) = "a10"
            Else
                PrintValue(2) = Format(PREmployee.DateOfBirth, "mm/dd/yyyy"):   FormatString(3) = "a10"
            End If
            PrintValue(4) = " ":                                                FormatString(4) = "~"
            FormatPrint
            Ln = Ln + 5
                     
            PrintValue(1) = "E M P L O Y E R   I N F O R M A T I O N":          FormatString(1) = "a50"
            PrintValue(2) = " ":                                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 4
            
            PrintValue(1) = "EMPLOYER FEDERAL EIN:":                            FormatString(1) = "a26"
            PrintValue(2) = PRCompany.FederalID:                                FormatString(2) = "a11"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            PrintValue(1) = "EMPLOYER NAME:":                                   FormatString(1) = "a26"
            PrintValue(2) = PRCompany.Name:                                     FormatString(2) = "a30"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            If Trim(PRCompany.Address2) = "" Then
                PrintValue(1) = "ADDRESS:":                                     FormatString(1) = "a26"
            Else
                PrintValue(1) = "ADDRESS 1:":                                   FormatString(1) = "a26"
            End If
            
            PrintValue(2) = PRCompany.Address1:                                 FormatString(2) = "a30"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            If Trim(PRCompany.Address2) <> "" Then
                PrintValue(1) = "ADDRESS 2:":                                   FormatString(1) = "a26"
                PrintValue(2) = PRCompany.Address2:                             FormatString(2) = "a30"
                PrintValue(3) = " ":                                            FormatString(3) = "~"
                FormatPrint
                Ln = Ln + 2
            End If
            
            PrintValue(1) = "CITY/STATE/ZIP:":                                  FormatString(1) = "a26"
            
            SQLString = "SELECT * from PRState where StateID = " & PRCompany.AddrStateID
            
            If Not PRState.GetBySQL(SQLString) Then
                MsgBox "Employer state not found !!!", vbExclamation, "State of Ohio New Hire Reporting Form 7048"
                GoBack
            End If
            
            PrintValue(2) = Trim(PRCompany.City) & ", " & PRState.StateAbbrev & "  " & PRCompany.ZipCode
            FormatString(2) = "a60"
            
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
            
            PrintValue(1) = "DATE: ":                                           FormatString(1) = "a26"
            PrintValue(2) = Format(Date, "mm/dd/yyyy"):                         FormatString(2) = "a30"
            PrintValue(3) = " ":                                                FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 2
                        
        End If
        
NewHireNext:
        If Not PREmployee.GetNext Then
            Exit Do
        End If
    
    Loop
            
    If FFlag = True Then
        MsgBox "No Employees found for the hire date range", vbInformation
        GoBack
    End If
            
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub QtrRpts(ByVal ReportList As String, _
                   ByVal QtrEnding As String, _
                   ByVal StateID As Long)

Dim StateString, ReportTitle As String
Dim BOY, YM1, YM2 As Long
Dim DataType As Byte
Dim RecType As Byte
Dim ID As Long
Dim FedUnempPct, StateUnempPct As Currency
Dim FedUnempAmt, StateUnempAmt As Currency
Dim FedUnempMax, StateUnempMax As Currency
Dim uFlag As Boolean

    PRTotal.CreateRS
    
    frmPRQtrlyRpts.Hide
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    FFSW = 0
    
    ' get State if assigned
    If StateID <> 0 Then
        If Not PRState.GetByID(StateID) Then
            MsgBox "StateID NF: " & StateID, vbExclamation
            GoBack
        End If
        StateString = "REPORTING FOR: " & PRState.StateName
        StateUnempMax = PRState.UnEmpMax
    Else
        StateString = "REPORTING FOR: ALL STATES"
    End If
    
    ' max wages
    FedUnempMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, frmPRQtrlyRpts.cmbYear)
    
    ' gather the data
    YM1 = frmPRQtrlyRpts.cmbYear * 100
    If frmPRQtrlyRpts.cmbQtr = 1 Then YM1 = YM1 + 1
    If frmPRQtrlyRpts.cmbQtr = 2 Then YM1 = YM1 + 4
    If frmPRQtrlyRpts.cmbQtr = 3 Then YM1 = YM1 + 7
    If frmPRQtrlyRpts.cmbQtr = 4 Then YM1 = YM1 + 10
    YM2 = YM1 + 2
    
    ' beginning of year - YYYYMM
    BOY = Int(YM1 / 100) * 100 + 1
    
    ' page set up based on the report listing selection
    ' get YTD for the unemployment report
    If ReportList = "QtrlyFedUnemp" Then
        ' get the percentages
        ' Federal from PRGlobal
        FedUnempPct = PRGlobal.GetAmount(PREquate.GlobalTypeFUNPct, CLng(Int(YM1 / 100)))
        ' State from PRCompany
        StateUnempPct = PRCompany.StateUnempPct
        SQLString = "SELECT * FROM PRHist WHERE YearMonth >= " & BOY & _
                    " AND YearMonth <= " & YM2
    Else
        SQLString = "SELECT * FROM PRHist WHERE YearMonth >= " & YM1 & _
                    " AND YearMonth <= " & YM2
    End If
    
    If StateID <> 0 Then
        SQLString = Trim(SQLString) & " AND PRHist.StateID = " & StateID
    End If
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No payroll history found!!", vbExclamation
        GoBack
    End If
    
    ct = 0
    Do
        
        ct = ct + 1
        If ct Mod 10 = 1 Then
            frmProgress.lblMsg2 = "Now Processing Record # " & Format(ct, "#,###,##0") & vbCr & _
                                  " Of: " & PRHist.Records
            frmProgress.Refresh
        End If
        
        ' skip 1099 employees
        If PREmployee.GetByID(PRHist.EmployeeID) Then
            If PREmployee.x1099Employee <> 0 Then
                GoTo NxtPRHist
            End If
        End If
        
        ' use "hard coded" 1/2/3 so when sorted
        ' it will print ee/Dept/Company
        For DataType = 1 To 3
            
            If DataType = 1 Then        ' employee info
                ID = PRHist.EmployeeID
            ElseIf DataType = 2 Then    ' dept info
                ID = PRHist.DepartmentID
            ElseIf DataType = 3 Then    ' company info
                ID = 99999
            End If
        
            ' skip if DeptID = 0 / always do comp totals
            If ID <> 0 Then
                
                ' create new total record???
                If Not PRTotal.tFind(DataType, ID) Then
                    PRTotal.Clear
                    PRTotal.RecType = DataType
                    PRTotal.RecID = ID
                    PRTotal.Save (Equate.RecAdd)
                End If
            
                If PRHist.YearMonth >= YM1 And PRHist.YearMonth <= YM2 Then
                    PRTotal.Gross = PRTotal.Gross + PRHist.Gross
                    PRTotal.SSWageBase = PRTotal.SSWageBase + PRHist.SSWageBase
                    PRTotal.SSWage = PRTotal.SSWage + PRHist.SSWage
                    PRTotal.SSTax = PRTotal.SSTax + PRHist.SSTax
                    PRTotal.MEDWage = PRTotal.MEDWage + PRHist.MEDWage
                    PRTotal.MedTax = PRTotal.MedTax + PRHist.MedTax
                    PRTotal.FWTWage = PRTotal.FWTWage + PRHist.FWTWage
                    PRTotal.FWTTax = PRTotal.FWTTax + PRHist.FWTTax
                    PRTotal.StateWage = PRTotal.StateWage + PRHist.SWTWage
                    PRTotal.StateTax = PRTotal.StateTax + PRHist.SWTTax
                    PRTotal.CityWage = PRTotal.CityWage + PRHist.CWTWage
                    PRTotal.CityTax = PRTotal.CityTax + PRHist.CWTTax
                    PRTotal.FUNWageBase = PRTotal.FUNWageBase + PRHist.FUNWageBase
                    PRTotal.FUNWage = PRTotal.FUNWage + PRHist.FUNWage
                    PRTotal.SUNWageBase = PRTotal.SUNWageBase + PRHist.SUNWageBase
                    PRTotal.SUNWage = PRTotal.SUNWage + PRHist.SUNWage
                End If
                
                PRTotal.YTDGross = PRTotal.YTDGross + PRHist.Gross
                PRTotal.YTDFUNWageBase = PRTotal.YTDFUNWageBase + PRHist.FUNWageBase
                PRTotal.YTDSUNWageBase = PRTotal.YTDSUNWageBase + PRHist.SUNWageBase
                        
                PRTotal.Save (Equate.RecPut)
            
            End If
        
        Next DataType
        
NxtPRHist:
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    Select Case ReportList
        Case "QtrlyFICAFWT"
            Pg = 0
            ReportTitle = "PAYROLL QUARTERLY FICA AND FWT REPORT"
            StateString = ""
        Case "QtrlyStateCity"
            Pg = 0
            ReportTitle = "PAYROLL QUARTERLY STATE AND CITY REPORT"
        Case "QtrlyFedUnemp"
            Pg = 0
            ReportTitle = "PAYROLL QUARTERLY FEDERAL UNEMPLOYMENT REPORT"
        Case "QtrlyTipsTaxes"
            Pg = 0
            ReportTitle = "PAYROLL QUARTERLY TIPS AND TAXES REPORT"
            StateString = ""
    End Select
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
                             
    ' print order ???
    ' loop the employee records
    PRTotal.TSortByString "RecType, RecId"
    If PRTotal.FindFirst = False Then
    End If
    
    Ln = 0
    
    Do
        
        If Ln = 0 Or Ln > MaxLines Or (LastType <> PRTotal.RecType And MaxLines - Ln < 8) Then
                                    
            SetFont 8, Equate.Portrait
            Columns = 115
            If Ln Then FormFeed
            Ln = Ln + 2
            Msg1 = QtrEnding
            PageHeader ReportTitle, Msg1, StateString, ""

            ' data header
            Ln = Ln + 2                ' Changed from Ln +1 to Ln + 2
                                 
            If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
            Else
                MsgBox "Employer State Not Found !!!", vbExclamation, "Quarterly Reports"
            End If

            QtrRptsHeaders (ReportList) ' Print Report Header
        End If
 
        '=============================================================================================
        '==========================      PRINT REPORT DETAIL     ===================================
        '=============================================================================================
    
        ' get the ee/dpt/er record
        If PRTotal.RecType = 1 Then
            If Not PREmployee.GetByID(PRTotal.RecID) Then
                MsgBox "Employee NF: " & PRTotal.RecID, vbExclamation
                GoBack
            End If
            SSString = PREmployee.SSString
            NameString = PREmployee.LFName
        ElseIf PRTotal.RecType = 3 Then                         '  Company Totals
            NameString = "COMPANY TOTALS"
            SSString = ""
        Else
            If Not PRDepartment.GetByID(PRTotal.RecID) Then     '  Department Totals
                MsgBox "Dept NF: " & PRTotal.RecID, vbExclamation
                GoBack
            End If
            SSString = "Dpt#: " & PRDepartment.DepartmentNumber
            NameString = PRDepartment.Name

        End If
              
        frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
        frmProgress.Show
    
        Frmt = "d14"
        
        Select Case ReportList
                  
            Case "QtrlyFICAFWT"
                
                If PRTotal.RecType >= 2 Then
                    Ln = Ln + 1
                End If
            
                If PRTotal.Gross = 0 And PRTotal.SSWage = 0 And PRTotal.MEDWage = 0 And PRTotal.FWTTax = 0 And PRTotal.SSTax = 0 And PRTotal.MedTax = 0 Then
                Else
                    
                    PrintValue(1) = SSString:           FormatString(1) = "a13"
                    PrintValue(2) = NameString:         FormatString(2) = "a40"
                    PrintValue(3) = " ":                FormatString(3) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = " ":                FormatString(1) = "a14"
                    PrintValue(2) = PRTotal.Gross:      FormatString(2) = Frmt
                    PrintValue(3) = PRTotal.SSWage:     FormatString(3) = Frmt
                    PrintValue(4) = PRTotal.MEDWage:    FormatString(4) = Frmt
                    PrintValue(5) = PRTotal.FWTTax:     FormatString(5) = Frmt
                    PrintValue(6) = PRTotal.SSTax:      FormatString(6) = "d13"
                    PrintValue(7) = PRTotal.MedTax:     FormatString(7) = "d13"
                    TotFICA = PRTotal.SSTax + PRTotal.MedTax
                    PrintValue(8) = TotFICA:            FormatString(8) = "d13"
                    PrintValue(9) = " ":                FormatString(9) = "~"
                    FormatPrint
                    Ln = Ln + 1
'                    If PRTotal.RecType = 2 Then
'                        Ln = Ln + 1
'                    End If
                End If
                
            Case "QtrlyStateCity"
            
                If PRTotal.RecType >= 2 Then
                    Ln = Ln + 1
                End If
                
                If PRTotal.StateWage = 0 And PRTotal.CityWage = 0 And PRTotal.StateTax = 0 And PRTotal.CityTax = 0 Then
                Else
                    PrintValue(1) = SSString:           FormatString(1) = "a13"
                    PrintValue(2) = NameString:         FormatString(2) = "a32"
                    PrintValue(3) = PRTotal.StateWage:  FormatString(3) = Frmt
                    PrintValue(4) = PRTotal.CityWage:   FormatString(4) = Frmt
                    PrintValue(5) = PRTotal.StateTax:   FormatString(5) = Frmt
                    PrintValue(6) = PRTotal.CityTax:    FormatString(6) = Frmt
                    PrintValue(7) = " ":                FormatString(7) = "~"
                    FormatPrint
                    Ln = Ln + 1
'                    If PRTotal.RecType = 2 Then
'                        Ln = Ln + 1
'                    End If
                End If
            Case "QtrlyFedUnemp"
            
                ' first the company totals
                If PRTotal.RecType = 2 And uFlag = False Then
                    Ln = Ln + 1
                    ' show the max wage and percent
                    X = "     TAX MAX WAGE - FED: " & Format(FedUnempMax, "$##,##0.00") & _
                        " @ " & Format(FedUnempPct, "##0.00") & "%"
                    PrintValue(1) = X:                  FormatString(1) = "a70"
                    PrintValue(2) = " ":                FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                                        
                    If StateID <> 0 Then
                        X = "     TAX MAX WAGE - " & PRState.StateAbbrev & ":  " & Format(StateUnempMax, "$##,##0.00") & _
                            " @ " & Format(StateUnempPct, "##0.00") & "%"
                        PrintValue(1) = X:                  FormatString(1) = "a70"
                        PrintValue(2) = " ":                FormatString(2) = "~"
                        FormatPrint
                        Ln = Ln + 1
                    End If
                    Ln = Ln + 1
                    uFlag = True
                End If
                
                If PRTotal.FUNWageBase = 0 And PRTotal.YTDFUNWageBase = 0 And PRTotal.FUNWage = 0 And PRTotal.SUNWage = 0 Then
                Else
                    
                    If PRTotal.RecType <> 1 And MaxLines - Ln < 5 Then
                        Msg1 = QtrEnding
                        FormFeed
                        PageHeader ReportTitle, Msg1, StateString, ""
                        QtrRptsHeaders (ReportList) ' Print Report Header
                    End If
                    
                    PrintValue(1) = SSString:               FormatString(1) = "a13"
                    PrintValue(2) = NameString:             FormatString(2) = "a32"
                    PrintValue(3) = PRTotal.FUNWageBase:    FormatString(3) = Frmt
                    PrintValue(4) = PRTotal.YTDFUNWageBase: FormatString(4) = Frmt
                    PrintValue(5) = PRTotal.FUNWage:        FormatString(5) = Frmt
                    PrintValue(6) = PRTotal.SUNWage:        FormatString(6) = Frmt
                    PrintValue(7) = " ":                    FormatString(7) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    If PRTotal.RecType <> 1 Then
                        StateUnempAmt = PRTotal.SUNWage * StateUnempPct / 100
                        FedUnempAmt = PRTotal.FUNWage * FedUnempPct / 100
                        PrintValue(1) = "":                 FormatString(1) = "a13"
                        PrintValue(2) = "TAX AMOUNT:":      FormatString(2) = "a60"
                        PrintValue(3) = FedUnempAmt:        FormatString(3) = Frmt
                        PrintValue(4) = StateUnempAmt:      FormatString(4) = Frmt
                        PrintValue(5) = " ":                FormatString(5) = "~"
                        FormatPrint
                        Ln = Ln + 2
                    End If
                    
                End If
                Case "QtrlyTipsTaxes"
'''''                    PrintValue(1) = SSString:           FormatString(1) = "a13"
'''''                    PrintValue(2) = PREmployee.FLName:  FormatString(2) = "a32"
'''''                    PrintValue(3) = " ":                FormatString(3) = "~"
'''''                    FormatPrint
'''''                    Ln = Ln + 1
'''''
'''''                    PrintValue(1) = " ":                FormatString(1) = "a19"
'''''
'''''              ' WAGEFIC
'''''              WageFic = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
'''''                          qYear, qYear, EndMonth, 0, 0, 0)
'''''              TotWageFic = TotWageFic + WageFic
'''''
'''''              PrintValue(2) = Format(WageFic, "##,###,##0.00")
'''''              FormatString(2) = "d13"
'''''
'''''              FindStr = "Dept=" & CStr(PRDepartment.DepartmentNumber)     ' CStr Function Converts numeric value to a string
'''''              trs.Find FindStr, 0, adSearchForward, 1
'''''              If trs.EOF Then
'''''                  trs.AddNew Array("Dept", "WageFIC", "WageMed", "TipsFIC", "TipsMed", "TaxSS", "TaxMed", "TaxFed"), _
'''''                  Array(PRDepartment.DepartmentNumber, 0, 0, 0, 0, 0, 0, 0)
'''''                  trs.UpdateBatch
'''''              End If
'''''
'''''              DWageFic = trs!WageFic
'''''              DWageFic = DWageFic + WageFic
'''''              trs.Fields("Dept") = trs!Dept
'''''              trs.Fields("WageFIC") = DWageFic
'''''
'''''              WageMed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageMed, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotWageMed = TotWageMed + WageMed
'''''
'''''              PrintValue(3) = Format(WageMed, "##,###,##0.00")
'''''              FormatString(3) = "d13"
'''''
'''''              DWageMed = trs!WageMed
'''''              DWageMed = DWageMed + WageMed
'''''              trs.Fields("WageMed") = DWageMed
'''''
'''''              'TIPSFIC
'''''              TipsFic = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotTipsFic = TotTipsFic + TipsFic
'''''
'''''              PrintValue(4) = Format(TipsFic, "##,##0.00")
'''''              FormatString(4) = "d9"
'''''
'''''              DTipsFic = trs!TipsFic
'''''              DTipsFic = DTipsFic + TipsFic
'''''              trs.Fields("TipsFIC") = DTipsFic
'''''
'''''              ' TIPSMED
'''''              TipsMed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotTipsMed = TotTipsMed + TipsMed
'''''
'''''              PrintValue(5) = Format(TipsMed, "##,##0.00")
'''''              FormatString(5) = "d9"
'''''
'''''              DTipsMed = trs!TipsMed
'''''              DTipsMed = DTipsMed + TipsMed
'''''              trs.Fields("TipsMed") = DTipsMed
'''''
'''''              TaxSS = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxSS, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotTaxSS = TotTaxSS + TaxSS
'''''
'''''              PrintValue(6) = Format(TaxSS, "##,##0.00")
'''''              FormatString(6) = "d9"
'''''
'''''              DTaxSS = trs!TaxSS
'''''              DTaxSS = DTaxSS + TaxSS
'''''              trs.Fields("TaxSS") = DTaxSS
'''''
'''''              TaxMed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxMed, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotTaxMed = TotTaxMed + TaxMed
'''''
'''''              PrintValue(7) = Format(TaxMed, "##,##0.00")
'''''              FormatString(7) = "d9"
'''''
'''''              DTaxMed = trs!TaxMed
'''''              DTaxMed = DTaxMed + TaxMed
'''''              trs.Fields("TaxMed") = DTaxMed
'''''
'''''              TaxFed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxFed, _
'''''                          qYear, qYear, StartMonth, EndMonth, 0, 0)
'''''              TotTaxFed = TotTaxFed + TaxFed
'''''
'''''              PrintValue(8) = Format(TaxFed, "##,##0.00")
'''''              FormatString(8) = "d9"
'''''
'''''              DTaxFed = trs!TaxFed
'''''              DTaxFed = DTaxFed + TaxFed
'''''              trs.Fields("TaxFed") = DTaxFed
'''''              trs.Update
'''''
'''''              PrintValue(9) = " "
'''''              FormatString(9) = "~"
'''''
'''''              FormatPrint
'''''              Ln = Ln + 1
              
        End Select

        LastType = PRTotal.RecType
        If Not PRTotal.GetNext Then Exit Do
        
    Loop

    frmProgress.Hide

    PRTotal.TClose

End Sub

Public Sub QtrRptsHeaders(ByVal ReportList As String)
    
    PrintValue(1) = "":                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Name):               FormatString(2) = "a30"
    PrintValue(3) = "":                                 FormatString(3) = "a40"
    PrintValue(4) = "REPORT DATE :  " & Format(Date, "mm/dd/yyyy "): FormatString(4) = "a25"
    PrintValue(5) = " ":                                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "":                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Address1):           FormatString(2) = "a30"
    PrintValue(3) = "":                                 FormatString(3) = "a40"
    PrintValue(4) = "EMPLOYER ID :  " & Trim(PRCompany.FederalID): FormatString(4) = "a25"
    PrintValue(5) = " ":                                FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "":                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Address2):           FormatString(2) = "a30"
    PrintValue(3) = "":                                 FormatString(3) = "a40"
    PrintValue(4) = "EMPLOYER ST.:  " & Trim(PRCompany.StateID): FormatString(4) = "a25"
    PrintValue(5) = " ":                                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
                   
    PrintValue(1) = "":                                                 FormatString(1) = "a10"
    PrintValue(2) = PRCompany.City & ", " & PRState.StateAbbrev & "  " & PRCompany.ZipCode
                                                                        FormatString(2) = "a30"
    PrintValue(3) = "":                                                 FormatString(3) = "a40"
    PrintValue(4) = " ":                                                FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 3
    
    Select Case ReportList
        Case "QtrlyFICAFWT"
            PrintValue(1) = "SOC. SEC.":                        FormatString(1) = "a11"
            PrintValue(2) = " ":                                FormatString(2) = "a2"
            PrintValue(3) = "GROSS":                            FormatString(3) = "r14"
            PrintValue(4) = " S.S.":                            FormatString(4) = "r14"
            PrintValue(5) = "MEDIC":                            FormatString(5) = "r14"
            PrintValue(6) = "FWT":                              FormatString(6) = "r14"
            PrintValue(7) = "S.S.":                             FormatString(7) = "r14"
            PrintValue(8) = "MEDIC":                            FormatString(8) = "r12"
            PrintValue(9) = "TOTAL":                            FormatString(9) = "r13"
            PrintValue(10) = " ":                               FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                           FormatString(1) = "a11"
            PrintValue(2) = " ":                                FormatString(2) = "a2"
            PrintValue(3) = "WAGE":                             FormatString(3) = "r14"
            PrintValue(4) = "WAGE":                             FormatString(4) = "r14"
            PrintValue(5) = "WAGE":                             FormatString(5) = "r14"
            PrintValue(6) = "TAX":                              FormatString(6) = "r14"
            PrintValue(7) = "TAX":                              FormatString(7) = "r13"
            PrintValue(8) = "TAX":                              FormatString(8) = "r12"
            PrintValue(9) = "FICA":                             FormatString(9) = "r13"
            PrintValue(10) = " ":                               FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
             
            PrintValue(1) = String(108, "-"):                   FormatString(1) = "a108"
            PrintValue(2) = " ":                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1

        Case "QtrlyStateCity"
            PrintValue(1) = " SOC. SEC.":                       FormatString(1) = "a13"
            PrintValue(2) = "EMPLOYEE":                         FormatString(2) = "a31"
            PrintValue(3) = "STATE":                            FormatString(3) = "r14"
            PrintValue(4) = " CITY":                            FormatString(4) = "r14"
            PrintValue(5) = "STATE ":                           FormatString(5) = "r15"
            PrintValue(6) = "CITY":                             FormatString(6) = "r13"
            PrintValue(7) = " ":                                FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "    NO.":                             FormatString(1) = "a15"
            PrintValue(2) = "NAME":                             FormatString(2) = "a30"
            PrintValue(3) = "WAGE":                             FormatString(3) = "r12"
            PrintValue(4) = "WAGE":                             FormatString(4) = "r15"
            PrintValue(5) = "TAX ":                             FormatString(5) = "r13"
            PrintValue(6) = "TAX":                              FormatString(6) = "r14"
            PrintValue(7) = " ":                                FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(100, "-"):                   FormatString(1) = "a118"
            PrintValue(2) = " ":                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
        Case "QtrlyFedUnemp"
            PrintValue(1) = "SOC. SEC.":                        FormatString(1) = "a13"
            PrintValue(2) = "EMPLOYEE":                         FormatString(2) = "a32"
            PrintValue(3) = "QTD WAGE":                         FormatString(3) = "r13"
            PrintValue(4) = "YTD WAGE":                         FormatString(4) = "r14"
            PrintValue(5) = "FEDERAL ":                         FormatString(5) = "r15"
            PrintValue(6) = "STATE":                            FormatString(6) = "r13"
            PrintValue(7) = " ":                                FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                           FormatString(1) = "a13"
            PrintValue(2) = "NAME":                             FormatString(2) = "a30"
            PrintValue(3) = "BASE":                             FormatString(3) = "r13"
            PrintValue(4) = "BASE":                             FormatString(4) = "r14"
            PrintValue(5) = "WAGE ":                            FormatString(5) = "r15"
            PrintValue(6) = "WAGE":                             FormatString(6) = "r14"
            PrintValue(7) = " ":                                FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(100, "-"):                   FormatString(1) = "a118"
            PrintValue(2) = " ":                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
        Case "QtrlyTipsTaxes"
            PrintValue(1) = "SOC. SEC.":                        FormatString(1) = "a11"
            PrintValue(2) = "FIC":                              FormatString(2) = "r13"
            PrintValue(3) = "MED":                              FormatString(3) = "r14"
            PrintValue(4) = "FIC ":                             FormatString(4) = "r13"
            PrintValue(5) = "MED":                              FormatString(5) = "r14"
            PrintValue(6) = "S.S.":                             FormatString(6) = "r13"
            PrintValue(7) = "MED":                              FormatString(7) = "r14"
            PrintValue(8) = "FWT ":                             FormatString(8) = "r14"
            PrintValue(9) = " ":                                FormatString(9) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                           FormatString(1) = "a11"
            PrintValue(2) = "NAME":                             FormatString(2) = "a30"
            PrintValue(3) = "WAGE":                             FormatString(3) = "r13"
            PrintValue(4) = "WAGE":                             FormatString(4) = "r14"
            PrintValue(5) = "TIPS":                             FormatString(5) = "r13"
            PrintValue(6) = "TIPS":                             FormatString(6) = "r14"
            PrintValue(7) = "TAX":                              FormatString(7) = "r14"
            PrintValue(8) = " TAX":                             FormatString(8) = "r14"
            PrintValue(9) = "TAX ":                             FormatString(9) = "r14"
            PrintValue(10) = " ":                               FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "-"):                   FormatString(1) = "a118"
            PrintValue(2) = " ":                                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
                        
    End Select
End Sub

Public Sub DepositListing(ByVal RangeType As Byte, _
                          ByVal BatchNumbr As Long, _
                          ByVal PEDate As Long, _
                          ByVal CheckDt As Long, _
                          ByVal StartDate As Long, _
                          ByVal EndDate As Long, _
                          ByVal OptDate As String)

    ' ===> add logic for Unemp & Wkc

Dim ReportTitle As String

Dim WkcAmount As Currency
Dim whSS, whMED, WageSS, WageMed, pctSS, pctMed, whFWT, whSWT, whCWT As Currency
Dim MedAddAmt As Currency
Dim MatchSS, MatchMed As Currency
Dim tlGross, tlDirDep, tlCheck, tlNet, tlEscDed As Currency
Dim FedDep, Escrow As Currency
Dim CheckEscrow As Boolean
Dim p1 As Currency
Dim DirDepCount, CheckCount As Long
Dim DirDepAdd As Currency

Dim rsSWT As New ADODB.Recordset
Dim StateCount As Integer
Dim StateString As String
    
    ReportTitle = "PAYROLL EMPLOYER DEPOSIT LISTING"

    ' track state withheld by state
    rsSWT.CursorLocation = adUseClient
    rsSWT.Fields.Append "StateID", adDouble
    rsSWT.Fields.Append "whSWT", adCurrency
    rsSWT.Open , , adOpenDynamic, adLockOptimistic

    ' get the unemployment percentages
    DepSTUnempPct = PRCompany.StateUnempPct
    
    DepFedUnempPct = 0
    
'    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeFUNPct
'    If PRGlobal.GetBySQL(SQLString) Then
'        DepFedUnempPct = PRGlobal.Amount
'    End If

    
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(BatchNumbr) Then
            MsgBox "PRBatch NF: " & BatchNumbr, vbExclamation
            GoBack
        End If
        
        ' 2015-04-25 - fix query
        DepFedUnempPct = PRGlobal.GetAmount(PREquate.GlobalTypeFUNPct, Year(PRBatch.PEDate))
        
        Msg2 = "BATCH: " & BatchNumbr & "   PERIOD ENDING DATE: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
        msg3 = "CHECK DATE: " & Format(PRBatch.CheckDate, "mm/dd/yyyy")
        SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & BatchNumbr
    Else
            
        ' 2015-04-25 - fix query
        DepFedUnempPct = PRGlobal.GetAmount(PREquate.GlobalTypeFUNPct, Year(StartDate))
        
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRHist WHERE PRHist.CheckDate >= " & CLng(StartDate) & _
                        " AND PRHist.CheckDate <= " & CLng(EndDate)
            Msg2 = "CHECK DATE FROM: " & Format(StartDate, "mm/dd/yyyy") & " To: " & Format(EndDate, "mm/dd/yyyy")
        
        
        Else
            SQLString = "SELECT * FROM PRHist WHERE PRHist.PEDate >= " & CLng(StartDate) & _
                        " AND PRHist.PEDate <= " & CLng(EndDate)
            Msg2 = "P/E DATE FROM: " & Format(StartDate, "mm/dd/yyyy") & " To: " & Format(EndDate, "mm/dd/yyyy")
            
        End If
    End If
    
    frmProgress.lblMsg1 = "Gathering data for deposit report for: " & PRCompany.Name
    frmProgress.Show
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Payroll History Data Found", vbInformation
        GoBack
    End If
    
    ' get the SS and MED pct
    SQLString = "SELECT * FROM PRGlobal WHERE PRGlobal.TypeCode = " & PREquate.GlobalTypeSSPct
    If PRGlobal.GetBySQL(SQLString) Then
        pctSS = PRGlobal.Amount
    Else
        pctSS = 6.2
    End If
    
    SQLString = "SELECT * FROM PRGlobal WHERE PRGlobal.TypeCode = " & PREquate.GlobalTypeMEDPct
    If PRGlobal.GetBySQL(SQLString) Then
        pctMed = PRGlobal.Amount
    Else
        pctMed = 1.45
    End If
    
    ' deductions to escrow ???
    CheckEscrow = False
    If frmDeposit.deds.RecordCount > 0 Then
        frmDeposit.deds.MoveFirst
        Do
            If frmDeposit.deds!UseDeduction = True Then
                CheckEscrow = True
                Exit Do
            End If
            frmDeposit.deds.MoveNext
            If frmDeposit.deds.EOF Then Exit Do
        Loop
    End If
    
    Do
        
        HistCount = HistCount + 1
        If HistCount Mod 100 = 1 Then
            frmProgress.lblMsg2 = "Now Processing Record # " & Format(HistCount, "#,###,##0") & vbCr & _
                                  " Of: " & Format(PRHist.Records, "#,###,##0")
            frmProgress.Refresh
        End If
        
        whSS = whSS + PRHist.SSTax
        whMED = whMED + PRHist.MedTax - PRHist.MedAddAmt
        MedAddAmt = MedAddAmt + PRHist.MedAddAmt
        whFWT = whFWT + PRHist.FWTTax
        
        ' match SS# logic
        ' 2011 ER is still 6.2% / EE is 4.2%
        If Year(PRHist.CheckDate) >= 2011 Then
            MatchSS = MatchSS + (Round(PRHist.SSWage * 0.062, 2))
        Else
            MatchSS = MatchSS + PRHist.SSTax
        End If
        
        ' State Withheld
        ' Data Entry restricts one state per paycheck
        whSWT = whSWT + PRHist.SWTTax
        ' store per state
        rsSWT.Find "StateID = " & PRHist.StateID, 0, adSearchForward, 1
        If rsSWT.EOF Then
            rsSWT.AddNew
            rsSWT!StateID = PRHist.StateID
            rsSWT!whSWT = 0
            rsSWT.Update
        End If
        rsSWT!whSWT = rsSWT!whSWT + PRHist.SWTTax
        rsSWT.Update
        
        whCWT = whCWT + PRHist.CWTTax
        WkcAmount = WkcAmount + PRHist.WkcAmount
            
        WageSS = WageSS + PRHist.SSWage
        WageMed = WageMed + PRHist.MEDWage
        
        tlGross = tlGross + PRHist.Gross
        tlDirDep = tlDirDep + PRHist.DirectDeposit
        tlCheck = tlCheck + PRHist.Net
        tlNet = tlNet + tlDirDep + PRHist.Net
        
        DepSTUnempAmt = DepSTUnempAmt + PRHist.SUNWage
        DepFedUnempAmt = DepFedUnempAmt + PRHist.FUNWage
        
        ' check count
        If PRHist.Net > 0 Then CheckCount = CheckCount + 1
        
        ' check deductions
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID
        If PRItemHist.GetBySQL(SQLString) Then
            Do
                If CheckEscrow = True Then
                    frmDeposit.deds.Find "ItemID = " & PRItemHist.EmployerItemID, 0, adSearchForward, 1
                    If Not frmDeposit.deds.EOF Then
                        frmDeposit.deds!Amount = frmDeposit.deds!Amount + PRItemHist.Amount
                        frmDeposit.deds!Count = frmDeposit.deds!Count + 1
                        frmDeposit.deds.Update
                        If frmDeposit.deds!DirDep = False Then
                            Escrow = Escrow + PRItemHist.Amount
                        Else
                            DirDepAdd = DirDepAdd + PRItemHist.Amount
                        End If
                    End If
                End If
                If PRItemHist.ItemType = PREquate.ItemTypeDirDepDed Then
                    DirDepCount = DirDepCount + 1
                End If
                If Not PRItemHist.GetNext Then Exit Do
            Loop
        End If
        If Not PRHist.GetNext Then Exit Do
    Loop
            
    ' calculated above in PRHist loop due to 2011 SS change
    ' MatchSS = Round(WageSS * pctSS / 100, 2)
    
    MatchMed = Round((WageMed - MedAddAmt) * pctMed / 100, 2)
    MatchMed = whMED
    
    DepSTUnempMatch = Round(DepSTUnempAmt * DepSTUnempPct / 100, 2)
    DepFedUnempMatch = Round(DepFedUnempAmt * DepFedUnempPct / 100, 2)
    
    FedDep = (whSS + whMED) * 2 + whFWT
    
    ' *** revised for 2011 SS Pct ***
    FedDep = (whMED * 2) + whSS + MatchSS + whFWT
    
    FedDep = (whMED * 2) + whSS + MatchSS + whFWT + MedAddAmt
    
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    SetEquates
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    Ln = Ln + 4
    
    PageHeader ReportTitle, Msg1, Msg2, msg3

    Ln = Ln + 2
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "TOTAL GROSS PAY:":                     FormatString(2) = "a56"
    PrintValue(3) = tlGross:                                FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "NUMBER OF RECORDS:":                   FormatString(2) = "a63"
    PrintValue(3) = HistCount:                              FormatString(3) = "n5"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 3
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "S.S. WITHHELD":                        FormatString(2) = "a56"
    PrintValue(3) = whSS:                                   FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "S.S. MATCH":                           FormatString(2) = "a56"
    PrintValue(3) = MatchSS:                                FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "S.S. TOTAL":                           FormatString(2) = "a56"
    PrintValue(3) = MatchSS + whSS:                         FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "MED WITHHELD":                         FormatString(2) = "a56"
    PrintValue(3) = whMED:                                  FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    If MedAddAmt <> 0 Then
        PrintValue(1) = "":                                     FormatString(1) = "a5"
        PrintValue(2) = "MED ADDL WITHHELD":                    FormatString(2) = "a56"
        PrintValue(3) = MedAddAmt:                                  FormatString(3) = "d13"
        PrintValue(4) = " ":                                    FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 2
    Else
        Ln = Ln + 1
    End If
    
    ' ********************************************************************************
    ' 12/24/09
    ' just add the SS/MED withheld for the match
    ' for employees flagged as non-taxable
    '     in PRHist - wage stored as taxable with no tax withheld
    '     for qtrly reporting???
'    PrintValue(1) = "":                                     FormatString(1) = "a5"
'    PrintValue(2) = "FICA MATCH":                           FormatString(2) = "a23"
'    PrintValue(3) = WageSS:                                 FormatString(3) = "d7"
'    PrintValue(4) = "x":                                    FormatString(4) = "a1"
'    PrintValue(5) = pctSS:                                  FormatString(5) = "d5"
'    PrintValue(6) = "%":                                    FormatString(6) = "a4"
'    PrintValue(7) = MatchSS:                                FormatString(7) = "d13"
'    PrintValue(8) = " ":                                    FormatString(8) = "~"
'    FormatPrint
'    Ln = Ln + 2
'
'    PrintValue(1) = " ":                                    FormatString(1) = "a5"
'    PrintValue(2) = "MED MATCH":                            FormatString(2) = "a23"
'    PrintValue(3) = WageMed:                                FormatString(3) = "d7"
'    PrintValue(4) = "x":                                    FormatString(4) = "a1"
'    PrintValue(5) = pctMed:                                 FormatString(5) = "d5"
'    PrintValue(6) = "%":                                    FormatString(6) = "a4"
'    PrintValue(7) = MatchMed:                               FormatString(7) = "d13"
'    PrintValue(8) = " ":                                    FormatString(8) = "~"
'    FormatPrint
'    Ln = Ln + 2
    ' ********************************************************************************
    
    ' calculated above - 2011 SS change
    ' calc while looping thru PRHist
    ' *** MatchSS = whSS
    
    MatchMed = whMED
    PrintValue(1) = " ":                                    FormatString(1) = "a5"
    PrintValue(2) = "MED MATCH":                            FormatString(2) = "a56"
    PrintValue(3) = MatchMed:                               FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    MatchMed = whMED
    PrintValue(1) = " ":                                    FormatString(1) = "a5"
    PrintValue(2) = "MED TOTAL":                            FormatString(2) = "a56"
    PrintValue(3) = whMED + MatchMed + MedAddAmt:           FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                                    FormatString(1) = "a5"
    PrintValue(2) = "FEDERAL TAX WITHHELD":                 FormatString(2) = "a56"
    PrintValue(3) = whFWT:                                  FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    ' line for each state withheld
    If rsSWT.RecordCount = 0 Then
        PrintValue(1) = "":                                     FormatString(1) = "a5"
        PrintValue(2) = "STATE TAX WITHHELD":                   FormatString(2) = "a56"
        PrintValue(3) = whSWT:                                  FormatString(3) = "d13"
        PrintValue(4) = " ":                                    FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 2
    Else
        StateCount = 0
        rsSWT.Sort = "StateID"
        rsSWT.MoveFirst
        Do
            If rsSWT!whSWT <> 0 Then
                
                If PRState.GetByID(rsSWT!StateID) = True Then
                    StateString = PRState.StateAbbrev
                Else
                    StateString = rsSWT!StateID
                End If
                StateString = StateString & " STATE TAX WITHHELD"
                
                PrintValue(1) = "":                                     FormatString(1) = "a5"
                PrintValue(2) = StateString:                            FormatString(2) = "a56"
                PrintValue(3) = rsSWT!whSWT:                            FormatString(3) = "d13"
                PrintValue(4) = " ":                                    FormatString(4) = "~"
                FormatPrint
                Ln = Ln + 1
        
            End If
            rsSWT.MoveNext
        Loop Until rsSWT.EOF
        Ln = Ln + 1
    End If
    
    PrintValue(1) = " ":                                    FormatString(1) = "a5"
    PrintValue(2) = "CITY TAX WITHHELD":                    FormatString(2) = "a56"
    PrintValue(3) = whCWT:                                  FormatString(3) = "d13"
    PrintValue(4) = " ":                                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2
    
    If frmDeposit.chkExclUnemp = 0 Then
    
        PrintValue(1) = " ":                                    FormatString(1) = "a5"
        PrintValue(2) = "STATE UNEMPLOYMENT":                   FormatString(2) = "a22"
        PrintValue(3) = DepSTUnempAmt:                          FormatString(3) = "d7"
        PrintValue(4) = "x":                                    FormatString(4) = "a2"
        PrintValue(5) = DepSTUnempPct:                          FormatString(5) = "d5"
        PrintValue(6) = "%":                                    FormatString(6) = "a4"
        PrintValue(7) = DepSTUnempMatch:                        FormatString(7) = "d13"
        PrintValue(8) = " ":                                    FormatString(8) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                                    FormatString(1) = "a5"
        PrintValue(2) = "FEDERAL UNEMPLOYMENT":                 FormatString(2) = "a22"
        PrintValue(3) = DepFedUnempAmt:                         FormatString(3) = "d7"
        PrintValue(4) = "x":                                    FormatString(4) = "a2"
        PrintValue(5) = DepFedUnempPct:                         FormatString(5) = "d5"
        PrintValue(6) = "%":                                    FormatString(6) = "a4"
        PrintValue(7) = DepFedUnempMatch:                       FormatString(7) = "d13"
        PrintValue(8) = " ":                                    FormatString(8) = "~"
        FormatPrint
        Ln = Ln + 2
    
    Else
    
        DepSTUnempAmt = 0
        DepSTUnempMatch = 0
        DepFedUnempAmt = 0
        DepFedUnempMatch = 0
        
    End If
    
    If WkcAmount <> 0 Then
        PrintValue(1) = " ":                                    FormatString(1) = "a5"
        PrintValue(2) = "WORKER'S COMPENSATION":                FormatString(2) = "a56"
        PrintValue(3) = WkcAmount:                              FormatString(3) = "d13"
        PrintValue(4) = " ":                                    FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 2
    End If
    
    Ln = Ln + 4
    
    ' ===== FED DEPOSIT =====
    Prvw.vsp.Font.Bold = True
        
    PrintValue(1) = "":                     FormatString(1) = "a5"
    
    If frmDeposit.chkFedTaxDep Then     ' FED tax deposit now
        PrintValue(2) = "FEDERAL DEPOSIT"
    Else
        PrintValue(2) = "MAKE FED DEPOSIT BY YOUR DEPOSIT DATE FOR:"
    End If
    FormatString(2) = "a41"
        
    PrintValue(3) = "":                                     FormatString(3) = "a15"
    PrintValue(4) = FedDep:                                 FormatString(4) = "d13"
    PrintValue(5) = " ":                                    FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    Prvw.vsp.Font.Bold = False
    
    ' ==== deductions for Escrow? ====
    If Escrow <> 0 Then
        
        PrintValue(1) = "     DEDUCTION ESCROW":            FormatString(1) = "a35"
        PrintValue(2) = "":                                 FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
        
        frmDeposit.deds.MoveFirst
        Do
                
            If frmDeposit.deds!UseDeduction And _
               frmDeposit.deds!Amount <> 0 And _
               frmDeposit.deds!DirDep = False Then
            
                If PRItem.GetByID(frmDeposit.deds!ItemID) Then
                    X = PRItem.Title
                Else
                    X = "Deduction: " & frmDeposit.deds!ItemID
                End If
                
                PrintValue(1) = "     " & X:                FormatString(1) = "a35"
                PrintValue(2) = "":                         FormatString(2) = "a10"
                PrintValue(3) = frmDeposit.deds!Amount:     FormatString(3) = "d13"
                PrintValue(4) = "":                         FormatString(4) = "~"
                FormatPrint
                Ln = Ln + 1
        
                tlEscDed = tlEscDed + frmDeposit.deds!Amount
        
            End If
        
            frmDeposit.deds.MoveNext
            If frmDeposit.deds.EOF Then Exit Do
        
        Loop
    End If
    
    ' ==== print total of deductions ====
    If tlEscDed <> 0 Then
    
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        PrintValue(2) = "TOTAL DEDUCTIONS TO BE ESCROWED":  FormatString(2) = "a41"
        PrintValue(3) = "":                                 FormatString(3) = "a15"
        PrintValue(4) = tlEscDed:                           FormatString(4) = "d13"
        PrintValue(5) = " ":                                FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
    
    End If
    
    ' ==== combine tax escrow and net pay ====
    If frmDeposit.chkCombNetPay Then
        
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        PrintValue(2) = "TOTAL OF ALL TAXES":               FormatString(2) = "a41"
        PrintValue(3) = "":                                 FormatString(3) = "a15"
        p1 = whSS + whMED + whFWT + whSWT + whCWT + MatchSS + MatchMed + WkcAmount + MedAddAmt
        PrintValue(4) = p1:                                 FormatString(4) = "d13"
        PrintValue(5) = " ":                                FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
    
        ' ===================================
    
        Prvw.vsp.Font.Bold = True
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        PrintValue(2) = "TO PAYROLL ACCOUNT":               FormatString(2) = "a41"
        PrintValue(3) = "":                                 FormatString(3) = "a15"
        p1 = whSS + whMED + whFWT + whSWT + whCWT + MatchSS + MatchMed + tlDirDep + tlCheck + tlEscDed + WkcAmount + MedAddAmt
        PrintValue(4) = p1:                                 FormatString(4) = "d13"
        PrintValue(5) = " ":                                FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        Prvw.vsp.Font.Bold = False
    
    Else        ' net and tax escrow separate
    
        Prvw.vsp.Font.Bold = True
    
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        PrintValue(2) = "TOTAL TAXES AND DEDUCTIONS TO BE ESCROWED":   FormatString(2) = "a41"
        PrintValue(3) = "":                                 FormatString(3) = "a15"
        p1 = whSS + whMED + whFWT + whSWT + whCWT + MatchSS + MatchMed + tlEscDed + MedAddAmt
        p1 = p1 + DepFedUnempMatch + DepSTUnempMatch + WkcAmount
        PrintValue(4) = p1:                                 FormatString(4) = "d13"
        PrintValue(5) = " ":                                FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        Prvw.vsp.Font.Bold = False
        
    End If
    
    Ln = Ln + 1
    
    ' ==== net payroll ====
    PrintValue(1) = "":                                     FormatString(1) = "a5"
    PrintValue(2) = "NET PAYROLL TOTALS":                   FormatString(2) = "a50"
    PrintValue(3) = "":                                     FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
    
    If tlDirDep <> 0 Then
    
        If tlCheck = 0 Then Prvw.vsp.Font.Bold = True
        If frmDeposit.chkCombNetPay Then Prvw.vsp.Font.Bold = False
    
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        
        X = "DIRECT DEPOSIT AMOUNT " & Format(DirDepCount, "#,##0") & " TRANSACTIONS"
        PrintValue(2) = X:                                  FormatString(2) = "a56"
        PrintValue(3) = tlDirDep:                           FormatString(3) = "d13"
        PrintValue(4) = "":                                 FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1
    
        Prvw.vsp.Font.Bold = False
    
    End If
    
    If DirDepAdd <> 0 Then
    
        frmDeposit.deds.MoveFirst
        Do
            If frmDeposit.deds!DirDep = True And frmDeposit.deds!Count <> 0 Then
                
                If PRItem.GetByID(frmDeposit.deds!ItemID) Then
                    X = PRItem.Title
                Else
                    X = "Deduction: " & frmDeposit.deds!ItemID
                End If
                
                X = Trim(X) & " " & frmDeposit.deds!Count & " TRANSACTIONS"
                
                PrintValue(1) = "":                         FormatString(1) = "a5"
                PrintValue(2) = X:                          FormatString(2) = "a56"
                PrintValue(3) = frmDeposit.deds!Amount:     FormatString(3) = "d13"
                PrintValue(4) = "":                         FormatString(4) = "~"
                FormatPrint
                Ln = Ln + 1
    
            End If
            
            frmDeposit.deds.MoveNext
        
        Loop Until frmDeposit.deds.EOF
    
    End If
    
    If tlCheck <> 0 Then
    
        If tlDirDep = 0 Then Prvw.vsp.Font.Bold = True
        If frmDeposit.chkCombNetPay Then Prvw.vsp.Font.Bold = False
    
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        X = "NET PAYROLL CHECK AMOUNT " & Format(CheckCount, "#,##0") & " CHECKS"
        PrintValue(2) = X:                                  FormatString(2) = "a56"
        PrintValue(3) = tlCheck:                            FormatString(3) = "d13"
        PrintValue(4) = "":                                 FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1
    
        Prvw.vsp.Font.Bold = False
    
    End If
    
    If (tlDirDep <> 0 And tlCheck <> 0) Or DirDepAdd <> 0 Then
        
        Prvw.vsp.Font.Bold = True
        If frmDeposit.chkCombNetPay Then Prvw.vsp.Font.Bold = False
        
        PrintValue(1) = "":                                 FormatString(1) = "a5"
        PrintValue(2) = "NET PAYROLL TOTAL":                FormatString(2) = "a56"
        p1 = tlCheck + tlDirDep + DirDepAdd
        PrintValue(3) = p1:                                 FormatString(3) = "d13"
        PrintValue(4) = "":                                 FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1
        
        Prvw.vsp.Font.Bold = False
    
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub GLUpdate()

Dim rsGL As New ADODB.Recordset
Dim trsGL As New ADODB.Recordset
Dim HistCount, I As Integer
Dim CompanyFlag As Boolean
Dim GLFlag As Boolean
Dim sGLType, sRelatedID As String
Dim Amt As Currency
Dim GLMsg As String

    frmProgress.Show
    frmProgress.lblMsg1 = "Setting Up Files"
    frmProgress.Refresh

    trsGL.CursorLocation = adUseClient
    
    trsGL.Fields.Append "GLType", adInteger
    trsGL.Fields.Append "RelatedID", adDouble
    trsGL.Fields.Append "GLItemType", adInteger
    trsGL.Fields.Append "ItemID", adDouble
    trsGL.Fields.Append "GLAccountNum", adDouble
    trsGL.Fields.Append "Amount", adCurrency
    
    trsGL.Open , , adOpenDynamic, adLockOptimistic

    ' load up the PRGLUpd Info
    CompanyFlag = False
    SQLString = "SELECT * FROM PRGLUpd ORDER BY GLType, RelatedID"
                
    If Not PRGLUpd.GetBySQL(SQLString) Then
        MsgBox "No PR to GL Info found!", vbExclamation, "PR to GL Update"
        Exit Sub
    End If
    
    Do
        trsGL.AddNew
        trsGL!GLType = PRGLUpd.GLType
        trsGL!RelatedID = PRGLUpd.RelatedID
        trsGL!GLItemType = PRGLUpd.GLItemType
        trsGL!ItemID = PRGLUpd.ItemID
        trsGL!GLAccountNum = PRGLUpd.GLAccountNum
        trsGL!Amount = 0
        trsGL.Update
        If PRGLUpd.GLType = PREquate.GLTypeCompany Then CompanyFlag = True
        If Not PRGLUpd.GetNext Then Exit Do
    Loop
    
    If Not CompanyFlag Then
        MsgBox "No Company GL Info Exists!", vbExclamation, "PR to GL Update"
        Exit Sub
    End If
    
    ' loop thru PR Hist and gather info
    SQLString = "SELECT * FROM PRHist WHERE YearMonth = " & frmGLUpdate.YM
    If frmGLUpdate.optUpdRecent Then
        SQLString = Trim(SQLString) & " AND GLUpdate = 0"
    End If

    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Payroll History Found!", vbExclamation, "PR to GL Update"
        Exit Sub
    End If
    
    frmProgress.lblMsg1 = "Now Updating Payroll History to GL " & frmGLUpdate.YM
    
    Do
        
        HistCount = HistCount + 1
        frmProgress.lblMsg2 = "History Record Count: " & HistCount
        frmProgress.Refresh
        
        ' filter down the trsGL (PRGLUpd) record set
        ' to the GLType (Employee or Dept or Company
        ' 1 = Employee
        ' 2 = Dept
        ' 3 = Company
        I = 1
        trsGL.Filter = adFilterNone
        Do
            Select Case I
                Case 1
                    trsGL.Filter = "GLType = " & I & " AND RelatedID = " & PRHist.EmployeeID
                    If PREmployee.GetByID(PRHist.EmployeeID) Then
                    End If
                    sGLType = "Employee# " & PREmployee.EmployeeNumber & " " & PREmployee.LFName
                Case 2
                    trsGL.Filter = "GLType = " & I & " AND RelatedID = " & PRHist.DepartmentID
                    If PRDepartment.GetByID(PRHist.DepartmentID) Then
                    End If
                    sGLType = "Dept# " & " " & PRDepartment.DepartmentNumber & " " & _
                              PRDepartment.Name
                Case 3
                    trsGL.Filter = "GLType = " & I
                    sGLType = "Company Record"
            End Select
            If trsGL.RecordCount > 0 Then Exit Do
            trsGL.Filter = adFilterNone
            I = I + 1
            If I = 4 Then
                MsgBox "No PRGLUpd Info Found!", vbExclamation, "PR to GL Update"
                Exit Sub
            End If
        Loop

        ' update the amounts from PRHist
        For I = 1 To 10
            Select Case I
                Case 1
                    J = PREquate.GLItemTypeSSTax
                    Amt = -PRHist.SSTax
                    GLMsg = "SS Tax"
                Case 2
                    J = PREquate.GLItemTypeSSMatch
                    If PRHist.YearMonth >= 201101 Then  ' employer match still 6.2%
                        Amt = -Round(PRHist.SSWage * 0.062, 2)
                    Else
                        Amt = -PRHist.SSTax
                    End If
                    GLMsg = "SS Match"
                Case 3
                    J = PREquate.GLItemTypeSSExp
                    If PRHist.YearMonth >= 201101 Then  ' employer match still 6.2%
                        Amt = Round(PRHist.SSWage * 0.062, 2)
                    Else
                        Amt = PRHist.SSTax
                    End If
                    GLMsg = "SS Expense"
                Case 4
                    J = PREquate.GLItemTypeMedTax
                    Amt = -PRHist.MedTax
                    GLMsg = "Med Tax"
                Case 5
                    J = PREquate.GLItemTypeMedMatch
                    Amt = -PRHist.MedTax
                    GLMsg = "Med Match"
                Case 6
                    J = PREquate.GLItemTypeMEDExp
                    Amt = PRHist.MedTax
                    GLMsg = "Med Exp"
                Case 7
                    J = PREquate.GLItemTypeFWTTax
                    Amt = -PRHist.FWTTax
                    GLMsg = "FWT Tax"
                Case 8
                    J = PREquate.GLItemTypeSWTTax
                    Amt = -PRHist.SWTTax
                    GLMsg = "SWT Tax"
                Case 9
                    J = PREquate.GLItemTypeCWTTax
                    Amt = -PRHist.CWTTax
                    GLMsg = "CWT Tax"
                Case 10
                    J = PREquate.GLItemTypeNet
                    Amt = -PRHist.Net - PRHist.DirectDeposit
                    GLMsg = "Net Pay"
            End Select
                   
            trsGL.Find "GLItemType = " & J, 0, adSearchForward, 1
            If trsGL.EOF Then
                MsgBox "GL Update Definition not found for: " & _
                       sGLType & vbCr & _
                       "Item Type: " & GLMsg, vbExclamation, "PR to GL Update"
                Exit Sub
            End If
            trsGL!Amount = trsGL!Amount + Amt
            trsGL.Update
        
        Next I
            
        ' get info from PRDist
        SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID
        If PRDist.GetBySQL(SQLString) Then
            Do
                I = 0
                Select Case PRDist.ItemType
                    Case PREquate.ItemTypeRegPay
                        I = PREquate.GLItemTypeRegPay
                        J = 0
                    Case PREquate.ItemTypeOvtPay
                        I = PREquate.GLItemTypeOvtPay
                        J = 0
                    Case PREquate.ItemTypeOE
                        I = PREquate.GLItemTypeOE
                        J = PRDist.EmployerItemID
                End Select
                If I = 0 Then
                    MsgBox "PRDist Item NF: " & PRDist.ItemType, vbExclamation, "PR to GL Update"
                    Exit Sub
                End If
                
                If J = 0 Then
                    trsGL.Find "GLItemType = " & I, 0, adSearchForward, 1
                    If trsGL.EOF Then
                        MsgBox "PRDist Item nf: " & I, vbExclamation, "PR to GL Update"
                        Exit Sub
                    End If
                    trsGL!Amount = trsGL!Amount + PRDist.Amount
                    trsGL.Update
                Else
                    GLFlag = False
                    trsGL.MoveFirst
                    Do
                        If trsGL!GLItemType = I And trsGL!ItemID = PRDist.EmployerItemID Then
                            trsGL!Amount = trsGL!Amount + PRDist.Amount
                            trsGL.Update
                            GLFlag = True
                            Exit Do
                        End If
                        trsGL.MoveNext
                    Loop Until trsGL.EOF
                    If Not GLFlag Then
                        ' get the employee record
                        If PREmployee.GetByID(PRDist.EmployeeID) Then
                        End If
                        ' get the department id
                        If PRDepartment.GetByID(PRDist.DepartmentID) Then
                        End If
                        ' get the employer item
                        If PRItem.GetByID(PRDist.EmployerItemID) Then
                        End If
                        
                        ' construct the error message
                        X = "GL Account not assigned: " & vbCr & _
                            "Dept #: " & PRDepartment.DepartmentNumber & " " & PRDepartment.Name & vbCr & _
                            "Item: " & PRItem.Title
                            
                        ' MsgBox "PRDist ID NF: " & PRDist.ItemID, vbExclamation, "PR to GL Update"
                        MsgBox X, vbExclamation
                        
                        Exit Sub
                    End If
                End If
                
                If Not PRDist.GetNext Then Exit Do
            Loop
        End If
                    
        ' get info from PRItemHist
        SQLString = "SELECT * FROM PRItemHist WHERE HistID = " & PRHist.HistID & _
                    " AND PRItemHist.ItemType <> " & PREquate.ItemTypeDirDepDed & _
                    " AND PRItemHist.EmployerItemID <> 0"
        If PRItemHist.GetBySQL(SQLString) Then
            Do
                GLFlag = False
                trsGL.MoveFirst
                Do
                    If trsGL!GLItemType = PREquate.GLItemTypeDed And trsGL!ItemID = PRItemHist.EmployerItemID Then
                        trsGL!Amount = trsGL!Amount - PRItemHist.Amount
                        trsGL.Update
                        GLFlag = True
                        Exit Do
                    End If
                    trsGL.MoveNext
                Loop Until trsGL.EOF
                If Not GLFlag Then
                    If PRItem.GetByID(PRItemHist.EmployerItemID) Then
                    End If
                    MsgBox "No GL account set for deduction: " & PRItem.Title, vbExclamation
                    ' MsgBox "PRItemHistID NF: " & PRItemHist.EmployerItemID & vbCr & PRItemHist.ItemHistID, vbExclamation
                    Exit Sub
                End If
                
                If Not PRItemHist.GetNext Then Exit Do
            Loop
        End If
                    
        ' update the PRHist flag
        If PRHist.GLUpdate = 0 Then
            PRHist.GLUpdate = 1
            PRHist.Save (Equate.RecPut)
        End If
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    trsGL.Filter = adFilterNone

    ' create GLBatch record
    SQLString = "SELECT BatchNumber FROM GLBatch ORDER BY BatchNumber DESC"
    rsInit SQLString, cn, rsGL
    If rsGL.BOF And rsGL.EOF Then
        I = 1
    Else
        rsGL.MoveFirst
        I = rsGL!BatchNumber + 1
    End If
    rsGL.Close
    
    GLBatch.OpenRS
    GLHistory.OpenRS
    
    GLBatch.Clear
    GLBatch.FiscalYear = frmGLUpdate.cmbFiscalYear.text
    GLBatch.Period = frmGLUpdate.cmbFiscalPeriod.ListIndex + 1
    GLBatch.BatchNumber = I
    GLBatch.Debits = 0
    GLBatch.Credits = 0
    GLBatch.Created = Now()
    GLBatch.CreateUser = User.ID
    GLBatch.UpdateUser = User.ID
    GLBatch.JournalSource = frmGLUpdate.JS
    
    ' loop thru the temp record set and update to GLHistory
    trsGL.Sort = "GLType, RelatedID, GLItemType, ItemID"
    trsGL.MoveFirst
    Do
        If trsGL!Amount <> 0 Then
            
            GLHistory.Clear
            GLHistory.Account = trsGL!GLAccountNum
            GLHistory.FiscalYear = GLBatch.FiscalYear
            GLHistory.Period = GLBatch.Period
            GLHistory.BatchNumber = GLBatch.BatchNumber
            GLHistory.Amount = trsGL!Amount
            If trsGL!GLType = PREquate.GLTypeEmployee Then
                PREmployee.GetByID (trsGL!RelatedID)
                GLHistory.Reference = "EE: " & PREmployee.EmployeeNumber
            ElseIf trsGL!GLType = PREquate.GLTypeDept Then
                PRDepartment.GetByID (trsGL!RelatedID)
                GLHistory.Reference = "DPT: " & PRDepartment.DepartmentNumber
            Else
                GLHistory.Reference = "COMPANY"
            End If
            
            GLHistory.Description = Mid(frmGLUpdate.txtDescription.text, 1, 20)
         '   GLHistory.Description = trsGL!GLItemType & " " & trsGL!ItemID
            
            GLHistory.SourceCode = 0
            
            GLHistory.JournalSource = GLBatch.JournalSource
            GLHistory.HisType = "A"
            GLHistory.UpdateFlag = True
            GLHistory.PostDate = Now()
            GLHistory.Save (Equate.RecAdd)
            
            ' update the batch totals
            GLBatch.Records = GLBatch.Records + 1
            If trsGL!Amount > 0 Then
                GLBatch.Debits = GLBatch.Debits + trsGL!Amount
            Else
                GLBatch.Credits = GLBatch.Credits + trsGL!Amount
            End If
    
        End If
        
        trsGL.MoveNext
    
    Loop Until trsGL.EOF
    
    GLBatch.Updated = Now()
    GLBatch.Save (Equate.RecAdd)
    
    MsgBox Format(HistCount, "##,###,##0") & " Payroll History Records have been updated" & vbCr & _
           "Hit OK to run the GL Update", vbInformation, "Payroll to GL Update"
    
    ' call to update program
    If BalintFolder = "" Then
        X = "\Balint\GLUtil.exe " & _
            "SysFile=\Balint\Data\GLSystem.mdb " & _
            "UserID=" & User.ID & " " & _
            "BackName=\Balint\GLMenu.exe " & _
            "ProgName=UpdateB " & _
            "Batch=" & GLBatch.BatchNumber
    Else
        X = "\Balint\GLUtil.exe " & _
            "SysFile=" + BalintFolder + "\Data\GLSystem.mdb " & _
            "UserID=" & User.ID & " " & _
            "BackName=\Balint\GLMenu.exe " & _
            "ProgName=UpdateB " & _
            "Batch=" & GLBatch.BatchNumber & " " & _
            "BalintFolder=" & BalintFolder
    End If
    
    TaskID = Shell(X, vbMaximizedFocus)
    End
    
End Sub

Public Sub CheckRegister(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal CheckDt As Long, _
                         ByVal StartDate As Long, _
                         ByVal EndDate As Long, _
                         ByVal OptDate As String, _
                         ByVal Bold As Boolean)
                         
Dim SQLString1 As String
Dim RecCnt, LastEE, EECount As Long
Dim tFlag As Boolean
Dim DedString As String
Dim TotHours, TotTaxes As Currency
Dim Lines As Byte
Dim I As Long
Dim jbCount As Long
Dim rsState As New ADODB.Recordset

    ' 2016-10-29 - rsState for state splits
    '    report at end of register if more than one state
    rsState.CursorLocation = adUseClient
    rsState.Fields.Append "StateID", adDouble
    rsState.Fields.Append "Gross", adDouble
    rsState.Fields.Append "StWage", adDouble
    rsState.Fields.Append "SWT", adDouble
    rsState.Open , , adOpenDynamic, adLockOptimistic

    ' max items per line
    ItemMax = 10
    ChkEmp = 0
    frmCheckReg.Hide
    ReportTitle = "PAYROLL CHECK REGISTER"
    PrtInit ("Land")
    ' LandSw = 1
    SetFont 8, Equate.LandScape
    PRTotal.CreateRS
    Columns = 145
    
    If Bold Then
        Prvw.vsp.Font.Bold = True
    End If
           
    ' does the client have any direct deposits set up?
    DirDepFlag = False
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDirDepDed
    If PRItem.GetBySQL(SQLString) Then DirDepFlag = True
           
    SQLString1 = "SELECT PREmployee.*, PRHist.*, PRDepartment.* " & _
                 " FROM (PREmployee " & _
                 " INNER JOIN PRHist ON PREmployee.EmployeeID = PRHist.EmployeeID) " & _
                 " INNER JOIN PRDepartment ON PRDepartment.DepartmentID = PREmployee.DepartmentID "

    ' take out dept JOIN
    SQLString1 = "SELECT PREmployee.*, PRHist.* " & _
                 " FROM (PREmployee " & _
                 " INNER JOIN PRHist ON PREmployee.EmployeeID = PRHist.EmployeeID) "

    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(BatchNumbr) Then
            MsgBox "PR Batch NF: " & BatchNumbr, vbExclamation
            GoBack
        End If
        SQLString1 = Trim(SQLString1) & " WHERE PRHist.BatchID = " & BatchNumbr
        DedString = " PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
        OptDate = "P/E DATE"
    Else
        If OptDate = "CHECK DATE" Then
            SQLString1 = Trim(SQLString1) & " WHERE PRHist.CheckDate >= " & (StartDate) & " AND " & _
                                    " PRHist.CheckDate <= " & (EndDate)
            DedString = " PRHist.CheckDate >= " & (StartDate) & " AND " & _
                                    " PRHist.CheckDate <= " & (EndDate)
            Msg1 = "CHECK DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
                                    
        Else
            SQLString1 = Trim(SQLString1) & " WHERE PRHist.PEDate >= " & (StartDate) & " AND " & _
                                    " PRHist.PEDate <= " & (EndDate)
            DedString = " PRHist.PEDate >= " & (StartDate) & " AND " & _
                                    " PRHist.PEDate <= " & (EndDate)
            Msg1 = "P/E DATE RANGE: " & Format(StartDate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If

    ' set up SQL statement based upon order checked
    If frmCheckReg.optCheckNo = True Then
        SQLString1 = Trim(SQLString1) & " ORDER BY PRHist.CheckNumber"
        ReportTitle = Trim(ReportTitle) & " BY CHECK NUMBER"
    ElseIf frmCheckReg.optEmpNo = True Then
        SQLString1 = Trim(SQLString1) & " ORDER BY PREmployee.EmployeeNumber, PRHist.HistID"
        ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NUMBER"
    Else                                                          ' order by Employee Name
        SQLString1 = Trim(SQLString1) & " ORDER BY PREmployee.LastName, PREmployee.FirstName, PRHist.HistID"
        ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NAME"
    End If
        
    rsInit SQLString1, cn, rrs       ' rrs vars get assigned in rsInit

    If rrs.EOF = True And rrs.BOF = True Then
        MsgBox "No Data Found !!!", vbExclamation, "Payroll Check Register"
        Prvw.vsp.EndDoc
        Exit Sub
    End If

    rrs.MoveFirst
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    ODTFlag = False
    LineCount = 3   ' lines per entry - always two data lines and the separator line
    If frmCheckReg.chkOEHrs Or frmCheckReg.chkOEAmt Or frmCheckReg.chkDed Then
        ODTFlag = True
        ChkRegGetHeaderData
        I = Int(Round(OECount / ItemMax, 1) + 0.9)
        If frmCheckReg.chkOEHrs Then LineCount = LineCount + I
        If frmCheckReg.chkOEAmt Then LineCount = LineCount + I
        If frmCheckReg.chkDed Then LineCount = LineCount + Int(Round(DEDCount / ItemMax, 1) + 0.9)
    End If
    
    rrs.MoveFirst

    jbCount = 0
    Recs = rrs.RecordCount

    Do
        
        jbCount = jbCount + 1
        If jbCount = 1 Or jbCount Mod 20 = 0 Then
            frmProgress.lblMsg2 = "Processing record: " & Format(jbCount, "##,###,##0") & _
                                  " Of: " & Format(Recs, "##,###,##0")
            frmProgress.Refresh
        End If
        
        ' employee select filter?
        If frmEmpSelect.AllEmployees = False Then
            Msg2 = frmEmpSelect.Count & " Selected Employees"
            SQLString = "EmpNo = " & rrs!EmployeeNumber
            frmEmpSelect.rsEmp.Find SQLString, 0, adSearchForward, 1
            If frmEmpSelect.rsEmp.EOF Then GoTo NextEmp
            If frmEmpSelect.rsEmp!Select = False Then GoTo NextEmp
        Else
            Msg2 = "All Employees"
        End If
        
        If Ln = 0 Or Ln > MaxLines - LineCount Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            ChkRegHeader
        End If

        ' ****************************************************************************
        ' 2016-10-29
        ' get all PRHist
        ' separate update for item types
        ' update all for state totals
        ' store per state
        rsState.Find "StateID = " & rrs!StateID, 0, adSearchForward, 1
        If rsState.EOF Then
            rsState.AddNew
            rsState!StateID = rrs!StateID
            rsState!Gross = 0
            rsState!StWage = 0
            rsState!SWT = 0
            rsState.Update
        End If
        
        rsState!Gross = rsState!Gross + rrs!Gross
        rsState!StWage = rsState!StWage + rrs!SWTWage
        rsState!SWT = rsState!SWT + rrs!SWTTax
        rsState.Update
        ' ****************************************************************************

        ' employee subtotal
        If frmCheckReg.chkEESubTotal And LastEE <> 0 And rrs![PREmployee.EmployeeID] <> LastEE Then
            If PRTotal.tFind(PREquate.GLTypeEmployee, 999999998) And EECount >= 1 Then
                If MaxLines - Ln <= 5 Then
                    FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    ChkRegHeader
                End If
                
                If PREmployee.GetByID(LastEE) Then
                    ChkRegPrtTotals PREmployee.EmployeeNumber & "-" & PREmployee.LFName
                Else
                    ChkRegPrtTotals "Employee ID Not Found: " & LastEE
                End If
                ChkRegPrintODT ODTTypeEE, 0, "", False
            Else
                PRTotal.Clear
                PRTotal.Save (Equate.RecPut)
                ChkRegClearODT ODTTypeEE, 0
            End If
            EECount = 0
            DirDepEE = 0
        End If
    
        LastEE = rrs![PREmployee.EmployeeID]
        EECount = EECount + 1
        RecCnt = RecCnt + 1

        '  RecType/IDNumber
                
        ' get the department ID
'        If PREmployee.DepartmentID <> 0 Then
'            If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
'                MsgBox "Department NF: " & PREmployee.EmployeeID & " " & PREmployee.DepartmentID, vbExclamation
'                GoBack
'            End If
'            DeptID = PRDepartment.DepartmentID
'            DeptNum = PRDepartment.DepartmentNumber
'        Else
'            DeptID = 0
'            DeptNum = 0
'        End If
        
        
        If rrs![PRHist.DepartmentID] <> 0 Then
            If Not PRDepartment.GetByID(rrs![PRHist.DepartmentID]) Then
                MsgBox "Department NF: " & PREmployee.EmployeeID & " " & rrs![PRHist.DepartmentID], vbExclamation
                GoBack
            End If
            DeptID = PRDepartment.DepartmentID
            DeptNum = PRDepartment.DepartmentNumber
        Else
            DeptID = 0
            DeptNum = 0
        End If
        
        ChkRegUpdateTotals PREquate.GLTypeEmployee, 999999998, 999999998
        ' ChkRegUpdateTotals PREquate.GLTypeDept, rrs![PREmployee.DepartmentID], rrs!DepartmentNumber
        ChkRegUpdateTotals PREquate.GLTypeDept, DeptID, DeptNum
        ChkRegUpdateTotals PREquate.GLTypeCompany, 999999999, 999999999
      
      ' PRINT DETAIL   ##############################################################
    
        TotTaxes = rrs!SSTax + rrs!MedTax + rrs!FWTTax + rrs!SWTTax + rrs!CWTTax
        TotHours = rrs!RegHours + rrs!OTHours + rrs!OEHours
        frmProgress.lblMsg2 = "Employee: " & rrs!EmployeeNumber & " - " & Trim(rrs!LastName) & ", " & (rrs!FirstName)
        frmProgress.Show
        
        If frmCheckReg.chkTotalsOnly Then
            GoTo SkipDetail
        Else
            PrintValue(1) = String(145, "-"):                           FormatString(1) = "a145"
            PrintValue(2) = " ":                                        FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1

            PrintValue(1) = rrs!CheckNumber:                            FormatString(1) = "n9"  ' ***
            PrintValue(2) = " ":                                        FormatString(2) = "a1"
            
            If OptDate = "CHECK DATE" Then
                PrintValue(3) = Format(rrs!CheckDate, "mm/dd/yyyy"):        FormatString(3) = "a10"
            Else
                PrintValue(3) = Format(rrs!PEDate, "mm/dd/yyyy"):           FormatString(3) = "a10"
            End If
            
            PrintValue(4) = DeptNum:                                    FormatString(4) = "n3"
            PrintValue(5) = " ":                                        FormatString(5) = "a1"
            X = rrs!EmployeeNumber & "-" & Trim(rrs!LastName) & ", " & rrs!FirstName
            PrintValue(6) = X:                                          FormatString(6) = "a33"
            PrintValue(7) = rrs!RegAmount:                              FormatString(7) = "d0"
            PrintValue(8) = rrs!OTAmount:                               FormatString(8) = "d0"
            PrintValue(9) = rrs!OEAmount:                               FormatString(9) = "d0"
            PrintValue(10) = rrs!Gross:                                 FormatString(10) = "d0"
            PrintValue(11) = rrs!Deductions:                            FormatString(11) = "d0"
            PrintValue(12) = rrs!Net:                                   FormatString(12) = "d0"
            PrintValue(13) = " ":                                       FormatString(13) = "~"
    
            If TextFileName <> "" Then
                Write #TextChannel2, PrintValue(1);
                Write #TextChannel2, PrintValue(3);
                Write #TextChannel2, PrintValue(4);
                Write #TextChannel2, rrs!EmployeeNumber;
                Write #TextChannel2, Trim(rrs!LastName) & ", " & Trim(rrs!FirstName);
                Write #TextChannel2, PrintValue(7);
                Write #TextChannel2, PrintValue(8);
                Write #TextChannel2, PrintValue(9);
                Write #TextChannel2, PrintValue(10);
                Write #TextChannel2, PrintValue(11);
                Write #TextChannel2, PrintValue(12);
            End If
            FormatPrint
            Ln = Ln + 1
                        
            PrintValue(1) = " ":                                        FormatString(1) = "a1"
            PrintValue(2) = rrs!SSTax:                                  FormatString(2) = "d0"
            PrintValue(3) = rrs!MedTax:                                 FormatString(3) = "d0"
            PrintValue(4) = rrs!FWTTax:                                 FormatString(4) = "d0"
            PrintValue(5) = rrs!SWTTax:                                 FormatString(5) = "d0"
            PrintValue(6) = rrs!CWTTax:                                 FormatString(6) = "d0"
            PrintValue(7) = TotTaxes:                                   FormatString(7) = "d0"
            PrintValue(8) = rrs!RegHours:                               FormatString(8) = "d0"
            PrintValue(9) = rrs!OTHours:                                FormatString(9) = "d0"
            PrintValue(10) = rrs!OEHours:                               FormatString(10) = "d0"
            PrintValue(11) = TotHours:                                  FormatString(11) = "d0"
            PrintValue(12) = " ":                                       FormatString(12) = "~"
            
            If TextFileName <> "" Then
                Write #TextChannel2, PrintValue(2);
                Write #TextChannel2, PrintValue(3);
                Write #TextChannel2, PrintValue(4);
                Write #TextChannel2, PrintValue(5);
                Write #TextChannel2, PrintValue(6);
                Write #TextChannel2, PrintValue(7);
                Write #TextChannel2, PrintValue(8);
                Write #TextChannel2, PrintValue(9);
                Write #TextChannel2, PrintValue(10);
                Write #TextChannel2, PrintValue(11);
            End If

            FormatPrint
            Ln = Ln + 1
        End If
            
SkipDetail:


        ' ======================================================================================
        ' get OE DED info, print and update totals
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & rrs!HistID & _
                    " AND PRDist.DistType = " & PREquate.DistTypeItem
        
        If PRDist.GetBySQL(SQLString) Then
            Do
                
                ' update the ODT records
                ChkRegUpdateODT ODTTypeHist, 0, ODTLineHr, PRDist.EmployerItemID, PRDist.Hours
                ChkRegUpdateODT ODTTypeHist, 0, ODTLineOE, PRDist.EmployerItemID, PRDist.Amount
                ChkRegUpdateODT ODTTypeEE, 0, ODTLineHr, PRDist.EmployerItemID, PRDist.Hours
                ChkRegUpdateODT ODTTypeEE, 0, ODTLineOE, PRDist.EmployerItemID, PRDist.Amount
                ChkRegUpdateODT ODTTypeER, 0, ODTLineHr, PRDist.EmployerItemID, PRDist.Hours
                ChkRegUpdateODT ODTTypeER, 0, ODTLineOE, PRDist.EmployerItemID, PRDist.Amount
                If rrs![PRHist.DepartmentID] <> 0 Then
                    ChkRegUpdateODT ODTTypeDpt, rrs![PRHist.DepartmentID], ODTLineHr, PRDist.EmployerItemID, PRDist.Hours
                    ChkRegUpdateODT ODTTypeDpt, rrs![PRHist.DepartmentID], ODTLineOE, PRDist.EmployerItemID, PRDist.Amount
                End If
                If Not PRDist.GetNext Then Exit Do
            
            Loop
        
        End If
        
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & rrs!HistID & _
                    " AND (PRItemHist.ItemType = " & PREquate.ItemTypeDED & _
                    " OR PRItemHist.ItemType = " & PREquate.ItemTypeDirDepDed & _
                    " OR PRItemHist.ItemType = " & PREquate.ItemTypeSDTax & ")"
        
        If PRItemHist.GetBySQL(SQLString) = True Then
            Do
                
                ' ****************************************************************
                ' Escott patch 10/01/09 - missing EER ItemID in PRItemHist ???
                If PRItemHist.ItemType <> PREquate.ItemTypeDirDepDed And PRItemHist.EmployerItemID = 0 Then
                    If PRItem.GetByID(PRItemHist.ItemID) Then
                        PRItemHist.EmployerItemID = PRItem.EmployerItemID
                        PRItemHist.Save (Equate.RecPut)
                    End If
                End If
                ' ****************************************************************
                
                ' update the ODT records
                If PRItemHist.ItemType = PREquate.ItemTypeDED Or PRItemHist.ItemType = PREquate.ItemTypeSDTax Then
                    ChkRegUpdateODT ODTTypeHist, 0, ODTLineDed, PRItemHist.EmployerItemID, PRItemHist.Amount
                    ChkRegUpdateODT ODTTypeEE, 0, ODTLineDed, PRItemHist.EmployerItemID, PRItemHist.Amount
                    ChkRegUpdateODT ODTTypeER, 0, ODTLineDed, PRItemHist.EmployerItemID, PRItemHist.Amount
                    If rrs![PRHist.DepartmentID] <> 0 Then
                        ChkRegUpdateODT ODTTypeDpt, rrs![PRHist.DepartmentID], ODTLineDed, PRItemHist.EmployerItemID, PRItemHist.Amount
                    End If
                Else        ' dir dep total
                    ChkRegUpdateODT ODTTypeHist, 0, ODTLineDed, 999999999, PRItemHist.Amount
                    ChkRegUpdateODT ODTTypeEE, 0, ODTLineDed, 999999999, PRItemHist.Amount
                    ChkRegUpdateODT ODTTypeER, 0, ODTLineDed, 999999999, PRItemHist.Amount
                    If rrs![PRHist.DepartmentID] <> 0 Then
                        ChkRegUpdateODT ODTTypeDpt, rrs![PRHist.DepartmentID], ODTLineDed, 999999999, PRItemHist.Amount
                    End If
                    DirDepTl = DirDepTl + PRItemHist.Amount
                End If
                
                If Not PRItemHist.GetNext Then Exit Do
            Loop
        End If
                    
        CheckTotal = CheckTotal + rrs!Net
        
        If frmCheckReg.chkTotalsOnly Then           ' Skip printing report body if Totals ONly
        Else
            ' >>>> print report body amounts
            ChkRegPrintODT ODTTypeHist, 0, "", True
            If TextFileName <> "" Then Write #TextChannel2,
        End If
        ' clear the amounts for the body of the report
        ChkRegClearODT ODTTypeHist, 0
        
                ' ======================================================================================
NextEmp:
        rrs.MoveNext
        If rrs.EOF Then
            ChkEmp = 1
            Exit Do
        End If

    Loop

    If Ln >= MaxLines - LineCount Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        ChkRegHeader
        FormatPrint
        Ln = Ln + 1
    End If

    ' last employee subtotal
    If frmCheckReg.chkEESubTotal And EECount >= 1 And PRTotal.tFind(PREquate.GLTypeEmployee, 999999998) Then
        If PREmployee.GetByID(LastEE) Then
            ChkRegPrtTotals PREmployee.EmployeeNumber & "-" & PREmployee.LFName
        Else
            ChkRegPrtTotals "Employee ID Not Found: " & LastEE
        End If
        ChkRegPrintODT ODTTypeEE, 0, "", False
    End If

    ' pg feed before totals?
    If frmCheckReg.chkSepTotPg Then
        FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        ChkRegHeader
        FormatPrint
        Ln = Ln + 1
    End If

    If Ln >= MaxLines - LineCount Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        ChkRegHeader
        FormatPrint
        Ln = Ln + 1

    End If

    ' turn off output file - don't export totals
    If TextFileName <> "" Then
        TextFileName = ""
        TextChannel2 = 0
    End If

    ' print dept totals
    SQLString = "SELECT * FROM PRDepartment ORDER BY DepartmentNumber"
    If PRDepartment.GetBySQL(SQLString) Then
        Do
            If PRTotal.tFind(PREquate.GLTypeDept, PRDepartment.DepartmentID) Then
                ChkRegPrtTotals PRDepartment.DepartmentNumber & " " & PRDepartment.Name
                ChkRegPrintODT ODTTypeDpt, PRDepartment.DepartmentID, "", False
            End If
            If Not PRDepartment.GetNext Then Exit Do
        Loop
    End If
    
    If Ln >= MaxLines - LineCount Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        ChkRegHeader
    End If

    ' print grand total
    If PRTotal.tFind(PREquate.GLTypeCompany, 999999999) Then
        ChkRegPrtTotals PRCompany.Name & " TOTALS"
        ChkRegPrintODT ODTTypeER, 0, "", False
    End If

    ' print net pay / dir dep total
    If DirDepFlag = True And DirDepTl > 0 Then
        If MaxLines - Ln < 8 Then
            FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            ChkRegHeader
            If frmCheckReg.chkTotalsOnly Then
                PrintValue(1) = String(145, "="):   FormatString(1) = "a145"
                PrintValue(2) = " ":                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
        End If
        
        Ln = Ln + 2
        Prvw.vsp.Font.Bold = True
        
        PrintValue(1) = "":                         FormatString(1) = "a20"
        PrintValue(2) = "NET PAY BREAKDOWN":        FormatString(2) = "a20"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        Prvw.vsp.Font.Bold = False
        
        PrintValue(1) = "":                         FormatString(1) = "a20"
        PrintValue(2) = "CHECKS NET:":              FormatString(2) = "a20"
        PrintValue(3) = CheckTotal:                 FormatString(3) = "d0"
        PrintValue(4) = " ":                        FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1

        PrintValue(1) = "":                         FormatString(1) = "a20"
        PrintValue(2) = "DIRECT DEP TOTAL:":        FormatString(2) = "a20"
        PrintValue(3) = DirDepTl:                   FormatString(3) = "d0"
        PrintValue(4) = " ":                        FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1

        Prvw.vsp.Font.Bold = True
        
        PrintValue(1) = "":                         FormatString(1) = "a20"
        PrintValue(2) = "TOTAL NET PAY:":           FormatString(2) = "a20"
        PrintValue(3) = DirDepTl + CheckTotal:      FormatString(3) = "d0"
        PrintValue(4) = " ":                        FormatString(4) = "~"
        FormatPrint
        Ln = Ln + 1

        Prvw.vsp.Font.Bold = False
    
    End If

    ' print state splits if necessary
    If rsState.RecordCount > 1 Then
        
        ' form feed?
        If Ln >= MaxLines - rsState.RecordCount + 2 Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            ChkRegHeader
        End If

        Ln = Ln + 1
        PrintValue(1) = "State Distribution":       FormatString(1) = "a20"
        PrintValue(2) = "State":                    FormatString(2) = "a10"
        PrintValue(3) = " ":                        FormatString(3) = "a3"
        PrintValue(4) = PadRight("Gross Wage", 14): FormatString(4) = "a14"
        PrintValue(5) = PadRight("State Wage", 14): FormatString(5) = "a14"
        PrintValue(6) = PadRight("State Tax", 14):  FormatString(6) = "a14"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1

        rsState.MoveFirst
        Do

            PRState.Clear
            If PRState.GetByID(rsState!StateID) Then
            End If

            PrintValue(1) = "":                         FormatString(1) = "a20"
            PrintValue(2) = Left(PRState.StateName, 10): FormatString(2) = "a10"
            PrintValue(3) = " ":                        FormatString(3) = "a3"
            PrintValue(4) = rsState!Gross:              FormatString(4) = "d0"
            PrintValue(5) = rsState!StWage:             FormatString(5) = "d0"
            PrintValue(6) = rsState!SWT:                FormatString(6) = "d0"
            PrintValue(7) = " ":                        FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1

            rsState.MoveNext
            If rsState.EOF Then Exit Do

        Loop
    
    End If

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub ChkRegGetHeaderData()
    
Dim HdrCount As Integer
    
    HdrCount = 0
    
    ' OE Ded Total Type
    ODTTypeHist = 1         ' for the body of the report
    ODTTypeEE = 2           ' for the employee subtotal option
    ODTTypeDpt = 3          ' department totals
    ODTTypeER = 4           ' company totals
    
    ' OE DED Line Type
    ODTLineHr = 1           ' hours line
    ODTLineOE = 2           ' oe amount line
    ODTLineDed = 3          ' deduction line
    
    trsODT.CursorLocation = adUseClient

    trsODT.Fields.Append "TotalType", adInteger
    trsODT.Fields.Append "TotalID", adInteger           ' dept id - zero for Hist/EE/EER
    trsODT.Fields.Append "LineType", adInteger
    trsODT.Fields.Append "ItemID", adDouble
    trsODT.Fields.Append "Title", adVarChar, 14, adFldIsNullable
    trsODT.Fields.Append "Amount", adCurrency
    trsODT.Fields.Append "Active", adInteger

    trsODT.Open , , adOpenDynamic, adLockOptimistic
    
    ' get the employer items
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeOE & " OR " & _
                " PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeSDTax & ")"
                
    If PRItem.GetBySQL(SQLString) Then
        
        Do
            ChkRegAddTotalRecs ODTTypeHist, 0
            ChkRegAddTotalRecs ODTTypeEE, 0
            ChkRegAddTotalRecs ODTTypeER, 0
                
            SQLString = "SELECT * FROM PRDepartment"
            If PRDepartment.GetBySQL(SQLString) Then
                Do
                    ChkRegAddTotalRecs ODTTypeDpt, PRDepartment.DepartmentID
                    If Not PRDepartment.GetNext Then Exit Do
                Loop
            End If
            HdrCount = HdrCount + 1
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If

    ' add dir dep to deductions?
    If DirDepFlag Then
        PRItem.Title = "DIR DEPOSIT"
        PRItem.Abbreviation = "DIR DEPOSIT"
        PRItem.ItemType = PREquate.ItemTypeDED
        PRItem.EmployerItemID = 999999999
        PRItem.ItemID = 999999999
        PRItem.Active = 1
        ChkRegAddTotalRecs ODTTypeHist, 0
        ChkRegAddTotalRecs ODTTypeEE, 0
        ChkRegAddTotalRecs ODTTypeER, 0
        SQLString = "SELECT * FROM PRDepartment"
        If PRDepartment.GetBySQL(SQLString) Then
            Do
                ChkRegAddTotalRecs ODTTypeDpt, PRDepartment.DepartmentID
                If Not PRDepartment.GetNext Then Exit Do
            Loop
        End If
        HdrCount = HdrCount + 1
    End If

    ' no items for this company
    If HdrCount = 0 Then ODTFlag = False

End Sub

Public Sub ChkRegHeader()
    Ln = Ln + 1
    
    PrintValue(1) = "CHECK NUM":                        FormatString(1) = "a10"
    
    If OptDate = "CHECK DATE" Then
        PrintValue(2) = "CHECK DATE":                   FormatString(2) = "a12"
    Else
        PrintValue(2) = "P/E DATE":                     FormatString(2) = "a10"
    End If
    
    PrintValue(3) = "DPT":                              FormatString(3) = "a4"
    PrintValue(4) = "EMPLOYEE #/NAME":                  FormatString(4) = "a33"
    PrintValue(5) = "REG PAY ":                         FormatString(5) = "r14"
    PrintValue(6) = "OT PAY ":                          FormatString(6) = "r14"
    PrintValue(7) = "OTH PAY ":                         FormatString(7) = "r14"
    PrintValue(8) = "GROSS PAY ":                       FormatString(8) = "r14"
    PrintValue(9) = "TOT DED ":                         FormatString(9) = "r14"
    PrintValue(10) = "NET PAY ":                        FormatString(10) = "r14"
    PrintValue(11) = " ":                               FormatString(11) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "SS TAX ":                          FormatString(1) = "r15"
    PrintValue(2) = "MED TAX ":                         FormatString(2) = "r14"
    PrintValue(3) = "FWT TAX ":                         FormatString(3) = "r14"
    PrintValue(4) = "SWT TAX ":                         FormatString(4) = "r14"
    PrintValue(5) = "CWT TAX ":                         FormatString(5) = "r14"
    PrintValue(6) = "TOT TAXES ":                       FormatString(6) = "r14"
    PrintValue(7) = "REG HRS ":                         FormatString(7) = "r14"
    PrintValue(8) = "OT HRS ":                          FormatString(8) = "r14"
    PrintValue(9) = "OTH HRS ":                         FormatString(9) = "r14"
    PrintValue(10) = "TOT HRS ":                        FormatString(10) = "r14"
    PrintValue(11) = " ":                               FormatString(11) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    If OptDate = "CHECK DATE" Then
        PrintValue(2) = "CHECK DATE":                   FormatString(2) = "a12"
    Else
        PrintValue(2) = "P/E DATE":                     FormatString(2) = "a10"
    End If
    
    If Pg = 1 And TextFileName <> "" Then
        Write #TextChannel2, "CHECK NUM"; PrintValue(2); "DEPT"; "EMPLOYEE #"; "EMPLOYEE NAME"; "REG PAY";
        Write #TextChannel2, "OT PAY"; "OTH PAY"; "GROSS PAY"; "TOT DED"; "Net PAY"; "SS TAX";
        Write #TextChannel2, "MED TAX", "FWT TAX"; "SWT TAX"; "CWT TAX"; "TOT TAXES"; ; "REG HRS";
        Write #TextChannel2, "OT HRS"; "OTH HRS"; "TOT HRS";
    End If

    ' OE / DED Headers
    If ODTFlag = True Then
        X = " "
        LastLine = 0
        ItemCount = 0
        If frmCheckReg.chkIncInactiveItems Then
            trsODT.Filter = "TotalType = 1"
        Else
            trsODT.Filter = "TotalType = 1 AND Active = 1"
        End If
        trsODT.Sort = "LineType, ItemID"
        trsODT.MoveFirst
        
        Do Until trsODT.EOF
            
            ItemCount = ItemCount + 1
            
            If (LastLine <> 0 And LastLine <> trsODT!LineType) Or ItemCount > ItemMax Then
                PrintFlag = False
                If frmCheckReg.chkOEHrs And LastLine = 1 Then PrintFlag = True
                If frmCheckReg.chkOEAmt And LastLine = 2 Then PrintFlag = True
                If frmCheckReg.chkDed And LastLine = 3 Then PrintFlag = True
                If PrintFlag = True Then
                    PrintValue(1) = X:                      FormatString(1) = "a141"
                    PrintValue(2) = "":                     FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                End If
                X = " "
                ItemCount = 1
            End If
            LastLine = trsODT!LineType
            
            X = X & Space(13 - Len(Trim(trsODT!Title))) & trsODT!Title & " "

            If Pg = 1 And TextFileName <> "" Then
                Write #TextChannel2, trsODT!Title;
            End If
            trsODT.MoveNext
        Loop
        If X <> " " Then
            PrintFlag = False
            If frmCheckReg.chkOEHrs And LastLine = 1 Then PrintFlag = True
            If frmCheckReg.chkOEAmt And LastLine = 2 Then PrintFlag = True
            If frmCheckReg.chkDed And LastLine = 3 Then PrintFlag = True
            If PrintFlag = True Then
                PrintValue(1) = X:          FormatString(1) = "a141"
                PrintValue(2) = "":         FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
            X = " "
        End If
        
    End If
    
'    PrintValue(1) = String(145, "="):                   FormatString(1) = "a145"
'    PrintValue(2) = " ":                                FormatString(2) = "~"
'    FormatPrint
'    Ln = Ln + 1
    
    If Pg = 1 And TextFileName <> "" Then
        Write #TextChannel2,
    End If
    
    trsODT.Filter = ""
    
End Sub

Private Sub ChkRegUpdateTotals(ByVal RecType As Byte, ByVal RecID As Long, ByVal IDNumber As Long)
            
    If Not PRTotal.tFind(RecType, RecID) Then
        PRTotal.Clear
        PRTotal.RecType = RecType
        PRTotal.RecID = RecID
        PRTotal.IDNumber = IDNumber
        PRTotal.Save (Equate.RecAdd)
    End If

    ' PRTotal.DepartmentID = rrs![PREmployee.DepartmentID]
    PRTotal.DepartmentID = IDNumber
    PRTotal.RegHours = PRTotal.RegHours + rrs!RegHours
    PRTotal.RegAmount = PRTotal.RegAmount + rrs!RegAmount
    PRTotal.OTHours = PRTotal.OTHours + rrs!OTHours
    PRTotal.OTAmount = PRTotal.OTAmount + rrs!OTAmount
    PRTotal.OEHours = PRTotal.OEHours + rrs!OEHours
    PRTotal.OEAmount = PRTotal.OEAmount + rrs!OEAmount
    PRTotal.SSWage = PRTotal.SSWage + rrs!SSWage
    PRTotal.SSTax = PRTotal.SSTax + rrs!SSTax
    PRTotal.MEDWage = PRTotal.MEDWage + rrs!MEDWage
    PRTotal.MedTax = PRTotal.MedTax + rrs!MedTax
    PRTotal.FWTWage = PRTotal.FWTWage + rrs!FWTWage
    PRTotal.FWTTax = PRTotal.FWTTax + rrs!FWTTax
    PRTotal.Deductions = PRTotal.Deductions + rrs!Deductions
    PRTotal.StateWage = PRTotal.StateWage + rrs!SWTWage
    PRTotal.StateTax = PRTotal.StateTax + rrs!SWTTax
    PRTotal.CityWage = PRTotal.CityWage + rrs!CWTWage
    PRTotal.CityTax = PRTotal.CityTax + rrs!CWTTax
    PRTotal.Gross = PRTotal.Gross + rrs!Gross
    PRTotal.Net = PRTotal.Net + rrs!Net
    PRTotal.Count = PRTotal.Count + 1
    
    PRTotal.Save (Equate.RecPut)

End Sub

Private Sub ChkRegPrtTotals(ByVal TotalTitle As String)

    If Ln >= MaxLines - LineCount Then
        If Ln Then FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        ChkRegHeader
    End If

    PrintValue(1) = String(145, "="):           FormatString(1) = "a145"
    PrintValue(2) = " ":                        FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

    TTotTaxes = PRTotal.SSTax + PRTotal.MedTax + PRTotal.StateTax + PRTotal.FWTTax + PRTotal.CityTax
    TTotHours = PRTotal.RegHours + PRTotal.OTHours + PRTotal.OEHours
    X = "Count: " & Format(PRTotal.Count, "###,##0"):

    PrintValue(1) = TotalTitle:                 FormatString(1) = "a30"
    PrintValue(2) = X:                          FormatString(2) = "a27"
    PrintValue(3) = PRTotal.RegAmount:          FormatString(3) = "d0"
    PrintValue(4) = PRTotal.OTAmount:           FormatString(4) = "d0"
    PrintValue(5) = PRTotal.OEAmount:           FormatString(5) = "d0"
    PrintValue(6) = PRTotal.Gross:              FormatString(6) = "d0"
    PrintValue(7) = PRTotal.Deductions:         FormatString(7) = "d0"
    PrintValue(8) = PRTotal.Net:                FormatString(8) = "d0"
    PrintValue(9) = " ":                        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a1"
    PrintValue(2) = PRTotal.SSTax:              FormatString(2) = "d0"
    PrintValue(3) = PRTotal.MedTax:             FormatString(3) = "d0"
    PrintValue(4) = PRTotal.FWTTax:             FormatString(4) = "d0"
    PrintValue(5) = PRTotal.StateTax:           FormatString(5) = "d0"
    PrintValue(6) = PRTotal.CityTax:            FormatString(6) = "d0"
    PrintValue(7) = TTotTaxes:                  FormatString(7) = "d0"
    PrintValue(8) = PRTotal.RegHours:           FormatString(8) = "d0"
    PrintValue(9) = PRTotal.OTHours:            FormatString(9) = "d0"
    PrintValue(10) = PRTotal.OEHours:           FormatString(10) = "d0"
    PrintValue(11) = TTotHours:                 FormatString(11) = "d0"
    PrintValue(12) = " ":                       FormatString(12) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PRTotal.Clear
    PRTotal.Save (Equate.RecPut)

End Sub

Private Sub ChkRegAddTotalRecs(ByVal TotalType As Integer, ByVal TotalID As Integer)

    trsODT.AddNew
    trsODT!TotalType = TotalType
    trsODT!TotalID = TotalID
    If PRItem.ItemType = PREquate.ItemTypeOE Then
        trsODT!LineType = ODTLineHr
    Else
        trsODT!LineType = ODTLineDed
    End If
    trsODT!ItemID = PRItem.ItemID
    If PRItem.Abbreviation = "" Then
        trsODT!Title = Mid(PRItem.Title, 1, 13)
    Else
        trsODT!Title = Mid(PRItem.Abbreviation, 1, 13)
    End If
    If PRItem.ItemType = PREquate.ItemTypeOE Then
        trsODT!Title = Mid(trsODT!Title & " HR", 1, 13)
    End If
    trsODT!Amount = 0
    trsODT!Active = PRItem.Active
    trsODT.Update
    
    If PRItem.ItemType = PREquate.ItemTypeOE Then
        trsODT.AddNew
        trsODT!TotalType = TotalType
        trsODT!TotalID = TotalID
        trsODT!LineType = ODTLineOE
        trsODT!ItemID = PRItem.ItemID
        If PRItem.Abbreviation = "" Then
            trsODT!Title = Mid(PRItem.Title, 1, 13)
        Else
            trsODT!Title = Mid(PRItem.Abbreviation, 1, 13)
        End If
        trsODT!Amount = 0
        trsODT!Active = PRItem.Active
        trsODT.Update
    End If
    
    ' keep a count of the OE and deducts used for the report
    If TotalType = ODTTypeHist Then
        If (Not frmCheckReg.chkIncInactiveItems And PRItem.Active) Or frmCheckReg.chkIncInactiveItems Then
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                OECount = OECount + 1
            Else
                DEDCount = DEDCount + 1
            End If
        End If
    End If
End Sub

Private Sub ChkRegUpdateODT(ByVal TotalType As Integer, _
                            ByVal TotalID As Long, _
                            ByVal LineType As Integer, _
                            ByVal ItemID As Long, _
                            ByVal Amount As Currency)
                             
    If ODTFlag = False Then Exit Sub
    
    trsODT.Filter = "TotalType = " & TotalType & _
                    " AND TotalID = " & TotalID & _
                    " AND LineType = " & LineType & _
                    " AND ItemID = " & ItemID
                     
    If trsODT.RecordCount = 0 Then
        MsgBox "TotalType: " & TotalType & vbCr & _
               " TotalID: " & TotalID & vbCr & _
               " LineType: " & LineType & vbCr & _
               " ItemID: " & ItemID & vbCr & _
               " EmployeeID: " & PRHist.EmployeeID & vbCr & _
               " Not Found", vbExclamation
        End
    End If
        
    trsODT!Amount = trsODT!Amount + Amount
    trsODT.Update
    trsODT.Filter = ""

End Sub

Private Sub ChkRegClearODT(ByVal TotalType As Integer, _
                           ByVal TotalID As Long)
                            
    If ODTFlag = False Then Exit Sub
    
    trsODT.Filter = "TotalType = " & TotalType & _
                    " AND TotalID = " & TotalID
                    
    trsODT.MoveFirst
    Do Until trsODT.EOF
        trsODT!Amount = 0
        trsODT.Update
        trsODT.MoveNext
    Loop
    trsODT.Filter = ""

End Sub

Private Sub ChkRegPrintODT(ByVal TotalType As Integer, _
                           ByVal TotalID As Long, _
                           ByVal Title As String, _
                           ByVal TextOutput As Boolean)
                            
    If ODTFlag = False Then
'        PrintValue(1) = String(145, "-"):           FormatString(1) = "a145"
'        PrintValue(2) = " ":                        FormatString(2) = "~"
'        FormatPrint
'        Ln = Ln + 1
        Exit Sub
    End If
    
    LastLine = 0
    ItemCount = 0
    I = 1
    PrintValue(1) = " ":        FormatString(1) = "a1"
                            
    ' print the body of the report data
    If frmCheckReg.chkIncInactiveItems Then
        trsODT.Filter = "TotalType = " & TotalType & " AND TotalID = " & TotalID
    Else
        trsODT.Filter = "TotalType = " & TotalType & " AND TotalID = " & TotalID & _
                        " AND Active = 1"
    End If
    trsODT.Sort = "LineType, ItemID"
    trsODT.MoveFirst
    Do Until trsODT.EOF
        
        ItemCount = ItemCount + 1

        If (LastLine <> 0 And LastLine <> trsODT!LineType) Or ItemCount > ItemMax Then
            PrintFlag = False
            If frmCheckReg.chkOEHrs And LastLine = 1 Then PrintFlag = True
            If frmCheckReg.chkOEAmt And LastLine = 2 Then PrintFlag = True
            If frmCheckReg.chkDed And LastLine = 3 Then PrintFlag = True
            If PrintFlag = True Then
                I = I + 1
                PrintValue(I) = " ":            FormatString(I) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
            ' clear the values
            ItemCount = 1
            I = 1

            PrintValue(1) = " ":                FormatString(1) = "a1"
        End If
        LastLine = trsODT!LineType
        
        I = I + 1
        PrintValue(I) = trsODT!Amount:          FormatString(I) = "d0"

        If TextFileName <> "" And TextOutput = True Then Write #TextChannel2, PrintValue(I);
        trsODT.MoveNext

    Loop
    
    If X <> " " Then
        PrintFlag = False
        If frmCheckReg.chkOEHrs And LastLine = 1 Then PrintFlag = True
        If frmCheckReg.chkOEAmt And LastLine = 2 Then PrintFlag = True
        If frmCheckReg.chkDed And LastLine = 3 Then PrintFlag = True
        If PrintFlag = True Then
            I = I + 1
            PrintValue(I) = " ":        FormatString(I) = "~"
            FormatPrint
            Ln = Ln + 1
            If TextFileName <> "" And TextOutput = True Then Write #TextChannel2, PrintValue(I);
        End If
        X = " "
    End If
    
'    PrintValue(1) = String(145, "-"):           FormatString(1) = "a145"
'    PrintValue(2) = " ":                        FormatString(2) = "~"
'    FormatPrint
'    Ln = Ln + 1

    trsODT.Filter = ""
                            
    ' clear the subtotals
    ChkRegClearODT TotalType, TotalID
                            
End Sub
Public Sub DptDistRpt(ByVal RangeType As Byte, _
                            ByVal BatchNumbr As Long, _
                            ByVal PEDate As Long, _
                            ByVal CheckDt As Long, _
                            ByVal StartDate As Long, _
                            ByVal EndDate As Long, _
                            ByVal OptDate As String)

Dim StartYM As Long
Dim EndYM As Long
Dim DptName As String
Dim DptNumber As Long
Dim LastID As Double
Dim LastDptName As String
Dim LastDptNumber As Long
Dim ReportTitle As String
Dim HourTl, WageTl, TaxTl As Currency
Dim tString As String

    frmDptDist.Hide
    
    ReportTitle = "Payroll Department Distribution Report"
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
            
    trs.CursorLocation = adUseClient
    trs.Fields.Append "TempID", adDouble
    trs.Fields.Append "DptID", adDouble
    trs.Fields.Append "DptNumber", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    trs.Fields.Append "EmployeeNumber", adDouble
    trs.Fields.Append "Hours", adDouble
    trs.Fields.Append "Wage", adCurrency
    trs.Fields.Append "Tax", adCurrency
    
    trs.Open , , adOpenDynamic, adLockOptimistic
           
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    MaxLines = 65
           
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.BatchID = " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRDist WHERE PRDist.CheckDate >= " & StartDate & _
            " AND PRDist.CheckDate <= " & EndDate
            Msg1 = "CHECK DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        Else
            SQLString = "SELECT * FROM PRDist WHERE PRDist.PEDate >= " & StartDate & _
            " AND PRDist.PEDate <= " & EndDate
            Msg1 = "PERIOD ENDING DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If

    If Not PRDist.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbExclamation, "Dpt Dist Tax Report"
        GoBack
    End If
    
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PR Batch Not Found: " & PRBatchID, vbExclamation
            GoBack
        End If
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & PRBatch.PEDate
    End If

    Do
        
        ' employee select filter?
        If frmEmpSelect.AllEmployees = False Then
            Msg2 = frmEmpSelect.Count & " Selected Employees"
            SQLString = "EmployeeID = " & PRDist.EmployeeID
            frmEmpSelect.rsEmp.Find SQLString, 0, adSearchForward, 1
            If frmEmpSelect.rsEmp.EOF Then GoTo NextDist
            If frmEmpSelect.rsEmp!Select = False Then GoTo NextDist
        Else
            Msg2 = "All Employees"
        End If
        
        ' skip non-taxable
        If PRDist.GrossWage = 0 And PRDist.CityTax = 0 Then GoTo NextDist
        
        DptDistAccum PRDist.DepartmentID, PRDist.Hours, PRDist.GrossWage, PRDist.CityTax

NextDist:
        ' trs.MoveFirst
        If Not PRDist.GetNext Then Exit Do
    
    Loop
    
    ' get the EE and City Numbers
    trs.MoveFirst
    Do Until trs.EOF
        If Not PREmployee.GetByID(trs!EmployeeID) Then
            MsgBox "Employee ID Not Found!", vbExclamation
            GoBack
        End If
        trs!EmployeeNumber = PREmployee.EmployeeNumber
        If trs!DptID <> 999999 Then
            If Not PRDepartment.GetByID(trs!DptID) Then
                MsgBox "PRDepartment ID Not Found! " & trs!DptID, vbExclamation
                End
            End If
            trs!DptNumber = PRDepartment.DepartmentNumber
        Else
            trs!DptNumber = 999999
        End If
            
        trs.MoveNext
    Loop
    
    If frmDptDist.optByDpt Then
        trs.Sort = "DptNumber, EmployeeNumber"
    Else
        trs.Sort = "EmployeeNumber, DptNumber"
    End If

    trs.MoveFirst
    LastID = 0
    Ln = 0
    
    If frmDptDist.optByDpt Then
        ReportTitle = "PAYROLL DEPT DISTRIBUTION REPORT BY DEPT"
    Else
        ReportTitle = "PAYROLL DEPT DISTRIBUTION REPORT BY EMPLOYEE"
    End If
    
    Do
        
        If Ln = 0 Or Ln > MaxLines Then
            If Ln <> 0 Then FormFeed
            DptDistHeader (ReportTitle)
            DptDistSubHeader Int(trs!TempID / 10 ^ 6), False
        End If

        If (LastID <> 0 And LastID <> Int(trs!TempID / 10 ^ 6)) Then
            
            If Ln > MaxLines - 5 Then
                FormFeed
                DptDistHeader ReportTitle
            End If
            
            DptDistSubHeader LastID, True
            HourTotal = 0
            WageTotal = 0
            TaxTotal = 0
            Ln = Ln + 1
            
            If frmDptDist.chkFormFeed Then
                FormFeed
                DptDistHeader ReportTitle
            End If
            
            DptDistSubHeader Int(trs!TempID / 10 ^ 6), False
        
        End If
        LastID = Int(trs!TempID / 10 ^ 6)

        RecCount = RecCount + 1
        frmProgress.lblMsg2 = "Dept Dist Report " & PRCompany.Name & " Records Processed: " & Format(RecCount, "#,###,##0")
        frmProgress.Show
        
        HourTotal = HourTotal + trs!Hours
        WageTotal = WageTotal + trs!Wage
        TaxTotal = TaxTotal + trs!Tax
        HourTl = HourTl + trs!Hours
        WageTl = WageTl + trs!Wage
        TaxTl = TaxTl + trs!Tax
        
        ' MOD does not work for large numbers
        TempID = trs!TempID - (Int(trs!TempID / 10 ^ 6) * 10 ^ 6)
        
        If frmDptDist.optByEmployee Then     ' find the department
            If TempID = 999999 Then
                Hdr1 = ""
                Hdr2 = "NON TAX"
            ElseIf Not PRDepartment.GetByID(TempID) Then
                Hdr1 = TempID
                Hdr2 = "DEPT NOT FOUND!"
            Else
                Hdr1 = PRDepartment.DepartmentNumber
                Hdr2 = PRDepartment.Name
            End If
        Else                                    ' by city - get the employee
            If Not PREmployee.GetByID(TempID) Then
                MsgBox "Employee ID Not Found: " & TempID, vbExclamation
                GoBack
            End If
            Hdr1 = PREmployee.EmployeeNumber
            Hdr2 = PREmployee.FLName
        End If
        
        PrintValue(1) = " ":                        FormatString(1) = "a6"
        PrintValue(2) = Hdr1:                       FormatString(2) = "r10"
        PrintValue(3) = " ":                        FormatString(3) = "a2"
        PrintValue(4) = Hdr2:                       FormatString(4) = "a21"
        PrintValue(5) = " ":                        FormatString(5) = "a2"
        PrintValue(6) = trs!Hours:                  FormatString(6) = "d0"
        PrintValue(7) = " ":                        FormatString(7) = "a2"
        PrintValue(8) = trs!Wage:                   FormatString(8) = "d0"
        PrintValue(9) = " ":                        FormatString(9) = "a2"
        PrintValue(10) = trs!Tax:                   FormatString(10) = "d0"
        PrintValue(11) = " ":                       FormatString(11) = "~"
        FormatPrint
        Ln = Ln + 1
        
        trs.MoveNext
        
        If trs.EOF Then
            Exit Do
        End If
    
    Loop
    
    ' print the last subtotal
    If MaxLines - Ln <= 4 Then
        FormFeed
        DptDistHeader ReportTitle
    End If
    DptDistSubHeader LastID, True
    Ln = Ln + 1
    
    ' grand totals
    If frmDptDist.chkFormFeed Then
        FormFeed
        DptDistHeader ReportTitle
    End If
    HourTotal = HourTl
    WageTotal = WageTl
    TaxTotal = TaxTl
    DptDistSubHeader 0, True
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub
Private Sub DptDistAccum(ByVal DptID As Long, _
                         ByVal DptHr As Double, _
                         ByVal DptWage As Currency, _
                         ByVal CtyTax As Currency)

Dim TempID As Double
Dim FFlag As Boolean
        
    If DptID = 0 Then DptID = 999999
    
    If frmDptDist.optByDpt Then
        TempID = DptID * 10 ^ 6 + PRDist.EmployeeID
    Else
        TempID = PRDist.EmployeeID * 10 ^ 6 + DptID
    End If
    
    FFlag = False
    
    If trs.RecordCount > 0 Then
        trs.MoveFirst
        Do
            If trs!TempID = TempID Then
                FFlag = True
                trs!Hours = trs!Hours + DptHr
                trs!Wage = trs!Wage + DptWage
                trs!Tax = trs!Tax + CtyTax
                trs.Update
                Exit Do
            End If
            trs.MoveNext
        Loop Until trs.EOF
    End If
    
    If FFlag = False Then
        trs.AddNew
        trs!TempID = TempID
        trs!DptID = DptID
        trs!EmployeeID = PRDist.EmployeeID
        trs!Hours = DptHr
        trs!Wage = DptWage
        trs!Tax = CtyTax
        trs.Update
    End If
    
End Sub
Private Sub DptDistHeader(ReportTitle)
 
    PageHeader ReportTitle, Msg1, Msg2, ""
    Ln = Ln + 1
    
    If frmDptDist.optByEmployee Then
        Hdr1 = "DPT NUMBER"
        Hdr2 = "DEPT NAME"
    Else
        Hdr1 = "EMP NUMBER"
        Hdr2 = "EMP NAME"
    End If
    
    PrintValue(1) = " ":                    FormatString(1) = "a6"
    PrintValue(2) = Hdr1:                   FormatString(2) = "a10"
    PrintValue(3) = " ":                    FormatString(3) = "a2"
    PrintValue(4) = Hdr2:                   FormatString(4) = "a21"
    PrintValue(5) = " ":                    FormatString(5) = "a2"
    PrintValue(6) = "HOURS ":               FormatString(6) = "r14"
    PrintValue(7) = " ":                    FormatString(7) = "a2"
    PrintValue(8) = "GROSS WAGE ":          FormatString(8) = "r14"
    PrintValue(9) = " ":                    FormatString(9) = "a2"
    PrintValue(10) = "CITY TAX ":           FormatString(10) = "r14"
    PrintValue(11) = " ":                   FormatString(11) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(98, "-"):        FormatString(1) = "a86"
    PrintValue(2) = " ":                    FormatString(2) = "~"
    
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub DptDistSubHeader(ByVal ID As Double, _
                             ByVal PrintTotal As Boolean)

    Prvw.vsp.Font.Bold = True

    If ID <> 0 Then
        If frmDptDist.optByDpt Then
            If ID = 999999 Then
                Hdr1 = ""
                Hdr2 = "NO DEPT"
            ElseIf Not PRDepartment.GetByID(ID) Then
                Hdr1 = ID
                Hdr2 = "DEPT NOT FOUND!"
            Else
                Hdr1 = PRDepartment.DepartmentNumber
                Hdr2 = PRDepartment.Name
            End If
        Else
            If Not PREmployee.GetByID(ID) Then
                MsgBox "Employee ID Not Found: " & ID, vbExclamation
                GoBack
            End If
            Hdr1 = PREmployee.EmployeeNumber
            Hdr2 = PREmployee.FLName
        End If
    Else
        Hdr1 = ""
        Hdr2 = "COMPANY TOTALS"
    End If
    
    If PrintTotal Then
        PrintValue(1) = " TOTAL":               FormatString(1) = "a6"
    Else
        PrintValue(1) = " ":                    FormatString(1) = "a6"
    End If
    
    PrintValue(2) = Hdr1:                   FormatString(2) = "r10"
    PrintValue(3) = " ":                    FormatString(3) = "a2"
    PrintValue(4) = Hdr2:                   FormatString(4) = "a23"
    
    If PrintTotal Then
        PrintValue(5) = HourTotal:          FormatString(5) = "d0"
        PrintValue(6) = " ":                FormatString(6) = "a2"
        PrintValue(7) = WageTotal:          FormatString(7) = "d0"
        PrintValue(8) = " ":                FormatString(8) = "a2"
        
        PrintValue(9) = TaxTotal:           FormatString(9) = "d0"
        PrintValue(10) = " ":               FormatString(10) = "~"
    Else
        PrintValue(5) = " ":                FormatString(5) = "~"
    End If
    FormatPrint
    Ln = Ln + 1
    
    Prvw.vsp.Font.Bold = False

End Sub

Public Sub CityTaxRpt(ByVal RangeType As Byte, _
                            ByVal BatchNumbr As Long, _
                            ByVal PEDate As Long, _
                            ByVal CheckDt As Long, _
                            ByVal StartDate As Long, _
                            ByVal EndDate As Long, _
                            ByVal OptDate As String)

Dim StartYM As Long
Dim EndYM As Long
Dim CityName As String
Dim CityNumber As Long
Dim LastID As Double
Dim LastCityName As String
Dim LastCityNumber As Long
Dim ReportTitle As String
Dim WageTl, TaxTl As Currency
Dim tString As String
Dim LastCourt As Integer

    frmCityTaxRpt.Hide
    SetEquates
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
            
    trs.CursorLocation = adUseClient
    trs.Fields.Append "TempID", adDouble
    trs.Fields.Append "CityID", adDouble
    trs.Fields.Append "CityNumber", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    trs.Fields.Append "EmployeeNumber", adDouble
    trs.Fields.Append "Wage", adCurrency
    trs.Fields.Append "Tax", adCurrency
    trs.Fields.Append "Courtesy", adInteger
    
    trs.Open , , adOpenDynamic, adLockOptimistic
           
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    MaxLines = 65
           
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.BatchID = " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRDist WHERE PRDist.CheckDate >= " & StartDate & _
            " AND PRDist.CheckDate <= " & EndDate
            Msg1 = "CHECK DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        Else
            SQLString = "SELECT * FROM PRDist WHERE PRDist.PEDate >= " & StartDate & _
            " AND PRDist.PEDate <= " & EndDate
            Msg1 = "PERIOD ENDING DATE FROM: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If

    If Not PRDist.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbExclamation, "City Tax Report"
        GoBack
    End If
    
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PR Batch Not Found: " & PRBatchID, vbExclamation
            GoBack
        End If
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & PRBatch.PEDate
    End If

    Do
        
        ' employee select filter?
        If frmEmpSelect.AllEmployees = False Then
            Msg2 = frmEmpSelect.Count & " Selected Employees"
            SQLString = "EmployeeID = " & PRDist.EmployeeID
            frmEmpSelect.rsEmp.Find SQLString, 0, adSearchForward, 1
            If frmEmpSelect.rsEmp.EOF Then GoTo NextDist
            If frmEmpSelect.rsEmp!Select = False Then GoTo NextDist
        Else
            Msg2 = "All Employees"
        End If
        
        ' skip non-taxable
        If PRDist.CityWage = 0 And PRDist.CityTax = 0 Then GoTo NextDist
        
        CityTaxAccum PRDist.CityID, PRDist.CityWage, PRDist.CityTax, 0
        If PRDist.CourtesyCityTax <> 0 Then
            CityTaxAccum PRDist.CourtesyCityID, PRDist.CityWage, PRDist.CourtesyCityTax, 1
        End If

NextDist:
        ' trs.MoveFirst
        If Not PRDist.GetNext Then Exit Do
    Loop
    
    ' get the EE and City Numbers
    trs.MoveFirst
    Do Until trs.EOF
        If Not PREmployee.GetByID(trs!EmployeeID) Then
            MsgBox "Employee ID Not Found!", vbExclamation
            GoBack
        End If
        trs!EmployeeNumber = PREmployee.EmployeeNumber
        If trs!CityID <> 999999 Then
            If Not PRCity.GetByID(trs!CityID) Then
                MsgBox "PRCity ID Not Found! " & trs!CityID, vbExclamation
                End
            End If
            trs!CityNumber = PRCity.CityNumber
        Else
            trs!CityNumber = 999999
        End If
            
        trs.MoveNext
    Loop
    
    If frmCityTaxRpt.optByCity Then
        trs.Sort = "CityNumber, Courtesy, EmployeeNumber"
    Else
        trs.Sort = "EmployeeNumber, CityNumber, Courtesy"
    End If

    trs.MoveFirst
    LastID = 0
    Ln = 0
    
    If frmCityTaxRpt.optByCity Then
        ReportTitle = "PAYROLL CITY TAX REPORT BY CITY"
    Else
        ReportTitle = "PAYROLL CITY TAX REPORT BY EMPLOYEE"
    End If
    
    LastCourt = -1
    
    Do
        
        If Ln = 0 Or Ln > MaxLines Then
            If Ln <> 0 Then FormFeed
            CityTaxHeader (ReportTitle)
            CityTaxSubHeader Int(trs!TempID / 10 ^ 6), False, trs!Courtesy
        End If

        If (LastID <> 0 And LastID <> Int(trs!TempID / 10 ^ 6)) _
            Or (LastCourt <> -1 And trs!Courtesy <> LastCourt) Then
            
            If Ln > MaxLines - 5 Then
                FormFeed
                CityTaxHeader ReportTitle
            End If
            
            CityTaxSubHeader LastID, True, LastCourt
            WageTotal = 0
            TaxTotal = 0
            Ln = Ln + 1
            
            If frmCityTaxRpt.chkFormFeed Then
                FormFeed
                CityTaxHeader ReportTitle
            End If
            
            CityTaxSubHeader Int(trs!TempID / 10 ^ 6), False, trs!Courtesy
        
        End If
        LastID = Int(trs!TempID / 10 ^ 6)
        LastCourt = trs!Courtesy

        RecCount = RecCount + 1
        frmProgress.lblMsg2 = "City Tax Report " & PRCompany.Name & " Records Processed: " & Format(RecCount, "#,###,##0")
        frmProgress.Show
        
        WageTotal = WageTotal + trs!Wage
        TaxTotal = TaxTotal + trs!Tax
        WageTl = WageTl + trs!Wage
        TaxTl = TaxTl + trs!Tax
        
        ' MOD does not work for large numbers
        TempID = trs!TempID - (Int(trs!TempID / 10 ^ 6) * 10 ^ 6)
        
        If frmCityTaxRpt.optByEmployee Then     ' find the city
            If TempID = 999999 Then
                Hdr1 = ""
                Hdr2 = "NON TAX"
            ElseIf Not PRCity.GetByID(TempID) Then
                Hdr1 = TempID
                Hdr2 = "CITY NOT FOUND!"
            Else
                Hdr1 = PRCity.CityNumber
                Hdr2 = PRCity.ShortName
                If PRState.GetByID(PRCity.StateID) Then
                    Hdr2 = Trim(Hdr2) & " " & PRState.StateAbbrev
                End If
            End If
        Else                                    ' by city - get the employee
            If Not PREmployee.GetByID(TempID) Then
                MsgBox "Employee ID Not Found: " & TempID, vbExclamation
                GoBack
            End If
            Hdr1 = PREmployee.EmployeeNumber
            Hdr2 = PREmployee.FLName
        End If
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = Hdr1:                       FormatString(2) = "r10"
        PrintValue(3) = " ":                        FormatString(3) = "a3"
        PrintValue(4) = Hdr2:                       FormatString(4) = "a35"
        PrintValue(5) = trs!Wage:                   FormatString(5) = "d0"
        PrintValue(6) = " ":                        FormatString(6) = "a5"
        PrintValue(7) = trs!Tax:                    FormatString(7) = "d0"
        PrintValue(8) = " ":                        FormatString(8) = "~"
        FormatPrint
        Ln = Ln + 1
        
        trs.MoveNext
        
        If trs.EOF Then
            Exit Do
        End If
    
    Loop
    
    ' print the last subtotal
    If MaxLines - Ln <= 4 Then
        FormFeed
        CityTaxHeader ReportTitle
    End If
    CityTaxSubHeader LastID, True, LastCourt
    Ln = Ln + 1
    
    ' grand totals
    If frmCityTaxRpt.chkFormFeed Then
        FormFeed
        CityTaxHeader ReportTitle
    End If
    WageTotal = WageTl
    TaxTotal = TaxTl
    CityTaxSubHeader 0, True, 0
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub
Private Sub CityTaxAccum(ByVal CtyID As Long, _
                         ByVal CtyWage As Currency, _
                         ByVal CtyTax As Currency, _
                         ByVal Courtesy As Integer)

Dim TempID As Double
Dim FFlag As Boolean
        
    ' separate courtesy tax?
    If frmCityTaxRpt.chkSepCourtesy = 0 Then Courtesy = 0
        
    If CtyID = 0 Then CtyID = 999999
    
    If frmCityTaxRpt.optByCity Then
        TempID = CtyID * 10 ^ 6 + PRDist.EmployeeID
    Else
        TempID = PRDist.EmployeeID * 10 ^ 6 + CtyID
    End If
    
    FFlag = False
    
    If trs.RecordCount > 0 Then
        trs.MoveFirst
        Do
            If trs!TempID = TempID And trs!Courtesy = Courtesy Then
                FFlag = True
                trs!Wage = trs!Wage + CtyWage
                trs!Tax = trs!Tax + CtyTax
                trs.Update
                Exit Do
            End If
            trs.MoveNext
        Loop Until trs.EOF
    End If
    
    If FFlag = False Then
        trs.AddNew
        trs!TempID = TempID
        trs!CityID = CtyID
        trs!EmployeeID = PRDist.EmployeeID
        trs!Wage = CtyWage
        trs!Tax = CtyTax
        trs!Courtesy = Courtesy
        trs.Update
    End If
    
End Sub
        
Private Sub CityTaxHeader(ReportTitle)
 
    PageHeader ReportTitle, Msg1, Msg2, ""
    Ln = Ln + 1
    
    If frmCityTaxRpt.optByEmployee Then
        Hdr1 = "CTY NUMBER"
        Hdr2 = "CITY NAME"
    Else
        Hdr1 = "EMP NUMBER"
        Hdr2 = "EMP NAME"
    End If
    
    PrintValue(1) = " ":                    FormatString(1) = "a7"
    PrintValue(2) = Hdr1:                   FormatString(2) = "a10"
    PrintValue(3) = " ":                    FormatString(3) = "a3"
    PrintValue(4) = Hdr2:                   FormatString(4) = "a35"
    PrintValue(5) = "CITY WAGE ":           FormatString(5) = "r14"
    PrintValue(6) = " ":                    FormatString(6) = "a5"
    PrintValue(7) = "CITY TAX ":            FormatString(7) = "r14"
    PrintValue(8) = " ":                    FormatString(8) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(92, "-"):        FormatString(1) = "a92"
    PrintValue(2) = " ":                    FormatString(2) = "~"
    
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub CityTaxSubHeader(ByVal ID As Double, _
                             ByVal PrintTotal As Boolean, _
                             ByVal Courtesy As Integer)

    Prvw.vsp.Font.Bold = True

    If ID <> 0 Then
        If frmCityTaxRpt.optByCity Then
            If ID = 999999 Then
                Hdr1 = ""
                Hdr2 = "NON TAX"
            ElseIf Not PRCity.GetByID(ID) Then
                Hdr1 = ID
                Hdr2 = "CITY NOT FOUND!"
            Else
                Hdr1 = PRCity.CityNumber
                If PRState.GetByID(PRCity.StateID) Then
                    Hdr2 = Trim(PRCity.CityName) & " " & PRState.StateAbbrev
                End If
                Hdr2 = Trim(Hdr2) & Format(PRCity.CityRate / 100, " #0.00%")
            End If
        Else
            If Not PREmployee.GetByID(ID) Then
                MsgBox "Employee ID Not Found: " & ID, vbExclamation
                GoBack
            End If
            Hdr1 = PREmployee.EmployeeNumber
            Hdr2 = PREmployee.FLName
        End If
    Else
        Hdr1 = ""
        Hdr2 = "COMPANY TOTALS"
    End If
    
    If PrintTotal Then
        PrintValue(1) = " TOTAL":               FormatString(1) = "a7"
    Else
        PrintValue(1) = " ":                    FormatString(1) = "a7"
    End If
    
    If Courtesy = 1 And PrintTotal = True Then
        Hdr2 = Trim(Hdr2) & " COURTESY"
    End If
    
    PrintValue(2) = Hdr1:                   FormatString(2) = "r10"
    PrintValue(3) = " ":                    FormatString(3) = "a3"
    PrintValue(4) = Hdr2:                   FormatString(4) = "a35"
    
    If PrintTotal Then
        PrintValue(5) = WageTotal:          FormatString(5) = "d0"
        PrintValue(6) = " ":                FormatString(6) = "a5"
        PrintValue(7) = TaxTotal:           FormatString(7) = "d0"
        PrintValue(8) = " ":                FormatString(8) = "~"
    Else
        PrintValue(5) = " ":                FormatString(5) = "~"
    End If
    FormatPrint
    Ln = Ln + 1
    
    Prvw.vsp.Font.Bold = False

End Sub

Public Sub EntryForm()                              ''''''''''''''''''''''''''''''''''''''''''

Dim PgFlag As Boolean
Dim LastLine, LineNumber, Colcount, LastType As Byte
Dim LastDpt As Long
Dim LastEmp As Integer
Dim LastPctAmt As Currency
Dim EntryString As String
Dim Amt As Currency
Dim FirstFlag As Boolean
Dim RateDifferenceType As Byte
Dim RateDifferenceAmount As Currency

Dim rsEE As New ADODB.Recordset

LastDpt = -1
LastLine = 0
LastEmp = 0
LastPctAmt = 0
Colcount = 0

    EntryString = "[___.__]"

    ' rs for employee - in case dept# sort order was chosen
    rsEE.CursorLocation = adUseClient
    rsEE.Fields.Append "EEID", adDouble
    rsEE.Fields.Append "EENumber", adDouble
    rsEE.Fields.Append "EEName", adVarChar, 80
    rsEE.Fields.Append "DptNumber", adDouble
    rsEE.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PREmployee WHERE InActive = 0"
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Active Employees Found!", vbExclamation, "PR Entry Form"
        GoBack
    End If
    
    frmProgress.Caption = "Data Entry Form Print"
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.Show
    
    Do
        
        frmProgress.lblMsg2 = "Gathering Data: " & PREmployee.EmployeeNumber & _
                              PREmployee.LFName
        frmProgress.Refresh
        
        rsEE.AddNew
        rsEE!EEID = PREmployee.EmployeeID
        rsEE!EENumber = PREmployee.EmployeeNumber
        rsEE!EEName = Mid(PREmployee.LFName, 1, 80)
        If PRDepartment.GetByID(PREmployee.DepartmentID) Then
            rsEE!DptNumber = PRDepartment.DepartmentNumber
        Else
            rsEE!DptNumber = 0
        End If
        rsEE.Update
        
        If Not PREmployee.GetNext Then Exit Do
    
    Loop

    ' ===========================================================
    ' === get OE & DED info for hdr and body
    
    trsEntry.CursorLocation = adUseClient
    
    trsEntry.Fields.Append "ItemType", adInteger
    trsEntry.Fields.Append "ItemID", adDouble
    trsEntry.Fields.Append "Title", adVarChar, 8, adFldIsNullable
    trsEntry.Fields.Append "Amount", adCurrency
    trsEntry.Fields.Append "Basis", adInteger
    
    trsEntry.Open , , adOpenDynamic, adLockOptimistic
        
    frmProgress.lblMsg2 = "Gathering Employer Item Data"
    frmProgress.Refresh

    SQLString = "SELECT * FROM PRItem WHERE " & _
                " PRItem.EmployeeID = 0 AND " & _
                " (PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeDED & ") " & _
                " ORDER BY PRItem.ItemType, PRItem.ItemID"
    
    If PRItem.GetBySQL(SQLString) = True Then
    
        Do
            
            trsEntry.AddNew
            trsEntry!ItemType = PRItem.ItemType
            trsEntry!ItemID = PRItem.ItemID
            trsEntry!Title = Mid(PRItem.Abbreviation, 1, 8)
            trsEntry!Amount = 0
            trsEntry.Update
            
            If Not PRItem.GetNext Then Exit Do
        
        Loop
    
    End If
    
    ' =========================================================

    PrtInit ("Port")
    ReportTitle = "PAYROLL DATA ENTRY FORM"
    Prvw.vsp.FontBold = True

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    ' loop thru employees - set the sort order from the form
    With frmEntry.cmbSortOrder
        If .ListIndex = PREquate.SortOrderNumber Then
            SQLString = "EENumber"
        ElseIf .ListIndex = PREquate.SortOrderName Then
            SQLString = "EEName"
        ElseIf .ListIndex = PREquate.SortOrderDeptNumber Then
            SQLString = "DptNumber, EENumber"
        ElseIf .ListIndex = PREquate.SortOrderDeptName Then
            SQLString = "DptNumber, EEName"
        Else
            MsgBox "Form Error???", vbExclamation, "PR Entry Form"
            GoBack
        End If
    End With
    
    rsEE.Sort = SQLString
    rsEE.MoveFirst
    
    SetFont 12, Equate.Portrait
    
'    PageHeader ReportTitle, Msg1, "", "", 1, frmEntry.chkUseGLName
'    '' PrintCompanyHeader (ReportList)
'    EntryHeader
'    If frmEntry.cmbSortOrder.ListIndex > 1 Then
'        If PRDepartment.GetByID(PREmployee.DepartmentID) = True Then
'            PrintValue(1) = "Department # " & PRDepartment.DepartmentNumber:    FormatString(1) = "a18"
'            PrintValue(2) = PRDepartment.Name:                                  FormatString(2) = "a30"
'            PrintValue(3) = " ":                                                FormatString(3) = "~"
'            FormatPrint
'            Ln = Ln + 2
'        End If
'    End If
    
    FirstFlag = True
    
    Do
        
        If Not PREmployee.GetByID(rsEE!EEID) Then
            MsgBox "Employee ID Err: " & rsEE!EEID, vbExclamation, "PR Entry Form"
            GoBack
        End If
        
        frmProgress.lblMsg2 = "Printing Entry Data: " & PREmployee.EmployeeNumber & _
                              PREmployee.LFName
        frmProgress.Refresh
        
        ' clear the amounts in the temp recordset
        If trsEntry.RecordCount > 0 Then
            trsEntry.MoveFirst
            Do Until trsEntry.EOF
                trsEntry!Amount = 0
                trsEntry.Update
                trsEntry.MoveNext
            Loop
        End If
                
        SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = " & PREmployee.EmployeeID & _
                    " AND (PRItem.ItemType = " & PREquate.ItemTypeOE & _
                    " OR PRItem.ItemType = " & PREquate.ItemTypeDED & ")"
        If PRItem.GetBySQL(SQLString) Then
            Do
                SQLString = "ItemID = " & PRItem.EmployerItemID
                trsEntry.Find SQLString, 0, adSearchForward, 1
                If trsEntry.EOF Then
                    ' 2017-12-27 - show EE#
                    MsgBox "EER item NF - EE #: " & PREmployee.EmployeeNumber & " Item " & PRItem.ItemID & " " & PRItem.Abbreviation, vbExclamation
                    GoBack
                End If
                
                ' rate difference?
                RateDifferenceType = 0
                RateDifferenceAmount = 0
                
                If PREmployee.Salaried = 0 And PRItem.ItemType = PREquate.ItemTypeOE Then
                    
                    If PRItem.UseEmployer = 1 Then
                        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 AND " & _
                                    "ItemID = " & PRItem.EmployerItemID
                        rsInit SQLString, cn, rs
                        If rs.RecordCount > 0 Then
                            RateDifferenceType = nNull(rs!RateDifference)
                            RateDifferenceAmount = nNull(rs!AmtPct)
                        End If
                    Else        ' employee defn
                        RateDifferenceType = PRItem.RateDifference
                        RateDifferenceAmount = PRItem.AmtPct
                    End If
                
                    If RateDifferenceType = PREquate.BasisAmount Then
                        trsEntry!Amount = PREmployee.HourlyAmount + RateDifferenceAmount
                    ElseIf RateDifferenceType = PREquate.BasisPercent Then
                        trsEntry!Amount = PREmployee.HourlyAmount
                        trsEntry!Amount = trsEntry!Amount + Round(PREmployee.HourlyAmount * RateDifferenceAmount / 100, 2)
                    Else
                        trsEntry!Amount = PRItem.AmtPct
                    End If
                Else
                    trsEntry!Amount = PRItem.AmtPct
                End If
                
                trsEntry!Basis = PRItem.Basis
                trsEntry.Update
                
                If Not PRItem.GetNext Then Exit Do
            Loop
        End If
        
        ' form feed???
        PgFlag = False
        If MaxLines - Ln <= LineCount Then PgFlag = True
        If frmEntry.cmbSortOrder.ListIndex > 1 And frmEntry.chkDeptSepPage = 1 Then
            If PREmployee.DepartmentID <> LastDpt And LastDpt <> -1 Then
                PgFlag = True
            End If
            If FirstFlag = True Then PgFlag = True
        End If
        If FirstFlag = True Then PgFlag = True
        LastDpt = PREmployee.DepartmentID
            
        If PgFlag = True Then
            If FirstFlag = False Then FormFeed
            SetFont 12, Equate.Portrait
            PageHeader ReportTitle, Msg1, "", "", 1
            EntryHeader
            If frmEntry.cmbSortOrder.ListIndex > 1 And frmEntry.chkDeptSepPage = 1 Then
                If PRDepartment.GetByID(PREmployee.DepartmentID) = True Then
                    PrintValue(1) = "Department # " & PRDepartment.DepartmentNumber:    FormatString(1) = "a18"
                    PrintValue(2) = PRDepartment.Name:                                  FormatString(2) = "a30"
                    PrintValue(3) = " ":                                                FormatString(3) = "~"
                    FormatPrint
                    Ln = Ln + 2
                End If
            Else
                Ln = Ln + 1
            End If
            '' PrintCompanyHeader (ReportList)
        End If
        FirstFlag = False
        
        ' print the employee line
        
        Prvw.vsp.Font.Bold = True
        
        X = Format(PREmployee.EmployeeNumber, "######0 ")
        PrintValue(1) = X:                  FormatString(1) = "a8"
        PrintValue(2) = PREmployee.LFName:  FormatString(2) = "a27"
             
        ' get the dept #
        If PRDepartment.GetByID(PREmployee.DepartmentID) Then
            X = Format(PRDepartment.DepartmentNumber, "##0 ")
        Else
            X = ""
        End If
        PrintValue(3) = X:                  FormatString(3) = "a4"
    
        If frmEntry.optSalHrly Or (frmEntry.optHrly And PREmployee.Salaried = 0) Then
            If PREmployee.Salaried Then
                Amt = PREmployee.SalaryAmount
            Else
                Amt = PREmployee.HourlyAmount
            End If
            PrintValue(4) = Amt:                FormatString(4) = "d11"
        Else
            If PREmployee.Salaried Then
                X = "SALARIED"
            Else
                X = "HOURLY"
            End If
            PrintValue(4) = X:                  FormatString(4) = "r11"
        End If
        
        PrintValue(5) = " ":                FormatString(5) = "a2"
        PrintValue(6) = EntryString:        FormatString(6) = "a9"     ' reg
        PrintValue(7) = EntryString:        FormatString(7) = "a9"     ' ovt
    
        ' print OE / DED titles
        LineCount = 1
        LastType = 0
        ItemCount = 0
        PrintNum = 7
        
        If trsEntry.RecordCount > 0 Then
            trsEntry.MoveFirst
            Do Until trsEntry.EOF
                        
                ItemCount = ItemCount + 1
                PrintNum = PrintNum + 1
                        
                ' new line
                ' at change in type or end of line (5 OE/DED)
                If (LastType <> 0 And LastType <> trsEntry!ItemType) Or ItemCount = 6 Then
                    PrintValue(PrintNum) = " ": FormatString(PrintNum) = "~"
                    FormatPrint
                    
                    Prvw.vsp.Font.Bold = False
                    
                    Ln = Ln + 1
                    If LineCount = 1 And frmEntry.chkHireDate = 1 Then
                        If PREmployee.DateHired <> 0 Then
                            X = "Hire Date: " & Format(PREmployee.DateHired, "mm/dd/yyyy")
                        Else
                            X = "Hire Date: "
                        End If
                    Else
                        X = ""
                    End If
                    PrintValue(1) = " ": FormatString(1) = "a84"
                    PrintValue(1) = X: FormatString(1) = "a70"
                    PrintNum = 2
                    ItemCount = 1
                    LineCount = LineCount + 1
                End If
        
                If trsEntry!Amount = 0 Then
                    PrintValue(PrintNum) = EntryString
                ElseIf frmEntry.chkDeds = 0 And trsEntry!ItemType = PREquate.ItemTypeDED Then
                    PrintValue(PrintNum) = EntryString
                ElseIf frmEntry.chkOtherEarns = 0 And trsEntry!ItemType = PREquate.ItemTypeOE Then
                    PrintValue(PrintNum) = EntryString
                Else
                    If trsEntry!Basis = PREquate.BasisPercent Then
                        PrintValue(PrintNum) = "[" & Format(trsEntry!Amount, "##0.00") & "%]"
                    Else
                        ' always show the deduction amount
                        If trsEntry!ItemType = PREquate.ItemTypeDED Then
                            PrintValue(PrintNum) = "[" & Format(trsEntry!Amount, "###0.00") & "]"
                        Else        ' OE - it depends .....
                            PrintValue(PrintNum) = "[" & Format(trsEntry!Amount, "###0.00") & "]"
                            If frmEntry.optNone = True Then     ' show no salary selected
                                PrintValue(PrintNum) = EntryString
                            ElseIf frmEntry.optHrly = True And PREmployee.Salaried = 1 Then ' don't show for salaried EE's
                                PrintValue(PrintNum) = EntryString
                            End If
                        End If
                    End If
                End If
                FormatString(PrintNum) = "a9"
        
                LastType = trsEntry!ItemType
                trsEntry.MoveNext
            
            Loop
        
        End If
        
        ' print the last one - always necessary (?)
        PrintNum = PrintNum + 1
        PrintValue(PrintNum) = " ": FormatString(PrintNum) = "~"
        FormatPrint
        Ln = Ln + 3
            
        rsEE.MoveNext
    
    Loop Until rsEE.EOF
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub EntryHeader()

    X = "a" & (40 - Len(Trim(frmEntry.tdbtxtHdrComment)) / 2)
    PrintValue(1) = " ": FormatString(1) = X
    PrintValue(2) = frmEntry.tdbtxtHdrComment: FormatString(2) = "a50"
    PrintValue(3) = "": FormatString(3) = "~"
    FormatPrint
    
    SetFont 8, Equate.Portrait
    
    Ln = Ln + 2

    PrintValue(1) = "EMP#":                 FormatString(1) = "a8"
    PrintValue(2) = "N A M E":              FormatString(2) = "a27"
    PrintValue(3) = "DPT ":                 FormatString(3) = "a4"
    PrintValue(4) = "RATE/SAL ":            FormatString(4) = "r11"
    PrintValue(5) = " ":                    FormatString(5) = "a2"
    PrintValue(6) = "REG HRS  ":            FormatString(6) = "r9"
    PrintValue(7) = "OVT HRS  ":            FormatString(7) = "r9"
    
    ' print OE / DED titles
    LineCount = 1
    LastType = 0
    ItemCount = 0
    PrintNum = 7
    
    If trsEntry.RecordCount > 0 Then
        trsEntry.MoveFirst
        Do Until trsEntry.EOF
                    
            ItemCount = ItemCount + 1
            PrintNum = PrintNum + 1
                    
            ' new line
            ' at change in type or end of line (5 OE/DED)
            If (LastType <> 0 And LastType <> trsEntry!ItemType) Or ItemCount = 6 Then
                PrintValue(PrintNum) = " ": FormatString(PrintNum) = "~"
                FormatPrint
                Ln = Ln + 1
                PrintValue(1) = " ": FormatString(1) = "a70"
                PrintNum = 2
                ItemCount = 1
                LineCount = LineCount + 1
            End If
    
            X = trsEntry!Title & "  "
            PrintValue(PrintNum) = X: FormatString(PrintNum) = "r9"
    
            LastType = trsEntry!ItemType
            trsEntry.MoveNext
        
        Loop
    
    End If
    
    ' print the last one - always necessary (?)
    PrintNum = PrintNum + 1
    PrintValue(PrintNum) = " ": FormatString(PrintNum) = "~"
    FormatPrint
    Ln = Ln + 1
    
    ' blank line beween employees
    LineCount = LineCount + 1
    
End Sub

Public Sub CheckRecon(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal StartDate As Long, _
                         ByVal EndDate As Long, _
                         ByVal OptDate As String)
                         
Dim SQLString1 As String
Dim EmpName As String
Dim TotAmt As Currency


    trs.CursorLocation = adUseClient
   
    trs.Fields.Append "Date", adDate
    trs.Fields.Append "Number", adInteger
    trs.Fields.Append "Amount", adCurrency
   
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    frmCheckRecon.Hide
    ReportTitle = "PAYROLL CHECK RECONCILIATION REPORT"
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    SetEquates

    SQLString = "SELECT * FROM PRHist"
 
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = Trim(SQLString) & " WHERE PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = Trim(SQLString) & " WHERE PRHist.CheckDate >= " & CLng(StartDate) & " AND " & _
                                    " PRHist.CheckDate <= " & CLng(EndDate)
            Msg1 = "CHECK DATE RANGE: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        ElseIf OptDate = "P/E DATE" Then
             SQLString = Trim(SQLString) & " WHERE PRHist.PEDate >= " & CLng(StartDate) & " AND " & _
                                    " PRHist.PEDate <= " & CLng(EndDate)
            Msg1 = "P/E DATE RANGE: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
                                    
    End If

    SQLString = Trim(SQLString) & " ORDER BY PRHist.CheckNumber"

    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History Found !!!", vbExclamation, "Payroll Check Reconciliation"
        GoBack
    End If

    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = Ln + 1
            CheckReconHdr
        End If

        SQLString1 = "Date = " & Format(PRHist.CheckDate, "mm/dd/yyyy")
        trs.Find SQLString1, 0, adSearchForward, 1
        
        If trs.EOF Then
            trs.AddNew Array("Date", "Number", "Amount"), _
            Array(PRHist.CheckDate, 0, 0)
            trs.UpdateBatch
        End If
        
        If Not PREmployee.GetBySQL("SELECT * FROM PREmployee WHERE PREmployee.EmployeeID = " & PRHist.EmployeeID) Then
            EmpName = "None"
        Else
            EmpName = PREmployee.FLName
        End If
        
        PrintValue(1) = PRHist.CheckNumber:         FormatString(1) = "n9"
        PrintValue(2) = Format(PRHist.CheckDate, "  mm/dd/yyyy  "):  FormatString(2) = "a14"
        PrintValue(3) = EmpName:                    FormatString(3) = "a34"
        PrintValue(4) = PRHist.Net:                 FormatString(4) = "d15"
        CheckAmt = CheckAmt + PRHist.Net
 
        PrintValue(5) = PRHist.DirectDeposit:       FormatString(5) = "d15"
        
        DepoAmt = DepoAmt + PRHist.DirectDeposit
        TotAmt = TotAmt + PRHist.Net + PRHist.DirectDeposit
        
        PrintValue(6) = " ":                        FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 1
        
        trs!Amount = trs!Amount + PRHist.Net + PRHist.DirectDeposit
        trs!Number = trs!Number + 1
        
        trs.Update
        NoRecords = NoRecords + 1
        If Not PRHist.GetNext Then Exit Do
    Loop
    Ln = Ln + 1
    PrintValue(1) = " FINAL TOTAL:":                FormatString(1) = "a56"
    PrintValue(2) = CheckAmt:                       FormatString(2) = "d85"
    PrintValue(3) = DepoAmt:                        FormatString(3) = "d19"
    PrintValue(4) = " ":                            FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2

    If Ln >= MaxLines - 15 Then
        FormFeed
        PageHeader ReportTitle, Msg1, "", ""
        Ln = Ln + 1
        CheckReconHdr
    Else
        Ln = Ln + 2
    End If
    
    PrintValue(1) = "-------------  SUMMARY  ---------------------"
    FormatString(1) = "a40"
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "Date ":                        FormatString(1) = "a13"
    PrintValue(2) = "# of Checks":                  FormatString(2) = "a18"
    PrintValue(3) = "Amount ":                      FormatString(3) = "a16"
    PrintValue(4) = " ":                            FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(85, "-"):                FormatString(1) = "a40"
    FormatString(1) = "a40"
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
        
    trs.Sort = "date"
    trs.MoveFirst

    Do
        If trs.EOF = True Then
            Exit Do
        End If
        
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = Ln + 1
            CheckReconHdr
        End If
        
        PrintValue(1) = Format(trs!Date, "mm/dd/yy"):   FormatString(1) = "a8"
        PrintValue(2) = " ":                        FormatString(2) = "a5"

        PrintValue(3) = trs!Number:                 FormatString(3) = "n8"
        PrintValue(4) = " ":                        FormatString(4) = "a5"
        
        PrintValue(5) = trs!Amount:                 FormatString(5) = "d12"
        PrintValue(6) = " ":                        FormatString(6) = "~"
        
        FormatPrint
        Ln = Ln + 1
        trs.MoveNext

    Loop

    Ln = Ln + 1
    PrintValue(1) = "TOTAL: ":                      FormatString(1) = "a13"
    PrintValue(2) = NoRecords:                      FormatString(2) = "n8"
        
    PrintValue(3) = " ":                            FormatString(3) = "a5"
    PrintValue(4) = TotAmt:                         FormatString(4) = "d12"
       
    PrintValue(5) = " ":                            FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
            
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub


Public Sub CheckReconHdr()
    PrintValue(1) = " ":                FormatString(1) = "a77"
    PrintValue(2) = "Direct":           FormatString(2) = "a6"
    PrintValue(3) = " ":                FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                FormatString(1) = "a63"
    PrintValue(2) = "Check":            FormatString(2) = "a14"
    PrintValue(3) = "Deposit":          FormatString(3) = "a15"
    PrintValue(4) = " ":                FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
                    
    PrintValue(1) = "Check No.":        FormatString(1) = "a11"
    
'    If OptDate = "CHECK DATE" Then
'        PrintValue(2) = "CHECK DATE":   FormatString(2) = "a12"
'    Else
'        PrintValue(2) = "P/E DATE":     FormatString(2) = "a12"
'    End If
    
    ' always use the check date
    PrintValue(2) = "CHECK DATE":   FormatString(2) = "a12"
    
    PrintValue(3) = "Payee":            FormatString(3) = "a41"
    PrintValue(4) = "Amount":           FormatString(4) = "a14"
    PrintValue(5) = "Amount":           FormatString(5) = "a15"
    PrintValue(6) = " ":                FormatString(6) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(85, "-"):   FormatString(1) = "a118"
    PrintValue(2) = " ":                FormatString(2) = "~"
    
    FormatPrint
    Ln = Ln + 1
End Sub

Public Sub DirectDepositRpt()
                            
Dim ReportTitle, LastABA As String
Dim LastEmployee As Long
Dim ddFlag As Boolean
Dim LastType, ItmID As Long
Dim FedID As String

    ' gather and verify employer info
    If frmDirectDep.chkOutputFile = 1 Then
        If Len(PRCompany.FederalID) <> 10 Then
            MsgBox "Invalid Employer Federal ID: " & PRCompany.FederalID, vbExclamation
            GoBack
        End If
        If Mid(PRCompany.FederalID, 3, 1) <> "-" Then
            MsgBox "Invalid Employer Federal ID: " & PRCompany.FederalID, vbExclamation
            GoBack
        End If
    
    End If
    
    FedID = Mid(PRCompany.FederalID, 1, 2) & Mid(PRCompany.FederalID, 4, 7)
    
    ' employer ABA verify
    ChkDigNo = ABACheckDigit(PRCompany.BankABA, frmDirectDep.chkOutputFile)
    If frmDirectDep.chkOutputFile = 1 Then
        If ChkDigNo > 9 Then         ' invalid ?
            MsgBox "Invalid Employer ABA Routing Number: " & PRCompany.BankABA, vbExclamation
            GoBack
        End If
    End If
        
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "Batch", adInteger
    trs.Fields.Append "EmployeeNo", adInteger
    trs.Fields.Append "Routing", adVarChar, 14, adFldIsNullable
    trs.Fields.Append "AcctNo", adVarChar, 17, adFldIsNullable
    trs.Fields.Append "Name", adVarChar, 50, adFldIsNullable
    trs.Fields.Append "AcctType", adVarChar, 3, adFldIsNullable
    trs.Fields.Append "Credit", adCurrency
    trs.Fields.Append "BankName", adVarChar, 20, adFldIsNullable
    trs.Fields.Append "ItemType", adInteger
    
    trs.Open , , adOpenDynamic, adLockOptimistic

    PrtInit ("Port")

    ' **** Header Info ****
    RptTitle = "CENTRALIZED DIRECT DEPOSIT REPORT"
        
    If frmDirectDep.chkOutputFile Then
        Msg2 = "Output File: " & Trim(frmDirectDep.tdbtxtFileName)
        msg3 = "Effective Date: " & Format(frmDirectDep.tdbEffDate, "mm/dd/yyyy")
    End If
    
    ' **** Header Info ****

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 10, Equate.Portrait
    Ln = 0
        
    ' loop thru the batch(es) selected
    frmDirectDep.rsDDBatch.MoveFirst
    Do
        
        If frmDirectDep.rsDDBatch!Select = True Then
                        
            ddPEDate = frmDirectDep.rsDDBatch!PEDate
            ddCheckDate = frmDirectDep.rsDDBatch!CheckDate
            
            SQLString = "SELECT * FROM PRItemHist" & _
                        " WHERE PRItemHist.BatchID = " & frmDirectDep.rsDDBatch!BatchNumber
            If PRItemHist.GetBySQL(SQLString) Then
    
                Do
                    
                    ' get the PRItem record
                    If PRItem.GetByID(PRItemHist.ItemID) = False Then
                        MsgBox "PRItem NF: " & PRItemHist.ItemID, vbExclamation
                        GoBack
                    End If
                    
                    ' use it if dir dep or flagged for dir dep rpt
                    ddFlag = False
                    If PRItemHist.ItemType = PREquate.ItemTypeDirDepDed Then
                        ddFlag = True
                    Else
                        ' use the employer defn?
                        If PRItem.UseEmployer = 1 Then
                            If PRItem.GetByID(PRItem.EmployerItemID) = False Then
                                MsgBox "Employer Item NF: " & PRItem.EmployerItemID, vbExclamation
                                GoBack
                            End If
                            If PRItem.DirDepRpt = 1 Then ddFlag = True
                            ' get the item back
                            If PRItem.GetByID(PRItemHist.ItemID) = False Then
                            End If
                        Else
                            If PRItem.DirDepRpt = 1 Then ddFlag = True
                        End If
                    End If
                                        
                    If ddFlag = True Then
                        trs.AddNew
                        If Not PREmployee.GetByID(PRItemHist.EmployeeID) Then    '  Get Employee Info
                            MsgBox "Employee NF: " & PRItemHist.EmployeeID, vbExclamation
                            GoBack
                        End If
                        trs!EmployeeNo = PREmployee.EmployeeNumber
                        trs!Name = Mid(PREmployee.LFName, 1, 50)
                            
                        If PRItem.ItemType = PREquate.ItemTypeDirDepDed Then
                            trs!Batch = PRItemHist.BatchID
                            trs!Routing = Mid(PRItem.DirDepABA, 1, 14)
                            trs!AcctNo = Mid(PRItem.DirDepAccount, 1, 17)
                            If PRItem.DirDepType = PREquate.DirDepTypeChecking Then
                                trs!AcctType = "CHK"
                            Else
                                trs!AcctType = "SVE"
                            End If
                            trs!Credit = PRItemHist.Amount
                            trs!BankName = Mid(PRItem.DirDepBank, 1, 20)
                        Else
                            trs!Batch = PRItemHist.BatchID
                            trs!Routing = ""
                            trs!AcctNo = PRItem.Comment
                            trs!AcctType = "DED"
                            trs!Credit = PRItemHist.Amount
                            If PRItem.UseEmployer = 1 Then
                                ItmID = PRItem.EmployerItemID
                                If PRItem.GetByID(ItmID) = False Then
                                    MsgBox "Employer Item NF: " & PRItem.EmployerItemID, vbExclamation
                                    GoBack
                                End If
                            End If
                            trs!BankName = PRItem.Title
                        End If
                            
                        trs!ItemType = PRItemHist.ItemType
                        trs.Update
                    
                    End If
                    If Not PRItemHist.GetNext Then Exit Do
                Loop
            
            End If

        End If

        frmDirectDep.rsDDBatch.MoveNext
        If frmDirectDep.rsDDBatch.EOF Then Exit Do
        
    Loop
    
    ' =============================================================================
    ' --- Init NACHA file and write File Header (1) and Batch Header (5)
    If frmDirectDep.chkOutputFile Then
        
        TChannel = FreeFile
        On Error Resume Next
        Open frmDirectDep.tdbtxtFileName For Output As #TChannel Len = 94
        If Err.Number <> 0 Then
            X = "Error Opening: " & frmDirectDep.tdbtxtFileName & vbCr & vbCr & _
                        " " & Err.Number & " " & Err.Description
            MsgBox X, vbExclamation
            GoBack
        End If
        On Error GoTo 0
        
        ' special file header?
        Dim SEH_Flag As Boolean
        If PRCompany.CompanyID = 37 And InStr(LCase(PRCompany.Name), "south east") Then
            SEH_Flag = True
        Else
            SEH_Flag = False
        End If
        If SEH_Flag = True Then
            SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDirDepFolder & _
                        " AND UserID = " & PRCompany.CompanyID
            If PRGlobal.GetBySQL(SQLString) = True Then
                If PRGlobal.Var2 <> "" Then
                    Print #TChannel, PRGlobal.Var2
                End If
            End If
        End If
        
        ' federal ID string
        If PRCompany.DirDepID1 = 1 Then
            FedID = "1"
        Else
            FedID = " "
        End If
        If PRCompany.DirDepUseAltID = 1 Then
            FedID = FedID & Format(PRCompany.DirDepAltID, "000000000")
        Else
            FedID = FedID & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
        End If
        
        ' 2017-02-21 - for Freedom Harley
        ' 2017-12-23 - Custom Auto Body too
        '   use alt FedID for recs 5 & 8 --- NOT rec 1
        Dim FedID1, FedID5, FedID8 As String
        FedID1 = FedID
        FedID5 = FedID
        FedID8 = FedID
        If PRCompany.CompanyID = 35 And InStr(LCase(PRCompany.Name), "freedom") Then
            ' huntington - ACH Company ID
            FedID1 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID5 = "9" & Format(PRCompany.DirDepAltID, "000000000")
            FedID8 = "9" & Format(PRCompany.DirDepAltID, "000000000")
        End If
        If PRCompany.CompanyID = 23 And InStr(LCase(PRCompany.Name), "custom auto") Then
            ' huntington - ACH Company ID
            FedID1 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID5 = "9" & Format(PRCompany.DirDepAltID, "000000000")
            FedID8 = "9" & Format(PRCompany.DirDepAltID, "000000000")
        End If
        If PRCompany.CompanyID = 9 And InStr(LCase(PRCompany.Name), "stark county") Then
            ' huntington - ACH Company ID
            FedID1 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID5 = "9" & Format(PRCompany.DirDepAltID, "000000000")
            FedID8 = "9" & Format(PRCompany.DirDepAltID, "000000000")
        End If
        If InStr(LCase(PRCompany.Name), "scott molders") Then
            ' huntington - ACH Company ID
            FedID1 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID5 = "9" & Format(PRCompany.DirDepAltID, "000000000")
            FedID8 = "9" & Format(PRCompany.DirDepAltID, "000000000")
        End If
        ' 2022-09-09
        If InStr(LCase(PRCompany.Name), "artsparks") Then
            ' huntington - ACH Company ID
            FedID1 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID5 = "9" & Format(PRCompany.DirDepAltID, "000000000")
            FedID8 = "9" & Format(PRCompany.DirDepAltID, "000000000")
        End If
        If PRCompany.CompanyID = 24 And InStr(LCase(PRCompany.Name), "conti") Then
            ' westfield
            FedID1 = " " & Format(PRCompany.DirDepAltID, "000000000")
            FedID5 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID8 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
        End If
        If PRCompany.CompanyID = 25 And InStr(LCase(PRCompany.Name), "central") Then
            ' westfield
            FedID1 = " " & Format(PRCompany.DirDepAltID, "000000000")
            FedID5 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
            FedID8 = " " & Left(PRCompany.FederalID, 2) & Mid(PRCompany.FederalID, 4, 7)
        End If
        
        ' 2015-08-20 - special for Conti - store for the balance entry "6"
        Dim cFedID As String
        cFedID = FedID
        
        ' 2015-08-20 - special switch for Conti
        Dim CompanyABA As String
        CompanyABA = Trim(PRCompany.BankABA)
        If InStr(1, LCase(PRCompany.Name), "conti") <> 0 Then
            If PRCompany.DirDepUseAltID = 1 Then
                CompanyABA = Trim(cFedID)
            End If
        End If
            
        ' 2015-08-17 - special for Conti Westfield Bank
        Dim ddCompName As String
        ddCompName = PRCompany.Name
        If InStr(1, LCase(PRCompany.Name), "conti") <> 0 Then
            If PRCompany.DirDepUseAltID = 1 Then
                ddCompName = "WESTFIELD BANK"
            End If
        End If
        
        X = "101" & _
            " " & PRCompany.BankABA & _
            Mid(FedID1, 1, 10) & _
            Format(Now(), "YYMMDDHHMM") & "A094101" & _
            PadString(PRCompany.BankName, 23) & _
            PadString(ddCompName, 23) & _
            Format(frmDirectDep.tdbEffDate, "yyyymmdd")
        
        Print #TChannel, X  ' Output text.
        
        ' Conti 01/07/2010 - no "1" before the FedID
        ' 2012-09-01 - optional batch header from PRGlobal
        If frmDirectDep.BatchHeader = "" Then
            BtchHeader = PadString(PRCompany.Name, 16)
        Else
            BtchHeader = PadString(frmDirectDep.BatchHeader, 16)
        End If
        
        X = "5200" & _
            BtchHeader & _
            "PAY ENDING  " & Format(ddPEDate, "yyyymmdd") & _
            Mid(FedID5, 1, 10) & _
            PadString("PPDPAYROLL", 13) & _
            Format(Now(), "yymmdd") & _
            Format(frmDirectDep.tdbEffDate, "yymmdd") & Space(3) & "1" & _
            Left(CompanyABA, 8) & _
            OutNumber(1, 7)
        
        Print #TChannel, X  ' Output text.
        
        WriteCt = 2
    End If
    
    ' --- Init NACHA file and write File Header (1) and Batch Header (5)
    ' =============================================================================
    
    ' Sort temporary recordset according to user sort selection
    If frmDirectDep.optEmpNo = True Then
        trs.Sort = "ItemType DESC, Batch, EmployeeNo"
    Else
        trs.Sort = "ItemType DESC, Batch, Name"
    End If
    
    LineCt = 0
    LastBatch = 0
    LastType = 0
    trs.MoveFirst
    
    Do
        
        If trs.EOF = True Then
            Exit Do
        End If
        
        If Ln = 0 Or Ln > MaxLines Or trs!Batch <> LastBatch Or trs!ItemType <> LastType Then
                        
            If Ln = 0 Or Ln > MaxLines Then
                DirectDepositRptHeader
            End If
                        
            ' break in item type
            If trs!ItemType <> LastType And LastType <> 0 Then
                DirectDepositSubTotals
                Ln = Ln + 2
            End If
                        
            ' break in batch number
            If (trs!Batch <> LastBatch And LastBatch <> 0) Then
                DirectDepositSubTotals
                Ln = Ln + 2
                DirectDepositBatchHeader trs!Batch
            End If
        
        End If
        
'======================================================================================
'                                   PRINT REPORT DETAIL
'======================================================================================

        frmProgress.lblMsg2 = "Employee: " & trs!EmployeeNo & " - " & Trim(trs!Name)
        frmProgress.Show
        
        If IsNull(trs!AcctNo) Then trs!AcctNo = 0
        If IsNull(trs!Routing) Then trs!Routing = 0
        If IsNull(trs!AcctType) Then trs!AcctType = 0
        If IsNull(trs!BankName) Then trs!BankName = 0
        
        TotCredAmt = TotCredAmt + trs!Credit
        SubCredAmt = SubCredAmt + trs!Credit
        
        PrintValue(1) = trs!EmployeeNo:         FormatString(1) = "n6"
        PrintValue(2) = " ":                    FormatString(2) = "a3"
        PrintValue(3) = trs!Routing:            FormatString(3) = "a11"
        PrintValue(4) = " ":                    FormatString(4) = "a2"
        PrintValue(5) = trs!AcctNo:             FormatString(5) = "a17"
        PrintValue(6) = " ":                    FormatString(6) = "a1"
        PrintValue(7) = trs!Name:               FormatString(7) = "a30"
        PrintValue(8) = " ":                    FormatString(8) = "a2"
        PrintValue(9) = trs!AcctType:           FormatString(9) = "a3"
        PrintValue(10) = " ":                   FormatString(10) = "a8"
        PrintValue(11) = trs!Credit:            FormatString(11) = "d8"
        PrintValue(12) = " ":                   FormatString(12) = "a1"
        PrintValue(13) = trs!BankName:          FormatString(13) = "a20"
        PrintValue(14) = " ":                   FormatString(14) = "~"
        FormatPrint
        Ln = Ln + 1
        LineCt = LineCt + 1
        SubLineCt = SubLineCt + 1
        
        LastBatch = trs!Batch
        LastType = trs!ItemType
        
        ' output detail to NACHA file
        If frmDirectDep.chkOutputFile Then
            If trs!AcctType = "CHK" Then
                ChkSve = 22
            Else
                ChkSve = 32
            End If
            
            SeqNo = SeqNo + 1
             
            ChkDigNo = ABACheckDigit(trs!Routing, frmDirectDep.chkOutputFile)
            If ChkDigNo > 9 Then End        ' invalid ?
            RteNo = CLng(Left(trs!Routing, 8))
            
            X = "6" & ChkSve & _
                Format(trs!Routing, "000000000") & _
                PadString(trs!AcctNo, 17) & _
                OutNumber(trs!Credit * 100, 10) & _
                PadString(trs!EmployeeNo, 15) & _
                PadString(trs!Name, 22) & "  0" & _
                Left(CompanyABA, 8) & Format(SeqNo, "0000000")
            Print #TChannel, X  ' Output text.
            WriteCt = WriteCt + 1
        
            '  Accumulate Hash Total
            Hash = Hash + RteNo
            DepositTotal = DepositTotal + trs!Credit
        
        End If
        trs.MoveNext
    
    Loop
    
    ' finish NACHA file
    If frmDirectDep.chkOutputFile Then
        
        ' Balanced file - write employer deposit
        If frmDirectDep.chkBalFile Then
            
            DebitTotal = DepositTotal
            SeqNo = SeqNo + 1
            
            ChkDigNo = ABACheckDigit(CompanyABA, frmDirectDep.chkOutputFile)
            If ChkDigNo > 9 Then End        ' invalid ?
            RteNo = CLng(Left(CompanyABA, 8))
            
            X = "627" & _
                Format(CompanyABA, "000000000") & _
                PadString(PRCompany.BankAccount, 17) & _
                OutNumber(DepositTotal * 100, 10) & _
                String(15, " ") & _
                PadString(PRCompany.Name, 22) & "  0" & _
                Left(CompanyABA, 8) & Format(SeqNo, "0000000")
            
            Print #TChannel, X  ' Output text.
            WriteCt = WriteCt + 1
        
            '  Accumulate Hash Total
            Hash = Hash + RteNo
        Else
            DebitTotal = 0
        End If
    
        ' batch control (8)
        HashString = Right(Format(Hash, String(20, "0")), 10)
            
        X = "8200" & _
            Format(SeqNo, "000000") & _
            Format(HashString, "0000000000") & _
            OutNumber(DebitTotal * 100, 12) & _
            OutNumber(DepositTotal * 100, 12) & _
            Mid(FedID8, 1, 10) & _
            Space(25) & _
            Left(CompanyABA, 8) & _
            OutNumber(1, 7)
        Print #TChannel, X  ' Output text.
        WriteCt = WriteCt + 1
        
        ' number of blocks in the file - even blocks of 10 lines
        WriteCt = WriteCt + 1   ' include the file control about to be written
        BlockCt = Int(WriteCt / 10)
        If WriteCt Mod 10 <> 0 Then BlockCt = BlockCt + 1
        
        ' file control (9)
        X = "9000001" & _
            Format(BlockCt, "000000") & _
            Format(SeqNo, "00000000") & _
            Format(HashString, "0000000000") & _
            OutNumber(DebitTotal * 100, 12) & _
            OutNumber(DepositTotal * 100, 12) & _
            Space(39)
        Print #TChannel, X  ' Output text.
    
        ' fill out the file to make even block of 10
        ' 2016-11-16 - don't fill if even multiple of 10
        If WriteCt Mod 10 <> 0 Then
            X = String(94, "9")
            For I = 1 To 10 - WriteCt Mod 10
                Print #TChannel, X  ' Output text.
            Next I
        End If
    
    End If
        
    DirectDepositSubTotals
    
    Ln = Ln + 2
    PrintValue(1) = " ":                            FormatString(1) = "a53"
    PrintValue(2) = "FINAL TOTAL CREDIT AMOUNT":    FormatString(2) = "a29"
    PrintValue(3) = TotCredAmt:                     FormatString(3) = "d9"
    PrintValue(4) = " ":                            FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " ":                            FormatString(1) = "a59"
    PrintValue(2) = "NUMBER OF DEPOSITS":           FormatString(2) = "a23"
    PrintValue(3) = LineCt:                         FormatString(3) = "n8"
    PrintValue(4) = " ":                            FormatString(4) = "~"
    FormatPrint

    trs.Close
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Private Sub DirectDepositSubTotals()
    
    If MaxLines - Ln <= 5 Then
        If Ln Then FormFeed
        DirectDepositRptHeader
    End If
    
    DirectDepositBatchHeader LastBatch      ' batch subtl
    
    PrintValue(1) = " ":                    FormatString(1) = "a51"
    PrintValue(2) = "BATCH " & LastBatch & " TOTAL CREDIT AMOUNT": FormatString(2) = "a31"
    PrintValue(3) = SubCredAmt:             FormatString(3) = "d9"
    PrintValue(4) = " ":                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " ":                    FormatString(1) = "a59"
    PrintValue(2) = "NUMBER OF EMPLOYEES":  FormatString(2) = "a23"
    PrintValue(3) = SubLineCt:              FormatString(3) = "n8"
    PrintValue(4) = " ":                    FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    SubCredAmt = 0
    SubLineCt = 0

End Sub

Private Sub DirectDepositBatchHeader(ByVal BatchNum As Long)
    
    If MaxLines - Ln <= 5 Then
        If Ln Then FormFeed
        DirectDepositRptHeader
    End If
        
    If Not PRBatch.GetByID(BatchNum) Then
        MsgBox "PRBatch Not Found: " & BatchNum, vbExclamation
        GoBack
    End If
    
    PrintValue(1) = "*** BATCH #: " & PRBatch.BatchID:                      FormatString(1) = "a25"
    PrintValue(2) = "PE DATE: " & Format(ddPEDate, "mm/dd/yy"):       FormatString(2) = "a19"
    PrintValue(3) = "CK DATE: " & Format(ddCheckDate, "mm/dd/yy"):    FormatString(3) = "a19"
    PrintValue(4) = "":                                                     FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub DirectDepositRptHeader()
    
    If Ln > 0 Then FormFeed
    
    SetFont 10, Equate.Portrait
    PageHeader RptTitle, Msg1, Msg2, msg3
    SetFont 8, Equate.Portrait
    Ln = Ln + 2

    PrintValue(1) = "EMPLOYEE":         FormatString(1) = "a8"
    PrintValue(2) = " ":                FormatString(2) = "a3"
    PrintValue(3) = "ROUTING":          FormatString(3) = "a7"
    PrintValue(4) = "":                 FormatString(4) = "a5"
    PrintValue(5) = "ACCT":             FormatString(5) = "a17"
    PrintValue(6) = " ":                FormatString(6) = "a31"
    PrintValue(7) = "ACCT":             FormatString(7) = "a4"
    PrintValue(8) = "":                 FormatString(8) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " NUMBER":          FormatString(1) = "a7"
    PrintValue(2) = " ":                FormatString(2) = "a2"
    PrintValue(3) = "AND TRANSIT":      FormatString(3) = "a11"
    PrintValue(4) = "":                 FormatString(4) = "a2"
    PrintValue(5) = "NUMBER":           FormatString(5) = "a17"
    PrintValue(6) = " ":                FormatString(6) = "a1"
    PrintValue(7) = "NAME":             FormatString(7) = "a25"
    PrintValue(8) = " ":                FormatString(8) = "a6"
    PrintValue(9) = "TYPE":             FormatString(9) = "a8"
    PrintValue(10) = "DEPOSIT AMT":     FormatString(10) = "a13"
    PrintValue(11) = "BANK NAME":       FormatString(11) = "a9"
    PrintValue(12) = " ":               FormatString(12) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = String(110, "="):   FormatString(1) = "a110"
    PrintValue(2) = " ":                FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Public Sub YECityHeader(ReportTitle)
    SetFont 8, Equate.Portrait
    Columns = 115
    PageHeader ReportTitle, Msg1, "", ""
    Ln = Ln + 1
    
    PrintValue(1) = "NUMBER":                                   FormatString(1) = "a7"
    PrintValue(2) = "EMPLOYEE NAME":                            FormatString(2) = "a35"
    PrintValue(3) = "SOC SEC #":                                FormatString(3) = "a16"
    PrintValue(4) = "ADDRESS":                                  FormatString(4) = "a32"
    PrintValue(5) = "YTD GROSS":                                FormatString(5) = "a17"
    PrintValue(6) = "YTD TAX":                                  FormatString(6) = "a10"
    PrintValue(7) = " ":                                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(115, "-"):                           FormatString(1) = "a115"
    PrintValue(2) = " ":                                        FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
       
End Sub

Public Sub YECityTax()

Dim ReportTitle As String

    frmCityTaxRpt.Hide
    SetEquates
    StartYM = qYear * 100 + 1
    EndYM = qYear * 100 + 12
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    trs.CursorLocation = adUseClient
    trs.Fields.Append "TempID", adDouble
    trs.Fields.Append "CityID", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    trs.Fields.Append "YTDGross", adCurrency
    trs.Fields.Append "YTDTax", adCurrency
    
    trs.Open , , adOpenDynamic, adLockOptimistic
           
    PrtInit ("Port")    ' "Port" = Portrait

    SetFont 8, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    Msg1 = "FOR YEAR: " & qYear
    ReportTitle = "PAYROLL YEARLY CITY TAX REPORT"
    SQLString = "Select * FROM PRDist WHERE PRDist.YearMonth >= " & StartYM & _
                " AND PRDist.YEARMONTH <= " & EndYM & _
                " AND (CityWage <> 0 OR CityTax <> 0)" & _
                " ORDER BY CheckDate"

    If Not PRDist.GetBySQL(SQLString) Then
        MsgBox "Data Not Found!!!", vbExclamation, "Payroll Yearly City Tax Report"
        GoBack
    End If
    Do
        
        If PRDist.CityID <> 0 Then
            TempID = PRDist.CityID * 10 ^ 6 + PRDist.EmployeeID
            SQLString = "TempID = " & TempID
            trs.Find SQLString, 0, adSearchForward, 1
            If trs.EOF Then
                trs.AddNew
                trs!TempID = TempID
                trs!EmployeeID = PRDist.EmployeeID
                trs!CityID = PRDist.CityID
                trs!YTDGross = 0
                trs!YTDTax = 0
                trs.Update
            End If
            trs!YTDGross = trs!YTDGross + PRDist.CityWage
            trs!YTDTax = trs!YTDTax + PRDist.CityTax
            trs.Update
        End If
        
        If PRDist.CourtesyCityID <> 0 And PRDist.CourtesyCityTax <> 0 Then
            TempID = PRDist.CourtesyCityID * 10 ^ 6 + PRDist.EmployeeID
            SQLString = "TempID = " & TempID
            trs.Find SQLString, 0, adSearchForward, 1
            If trs.EOF Then
                trs.AddNew
                trs!TempID = TempID
                trs!EmployeeID = PRDist.EmployeeID
                trs!CityID = PRDist.CourtesyCityID
                trs!YTDGross = 0
                trs!YTDTax = 0
                trs.Update
            End If
            trs!YTDGross = trs!YTDGross + PRDist.CityWage
            trs!YTDTax = trs!YTDTax + PRDist.CourtesyCityTax
            trs.Update
        End If
        
        If Not PRDist.GetNext Then Exit Do
        
    Loop
    
    trs.Sort = "TempID"
    trs.MoveFirst
    LastCityID = 0
    Ln = 0
    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            YECityHeader (ReportTitle)
        End If
        
        If LastCityID = 0 Or LastCityID <> trs!CityID Then
            If PRCity.GetBySQL("Select * from PRCity where PRCity.CityID = " & trs!CityID) Then
                CityName = PRCity.CityName
                CityNumber = PRCity.CityNumber
            Else
                CityName = CityNumber & " Not Found!"
            End If
            
            PrintValue(1) = CityNumber:                         FormatString(1) = "a5"
            PrintValue(2) = "   *** REPORT FOR CITY OF:  " & Trim(CityName) & "  ***"
                                                                FormatString(2) = "a60"
            PrintValue(3) = " ":                                FormatString(3) = "~"
            
            FormatPrint
            Ln = Ln + 2
        End If
               
        If Not PREmployee.GetByID(trs!EmployeeID) Then
            MsgBox "Employee Info Not Found!!!", vbExclamation, "Payroll Yearly City Tax Report"
            GoBack
        End If
        
        frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.LFName)
        frmProgress.Show
        
        PrintValue(1) = PREmployee.EmployeeNumber:              FormatString(1) = "a7"
        PrintValue(2) = Trim(PREmployee.LFName):                FormatString(2) = "a35"
        PrintValue(3) = PREmployee.SSString:                    FormatString(3) = "a16"
        PrintValue(4) = Trim(PREmployee.Address1):              FormatString(4) = "a28"
        PrintValue(5) = trs!YTDGross:                           FormatString(5) = "d16"
        PrintValue(6) = trs!YTDTax:                             FormatString(6) = "d16"
        PrintValue(7) = " ":                                    FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1
        
        SYTDGROSS = SYTDGROSS + trs!YTDGross
        TYTDGross = TYTDGross + trs!YTDGross
        SYTDTAX = SYTDTAX + trs!YTDTax
        TYTDTAX = TYTDTAX + trs!YTDTax

        If Trim(PREmployee.Address2) <> "" Then
            PrintValue(1) = " ":                        FormatString(1) = "a58"
            PrintValue(2) = Trim(PREmployee.Address2):  FormatString(2) = "a30"
            PrintValue(3) = " ":                        FormatString(3) = "~"
            FormatPrint
            Ln = Ln + 1
        End If
        
        PrintValue(1) = " ":                                    FormatString(1) = "a58"
        PrintValue(2) = Trim(PREmployee.City) & "  " & PREmployee.State & "  " & PREmployee.ZipCode
                                                                FormatString(2) = "a40"
        PrintValue(3) = " ":                                    FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 2
              
        LastCityID = trs!CityID
        LastCityNumber = CityNumber
        LastCityName = CityName
        
        trs.MoveNext
        
        If trs.EOF Then
            YECityTaxTotals
            Exit Do
        End If
    
        If LastCityID <> trs!CityID Then
            YECityTaxTotals
        End If
        
    Loop
    
    SYTDGROSS = TYTDGross
    SYTDTAX = TYTDTAX
    LastCityNumber = 0
    LastCityName = Trim(PRCompany.Name)
    YECityTaxTotals
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Private Sub YECityTaxTotals()
            
    Ln = Ln + 1
    PrintValue(1) = LastCityNumber & " - " & Trim(LastCityName) & " TOTALS"
                                                        FormatString(1) = "a88"
    PrintValue(2) = SYTDGROSS:                          FormatString(2) = "d12"
    PrintValue(3) = " ":                                FormatString(3) = "a2"
    PrintValue(4) = SYTDTAX:                            FormatString(4) = "d12"
    PrintValue(5) = " ":                                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    SYTDGROSS = 0
    SYTDTAX = 0
    
    If LastCityNumber <> 0 Then
        FormFeed
        YECityHeader (ReportTitle)
        FormatPrint
        Ln = Ln + 2
    End If

End Sub


'=======================================   WAGE REVIEW    ======================================

Public Sub OHBUCJournal()

    Ln = 0
    SetEquates
    NumEmployees = 0
    PrtInit ("Port")
    ReportTitle = "EMPLOYER'S REPORT OF WAGES - JOURNAL"
    SetFont 10, Equate.Portrait
    
    ' set up SQL statement based upon order requested
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If frmOHBUC.cmbQtr = 1 Then
      QtrEnding = "QUARTER ENDING: 03/31/" & frmOHBUC.cmbYear
    ElseIf frmOHBUC.cmbQtr = 2 Then
      QtrEnding = "QUARTER ENDING: 06/30/" & frmOHBUC.cmbYear
    ElseIf frmOHBUC.cmbQtr = 3 Then
      QtrEnding = "QUARTER ENDING: 09/30/" & frmOHBUC.cmbYear
    ElseIf frmOHBUC.cmbQtr = 4 Then
      QtrEnding = "QUARTER ENDING: 12/31/" & frmOHBUC.cmbYear
    End If
    
    Msg1 = QtrEnding
    
    frmOHBUC.rs.MoveFirst
    If frmOHBUC.optEmployee Then
        frmOHBUC.rs.Sort = "EmpID"
    Else
        frmOHBUC.rs.Sort = "SSN"
    End If
    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            SetFont 10, Equate.Portrait
            
            ' data header
            Ln = Ln + 2                ' Changed from Ln +1 to Ln + 2
            
            PrintValue(1) = " ":                            FormatString(1) = "a3"
            PrintValue(2) = "SS NUMBER":                    FormatString(2) = "a11"
            PrintValue(3) = " ":                            FormatString(3) = "a2"
            PrintValue(4) = "EMPLOYEE NAME":                FormatString(4) = "a20"
            PrintValue(5) = " ":                            FormatString(5) = "a7"
            PrintValue(6) = "GROSS WAGE":                   FormatString(6) = "a10"
            PrintValue(7) = " ":                            FormatString(7) = "a3"
            PrintValue(8) = "WKS":                          FormatString(8) = "a3"
            PrintValue(9) = " ":                            FormatString(9) = "~"
            FormatPrint
'            Ln = Ln + 1
            PrintValue(1) = " ":                            FormatString(1) = "3"
            PrintValue(1) = String(99, "="):                FormatString(1) = "a99"
            PrintValue(2) = " ":                            FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 2
        End If
            
            frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.LFName)
            frmProgress.Show
        
'            If PREmployee.SSN = 0 Then GoTo cycle3
            
            PrintValue(1) = " ":                            FormatString(1) = "a3"
            PrintValue(2) = PREmployee.SSString:            FormatString(2) = "a11"
            NumEmployees = NumEmployees + 1
            PrintValue(3) = " ":                            FormatString(3) = "a2"
            PrintValue(4) = frmOHBUC.rs!EmpName:            FormatString(4) = "a25"
                                   
            TotWageGross = TotWageGross + frmOHBUC.rs!Gross
            PrintValue(5) = frmOHBUC.rs!Gross:              FormatString(5) = "d13"
            PrintValue(6) = " ":                            FormatString(6) = "a3"
            PrintValue(7) = frmOHBUC.rs!NoWeeks:            FormatString(7) = "a3"
            PrintValue(8) = " ":                            FormatString(8) = "~"
            FormatPrint
            Ln = Ln + 1
'cycle3:
        frmOHBUC.rs.MoveNext
        If frmOHBUC.rs.EOF Then Exit Do
    Loop
    If Ln > MaxLines Then
      
        If Ln Then FormFeed
        Ln = Ln + 3
        PageHeader ReportTitle, " ", "", ""
        SetFont 10, Equate.Portrait
    End If
    ' data header
    Ln = Ln + 2                ' Changed from Ln +1 to Ln + 2
    
    PrintValue(1) = " ":                            FormatString(1) = "a3"
    PrintValue(2) = "No. Employees":                FormatString(2) = "a13"
    PrintValue(3) = NumEmployees:                   FormatString(3) = "n5"
    PrintValue(4) = "":                             FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                            FormatString(1) = "a3"
    PrintValue(2) = "Total Gross":                  FormatString(2) = "a13"
    PrintValue(3) = TotWageGross:                   FormatString(3) = "d14"
    PrintValue(4) = " ":                            FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
   
End Sub

'=======================================   SUPPLEMENTAL     ======================================

Public Sub OHBUCPurple(ByVal FormColor As String, ByVal FromRed As Boolean)

Dim AmtString As String
Dim ContSw As Byte

    ' setup page if red page not printed first
    If FromRed = False Then
        SetEquates
        PrtInit ("Port")
        ReportTitle = "labels "
        SetFont 11, Equate.Portrait
        PrtTitle = frmOHBUC.txtTitle
        NumEmployees = 0
        CurrPg = 0
        LnCnt = 0
        Ln = 0
        
        frmOHBUC.rs.MoveLast
        ' get the number of pages
        ii = frmOHBUC.rs.RecordCount
        NumPages = Int(ii / 20)
        If ii Mod 20 <> 0 Then NumPages = NumPages + 1
        With frmOHBUC.tdbnumStartPageNum
            If .Value > 1 Then
                NumPages = NumPages + .Value - 1
                CurrPg = .Value - 1
            End If
        End With
    End If
    
    frmOHBUC.rs.MoveFirst
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    Else
        StateAbbrev = ""
    End If
                           
    Do
        EmpCt = EmpCt + 1
        
        If FromRed = True Then
            If EmpCt > 20 And ContSw = 0 Then
                CurrPg = 1
                Ln = 0
                EmpCt = 1
                NumEmployees = 0
                ContSw = 1
            ElseIf ContSw = 0 Then
                GoTo CycleIt
            End If
        End If
        If Ln = 0 Then
            Ln = Ln + 5
            LnCnt = 0
           
            PosPrint 9200, 900, qQuarter
            PosPrint 9850, 900, Format(qYear Mod 100, "00")
            PosPrint 600, 1100, PRCompany.StateUnempID
            PosPrint 1000, 1650, PRCompany.Name
            PosPrint 1000, 1850, PRCompany.Address1
            
            ii = 1850
            If PRCompany.Address2 <> "" Then
                ii = ii + 200
                PosPrint 1000, ii, PRCompany.Address2
            End If
            ii = ii + 200
            PosPrint 1000, ii, PRCompany.City & ", " & PRState.StateAbbrev & " " & PRCompany.ZipCode
            
            ' vertical position init
            ii = 3550
            
        End If
        
        If EmpCt = 6 Then
            ii = ii - 50
        ElseIf EmpCt = 11 Then
            ii = ii - 75
        ElseIf EmpCt = 16 Then
            ii = ii - 75
        ElseIf EmpCt = 18 Then
            ii = ii - 25
        End If

        PosPrint 400, ii, Format(frmOHBUC.rs!SSN, "000-00-0000")
        PosPrint 2600, ii, frmOHBUC.rs!EmpName
        
        AmtString = Format(frmOHBUC.rs!Gross, "##,###,##0.00")
        PosPrint 6250, ii, PadRight(Format(frmOHBUC.rs!Gross, "##,###,##0.00"), 13)
        
        PosPrint 8350, ii, PadRight(frmOHBUC.rs!NoWeeks, 2)
        ii = ii + 500       ' increment the vertical position
        
        TotWageGross = TotWageGross + frmOHBUC.rs!Gross
CycleIt:
        NumEmployees = NumEmployees + 1
        frmOHBUC.rs.MoveNext
        
        If frmOHBUC.rs.EOF Or NumEmployees = 20 Then
            EmpCt = 0
            CurrPg = CurrPg + 1
            
            ' print Underlines for bottom of form fields
            If FormColor = "Plain" Then
                PosPrint 300, 14200, String(25, "_")
                PosPrint 1100, 14450, "SIGNATURE"
                
                PosPrint 4000, 14200, String(25, "_")
                PosPrint 5500, 14450, "TITLE"
                        
                PosPrint 7800, 14200, String(13, "_")
                PosPrint 8400, 14450, "DATE"
            End If
            
            PosPrint 6250, 13350, PadRight(Format(TotWageGross, "##,###,##0.00"), 13)
            PosPrint 8350, 13350, CurrPg
            PosPrint 9080, 13350, NumPages
            PosPrint 4200, 14100, PrtTitle
            PosPrint 8000, 14100, Format(PrtDate, "mm/dd/yyyy")
            
            If frmOHBUC.rs.EOF = False Then FormFeed
            Ln = 0
            NumEmployees = 0
            TotWageGross = 0
        End If

        If frmOHBUC.rs.EOF Then Exit Do
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
   
End Sub

Public Sub OHBUCRed(ByVal FormColor As String)       '  Employer's Report of Wages Form  (RED)

Dim AmtString As String

    SetEquates
    NumEmployees = 0
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 11, Equate.Portrait
    PrtTitle = frmOHBUC.txtTitle
    frmOHBUC.rs.MoveLast
    
    ' get the number of pages
    ii = frmOHBUC.rs.RecordCount
    NumPages = Int(ii / 20)
    If ii Mod 20 <> 0 Then NumPages = NumPages + 1
    
    frmOHBUC.rs.MoveFirst
    '  Get total gross amount for header
    Do
        TotGross = TotGross + frmOHBUC.rs!Gross
        frmOHBUC.rs.MoveNext
        If frmOHBUC.rs.EOF Then Exit Do
    Loop
        
    NumEmployees = 0
    CurrPg = 0
    LnCnt = 0
    Ln = 0
    
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    Else
        StateAbbrev = ""
    End If
    frmOHBUC.rs.MoveFirst
    Do

        If Ln = 0 Then
            Ln = Ln + 5
            LnCnt = 0
           
            PosPrint 7500, 2050, NumPages
            PosPrint 9150, 2050, PadRight(Format(TotGross, "##,###,##0.00"), 13)
            ' vertical position init
            ii = 3450
        End If

        EmpCt = EmpCt + 1
        If EmpCt = 6 Then
            ii = ii - 50
        ElseIf EmpCt = 11 Then
            ii = ii - 50
        ElseIf EmpCt = 16 Then
            ii = ii - 75
        ElseIf EmpCt = 18 Then
            ii = ii - 25
        End If
            
        PosPrint 400, ii, Format(frmOHBUC.rs!SSN, "000-00-0000")
        PosPrint 2600, ii, frmOHBUC.rs!EmpName
        
        AmtString = Format(frmOHBUC.rs!Gross, "##,###,##0.00")
        PosPrint 6250, ii, PadRight(Format(frmOHBUC.rs!Gross, "##,###,##0.00"), 13)
        
        PosPrint 8350, ii, PadRight(frmOHBUC.rs!NoWeeks, 2)
        ii = ii + 500       ' increment the vertical position
        
        TotWageGross = TotWageGross + frmOHBUC.rs!Gross
        NumEmployees = NumEmployees + 1
        
        frmOHBUC.rs.MoveNext
        
        If frmOHBUC.rs.EOF Or NumEmployees = 20 Then
            EmpCt = 0
            CurrPg = CurrPg + 1
            PosPrint 6250, 13250, PadRight(Format(TotWageGross, "##,###,##0.00"), 13)
            PosPrint 8350, 13250, CurrPg
            PosPrint 9080, 13250, NumPages
            TotWageGross = 0
            If frmOHBUC.rs.EOF = False Then FormFeed
            If frmOHBUC.rs.RecordCount = NumEmployees Then
                Exit Do
            Else
                OHBUCPurple FormColor, True
            End If
        End If

        If frmOHBUC.rs.EOF Then Exit Do
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
   
End Sub

Public Sub OHBUCRed201009()       '  Employer's Report of Wages Form  (RED)

Dim AmtString As String

Dim xx, yy, zz As Long

    SetEquates
    NumEmployees = 0
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 12, Equate.Portrait
    PrtTitle = frmOHBUC.txtTitle
    frmOHBUC.rs.MoveLast
    
    NumEmployees = 0
    frmOHBUC.rs.MoveFirst
    '  Get total gross amount for header
    Do
        TotGross = TotGross + frmOHBUC.rs!Gross
        If frmOHBUC.rs!Gross > 0 Then NumEmployees = NumEmployees + 1
        frmOHBUC.rs.MoveNext
        If frmOHBUC.rs.EOF Then Exit Do
    Loop
        
    ' get the number of pages
    ' 15 lines
    ' continuation sheets same format as red form
    NumPages = Int(NumEmployees / 15)
    If NumEmployees Mod 15 <> 0 Then NumPages = NumPages + 1
    PageNum = frmOHBUC.tdbnumStartPageNum
    NumPages = NumPages + PageNum - 1
    
    LnCnt = 0
    Ln = 0
    ThisPage_Count = 0
    ThisPage_Amount = 0
    
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    Else
        StateAbbrev = ""
    End If
    
    With frmOHBUC
    
        .rs.MoveFirst
        
        Do

            If .rs!Gross <= 0 Then GoTo NextBUC

            If Ln = 0 Then
            
                ' print company info header ?
                If PageNum <> 1 Or (PageNum = 1 And .chkRed = 0) Then
                    
                    SetFont 11, Equate.Portrait
            
                    yy = 1800
                    PosPrint 690, yy, PRCompany.StateUnempID
                    PosPrint 5300, yy, PRCompany.FederalID
                    PosPrint 8650, yy, .cmbQtr
                    PosPrint 9700, yy, .cmbYear
                    
                    yy = 2300
                    PosPrint 690, yy, PRCompany.Name
                    
                    Dim Mth1, Mth2, Mth3 As String
                    If .cmbQtr.text = "1" Then
                        Mth1 = "JANUARY"
                        Mth2 = "FEBRUARY"
                        Mth3 = "MARCH"
                    ElseIf .cmbQtr.text = "2" Then
                        Mth1 = "APRIL"
                        Mth2 = "MAY"
                        Mth3 = "JUNE"
                    ElseIf .cmbQtr.text = "3" Then
                        Mth1 = "JULY"
                        Mth2 = "AUGUST"
                        Mth3 = "SEPTEMBER"
                    Else
                        Mth1 = "OCTOBER"
                        Mth2 = "NOVEMBER"
                        Mth3 = "DECEMBER"
                    End If
                            
                    yy = 3800
                    xx = 700
                    zz = 1600
                    PosPrint xx, yy, Mth1
                    xx = xx + zz
                    PosPrint xx, yy, Mth2
                    xx = xx + zz
                    PosPrint xx, yy, Mth3
                    
                    SetFont 12, Equate.Portrait
            
                End If
            
                Ln = Ln + 5
                LnCnt = 0
                
                yy = 2800
                ' number of pages
                OHBox 510, yy, Format(NumPages, "000")
                 
                ' total number of employees
                OHBox 3640, yy, Format(NumEmployees, "00000")
                 
                ' total gross
                OHBox 7780, yy, Format(TotGross * 100, "000000000000")
                            
                ' employee counts
                yy = 4200
                xx = 460
                zz = 1620
                OHBox xx, yy, Format(.tdbEmpCount1, "0000")
                xx = xx + zz
                OHBox xx, yy, Format(.tdbEmpCount2, "0000")
                xx = xx + zz
                OHBox xx, yy, Format(.tdbEmpCount3, "0000")
            
                yy = 5200
                zz = 482
        
            End If      ' header
        
            OHBox 323, yy, Format(.rs!SSN, "000000000")
            OHBox 3330, yy, .rs!LastName, 10
            OHBox 6700, yy, .rs!FirstName, 1
            OHBox 7200, yy, .rs!MidInit, 1
            OHBox 7600, yy, Format(.rs!Gross * 100, "0000000000")
            OHBox 11060, yy, Format(.rs!NoWeeks, "00")
        
            ThisPage_Count = ThisPage_Count + 1
            ThisPage_Amount = ThisPage_Amount + .rs!Gross
        
            yy = yy + zz
        
            ' next page
            If ThisPage_Count = 15 Then
                If PageNum <> 1 Or (PageNum = 1 And .chkRed = 0) Then
                    OHBUC201009_Footer True
                Else
                    OHBUC201009_Footer False
                End If
                FormFeed
            End If
        
NextBUC:
            frmOHBUC.rs.MoveNext
            If frmOHBUC.rs.EOF Then Exit Do
    
        Loop
        
    End With
    
    If ThisPage_Count <> 0 Then
        If PageNum <> 1 Or (PageNum = 1 And frmOHBUC.chkRed = 0) Then
            OHBUC201009_Footer True
        Else
            OHBUC201009_Footer False
        End If
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
   
End Sub

Private Sub OHBUC201009_Footer(ByVal RedForm As Boolean)

Dim xx, yy, zz As Integer
Dim Under As String

    ' print this page totals
    yy = 12600
    OHBox 2860, yy, Format(ThisPage_Count, "00")
    OHBox 6935, yy, Format(ThisPage_Amount * 100, "000000000000")
    
    ' page count
    yy = 13250
    PosPrint 10100, yy, PageNum
    PosPrint 11000, yy, NumPages
    PageNum = PageNum + 1
    
    ' underlines if not on red form
    If RedForm = True Then
        
        ' title of field below the underline
        ' vertical incr
        zz = 240
        
        ' signature line
        yy = 13500
        Under = String(38, "_")
        PosPrint 250, yy, Under
        PosPrint 250, yy + zz, "Signed"
        
        ' title & date
        Under = String(23, "_")
        yy = 14260
        PosPrint 250, yy, Under
        PosPrint 250, yy + zz, "Title"
        
        Under = String(15, "_")
        PosPrint 3680, yy, Under
        PosPrint 3680, yy + zz, "Date"
        
    End If
    
    ' title and date
    yy = 14250
    PosPrint 350, yy, frmOHBUC.txtTitle
    PosPrint 3900, yy, Format(frmOHBUC.TDBDate1, "mm/dd/yyyy")

    ThisPage_Count = 0
    ThisPage_Amount = 0

End Sub

Private Sub OHBox(ByVal xPos As Integer, _
                    ByVal yPos As Integer, _
                    ByVal InString As String, _
                    Optional MaxLen As Integer)

Dim HorzSpace As Integer
Dim StrPos As Integer
Dim EndLoop As Integer

    If MaxLen <> 0 Then
        EndLoop = MaxLen
    Else
        EndLoop = Len(Trim(InString))
    End If

    HorzSpace = 330
    For StrPos = 1 To EndLoop
        PosPrint xPos, yPos, Mid(InString, StrPos, 1)
        xPos = xPos + HorzSpace
    Next StrPos

End Sub

Public Sub OHBUCPurple201009(ByVal FormColor As String, ByVal FromRed As Boolean)

Dim AmtString As String
Dim ContSw As Byte

    ' setup page if red page not printed first
    If FromRed = False Then
        SetEquates
        PrtInit ("Port")
        ReportTitle = "labels "
        SetFont 11, Equate.Portrait
        PrtTitle = frmOHBUC.txtTitle
        NumEmployees = 0
        CurrPg = 0
        LnCnt = 0
        Ln = 0
        
        frmOHBUC.rs.MoveLast
        ' get the number of pages
        ii = frmOHBUC.rs.RecordCount
        NumPages = Int(ii / 20)
        If ii Mod 20 <> 0 Then NumPages = NumPages + 1
        With frmOHBUC.tdbnumStartPageNum
            If .Value > 1 Then
                NumPages = NumPages + .Value - 1
                CurrPg = .Value - 1
            End If
        End With
    End If
    
    frmOHBUC.rs.MoveFirst
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    Else
        StateAbbrev = ""
    End If
                           
    Do
        EmpCt = EmpCt + 1
        
        If FromRed = True Then
            If EmpCt > 20 And ContSw = 0 Then
                CurrPg = 1
                Ln = 0
                EmpCt = 1
                NumEmployees = 0
                ContSw = 1
            ElseIf ContSw = 0 Then
                GoTo CycleIt
            End If
        End If
        If Ln = 0 Then
            Ln = Ln + 5
            LnCnt = 0
           
            PosPrint 9200, 900, qQuarter
            PosPrint 9850, 900, Format(qYear Mod 100, "00")
            PosPrint 600, 1100, PRCompany.StateUnempID
            PosPrint 1000, 1650, PRCompany.Name
            PosPrint 1000, 1850, PRCompany.Address1
            
            ii = 1850
            If PRCompany.Address2 <> "" Then
                ii = ii + 200
                PosPrint 1000, ii, PRCompany.Address2
            End If
            ii = ii + 200
            PosPrint 1000, ii, PRCompany.City & ", " & PRState.StateAbbrev & " " & PRCompany.ZipCode
            
            ' vertical position init
            ii = 3550
            
        End If
        
        If EmpCt = 6 Then
            ii = ii - 50
        ElseIf EmpCt = 11 Then
            ii = ii - 75
        ElseIf EmpCt = 16 Then
            ii = ii - 75
        ElseIf EmpCt = 18 Then
            ii = ii - 25
        End If

        PosPrint 400, ii, Format(frmOHBUC.rs!SSN, "000-00-0000")
        PosPrint 2600, ii, frmOHBUC.rs!EmpName
        
        AmtString = Format(frmOHBUC.rs!Gross, "##,###,##0.00")
        PosPrint 6250, ii, PadRight(Format(frmOHBUC.rs!Gross, "##,###,##0.00"), 13)
        
        PosPrint 8350, ii, PadRight(frmOHBUC.rs!NoWeeks, 2)
        ii = ii + 500       ' increment the vertical position
        
        TotWageGross = TotWageGross + frmOHBUC.rs!Gross
CycleIt:
        NumEmployees = NumEmployees + 1
        frmOHBUC.rs.MoveNext
        
        If frmOHBUC.rs.EOF Or NumEmployees = 20 Then
            EmpCt = 0
            CurrPg = CurrPg + 1
            
            ' print Underlines for bottom of form fields
            If FormColor = "Plain" Then
                PosPrint 300, 14200, String(25, "_")
                PosPrint 1100, 14450, "SIGNATURE"
                
                PosPrint 4000, 14200, String(25, "_")
                PosPrint 5500, 14450, "TITLE"
                        
                PosPrint 7800, 14200, String(13, "_")
                PosPrint 8400, 14450, "DATE"
            End If
            
            PosPrint 6250, 13350, PadRight(Format(TotWageGross, "##,###,##0.00"), 13)
            PosPrint 8350, 13350, CurrPg
            PosPrint 9080, 13350, NumPages
            PosPrint 4200, 14100, PrtTitle
            PosPrint 8000, 14100, Format(PrtDate, "mm/dd/yyyy")
            
            If frmOHBUC.rs.EOF = False Then FormFeed
            Ln = 0
            NumEmployees = 0
            TotWageGross = 0
        End If

        If frmOHBUC.rs.EOF Then Exit Do
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
   
End Sub


