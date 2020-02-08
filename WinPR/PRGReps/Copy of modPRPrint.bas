Attribute VB_Name = "modPRPrint"
Public BMo1Tax As Currency
Public BMo2Tax As Currency
Public BMo3Tax As Currency
Public BMoTax As Currency
Public TotTaxLiability As Currency
Dim ReportTitle As String
Dim ENumber As Long
Dim VarLength As Byte
Dim TextTitle As String
Dim TitleChars As String
Dim SpreadNumber As Long
Dim NumOfSpaces As String
Dim ANumber As String
Dim TitleCount As Long
Dim TitleString As String
Dim CoName As String
Dim BankName As String
Dim FOOKAmt As Currency
Dim PrtLine
Dim FirstSw As Boolean
Dim TChannel As Integer
Dim RteNo As String
Dim SeqNo As Long
Dim BatNo As String
Dim DepositTotal As Currency
Dim FedID As String
Dim EEABA As Long
Dim Hash As Double
Dim WriteCt As Long

Dim TotalGrossPay As Currency
Dim RecCtr As Long
Dim DepFICAWH As Currency
Dim DepMEDWH As Currency
Dim DepFICAMatch As Currency
Dim DepMEDMatch As Currency
Dim DepFedTaxWH As Currency
Dim DepSTTaxWH As Currency
Dim DepCityTaxWH As Currency
Dim DepSTUnemp As Currency
Dim DepFedUnemp As Currency
Dim DepFedDep As Currency
Dim DepTotEscrowed As Currency
Dim DepFICAAmt As Currency
Dim DepFICAPct As Currency
Dim DepMedAmt As Currency
Dim DepMedPct As Currency
Dim DepSTUnempAmt As Currency
Dim DepSTUnempPct As Currency
Dim DepFedUnempAmt As Currency
Dim DepFedUnempPct As Currency

Dim LastBatch As Long
Dim LineCt As Long
Dim TotCredAmt As Currency
Dim SubCredAmt As Currency
Dim SubLineCt As Long

Dim w, Y, z As String
Dim q, r As String
Dim i, j, k As Integer
Dim c As Currency
Dim SString As String

Dim QtrRptString As String
Dim ColumnCount As Integer
Dim DeptID As Long
Dim ExitSw As Boolean
Dim Pg As Integer
Dim ContFlg, PgBrk As Boolean
Dim Dgt As String
Dim LRow As Integer
Dim RBC As String
Dim LastEmpLName As String
Dim LastEmpFName As String
Dim LastChkDate As String
Dim LastEmpNo As Long
Dim LastEmpName As String
Dim EmpFlag As Boolean
Dim TotalFlag As Boolean
Dim EmpID As Long
Dim TotWageGross As Currency

Dim TotWage As Currency
Dim TotWageSS As Currency
Dim TotWageMed As Currency
Dim TotTaxFed As Currency
Dim TotTaxSS As Currency
Dim TotTaxMed As Currency
Dim QTDWageGross As Currency
Dim YTDWageGross As Currency
Dim FinalFica As Currency

Dim TotTaxes As Currency
Dim TOTHours As Single

Dim TotWageState As Currency
Dim TotWageCity As Currency
Dim TotTaxState As Currency
Dim TotTaxCity As Currency

Dim TotQTDWageGross As Currency
Dim TotYTDWageGross As Currency
Dim TotWageFed As Currency
Dim TotWageFIC As Currency

Dim TipsFIC As Currency
Dim TotTipsFIC As Currency
Dim TotTipsMed As Currency
Dim TotWeeks As Integer
Dim NumPages As Integer
Dim CurrPg As Integer

Dim YMStartDate As String
Dim YMEndDate As Date
Dim SYMYear As Long
Dim SYMMonth As String
Dim EYMYear As Long
Dim EYMMonth As String

Dim NoRecords As Long
Dim CheckAmt As Currency
Dim DepoAmt As Currency

Dim PadString As String
Dim CustLen As Long
Dim AmtLen As Long
Dim PadNumber As Long
Dim PadSpaces As Long
Dim CustName As String
Dim DPadNumber As Long

Public Msg1 As String
'Public RangeType As Byte

Public PEDate As Long
Public CheckDt As Long
Public GrossPay As Currency
Public NetPay As Currency
Public qYear As Long
Public qQuarter As Byte
Public StartMonth As Long
Public Quarter As Long
Public EndMonth As Long
Public StateAbbr As String

Public p1 As Currency
Public p2 As Currency
Public TrsDept As Double
Public FindStr As String

Public LastEmpNumber As Long
Public LastCityNumber As Long
Public NumEmployees  As Long
Public NumThirteen As Byte
Public EmpName As String
Public CityName As String
Public CityNumber As Long
Public SMTDGross As Currency
Public SMTDTax As Currency
Public SQTDGross As Currency
Public SQTDTax As Currency
Public SYTDGROSS As Currency
Public SYTDTAX As Currency

Public TMTDGross As Currency
Public TMTDTax As Currency
Public TQTDGross As Currency
Public TQTDTax As Currency
Public TYTDGross As Currency
Public TYTDTAX As Currency

Public rs As New ADODB.Recordset
'Public CurrYear As Long
'Public CurrDate As Date
'Public PrtDate As Long
Public TxtDate As Long
Public PrtTitle As String
Public PrtEin1 As Byte
Public PrtEin2 As Byte
Public PrtEin3 As Byte
Public PrtEin4 As Byte
Public PrtEin5 As Byte
Public PrtEin6 As Byte
Public PrtEin7 As Byte
Public PrtEin8 As Byte
Public PrtEin9 As Byte

Public PrtName As String
Public PrtYr1 As Byte
Public PrtYr2 As Byte
Public PrtYr3 As Byte
Public PrtYr4 As Byte
Public GridCnt As Byte
Public RowCnt As Byte
Public RecCnt As Long

Public OEHours, OEAmount, DEDAmount As Byte
Public OEHrsPrt, OEAmtPrt, DEDAmtPrt As Byte
Public sqlstring1 As String
Public WrittenAmount As String
Public DedString As String

Public rrs As ADODB.Recordset
Public ers As New ADODB.Recordset
Public trs As New ADODB.Recordset
Public trsDED As New ADODB.Recordset
Public trsDEDTot As New ADODB.Recordset
Public ItemCount As Long

Public Sub EEList(ByVal ReportType As String)
Dim ReportTitle
Dim LabelColumns As Long
Dim LabelRows As Integer
Dim MaxLabels As Integer
Dim LabelCount As Integer
LabelCount = 0

    SetEquates
    frmLists.Hide
    Msg2 = "Date: " & Format(Now, "mm/dd/yyyy")
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    ' page set up based on the report type
    Select Case ReportType
    
        Case "NumberName"
            PrtInit ("Port")    ' "Port" = Portrait
            ReportTitle = "EMPLOYEE NUMBER/NAME LIST"
            SetFont 10, Equate.Portrait
        Case "DetailList"
            PrtInit ("Land")    ' "Land" = Landscape
            LandSW = 1
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
            PrtInit ("Land")    ' "Land" = Landscape
            LandSW = 1
            ReportTitle = "RATE TAX LISTING"
            SetFont 8, Equate.LandScape
        Case "TimeCardLabels"
            PrtInit ("Port")
            YUnits = 235        '  corrects the lining up of data on the labels
            ReportTitle = "labels "
            SetFont 10, Equate.Portrait
        Case "MailingLabels"
            PrtInit ("Port")
            YUnits = 235
            ReportTitle = "labels "
            SetFont 10, Equate.Portrait
    End Select
                                                                
    ' set up SQL statement based upon order requested
    
    If frmLists.optNumber Then
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NO."
        End If
        SQLString = "Select * from PREmployee ORDER BY EmployeeNumber"
    ElseIf frmLists.optName Then
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NAME"
        End If
        SQLString = "Select * from PREmployee ORDER BY PREmployee.LastName, PREmployee.FirstName"
    Else
        If ReportTitle <> "labels " Then
            ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE ZIP CODE"
        End If
        SQLString = "Select * from PREmployee ORDER BY ZipCode"
    End If
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Found !!!", vbCritical, "Employee Lists and Labels"
        Exit Sub
    End If
    
    If DptCt = 0 Then
        Msg1 = "All Departments"
    Else
        Msg1 = "All Departments are Not Included"
    End If
    
    If frmLists.chkInactive = 1 And frmLists.chkSalaried = 1 Then
        Msg1 = Trim(Msg1) & " - Includes Inactive and Salaried Employees"
    ElseIf frmLists.chkInactive = 1 Then
        Msg1 = Trim(Msg1) & " - Includes Inactive Employees"
    ElseIf frmLists.chkSalaried = 1 Then
        Msg1 = Trim(Msg1) & " - Includes Salaried Employees"
    End If
    
    Do
    
         ' **** department filter
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            MsgBox "Department Info Not Found!!!", vbCritical, "Payroll Employee Lists"
            End
        End If
        
        FindStr = "Dept=" & CStr(PRDepartment.DepartmentNumber)
        Dpts.Find FindStr, 0, adSearchForward, 1
        If Dpts.EOF Then
            GoTo Cycle1
        End If
        
        ' **** inactive filter
        If frmLists.chkInactive = 0 And PREmployee.Inactive = 1 Then
            GoTo Cycle1
        End If

        ' **** salaried filter
        If frmLists.chkSalaried = 0 And PREmployee.Salaried = 1 Then
            GoTo Cycle1
        End If
        
        If Ln = 0 Or Ln > MaxLines Then
         
            If Ln Then FormFeed
            If ReportTitle = "labels " Then
'                Ln = Ln + 2
            Else
                PageHeader ReportTitle, Msg1, Msg2, ""
            End If
            ' data header
            Ln = Ln + 2
        
            Select Case ReportType
                
                Case "NumberName"                   ''''''''''   Number/Name
                    PrintValue(1) = "EMPLOYEE NUMBER"
                    FormatString(1) = "a15"
         
                    PrintValue(2) = " "
                    FormatString(2) = "a3"
         
                    PrintValue(3) = "EMPLOYEE NAME"
                    FormatString(3) = "a40"
                    
                    PrintValue(4) = " "
                    FormatString(4) = "~"
            
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "-----------------------------------------------------------------------------------------------------------------"
                    FormatString(1) = "a90"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
'=============================================================================================
                    
                Case "DetailList"                   ''''''''''   Detail Listing
                    SetFont 8, Equate.LandScape
                    PrintValue(1) = "EMPLOYEE NO."       '  First Heading
                    FormatString(1) = "a11"
         
                    PrintValue(2) = " "
                    FormatString(2) = "a2"
         
                    PrintValue(3) = "DEPARTMENT"
                    FormatString(3) = "a12"
                    
                    PrintValue(4) = " "
                    FormatString(4) = "a4"
       
                    PrintValue(5) = "EMPLOYEE NAME"
                    FormatString(5) = "a35"
                    
                    PrintValue(6) = " "
                    FormatString(6) = "a1"
                    
                    PrintValue(7) = "ADDRESS"
                    FormatString(7) = "a30"
                    
                    PrintValue(8) = " "
                    FormatString(8) = "a4"
                    
                    PrintValue(9) = "CITY"
                    FormatString(9) = "a20"
                    
                    PrintValue(10) = " "
                    FormatString(10) = "a1"
                    
                    PrintValue(11) = "STATE"
                    FormatString(11) = "a5"
                    
                    PrintValue(12) = " "
                    FormatString(12) = "a6"
                    
                    PrintValue(13) = "ZIP"
                    FormatString(13) = "a20"
                                                                                
                 '*** Print SS Number?
                    If frmLists.chkSSN Then
                        PrintValue(14) = "SS NUMBER "
                        FormatString(14) = "a9"

                        PrintValue(15) = " "
                        FormatString(15) = "~"
                    Else

                        PrintValue(14) = " "
                        FormatString(14) = "~"
                    End If
         
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = " "         '  Print 2nd heading
                    FormatString(1) = "a4"
                    
                    PrintValue(2) = "DATE LAST"       '  PAID
                    FormatString(2) = "a9"
         
                    PrintValue(3) = " "
                    FormatString(3) = "a4"
         
                    PrintValue(4) = "DATE"
                    FormatString(4) = "a4"
                    
                    PrintValue(5) = " "
                    FormatString(5) = "a8"
       
                    PrintValue(6) = "DATE LAST"       ' RAISE
                    FormatString(6) = "a9"
                    
                    PrintValue(7) = " "
                    FormatString(7) = "a3"
                    
                    PrintValue(8) = "DATE LAST"       ' REVIEW
                    FormatString(8) = "a9"
                    
                    PrintValue(9) = " "
                    FormatString(9) = "a3"
                    
                    PrintValue(10) = "DATE LAST"      ' LAYOFF
                    FormatString(10) = "a9"
                    
                    PrintValue(11) = " "
                    FormatString(11) = "a3"
                    
                    PrintValue(12) = "DATE LAST"      ' RECALL
                    FormatString(12) = "a9"
                    
                    PrintValue(13) = " "
                    FormatString(13) = "a5"
                    
                    PrintValue(14) = "DATE"           ' TERMINATED
                    FormatString(14) = "a4"
                    
                    PrintValue(15) = " "
                    FormatString(15) = "a6"
                    
                    PrintValue(16) = "DATE OF"        ' BIRTH
                    FormatString(16) = "a7"
                    
                    PrintValue(17) = " "
                    FormatString(17) = "a3"
                    
                    PrintValue(18) = "TERM"
                    FormatString(18) = "a4"
                    
                    PrintValue(19) = " "
                    FormatString(19) = "a17"
                    
                    PrintValue(20) = "RACE"
                    FormatString(20) = "a4"
                    
                    PrintValue(21) = " "
                    FormatString(21) = "a5"
                  
                    PrintValue(22) = "MARITAL"        ' STATUS
                    FormatString(22) = "a7"
                    
                    PrintValue(23) = " "
                    FormatString(23) = "a5"
                    
                    PrintValue(24) = "EDU"             ' LEVEL
                    FormatString(24) = "a3"
                    
                    PrintValue(25) = " "
                    FormatString(25) = "a5"
                    
                    PrintValue(26) = "SHIFT"          ' CODE
                    FormatString(26) = "a5"
                    
                    PrintValue(27) = " "
                    FormatString(27) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                                        
                    PrintValue(1) = " "         '  Print 3rd Heading
                    FormatString(1) = "a7"
                    
                    PrintValue(2) = "PAID"      'Date
                    FormatString(2) = "a4"
         
                    PrintValue(3) = " "
                    FormatString(3) = "a6"
         
                    PrintValue(4) = "HIRED"     'Date
                    FormatString(4) = "a5"
                    
                    PrintValue(5) = " "
                    FormatString(5) = "a9"
       
                    PrintValue(6) = "RAISE"     'Date
                    FormatString(6) = "a5"
                    
                    PrintValue(7) = " "
                    FormatString(7) = "a6"
                    
                    PrintValue(8) = "REVIEW"     'Date
                    FormatString(8) = "a6"
                    
                    PrintValue(9) = " "
                    FormatString(9) = "a6"
                    
                    PrintValue(10) = "LAYOFF"    'Date
                    FormatString(10) = "a6"
                    
                    PrintValue(11) = " "
                    FormatString(11) = "a6"
                    
                    PrintValue(12) = "RECALL"   'Date
                    FormatString(12) = "a6"
                    
                    PrintValue(13) = " "
                    FormatString(13) = "a4"
                    
                    PrintValue(14) = "TERMINATED"   'Date
                    FormatString(14) = "a10"
                    
                    PrintValue(15) = " "
                    FormatString(15) = "a4"
                    
                    PrintValue(16) = "BIRTH"
                    FormatString(16) = "a5"
                                        
                    PrintValue(17) = " "
                    FormatString(17) = "a3"
                                                            
                    PrintValue(18) = "REASON"    'For Termination
                    FormatString(18) = "a6"
                    
                    PrintValue(19) = " "
                    FormatString(19) = "a10"
                    
                    PrintValue(20) = "SEX"
                    FormatString(20) = "a3"
                    
                    PrintValue(21) = " "
                    FormatString(21) = "a3"
                    
                    PrintValue(22) = "CODE"     ' Race
                    FormatString(22) = "a4"
                    
                    PrintValue(23) = " "
                    FormatString(23) = "a5"
                  
                    PrintValue(24) = "STATUS"   ' Marital
                    FormatString(24) = "a6"
                    
                    PrintValue(25) = " "
                    FormatString(25) = "a5"
                    
                    PrintValue(26) = "LEVEL"    ' Education
                    FormatString(26) = "a5"
                    
                    PrintValue(27) = " "
                    FormatString(27) = "a4"
                    
                    PrintValue(28) = "CODE"     ' Shift Code
                    FormatString(28) = "a4"
                    
                    PrintValue(29) = " "
                    FormatString(29) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "======================================================================================" _
                                  & "======================================================================================" _
                                    & "=================="
                    FormatString(1) = "a155"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
'=============================================================================================
                Case "EmployeeRateList"                 ''''''''''   Employee Rate Listing
                    PrintValue(1) = " "
                    FormatString(1) = "a84"
                    
                    PrintValue(2) = "HOURLY/"
                    FormatString(2) = "a7"
                                   
                    PrintValue(3) = " "
                    FormatString(3) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "EMPL #"
                    FormatString(1) = "a9"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "a0"

                    PrintValue(3) = "EMPLOYEE NAME"
                    FormatString(3) = "a40"

                    PrintValue(4) = " "
                    FormatString(4) = "a3"

                    PrintValue(5) = "DEPT"
                    FormatString(5) = "a4"

                    PrintValue(6) = " "
                    FormatString(6) = "a3"

                    PrintValue(7) = "DEPT NAME"
                    FormatString(7) = "a10"

                    PrintValue(8) = " "
                    FormatString(8) = "a7"

                    PrintValue(9) = "RATE"
                    FormatString(9) = "a7"

                    PrintValue(10) = " "
                    FormatString(10) = "a1"

                    PrintValue(11) = "SALARY"
                    FormatString(11) = "a6"
               
                    PrintValue(12) = " "
                    FormatString(12) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "==========================================================================================="
                    FormatString(1) = "a91"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
'=============================================================================================
                Case "SSNFormat"                    ''''''''''   SSN Format
                    PrintValue(1) = "EMPLOYEE NO."
                    FormatString(1) = "a12"

                    PrintValue(2) = " "
                    FormatString(2) = "a3"

                    PrintValue(3) = "EMPLOYEE NAME"
                    FormatString(3) = "a40"

                    PrintValue(4) = " "
                    FormatString(4) = "a3"
                    
                    PrintValue(5) = "SS NUMBER "
                    FormatString(5) = "a11"
                     
                    PrintValue(6) = " "
                    FormatString(6) = "a3"

                    PrintValue(7) = "BIRTH DATE"
                    FormatString(7) = "a10"

                    PrintValue(8) = " "
                    FormatString(8) = "a3"
                    
                    PrintValue(9) = "GENDER "
                    FormatString(9) = "a6"
                     
                    PrintValue(10) = " "
                    FormatString(10) = "~"
            
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "==========================================================================================="
                    FormatString(1) = "a91"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
'=============================================================================================

                Case "RateTaxList"                    ''''''''''   Rate Tax Listing
                    SetFont 8, Equate.LandScape
                    PrintValue(1) = " "         ' Header 1
                    FormatString(1) = "a61"
                    
                    PrintValue(2) = "NO"        ' SS Tax
                    FormatString(2) = "a4"
                                   
                    PrintValue(3) = "NO"        ' MED Tax
                    FormatString(3) = "a4"
                    
                    PrintValue(4) = "NO"        ' FED Tax
                    FormatString(4) = "a4"
                                   
                    PrintValue(5) = "NO"        ' ST Tax
                    FormatString(5) = "a4"
                                   
                    PrintValue(6) = "NO"        ' Cty Tax
                    FormatString(6) = "a5"
                                   
                    PrintValue(7) = "NO"        ' FED Unemp
                    FormatString(7) = "a7"
                    
                    PrintValue(8) = "NO"        ' State Unemp
                    FormatString(8) = "a4"
                                   
                    PrintValue(9) = " "
                    FormatString(9) = "~"
                                   
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = " "         ' Header 2
                    FormatString(1) = "a52"
                                                                
                    PrintValue(2) = "HOURLY/"
                    FormatString(2) = "a8"
                                                                               
                    PrintValue(3) = " "
                    FormatString(3) = "a1"
                                                                               
                    PrintValue(4) = "SS"        ' TAX
                    FormatString(4) = "a4"
                                   
                    PrintValue(5) = "MED"       ' TAX
                    FormatString(5) = "a4"
                                   
                    PrintValue(6) = "FED"       ' TAX
                    FormatString(6) = "a4"
                                   
                    PrintValue(7) = "ST"        ' TAX
                    FormatString(7) = "a4"
                                   
                    PrintValue(8) = "CTY"       ' TAX
                    FormatString(8) = "a5"
                                   
                    PrintValue(9) = "FED"       ' UNEMP
                    FormatString(9) = "a5"
                                                                                                              
                    PrintValue(10) = "STATE"     ' UNEMP
                    FormatString(10) = "a6"
                                                                                                              
                    PrintValue(11) = "FWT"       ' MAR
                    FormatString(11) = "a4"
                                   
                    PrintValue(12) = "    FWT"       ' EXMP
                    FormatString(12) = "a7"
                         
                    PrintValue(13) = " "
                    FormatString(13) = "a5"
                    
                    PrintValue(14) = "FWT"        ' XAMT
                    FormatString(14) = "a5"
                    
                    PrintValue(15) = " "
                    FormatString(15) = "a3"
                                   
                    PrintValue(16) = "SWT"       ' MAR
                    FormatString(16) = "a6"
                                   
                    PrintValue(17) = "   SWT"       ' EXMP
                    FormatString(17) = "a7"

                    PrintValue(18) = " "
                    FormatString(18) = "a4"
                    
                    PrintValue(19) = "SWT"       ' XAMT
                    FormatString(19) = "a3"
                                   
                    PrintValue(20) = " "
                    FormatString(20) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "EMP #"        ' Header 3
                    FormatString(1) = "a5"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "a2"

                    PrintValue(3) = "NAME"
                    FormatString(3) = "a18"

                    PrintValue(4) = " "
                    FormatString(4) = "a3"

                    PrintValue(5) = "DEPARTMENT"
                    FormatString(5) = "a10"

                    PrintValue(6) = " "
                    FormatString(6) = "a9"

                    PrintValue(7) = "RATE"
                    FormatString(7) = "a5"

                    PrintValue(8) = " "
                    FormatString(8) = "a0"

                    PrintValue(9) = "SALARY"
                    FormatString(9) = "a8"
 
                    PrintValue(10) = " "
                    FormatString(10) = "a1"
                    
                    PrintValue(11) = "TAX"
                    FormatString(11) = "a4"
                                   
                    PrintValue(12) = "TAX"
                    FormatString(12) = "a4"
                    
                    PrintValue(13) = "TAX"
                    FormatString(13) = "a4"
                                   
                    PrintValue(14) = "TAX"
                    FormatString(14) = "a4"
                                   
                    PrintValue(15) = "TAX"
                    FormatString(15) = "a4"
                                   
                    PrintValue(16) = "UNEMP"
                    FormatString(16) = "a6"
                                   
                    PrintValue(17) = "UNEMP"
                    FormatString(17) = "a6"
                                                                      
                    PrintValue(18) = "MAR"
                    FormatString(18) = "a3"
                                   
                    PrintValue(19) = "     EXMP"
                    FormatString(19) = "a10"
                                   
                    PrintValue(20) = " "
                    FormatString(20) = "a3"
                    
                    PrintValue(21) = "XAMT"
                    FormatString(21) = "a5"
                                   
                    PrintValue(22) = " "
                    FormatString(22) = "a3"
                    
                    PrintValue(23) = "MAR"
                    FormatString(23) = "a6"
                                  
                    PrintValue(24) = "   EXMP"
                    FormatString(24) = "a8"
                                                       
                    PrintValue(25) = " "
                    FormatString(25) = "a3"
                    
                    PrintValue(26) = "XAMT"
                    FormatString(26) = "a5"
                                                                      
                    PrintValue(27) = " "
                    FormatString(27) = "~"
                                        
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = "================================================================================================================================================"
                    FormatString(1) = "a150"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
            
            End Select
        End If
                 
'=============================================================================================
'==========================      PRINT EEList REPORT DETAIL     ==============================
'=============================================================================================

        Select Case ReportType
            Case "NumberName"                           ''''''''''   Number/Name
    
                PrintValue(1) = PREmployee.EmployeeNumber
                FormatString(1) = "a9"
                
                PrintValue(2) = " "
                FormatString(2) = "a9"
            
                PrintValue(3) = PREmployee.LFName
                FormatString(3) = "a40"
                                         
                PrintValue(4) = " "
                FormatString(4) = "~"

                FormatPrint
                Ln = Ln + 1
'=============================================================================================
            Case "DetailList"                           ''''''''''   Detail Listing
                                                '  Print 1st Detail Line
                frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
                frmProgress.Show
                
                MaxLines = 47
                PrintValue(1) = PREmployee.EmployeeNumber
                FormatString(1) = "a9"

                PrintValue(2) = " "
                FormatString(2) = "a4"
                
                PrintValue(3) = PRDepartment.DepartmentNumber
                FormatString(3) = "n2"
                
                PrintValue(4) = " - "
                FormatString(4) = "a3"
                
                PrintValue(5) = Mid(PRDepartment.Name, 1, 8)
                FormatString(5) = "a8"
                               
                PrintValue(6) = " "
                FormatString(6) = "a3"
                                
                PrintValue(7) = RTrim(PREmployee.LFName)
                FormatString(7) = "a35"
                
                PrintValue(8) = " "
                FormatString(8) = "a1"
                
                PrintValue(9) = RTrim(PREmployee.Address1)
                FormatString(9) = "a30"
                
                PrintValue(10) = " "
                FormatString(10) = "a3"
                
                PrintValue(11) = RTrim(PREmployee.City)
                FormatString(11) = "a20"
                
                PrintValue(12) = " "
                FormatString(12) = "a2"
                
                PrintValue(13) = PREmployee.State
                FormatString(13) = "a2"
                
                PrintValue(14) = " "
                FormatString(14) = "a7"
                
                PrintValue(15) = PREmployee.ZipCode
                FormatString(15) = "a5"

                 '*** Print SS Number?
                If frmLists.chkSSN Then
                    PrintValue(16) = " "
                    FormatString(16) = "a3"
                    
                    PrintValue(17) = Format(PREmployee.SSN, "000-00-0000")
                    FormatString(17) = "a11"
                    
                    PrintValue(18) = " "
                    FormatString(18) = "~"
                Else
           
                    PrintValue(16) = " "
                    FormatString(16) = "~"
                End If
                
                FormatPrint
                Ln = Ln + 1

                PrintValue(1) = " "         '  Print 2nd Detail Line
                FormatString(1) = "a4"
                 
                If PREmployee.DateLastPaid <> 0 Then
                    PrintValue(2) = Format(PREmployee.DateLastPaid, "mm/dd/yy")
                    FormatString(2) = "a8"
                    PrintValue(3) = " "
                    FormatString(3) = "a4"
                Else
                    PrintValue(2) = " "
                    FormatString(2) = "a8"
                    PrintValue(3) = " "
                    FormatString(3) = "a4"
                End If
      
                If PREmployee.DateHired <> 0 Then
                    PrintValue(4) = Format(PREmployee.DateHired, "mm/dd/yy")
                    FormatString(4) = "a8"
                    PrintValue(5) = " "
                    FormatString(5) = "a5"
                Else
                    PrintValue(4) = " "
                    FormatString(4) = "a8"
                    PrintValue(5) = " "
                    FormatString(5) = "a5"
                End If
    
                If PREmployee.DateLastRaise <> 0 Then
                    PrintValue(6) = Format(PREmployee.DateLastRaise, "mm/dd/yy")
                    FormatString(6) = "a8"
                    PrintValue(7) = " "
                    FormatString(7) = "a4"
                Else
                    PrintValue(6) = " "
                    FormatString(6) = "a8"
                    PrintValue(7) = " "
                    FormatString(7) = "a4"
                End If
                 
                If PREmployee.DateLastReview <> 0 Then
                    PrintValue(8) = Format(PREmployee.DateLastReview, "mm/dd/yy")
                    FormatString(8) = "a8"
                    PrintValue(9) = " "
                    FormatString(9) = "a4"
                Else
                    PrintValue(8) = " "
                    FormatString(8) = "a8"
                    PrintValue(9) = " "
                    FormatString(9) = "a4"
                End If
                                                  
                If PREmployee.DateLastLayoff <> 0 Then
                    PrintValue(10) = Format(PREmployee.DateLastLayoff, "mm/dd/yy")
                    FormatString(10) = "a8"
                    PrintValue(11) = " "
                    FormatString(11) = "a4"
                Else
                    PrintValue(10) = " "
                    FormatString(10) = "a8"
                    PrintValue(11) = " "
                    FormatString(11) = "a4"
                End If
                                                  
                If PREmployee.DateLastRecall <> 0 Then
                    PrintValue(12) = Format(PREmployee.DateLastRecall, "mm/dd/yy")
                    FormatString(12) = "a8"
                    PrintValue(13) = " "
                    FormatString(13) = "a4"
                Else
                    PrintValue(12) = " "
                    FormatString(12) = "a8"
                    PrintValue(13) = " "
                    FormatString(13) = "a4"
                End If
                 
                If PREmployee.DateTerminated <> 0 Then
                    PrintValue(14) = Format(PREmployee.DateTerminated, "mm/dd/yy") & " - " & PREmployee.TermReason
                    FormatString(14) = "a8"
                    PrintValue(15) = " "
                    FormatString(15) = "a4"
                Else
                    PrintValue(14) = " "
                    FormatString(14) = "a8"
                    PrintValue(15) = " "
                    FormatString(15) = "a4"
                End If
                
                If PREmployee.DateOfBirth <> 0 Then
                    PrintValue(16) = Format(PREmployee.DateOfBirth, "mm/dd/yy")
                    FormatString(16) = "a8"
                    PrintValue(17) = " "
                    FormatString(17) = "a4"
                Else
                    PrintValue(16) = " "
                    FormatString(16) = "a8"
                    PrintValue(17) = " "
                    FormatString(17) = "a4"
                End If
                
                If PREmployee.TermReason <> 0 Then
                    PrintValue(18) = PREmployee.TermReason
                    FormatString(18) = "a6"
                    PrintValue(19) = " "
                    FormatString(19) = "a8"
                Else
                    PrintValue(18) = " "
                    FormatString(18) = "a6"
                    PrintValue(19) = " "
                    FormatString(19) = "a8"
                End If
                                 
                If Trim(PREmployee.Sex) <> "" Then
                    PrintValue(20) = PREmployee.Sex
                    FormatString(20) = "a2"
                    PrintValue(21) = " "
                    FormatString(21) = "a3"
                Else
                    PrintValue(20) = " "
                    FormatString(20) = "a2"
                    PrintValue(21) = " "
                    FormatString(21) = "a3"
                End If
                                
                If PREmployee.RaceCode <> 0 Then
                    PrintValue(22) = PREmployee.RaceCode
                    FormatString(22) = "a6"
                    PrintValue(23) = " "
                    FormatString(23) = "a5"
                Else
                    PrintValue(22) = " "
                    FormatString(22) = "a6"
                    PrintValue(23) = " "
                    FormatString(23) = "a5"
                End If
                                
                If Trim(PREmployee.MaritalStatus) <> "" Then
                    PrintValue(24) = PREmployee.MaritalStatus
                    FormatString(24) = "a7"
                    PrintValue(25) = " "
                    FormatString(25) = "a4"
                Else
                    PrintValue(24) = " "
                    FormatString(24) = "a7"
                    PrintValue(25) = " "
                    FormatString(25) = "a4"
                End If
                                
                If PREmployee.EducationLevel <> 0 Then
                    PrintValue(26) = PREmployee.EducationLevel
                    FormatString(26) = "a5"
                    PrintValue(27) = " "
                    FormatString(27) = "a4"
                Else
                    PrintValue(26) = " "
                    FormatString(26) = "a5"
                    PrintValue(27) = " "
                    FormatString(27) = "a4"
                End If
                                
                If PREmployee.ShiftCode <> 0 Then
                    PrintValue(28) = PREmployee.ShiftCode
                    FormatString(28) = "a7"
                    PrintValue(29) = " "
                    FormatString(29) = "a4"
                Else
                    PrintValue(28) = " "
                    FormatString(28) = "a7"
                    PrintValue(29) = " "
                    FormatString(29) = "a4"
                End If
                
                If PREmployee.WkcCat <> 0 Then        ''''''''''''########   WORKCOMPNUM  ######''''''''
                    PrintValue(30) = PREmployee.WkcCat
                    FormatString(30) = "a"
                Else
                    PrintValue(30) = " "
                    FormatString(30) = "a4"
                End If
                
                PrintValue(31) = " "
                FormatString(31) = "~"
                 
                FormatPrint
                Ln = Ln + 1
                                                        
                PrintValue(1) = "Gen Info: "         '  Print 3rd Detail Line
                FormatString(1) = "a11"
                
                If PREmployee.Inactive = 1 Then
                    PrintValue(2) = "Inactive: Y"
                Else
                    PrintValue(2) = "Inactive: N"
                End If
                FormatString(2) = "a11"
                
                PrintValue(3) = " "
                FormatString(3) = "a2"
                
                If PREmployee.Salaried = 1 Then
                    PrintValue(4) = "Salaried: Y"
                    FormatString(4) = "a12"
                    
                    PrintValue(5) = " "
                    FormatString(5) = "a2"
                    
                    PrintValue(6) = "Salary Amt: "
                    FormatString(6) = "a12"
                
                    PrintValue(7) = Format(PREmployee.SalaryAmount, "###,##0.00")
                    FormatString(7) = "d10"

                Else
                    PrintValue(4) = "Salaried: N"
                    FormatString(4) = "a12"
                    
                    PrintValue(5) = " "
                    FormatString(5) = "a2"
                    
                    PrintValue(6) = "Hourly Amt: "
                    FormatString(6) = "a12"
                    
                    PrintValue(7) = Format(PREmployee.HourlyAmount, "##0.00")
                    FormatString(7) = "d6"
                End If
                
                PrintValue(8) = " "
                FormatString(8) = "a2"
                
                PrintValue(9) = "Pays Per Yr: "
                FormatString(9) = "a13"
                  
                PrintValue(10) = PREmployee.PaysPerYear
                FormatString(10) = "n2"
                
                PrintValue(11) = " "
                FormatString(11) = "a2"
                                
                If PREmployee.DefaultCityID > 0 Then
                    If PRCity.GetBySQL("Select * from PRCity where PRCity.CityID = " & PREmployee.DefaultCityID) Then
                        PrintValue(12) = "City: "
                        FormatString(12) = "a6"
                        
                        PrintValue(13) = Trim(PRCity.CityName)
                        FormatString(13) = "a15"

                        PrintValue(14) = " "
                        FormatString(14) = "a2"
                        
                        PrintValue(15) = "Rate: "
                        FormatString(15) = "a6"
                        
                        PrintValue(16) = Format(PRCity.CityRate, "##0.00")
                        FormatString(16) = "d6"
                        
                        PrintValue(17) = " "
                        FormatString(17) = "a2"
                        
                        If PRCity.StateID > 0 Then
                            PrintValue(18) = "State: "
                            FormatString(18) = "a7"
                        
                            If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCity.StateID) Then
                                PrintValue(19) = PRState.StateAbbrev
                                FormatString(19) = "a2"
                            End If
                        End If
                    End If
                Else
                    PrintValue(12) = " "
                    FormatString(12) = "~"
                End If
                
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = "Tax Base: "         '  Print 4th Detail Line
                FormatString(1) = "a11"
                
                If PREmployee.FWTMarried = 1 Then
                    PrintValue(2) = "FWT Married: Y"
                Else
                    PrintValue(2) = "FWT Married: N"
                End If
                FormatString(2) = "a14"
                
                PrintValue(3) = " "
                FormatString(3) = "a2"
            
                If PREmployee.FWTBasis = PREquate.BasisExemptions Then
                    PrintValue(4) = "FWT Exemps: "
                    FormatString(4) = "a12"
                    PrintValue(5) = Format(PREmployee.FWTAmount, "#0")
                    FormatString(5) = "n2"
                    PrintValue(6) = " "
                    FormatString(6) = "a2"
                ElseIf PREmployee.FWTBasis = PREquate.BasisPercent Then
                    PrintValue(4) = "FWT: "
                    FormatString(4) = "a5"
                    PrintValue(5) = Format(PREmployee.FWTAmount, "###.#0")
                    FormatString(5) = "d6"
                    PrintValue(6) = "%  "
                    FormatString(6) = "a3"
                End If
                
                If PREmployee.FWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(7) = "FWT Extra: "
                    FormatString(7) = "a11"
                    
                    PrintValue(8) = Format(PREmployee.FWTAmount, "##0.00")
                    FormatString(8) = "d6"
                    
                    PrintValue(9) = "% "
                    FormatString(9) = "a3"
                ElseIf PREmployee.FWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(7) = "FWT Extra: "
                    FormatString(7) = "a11"
                    
                    PrintValue(8) = "$ " & Format(PREmployee.FWTAmount, "#,##0.00")
                    FormatString(8) = "d8"
                
                    PrintValue(9) = " "
                    FormatString(9) = "a2"
                End If
                
                If PREmployee.SWTMarried = 1 Then
                    PrintValue(10) = "SWT Married: Y"
                Else
                    PrintValue(10) = "SWT Married: N"
                End If
                FormatString(10) = "a14"
                
                PrintValue(11) = " "
                FormatString(11) = "a2"
                
                If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                    PrintValue(12) = "SWT Exemps: "
                    FormatString(12) = "a12"
                    PrintValue(13) = Format(PREmployee.SWTAmount, "#0")
                    FormatString(13) = "n2"
                    PrintValue(14) = " "
                    FormatString(14) = "a2"
                ElseIf PREmployee.SWTBasis = PREquate.BasisPercent Then
                    PrintValue(12) = "SWT: "
                    FormatString(12) = "a5"
                    PrintValue(13) = Format(PREmployee.SWTAmount, "###.#0")
                    FormatString(13) = "d6"
                    PrintValue(14) = "%  "
                    FormatString(14) = "a3"
                End If
                
                If PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(15) = "SWT Extra: "
                    FormatString(15) = "a11"
                    
                    PrintValue(16) = Format(PREmployee.SWTAmount, "##0.00")
                    FormatString(16) = "d6"
                    
                    PrintValue(17) = "%"
                    FormatString(17) = "a1"
                    
                    PrintValue(18) = " "
                    FormatString(18) = "~"
                ElseIf PREmployee.SWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(15) = "SWT Extra: "
                    FormatString(15) = "a11"
                    
                    PrintValue(16) = "$ " & Format(PREmployee.SWTAmount, "#,##0.00")
                    FormatString(16) = "d8"
                    
                    PrintValue(17) = " "
                    FormatString(17) = "~"
                End If
                                
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = "Tax Flags: "         '  Print 5th Detail Line
                FormatString(1) = "a11"
                
                If PREmployee.NoSSTax = 1 Then
                    PrintValue(2) = "No SS Tax? " & "Y"
                Else
                    PrintValue(2) = "No SS Tax? " & "N"
                End If
                FormatString(2) = "a12"
                
                PrintValue(3) = " "
                FormatString(3) = "a4"
            
                If PREmployee.NoMedTax = 1 Then
                    PrintValue(4) = "No Med Tax? " & "Y"
                Else
                    PrintValue(4) = "No Med Tax? " & "N"
                End If
                FormatString(4) = "a13"
                
                PrintValue(5) = " "
                FormatString(5) = "a4"
                
                If PREmployee.NoFedTax = 1 Then
                    PrintValue(6) = "No Fed Tax? " & "Y"
                Else
                    PrintValue(6) = "No Fed Tax? " & "N"
                End If
                FormatString(6) = "a13"
                
                PrintValue(7) = " "
                FormatString(7) = "a4"
                
                If PREmployee.NoStateTax = 1 Then
                    PrintValue(8) = "No State Tax? " & "Y"
                Else
                    PrintValue(8) = "No State Tax? " & "N"
                End If
                FormatString(8) = "a15"
                
                PrintValue(9) = " "
                FormatString(9) = "a4"
                
                If PREmployee.NoCityTax = 1 Then
                    PrintValue(10) = "No City Tax? " & "Y"
                Else
                    PrintValue(10) = "No City Tax? " & "N"
                End If
                FormatString(10) = "a14"
                
                PrintValue(11) = " "
                FormatString(11) = "a4"
                
                If PREmployee.NoFedUnemp = 1 Then
                    PrintValue(12) = "No Fed Unemp? " & "Y"
                Else
                    PrintValue(12) = "No Fed Unemp? " & "N"
                End If
                FormatString(12) = "a15"
                                
                PrintValue(13) = " "
                FormatString(13) = "a4"
                
                If PREmployee.NoStateUnemp = 1 Then
                    PrintValue(14) = "No State Unemp? " & "Y"
                Else
                    PrintValue(14) = "No State Unemp? " & "N"
                End If
                FormatString(14) = "a17"
    
                PrintValue(15) = " "
                FormatString(15) = "~"
                
                FormatPrint
                Ln = Ln + 1
                 
                PrintValue(1) = "----------------------------------------------------------------------" _
                              & "----------------------------------------------------------------------" _
                              & "---------------"
                FormatString(1) = "a155"
                 
                PrintValue(2) = " "
                FormatString(2) = "~"
                 
                FormatPrint
                Ln = Ln + 1
             
            Case "EmployeeRateList"                         ''''''''''   Employee Rate Listing
                frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
                frmProgress.Show
                
                PrintValue(1) = PREmployee.EmployeeNumber
                FormatString(1) = "a9"
                
                PrintValue(2) = " "
                FormatString(2) = "a0"
                
                PrintValue(3) = PREmployee.LFName
                FormatString(3) = "a40"
                
                PrintValue(4) = " "
                FormatString(4) = "a3"
                
                PrintValue(5) = PRDepartment.DepartmentNumber
                FormatString(5) = "n4"
                
                PrintValue(6) = " - "
                FormatString(6) = "a3"
                                                        
                PrintValue(7) = Mid(PRDepartment.Name, 1, 8)
                FormatString(7) = "a8"
                               
                PrintValue(8) = " "
                FormatString(8) = "a0"
                
                If PREmployee.Salaried = 1 Then
                    PrintValue(9) = Format(PREmployee.SalaryAmount, "#,##0.00")
                Else
                    PrintValue(9) = Format(PREmployee.HourlyAmount, "#,##0.00")
                End If
                FormatString(9) = "d8"
                
                PrintValue(10) = " "
                FormatString(10) = "a3"
                
                If PREmployee.Salaried = 1 Then
                    PrintValue(11) = "SALARY"
                    FormatString(11) = "a6"
                Else
                    PrintValue(11) = "HOURLY"
                    FormatString(11) = "a6"
                End If
                
                PrintValue(12) = " "
                FormatString(12) = "~"
                
                FormatPrint
                Ln = Ln + 1
'=============================================================================================
            Case "SSNFormat"                            ''''''''''   SSN Format
                frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
                frmProgress.Show
                
                PrintValue(1) = PREmployee.EmployeeNumber
                FormatString(1) = "a9"
                
                PrintValue(2) = " "
                FormatString(2) = "a6"
                
                PrintValue(3) = PREmployee.LFName
                FormatString(3) = "a40"
                
                PrintValue(4) = " "
                FormatString(4) = "a3"
                
                PrintValue(5) = Format(PREmployee.SSN, "000-00-0000")
                FormatString(5) = "a12"
                 
                PrintValue(6) = " "
                FormatString(6) = "a3"
                
                If PREmployee.DateOfBirth <> 0 Then
                    PrintValue(7) = PREmployee.DateOfBirth
                Else
                    PrintValue(7) = " "
                End If
                FormatString(7) = "a10"
                
                PrintValue(8) = " "
                FormatString(8) = "a3"
                
                PrintValue(9) = PREmployee.Sex
                FormatString(9) = "a6"
                
                PrintValue(10) = " "
                FormatString(10) = "~"
                
                FormatPrint
                Ln = Ln + 1
'=============================================================================================
            Case "RateTaxList"                          ''''''''''   Rate Tax Listing
             
                frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
                frmProgress.Show
                
                PrintValue(1) = PREmployee.EmployeeNumber
                FormatString(1) = "n5"
                
                PrintValue(2) = " "
                FormatString(2) = "a2"
                
                PrintValue(3) = PREmployee.LFName
                FormatString(3) = "a18"
                
                PrintValue(4) = " "
                FormatString(4) = "a3"
                
                PrintValue(5) = PRDepartment.DepartmentNumber & "-" & Trim(Mid(PRDepartment.Name, 1, 8))
                FormatString(5) = "a10"
                
                PrintValue(6) = " "
                FormatString(6) = "a0"
                
                If PREmployee.Salaried = 1 Then
                    PrintValue(7) = Format(PREmployee.SalaryAmount, "#,##0.00")
                    FormatString(7) = "d8"
                    PrintValue(8) = "SALARY"
                    FormatString(8) = "a7"
                Else
                    PrintValue(7) = Format(PREmployee.HourlyAmount, "#,##0.00")
                    FormatString(7) = "d8"
                    PrintValue(8) = "HOURLY"
                    FormatString(8) = "a7"
                End If
                
                PrintValue(9) = " "
                FormatString(9) = "a3"
                
                If PREmployee.NoSSTax = 1 Then
                   PrintValue(10) = "Y"
                Else
                   PrintValue(10) = "N"
                End If
                FormatString(10) = "a4"
                               
                If PREmployee.NoMedTax = 1 Then
                   PrintValue(11) = "Y"
                Else
                   PrintValue(11) = "N"
                End If
                FormatString(11) = "a4"
                
                If PREmployee.NoFedTax = 1 Then
                   PrintValue(12) = "Y"
                Else
                   PrintValue(12) = "N"
                End If
                FormatString(12) = "a4"
                               
                If PREmployee.NoStateTax = 1 Then
                   PrintValue(13) = "Y"
                Else
                   PrintValue(13) = "N"
                End If
                FormatString(13) = "a4"
                               
                PrintValue(14) = " "
                FormatString(14) = "a0"
                
                If PREmployee.NoCityTax = 1 Then
                   PrintValue(15) = "Y"
                Else
                   PrintValue(15) = "N"
                End If
                FormatString(15) = "a4"
                               
                PrintValue(16) = " "
                FormatString(16) = "a1"
                
                If PREmployee.NoFedUnemp = 1 Then
                   PrintValue(17) = "Y"
                Else
                   PrintValue(17) = "N"
                End If
                FormatString(17) = "a4"
                               
                PrintValue(18) = " "
                FormatString(18) = "a2"
                
                If PREmployee.NoStateUnemp = 1 Then
                   PrintValue(19) = "Y"
                Else
                   PrintValue(19) = "N"
                End If
                FormatString(19) = "a5"
                               
                If PREmployee.FWTMarried = 1 Then
                   PrintValue(20) = "Y"
                Else
                   PrintValue(20) = "N"
                End If
                FormatString(20) = "a3"
                                                                  
                If PREmployee.FWTBasis = PREquate.BasisExemptions Then      '  EXEMPTIONS
                    PrintValue(21) = Format(PREmployee.FWTAmount, "####0")
                    FormatString(21) = "r8"
                    
                    PrintValue(22) = ""
                    FormatString(22) = "a0"
                ElseIf PREmployee.FWTBasis = PREquate.BasisPercent Then
                    PrintValue(21) = Format(PREmployee.FWTAmount, "##.00")
                    FormatString(21) = "r7"
                    
                    PrintValue(22) = "%"
                    FormatString(22) = "a1"
                End If
                
                If PREmployee.FWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(23) = Format(PREmployee.FWTExtraAmount, "##0.00")
                    FormatString(23) = "r8"
                    
                    PrintValue(24) = ""
                    FormatString(24) = "a0"
                ElseIf PREmployee.FWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(23) = Format(PREmployee.FWTExtraAmount, "##.00")
                    FormatString(23) = "r7"
                    
                    PrintValue(24) = "%"
                    FormatString(24) = "a1"
                End If

                PrintValue(25) = ""
                FormatString(25) = "a5"
                
                If PREmployee.SWTMarried = 1 Then
                   PrintValue(26) = "Y"
                Else
                   PrintValue(26) = "N"
                End If
                FormatString(26) = "a1"
                             
                PrintValue(27) = " "
                FormatString(27) = "a3"
                
                If PREmployee.SWTBasis = PREquate.BasisExemptions Then      '  EXEMPTIONS
                    PrintValue(28) = Format(PREmployee.SWTAmount, "##0")
                    FormatString(28) = "r8"
                    
                    PrintValue(29) = ""
                    FormatString(29) = "a0"
                ElseIf PREmployee.SWTBasis = PREquate.BasisPercent Then
                    PrintValue(28) = Format(PREmployee.SWTAmount, "##.00")
                    FormatString(28) = "r7"
                    
                    PrintValue(29) = "%"
                    FormatString(29) = "a1"
                End If
                
                If PREmployee.SWTExtraBasis = PREquate.BasisAmount Then
                    PrintValue(30) = Format(PREmployee.SWTExtraAmount, "##0.00")
                    FormatString(30) = "r8"
                    
                    PrintValue(31) = ""
                    FormatString(31) = "a0"
                ElseIf PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
                    PrintValue(30) = Format(PREmployee.SWTExtraAmount, "##.00")
                    FormatString(30) = "r7"
                    
                    PrintValue(31) = "%"
                    FormatString(31) = "a1"
                End If
                                                                  
                PrintValue(32) = " "
                FormatString(32) = "~"
                
                FormatPrint
                Ln = Ln + 1
                
'=======================================   TIME CARD LABELS  ======================================

            Case "TimeCardLabels"
                LabelColumns = 1
                If NoLabels = 0 Then
                    Ln = Ln + 2
                    LabelString(1, 1) = "EMPLOYEE # " & PREmployee.EmployeeNumber
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    LabelString(1, 1) = RTrim(PREmployee.FLName)
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    LabelString(1, 1) = "PERIOD ENDING DATE: " & frmLists.PEDate
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                                              
                    LabelString(1, 1) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(Mid(PRDepartment.Name, 1, 8))
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                   
                    FormatPrint
                    Ln = Ln + 1
                    LabelRows = LabelRows + 1

                    LabelCount = LabelCount + 1
                    If LabelCount = 10 Then
                        LabelCount = 0
                        FormFeed
                    End If
                   
                ElseIf NoLabels = 1 Then
                    LabelColumns = 2
                    ColumnCount = ColumnCount + 1
                    Label2String(1, ColumnCount) = "EMPLOYEE # " & PREmployee.EmployeeNumber
                    Label2String(2, ColumnCount) = RTrim(PREmployee.FLName)
                    Label2String(3, ColumnCount) = "P/E DATE: " & NewPEDate
                    Label2String(4, ColumnCount) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name)

                    If ColumnCount = LabelColumns Then
                        Ln = Ln + 2
                         
                        For LRow = 1 To 4
                            ColumnCount = 0
                            PrintValue(1) = Label2String(LRow, 1)
                            FormatString(1) = "a35"
                            
                            PrintValue(2) = Label2String(LRow, 2)
                            FormatString(2) = "a35"
                            
                            PrintValue(3) = Label2String(LRow, 3)
                            FormatString(3) = "a35"
                            
                            PrintValue(4) = Label2String(LRow, 4)
                            FormatString(4) = "a35"
                            
                            PrintValue(5) = " "
                            FormatString(5) = "~"
                            
                            FormatPrint
                            Ln = Ln + 1
                            LabelRows = LabelRows + 1
                        Next LRow
                    End If
                    LabelCount = LabelCount + 1
                    If LabelCount = 20 Then
                        LabelCount = 0
                        FormFeed
                    End If
                ElseIf NoLabels = 2 Then
                    LabelColumns = 3
                    ColumnCount = ColumnCount + 1
                    Label2String(1, ColumnCount) = "EMPLOYEE # " & PREmployee.EmployeeNumber
                    Label2String(2, ColumnCount) = RTrim(PREmployee.FLName)
                    Label2String(3, ColumnCount) = "P/E DATE: " & NewPEDate
                    Label2String(4, ColumnCount) = "DEPT: " & PRDepartment.DepartmentNumber & " - " & RTrim(PRDepartment.Name)

                    If ColumnCount = LabelColumns Then
                        Ln = Ln + 2
                        
                        For LRow = 1 To 4
                            ColumnCount = 0
                            PrintValue(1) = Label2String(LRow, 1)
                            FormatString(1) = "a35"
                            
                            PrintValue(2) = Label2String(LRow, 2)
                            FormatString(2) = "a35"
                            
                            PrintValue(3) = Label2String(LRow, 3)
                            FormatString(3) = "a35"
                            
                            PrintValue(4) = Label2String(LRow, 4)
                            FormatString(4) = "a35"
                            
                            PrintValue(5) = " "
                            FormatString(5) = "~"
                            
                            FormatPrint
                            Ln = Ln + 1
                            LabelRows = LabelRows + 1
                        Next LRow
                    End If
                    LabelCount = LabelCount + 1
                    If LabelCount = 30 Then
                        LabelCount = 0
                        FormFeed
                    End If
                End If
                    
'=======================================   MAILING LABELS  ======================================
                    
            Case "MailingLabels"

                LabelColumns = 1
                If NoLabels = 0 Then
                    Ln = Ln + 2
                    LabelString(1, 1) = PREmployee.FLName
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    LabelString(1, 1) = PREmployee.Address1
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    LabelString(1, 1) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 2
                    LabelRows = LabelRows + 1

                    LabelCount = LabelCount + 1
                    If LabelCount = 10 Then
                        LabelCount = 0
                        FormFeed
                    End If
                    
                ElseIf NoLabels = 1 Then
                    LabelColumns = 2
                    ColumnCount = ColumnCount + 1
                    LabelString(1, ColumnCount) = PREmployee.FLName
                    LabelString(2, ColumnCount) = PREmployee.Address1
                    LabelString(3, ColumnCount) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                    
                    If ColumnCount = LabelColumns Then
                        Ln = Ln + 2
                         
                        For LRow = 1 To 4
                            ColumnCount = 0
                            PrintValue(1) = LabelString(LRow, 1)
                            FormatString(1) = "a35"
                            
                            PrintValue(2) = LabelString(LRow, 2)
                            FormatString(2) = "a35"
                            
                            PrintValue(3) = LabelString(LRow, 3)
                            FormatString(3) = "a35"
                            
                            PrintValue(4) = " "
                            FormatString(4) = "~"
                            
                            FormatPrint
                            Ln = Ln + 1
                            LabelRows = LabelRows + 1
                        Next LRow
                    End If
                    LabelCount = LabelCount + 1
                    If LabelCount = 20 Then
                        LabelCount = 0
                        FormFeed
                    End If
                   
                ElseIf NoLabels = 2 Then
                    LabelColumns = 3
                    ColumnCount = ColumnCount + 1
                    LabelString(1, ColumnCount) = PREmployee.FLName
                    LabelString(2, ColumnCount) = PREmployee.Address1
                    LabelString(3, ColumnCount) = RTrim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
                    
                    If ColumnCount = LabelColumns Then
                        Ln = Ln + 2
                         
                        For LRow = 1 To 4
                            ColumnCount = 0
                            PrintValue(1) = LabelString(LRow, 1)
                            FormatString(1) = "a35"
                            
                            PrintValue(2) = LabelString(LRow, 2)
                            FormatString(2) = "a35"
                            
                            PrintValue(3) = LabelString(LRow, 3)
                            FormatString(3) = "a35"
                            
                            PrintValue(4) = " "
                            FormatString(4) = "~"
                            
                            FormatPrint
                            Ln = Ln + 1
                            LabelRows = LabelRows + 1
                        Next LRow
                    End If
                    LabelCount = LabelCount + 1
                    If LabelCount = 30 Then
                        LabelCount = 0
                        FormFeed
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
            YUnits = 245
            If NoLabels = 2 Then
                Ln = Ln + 3
                If ColumnCount = 1 Then
                
                PrintValue(1) = LabelString(1, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = LabelString(2, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = LabelString(3, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
            ElseIf ColumnCount = 2 Then
                PrintValue(1) = LabelString(1, 1)     '(Row, Col)
                FormatString(1) = "a35"
                
                PrintValue(2) = LabelString(1, 2)
                FormatString(2) = "a35"
                
                PrintValue(3) = " "
                FormatString(3) = "~"
                
                FormatPrint
                Ln = Ln + 1
                PrintValue(1) = LabelString(2, 1)
                FormatString(1) = "a35"
                
                PrintValue(2) = LabelString(2, 2)
                FormatString(2) = "a35"
                
                PrintValue(3) = " "
                FormatString(3) = "~"
                
                FormatPrint
                Ln = Ln + 1
                PrintValue(1) = LabelString(3, 1)
                FormatString(1) = "a35"
                
                PrintValue(2) = LabelString(3, 2)
                FormatString(2) = "a35"
                
                PrintValue(3) = " "
                FormatString(3) = "~"
                
                FormatPrint
                
            End If
          
            ElseIf NoLabels = 1 Then
                If ColumnCount > 0 Then
                    Ln = Ln + 2
                
                    PrintValue(1) = LabelString(1, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = LabelString(2, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                       
                    PrintValue(1) = LabelString(3, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    
                    FormatPrint
                End If
            End If
        Case "TimeCardLabels"
            If NoLabels = 2 Then
                Ln = Ln + 3
                If ColumnCount = 1 Then
                
                    PrintValue(1) = Label2String(1, 1)
                    FormatString(1) = "a35"
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    
                    PrintValue(1) = Label2String(2, 1)
                    FormatString(1) = "a35"
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                       
                    PrintValue(1) = Label2String(3, 1)
                    FormatString(1) = "a35"
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                                         
                    PrintValue(1) = Label2String(4, 1)
                    FormatString(1) = "a35"
                    PrintValue(2) = " "
                    FormatString(2) = "~"
                    FormatPrint
                ElseIf ColumnCount = 2 Then
                    PrintValue(1) = Label2String(1, 1)     '(Row, Col)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = Label2String(1, 2)
                    FormatString(2) = "a35"
                    
                    PrintValue(3) = " "
                    FormatString(3) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    PrintValue(1) = Label2String(2, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = Label2String(2, 2)
                    FormatString(2) = "a35"
                    
                    PrintValue(3) = " "
                    FormatString(3) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    PrintValue(1) = Label2String(3, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = Label2String(3, 2)
                    FormatString(2) = "a35"
                    
                    PrintValue(3) = " "
                    FormatString(3) = "~"
                    
                    FormatPrint
                    Ln = Ln + 1
                    PrintValue(1) = Label2String(4, 1)
                    FormatString(1) = "a35"
                    
                    PrintValue(2) = Label2String(4, 2)
                    FormatString(2) = "a35"
                    
                    PrintValue(3) = " "
                    FormatString(3) = "~"
                    
                    FormatPrint
          
                End If
          
            ElseIf NoLabels = 1 Then
                If ColumnCount > 0 Then
                Ln = Ln + 2
                
                PrintValue(1) = Label2String(1, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = Label2String(2, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = Label2String(3, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = Label2String(4, 1)
                FormatString(1) = "a35"
                PrintValue(2) = " "
                FormatString(2) = "~"
                FormatPrint
            End If
        End If
    End Select
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

'=======================================   PAGE HEADER  ======================================

Private Sub PageHeader(ByVal ReportName As String, _
                       ByVal Msg1 As String, _
                       ByVal Msg2 As String, _
                       ByVal msg3 As String)
                       
    Ln = 0
    Pg = Pg + 1
   
    ' 29 characters for fixed left and right portion of first header line
    '    1             8       1   8                    10         1
    ' first line - system date & time / company name / page #
    x = PRCompany.Name
    Y = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
    z = "Page: " & Format(Pg, "####")
   
    If Len(x) > Columns - 17 Then
       x = Mid(Trim(PRCompany.Name), 1, Columns - 27)
    End If
           
    If LandSW = 1 Then
        PosPrint 200, 220, Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
        PosPrint 13700, 220, "Page: " & Format(Pg, "###0")
    Else
        i = ((Columns - Len(x)) / 2) - 19
        w = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss") & _
            Space(i) & x
        i = Columns - Len(w) - 10
        w = w & Space(i) & "Page: " & Format(Pg, "###0")

    End If
    
    PrtCenter Ln, w
    Ln = Ln + 1
   
    If ReportName <> "" Then
        PrtCenter 0, ReportName
        Ln = Ln + 1
    End If
           
    If QtrEnding <> "" Then
       PrtCenter Ln, QtrEnding
       Ln = Ln + 1
    End If
   
    If Msg1 <> "" Then
       PrtCenter Ln, Msg1
       Ln = Ln + 1
    End If
   
    If Msg2 <> "" Then
       PrtCenter Ln, Msg2
       Ln = Ln + 1
    End If

    If msg3 <> "" Then
       PrtCenter Ln, msg3
       Ln = Ln + 1
    End If

    Ln = Ln + 1

End Sub

'=======================================   QTRLY REPORTS  ======================================

Public Sub QtrRpts(ByVal ReportList As String)

    SetEquates
    trs.CursorLocation = adUseClient
    trs.Fields.Append "DeptID", adDouble:           trs.Fields.Append "WageGross", adCurrency
    trs.Fields.Append "WageSS", adCurrency:         trs.Fields.Append "WageMed", adCurrency
    trs.Fields.Append "TaxFed", adCurrency:         trs.Fields.Append "TaxSS", adCurrency
    trs.Fields.Append "TaxMed", adCurrency:         trs.Fields.Append "WageState", adCurrency
    trs.Fields.Append "WageCity", adCurrency:       trs.Fields.Append "TaxState", adCurrency
    trs.Fields.Append "TaxCity", adCurrency:        trs.Fields.Append "QTDWageGross", adCurrency
    trs.Fields.Append "YTDWageGross", adCurrency:   trs.Fields.Append "WageFed", adCurrency
    trs.Fields.Append "WageFIC", adCurrency:        trs.Fields.Append "TipsFIC", adCurrency
    trs.Fields.Append "TipsMed", adCurrency:        trs.Fields.Append "FicaTotal", adCurrency
    trs.Open , , adOpenDynamic, adLockOptimistic
    frmPRQtrlyRpts.Hide
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    
    ' page set up based on the report listing selection

    Select Case ReportList
    
    Case "QtrlyFICAFWT"
            PrtInit ("Port")
            ReportTitle = "Payroll Quar1terly FICA And FWT Report"
            SetFont 8, Equate.Portrait
    Case "QtrlyStateCity"
            PrtInit ("Port")
            ReportTitle = "Payroll Quarterly State and City Report"
            SetFont 8, Equate.Portrait
    Case "QtrlyFedUnemp"
            PrtInit ("Port")
            ReportTitle = "Payroll Quarterly Federal Unemployment Report"
            SetFont 8, Equate.Portrait
    Case "QtrlyTipsTaxes"
            PrtInit ("Port")
            ReportTitle = "Payroll Quarterly Tips and Taxes Report"
            SetFont 8, Equate.Portrait
    End Select
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
                             
    ' set up SQL statement based upon order requested

    SQLString = "SELECT * FROM PREmployee ORDER BY SSN"
    
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Found !!!", vbCritical, "Payroll Quarterly Reports"
        Exit Sub
    End If
    Ln = 0
    
    Do
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            MsgBox "Department Not Found!!!", vbCritical, "Qtrly FICA FWT"
            End
        End If
        
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            SetFont 8, Equate.Portrait

            ' data header
            Ln = Ln + 2
            PrintCompanyHeader (ReportList)
        End If
        
 '       If PREmployee.SSN = 0 Then GoTo cycle2
                       
'=============================================================================================
'==========================      QTRLY REPORTS -  DETAIL     =================================
'=============================================================================================
         
        frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.FLName)
        frmProgress.Show
        
        Select Case ReportList
            Case "QtrlyFICAFWT"             '''''' QUARTERLY FICA FWT ''''''
                
                If PREmployee.SSN <> 0 Then
                     PrintValue(1) = Format(PREmployee.SSN, "000-00-0000"): FormatString(1) = "a11"
                Else
                     PrintValue(1) = "":                                    FormatString(1) = "a11"
                End If
                PrintValue(2) = " ":                                        FormatString(2) = "a2"
                PrintValue(3) = PREmployee.FLName:                          FormatString(3) = "a40"
                PrintValue(4) = " ":                                        FormatString(4) = "~"
                FormatPrint
                Ln = Ln + 1
                
            WageGross = 10
'             WageGross = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                          qyear, qyear, EndMonth, 0, 0, 0)
                FindStr = "DeptID=" & CStr(PRDepartment.DepartmentID)
                trs.Find FindStr, 0, adSearchForward, 1
                
                If trs.EOF Then
                    trs.AddNew Array("DeptID", "WageGross", "WageSS", "WageMed", "TaxFed", "TaxSS", "TaxMed"), _
                    Array(PRDepartment.DepartmentID, 0, 0, 0, 0, 0, 0)
                    trs.UpdateBatch
                End If
                PrintValue(1) = " ":                                        FormatString(1) = "a20"
                DWageGross = trs!WageGross:                                 DWageGross = DWageGross + WageGross  ' Department Total
                trs.Fields("DeptID") = trs!DeptID:                          trs.Fields("WageGross") = DWageGross
                PrintValue(2) = Format(WageGross, "##,###,##0.00"):         FormatString(2) = "d13"
                TotWageGross = TotWageGross + WageGross
                            
            WageSS = 20
'             WageSS = GetPRAmount(PREmployee.EmployeeID, PREquate.WageSS, _
                       qyear, qyear, EndMonth, 0, 0, 0)
                DWageSS = trs!WageSS:                                       DWageSS = DWageSS + WageSS
                trs.Fields("WageSS") = DWageSS:                             TotWageSS = TotWageSS + WageSS
                PrintValue(3) = Format(WageSS, "##,###,##0.00"):            FormatString(3) = "d13"
                
            WageMed = 30
'             WageMed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageMed, _
                       qyear, qyear, EndMonth, 0, 0, 0)
                DWageMed = trs!WageMed:                                     DWageMed = DWageMed + WageMed
                trs.Fields("WageMed") = DWageMed:                           TotWageMed = TotWageMed + WageMed
                PrintValue(4) = Format(WageMed, "##,###,##0.00"):           FormatString(4) = "d13"
                                
            TaxFed = 40
'             TaxFed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxFed, _
                        qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxFed = trs!TaxFed:                                       DTaxFed = DTaxFed + TaxFed
                trs.Fields("TaxFed") = DTaxFed:                             TotTaxFed = TotTaxFed + TaxFed
                PrintValue(5) = Format(TaxFed, "#,###,##0.00"):             FormatString(5) = "d12"
                
            TaxSS = 50
'             TaxSS = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxSS, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxSS = trs!TaxSS:                                         DTaxSS = DTaxSS + TaxSS
                trs.Fields("TaxSS") = DTaxSS:                               TotTaxSS = TotTaxSS + TaxSS
                PrintValue(6) = Format(TaxSS, "#,###,##0.00"):              FormatString(6) = "d12"
                
            TaxMed = 60
'             TaxMed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxMed, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxMed = trs!TaxMed:                                       DTaxMed = DTaxMed + TaxMed
                trs.Fields("TaxMed") = DTaxMed:                             TotTaxMed = TotTaxMed + TaxMed
                PrintValue(7) = Format(TaxMed, "#,###,##0.00"):             FormatString(7) = "d12"
                TotFica = TaxSS + TaxMed:                                   FinalFica = FinalFica + TotFica
                PrintValue(8) = Format(TotFica, "###,###,##0.00"):          FormatString(8) = "d14"
                PrintValue(9) = " ":                                        FormatString(9) = "~"
                trs.Update
                FormatPrint
                Ln = Ln + 1
                            
            Case "QtrlyStateCity"           '''''' QUARTERLY STATE CITY ''''''

                If PREmployee.SSN <> 0 Then
                     PrintValue(1) = Format(PREmployee.SSN, "000-00-0000"): FormatString(1) = "a11"
                Else
                     PrintValue(1) = "":                                    FormatString(1) = "a11"
                End If
                                
                PrintValue(2) = " ":                                        FormatString(2) = "a2"
                PrintValue(3) = PREmployee.FLName:                          FormatString(3) = "a30"
                PrintValue(4) = " ":                                        FormatString(4) = "a1"
                
        WageState = 10
'        WageState = GetPRAmount(PREmployee.EmployeeID, PREquate.WageState, _
                          qyear, qyear, EndMonth, 0, 0, 0)
                FindStr = "DeptID =" & PRDepartment.DepartmentID
                trs.Find FindStr, 0, adSearchForward, 1
                PrintValue(5) = Format(WageState, "##,###,##0.00"):         FormatString(5) = "d13"
                If trs.EOF Then
                    trs.AddNew Array("DeptID", "WageState", "WageCity", "TaxState", "TaxCity"), _
                    Array(PRDepartment.DepartmentID, 0, 0, 0, 0)
                    trs.UpdateBatch
                End If
                trs.Fields("DeptID") = trs!DeptID:                          dWageState = trs!WageState
                dWageState = dWageState + WageState:                        trs.Fields("WageState") = dWageState
                TotWageState = TotWageState + WageState
                
        WageCity = 20
 '       WageCity = GetPRAmount(PREmployee.EmployeeID, PREquate.WageCity, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DWageCity = trs!WageCity:                                   DWageCity = DWageCity + WageCity:
                trs.Fields("WageCity") = DWageCity:                         TotWageCity = TotWageCity + WageCity
                PrintValue(6) = Format(WageCity, "##,###,##0.00"):          FormatString(6) = "d13"
                        
        TaxState = 30
'        TaxState = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxState, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxState = trs!TaxState:                                   DTaxState = DTaxState + TaxState
                trs.Fields("TaxState") = DTaxState:                         TotTaxState = TotTaxState + TaxState
                PrintValue(7) = Format(TaxState, "##,###,##0.00"):          FormatString(7) = "d13"
                
        TaxCity = 40
'        TaxCity = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxCity, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxCity = trs!TaxCity:                                     DTaxCity = DTaxCity + TaxCity
                trs.Fields("TaxCity") = DTaxCity:                           TotTaxCity = TotTaxCity + TaxCity
                PrintValue(8) = Format(TaxCity, "##,###,##0.00"):           FormatString(8) = "d13"
                PrintValue(9) = " ":                                        FormatString(9) = "~"
                trs.Update
                FormatPrint
                Ln = Ln + 1
                
            Case "QtrlyFedUnemp"                '''''' QUARTERLY FED UNEMPLOYMENT ''''''

                If PREmployee.SSN <> 0 Then
                     PrintValue(1) = Format(PREmployee.SSN, "000-00-0000"): FormatString(1) = "a13"
                Else
                     PrintValue(1) = "":                                    FormatString(1) = "a13"
                End If
                
                PrintValue(2) = PREmployee.FLName:                          FormatString(2) = "a30"
                PrintValue(3) = " ":                                        FormatString(3) = "a2"
               
                FindStr = "DeptID=" & PRDepartment.DepartmentID
                trs.Find FindStr, 0, adSearchForward, 1
                If trs.EOF Then
                    trs.AddNew Array("DeptID", "QTDWageGross", "YTDWageGross", "WageFed", "WageState"), _
                    Array(PRDepartment.DepartmentID, 0, 0, 0, 0)
                    trs.UpdateBatch
                End If
                
                trs.Fields("DeptID") = trs!DeptID
                
        QTDWageGross = 10
'        QTDWageGross = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                            qyear, qyear, StartMonth, EndMonth, 0, 0)
                DQTDWageGross = trs!QTDWageGross:                           DQTDWageGross = DQTDWageGross + QTDWageGross
                trs.Fields("QTDWageGross") = DQTDWageGross:                 TotQTDWageGross = TotQTDWageGross + QTDWageGross
                PrintValue(4) = Format(QTDWageGross, "##,###,##0.00"):      FormatString(4) = "d13"
                PrintValue(5) = " ":                                        FormatString(5) = "a3"

        YTDWageGross = 20
'       YTDWageGross = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                          qyear, qyear, 1, EndMonth, 0, 0)
                DYTDWageGross = trs!YTDWageGross:                           DYTDWageGross = DYTDWageGross + YTDWageGross
                trs.Fields("YTDWageGross") = DYTDWageGross:                 TotYTDWageGross = TotYTDWageGross + YTDWageGross
                PrintValue(6) = Format(YTDWageGross, "##,###,##0.00"):      FormatString(6) = "d13"
                PrintValue(7) = " ":                                        FormatString(7) = "a3"
              
        WageFed = 30
'        WageFed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageFed, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DWageFed = trs!WageFed:                                     DWageFed = DWageFed + WageFed
                trs.Fields("WageFed") = DWageFed:                           TotWageFed = TotWageFed + WageFed
                PrintValue(8) = Format(WageFed, "##,###,##0.00"):           FormatString(8) = "d13"
                PrintValue(9) = " ":                                        FormatString(9) = "a3"
              
        WageState = 40
'        WageState = GetPRAmount(PREmployee.EmployeeID, PREquate.WageState, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(10) = Format(WageState, "##,###,##0.00")
                FormatString(10) = "d13"
                
                dWageState = trs!WageState:                                 dWageState = dWageState + WageState
                trs.Fields("WageState") = dWageState:                       TotWageState = TotWageState + WageState
                PrintValue(11) = " ":                                       FormatString(11) = "~"
                trs.Update
                FormatPrint
                Ln = Ln + 1
                
            Case "QtrlyTipsTaxes"               '''''' QUARTERLY TIPS AND TAXES ''''''
                If PREmployee.SSN <> 0 Then
                    PrintValue(1) = Format(PREmployee.SSN, "000-00-0000"):  FormatString(1) = "a11"
                Else
                    PrintValue(1) = "":                                     FormatString(1) = "a11"
                End If
                PrintValue(2) = " ":                                        FormatString(2) = "a2"
                PrintValue(3) = PREmployee.FLName:                          FormatString(3) = "a40"
                PrintValue(4) = " ":                                        FormatString(4) = "~"
                FormatPrint
                Ln = Ln + 1
                
        WageFIC = 10
'        WageFIC = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                          qyear, qyear, EndMonth, 0, 0, 0)
                PrintValue(1) = " ":                                        FormatString(1) = "a19"
                PrintValue(2) = Format(WageFIC, "##,###,##0.00"):           FormatString(2) = "d13"
               
                FindStr = "DeptID =" & CStr(PRDepartment.DepartmentID)
                trs.Find FindStr, 0, adSearchForward, 1
                If trs.EOF Then
                    trs.AddNew Array("DeptID", "WageFIC", "WageMed", "TipsFIC", "TipsMed", "TaxSS", "TaxMed", "TaxFed"), _
                    Array(PRDepartment.DepartmentID, 0, 0, 0, 0, 0, 0, 0)
                    trs.UpdateBatch
                End If
                
                DWageFIC = trs!WageFIC:                                     DWageFIC = DWageFIC + WageFIC
                trs.Fields("DeptID") = trs!DeptID:                            TotWageFIC = TotWageFIC + WageFIC
                trs.Fields("WageFIC") = DWageFIC
                            
        WageMed = 20
'        WageMed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageMed, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(3) = Format(WageMed, "##,###,##0.00"):           FormatString(3) = "d13"
                DWageMed = trs!WageMed:                                     DWageMed = DWageMed + WageMed
                trs.Fields("WageMed") = DWageMed:                           TotWageMed = TotWageMed + WageMed
              
        TipsFIC = 30
'        TipsFIC = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(4) = Format(TipsFIC, "##,##0.00"):               FormatString(4) = "d9"
                DTipsFIC = trs!TipsFIC:                                     DTipsFIC = DTipsFIC + TipsFIC
                trs.Fields("TipsFIC") = DTipsFIC:                           TotTipsFIC = TotTipsFIC + TipsFIC

        TipsMed = 40
'        TipsMed = GetPRAmount(PREmployee.EmployeeID, PREquate.WageGross, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(5) = Format(TipsMed, "##,##0.00"):               FormatString(5) = "d9"
                DTipsMed = trs!TipsMed:                                     DTipsMed = DTipsMed + TipsMed
                trs.Fields("TipsMed") = DTipsMed:                           TotTipsMed = TotTipsMed + TipsMed

        TaxSS = 50
'        TaxSS = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxSS, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(6) = Format(TaxSS, "##,##0.00"):                 FormatString(6) = "d9"
                DTaxSS = trs!TaxSS:                                         DTaxSS = DTaxSS + TaxSS
                trs.Fields("TaxSS") = DTaxSS:                               TotTaxSS = TotTaxSS + TaxSS

        TaxMed = 60
'        TaxMed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxMed, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                PrintValue(7) = Format(TaxMed, "##,##0.00"):                FormatString(7) = "d9"
                DTaxMed = trs!TaxMed:                                       DTaxMed = DTaxMed + TaxMed
                trs.Fields("TaxMed") = DTaxMed:                             TotTaxMed = TotTaxMed + TaxMed

        TaxFed = 70
'        TaxFed = GetPRAmount(PREmployee.EmployeeID, PREquate.TaxFed, _
                          qyear, qyear, StartMonth, EndMonth, 0, 0)
                DTaxFed = trs!TaxFed:                                       DTaxFed = DTaxFed + TaxFed
                trs.Fields("TaxFed") = DTaxFed:                             TotTaxFed = TotTaxFed + TaxFed
                PrintValue(8) = Format(TaxFed, "##,##0.00"):                FormatString(8) = "d9"
                PrintValue(9) = " ":                                        FormatString(9) = "~"
                trs.Update
                FormatPrint
                Ln = Ln + 1
    End Select
cycle2:

    If Not PREmployee.GetNext Then
        Exit Do
    End If
    
    Loop

End Sub

'=======================================   COMPANY HEADER   ======================================

Public Sub PrintCompanyHeader(ByVal ReportList As String)

    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
       MsgBox "Company Info Not Found!!!", vbCritical, "Qtrly Reports"
       End
    End If
    
    PrintValue(1) = "":                                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Name):                               FormatString(2) = "a30"
    PrintValue(3) = "":                                                 FormatString(3) = "a40"
    PrintValue(4) = "REPORT DATE :  " & Format(Date, "mm/dd/yyyy "):    FormatString(4) = "a25"
    PrintValue(5) = " ":                                                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "":                                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Address1):                           FormatString(2) = "a30"
    PrintValue(3) = "":                                                 FormatString(3) = "a40"
    PrintValue(4) = "EMPLOYER ID :  " & Trim(PRCompany.FederalID):      FormatString(4) = "a25"
    PrintValue(5) = " ":                                                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "":                                                 FormatString(1) = "a10"
    PrintValue(2) = Trim(PRCompany.Address2):                           FormatString(2) = "a30"
    PrintValue(3) = "":                                                 FormatString(3) = "a40"
    PrintValue(4) = "EMPLOYER ST.:  " & Trim(PRCompany.StateID):        FormatString(4) = "a25"
    PrintValue(5) = " ":                                                FormatString(5) = "~"
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
'=======================================   FICA/FWT ======================================
        Case "QtrlyFICAFWT"
         
            PrintValue(1) = "SOC. SEC.":                                FormatString(1) = "a11"
            PrintValue(2) = " ":                                        FormatString(2) = "a2"
            PrintValue(3) = "EMPL.":                                    FormatString(3) = "a8"
            PrintValue(4) = " ":                                        FormatString(4) = "a7"
            PrintValue(5) = "GROSS":                                    FormatString(5) = "a12"
            PrintValue(6) = " ":                                        FormatString(6) = "a2"
            PrintValue(7) = " S.S.":                                    FormatString(7) = "a7"
            PrintValue(8) = " ":                                        FormatString(8) = "a7"
            PrintValue(9) = "MEDIC ":                                   FormatString(9) = "a13"
            PrintValue(10) = " ":                                       FormatString(10) = "a3"
            PrintValue(11) = "FWT":                                     FormatString(11) = "a10"
            PrintValue(12) = " ":                                       FormatString(12) = "a3"
            PrintValue(13) = "S.S.":                                    FormatString(13) = "a8"
            PrintValue(14) = " ":                                       FormatString(14) = "a5"
            PrintValue(15) = "MEDIC":                                   FormatString(15) = "a10"
            PrintValue(16) = " ":                                       FormatString(16) = "a4"
            PrintValue(17) = "TOTAL ":                                  FormatString(17) = "a12"
            PrintValue(18) = " ":                                       FormatString(18) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                                   FormatString(1) = "a11"
            PrintValue(2) = " ":                                        FormatString(2) = "a2"
            PrintValue(3) = "NAME":                                     FormatString(3) = "a8"
            PrintValue(4) = " ":                                        FormatString(4) = "a7"
            PrintValue(5) = "WAGE":                                     FormatString(5) = "a10"
            PrintValue(6) = " ":                                        FormatString(6) = "a5"
            PrintValue(7) = "WAGE":                                     FormatString(7) = "a10"
            PrintValue(8) = " ":                                        FormatString(8) = "a3"
            PrintValue(9) = "WAGE ":                                    FormatString(9) = "a11"
            PrintValue(10) = " ":                                       FormatString(10) = "a5"
            PrintValue(11) = "TAX":                                     FormatString(11) = "a10"
            PrintValue(12) = " ":                                       FormatString(12) = "a3"
            PrintValue(13) = "TAX":                                     FormatString(13) = "a10"
            PrintValue(14) = " ":                                       FormatString(14) = "a3"
            PrintValue(15) = " TAX":                                    FormatString(15) = "a10"
            PrintValue(16) = " ":                                       FormatString(16) = "a4"
            PrintValue(17) = "FICA ":                                   FormatString(17) = "a12"
            PrintValue(18) = " ":                                       FormatString(18) = "~"
            FormatPrint
            
            PrintValue(1) = String(118, "_"):                           FormatString(1) = "a118"
            PrintValue(2) = " ":                                        FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
'=======================================   STATE/CITY  ======================================
        Case "QtrlyStateCity"
                        
            PrintValue(1) = "":                                     FormatString(1) = "a5"
            PrintValue(1) = "SOC. SEC.":                            FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a3"
            PrintValue(3) = "EMPLOYEE":                             FormatString(3) = "a30"
            PrintValue(4) = " ":                                    FormatString(4) = "a8"
            PrintValue(5) = "STATE":                                FormatString(5) = "a10"
            PrintValue(6) = " ":                                    FormatString(6) = "a4"
            PrintValue(7) = " CITY":                                FormatString(7) = "a10"
            PrintValue(8) = " ":                                    FormatString(8) = "a4"
            PrintValue(9) = "STATE ":                               FormatString(9) = "a10"
            PrintValue(10) = " ":                                   FormatString(10) = "a5"
            PrintValue(11) = "CITY":                                FormatString(11) = "a9"
            PrintValue(12) = " ":                                   FormatString(12) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                               FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a3"
            PrintValue(3) = "NAME":                                 FormatString(3) = "a30"
            PrintValue(4) = " ":                                    FormatString(4) = "a8"
            PrintValue(5) = "WAGE":                                 FormatString(5) = "a10"
            PrintValue(6) = " ":                                    FormatString(6) = "a5"
            PrintValue(7) = "WAGE":                                 FormatString(7) = "a10"
            PrintValue(8) = " ":                                    FormatString(8) = "a4"
            PrintValue(9) = "TAX ":                                 FormatString(9) = "a10"
            PrintValue(10) = " ":                                   FormatString(10) = "a4"
            PrintValue(11) = "TAX":                                 FormatString(11) = "a10"
            PrintValue(12) = " ":                                   FormatString(12) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "_"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
'=======================================   FED UNEMPLOYMENT  ======================================
        Case "QtrlyFedUnemp"

            PrintValue(1) = "SOC. SEC.":                            FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a2"
            PrintValue(3) = "EMPLOYEE":                             FormatString(3) = "a30"
            PrintValue(4) = " ":                                    FormatString(4) = "a4"
            PrintValue(5) = "QTD GROSS":                            FormatString(5) = "a10"
            PrintValue(6) = " ":                                    FormatString(6) = "a6"
            PrintValue(7) = " YTD GROSS":                           FormatString(7) = "a10"
            PrintValue(8) = " ":                                    FormatString(8) = "a10"
            PrintValue(9) = "FEDERAL ":                             FormatString(9) = "a10"
            PrintValue(10) = " ":                                   FormatString(10) = "a9"
            PrintValue(11) = "STATE":                               FormatString(11) = "a10"
            PrintValue(12) = " ":                                   FormatString(12) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                               FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a3"
            PrintValue(3) = "NAME":                                 FormatString(3) = "a30"
            PrintValue(4) = " ":                                    FormatString(4) = "a6"
            PrintValue(5) = "WAGE":                                 FormatString(5) = "a10"
            PrintValue(6) = " ":                                    FormatString(6) = "a7"
            PrintValue(7) = "WAGE":                                 FormatString(7) = "a10"
            PrintValue(8) = " ":                                    FormatString(8) = "a8"
            PrintValue(9) = "WAGE ":                                FormatString(9) = "a10"
            PrintValue(10) = " ":                                   FormatString(10) = "a8"
            PrintValue(11) = "WAGE":                                FormatString(11) = "a10"
            PrintValue(12) = " ":                                   FormatString(12) = "~"
            FormatPrint
            Ln = Ln + 1
        
            PrintValue(1) = String(118, "_"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
'=======================================   TIPS/TAXES   ======================================
        Case "QtrlyTipsTaxes"
                                    
            PrintValue(1) = "SOC. SEC.":                            FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a2"
            PrintValue(3) = "EMPL.":                                FormatString(3) = "a8"
            PrintValue(4) = " ":                                    FormatString(4) = "a7"
            PrintValue(5) = "FIC":                                  FormatString(5) = "a12"
            PrintValue(6) = " ":                                    FormatString(6) = "a2"
            PrintValue(7) = "MED":                                  FormatString(7) = "a7"
            PrintValue(8) = " ":                                    FormatString(8) = "a7"
            PrintValue(9) = "FIC ":                                 FormatString(9) = "a13"
            PrintValue(10) = " ":                                   FormatString(10) = "a1"
            PrintValue(11) = "MED":                                 FormatString(11) = "a10"
            PrintValue(12) = " ":                                   FormatString(12) = "a5"
            PrintValue(13) = "S.S.":                                FormatString(13) = "a8"
            PrintValue(14) = " ":                                   FormatString(14) = "a6"
            PrintValue(15) = "MED":                                 FormatString(15) = "a10"
            PrintValue(16) = " ":                                   FormatString(16) = "a4"
            PrintValue(17) = "FWT ":                                FormatString(17) = "a12"
            PrintValue(18) = " ":                                   FormatString(18) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "   NO.":                               FormatString(1) = "a11"
            PrintValue(2) = " ":                                    FormatString(2) = "a2"
            PrintValue(3) = "NAME":                                 FormatString(3) = "a8"
            PrintValue(4) = " ":                                    FormatString(4) = "a7"
            PrintValue(5) = "WAGE":                                 FormatString(5) = "a10"
            PrintValue(6) = " ":                                    FormatString(6) = "a4"
            PrintValue(7) = "WAGE":                                 FormatString(7) = "a10"
            PrintValue(8) = " ":                                    FormatString(8) = "a4"
            PrintValue(9) = "TIPS":                                 FormatString(9) = "a11"
            PrintValue(10) = " ":                                   FormatString(10) = "a3"
            PrintValue(11) = "TIPS":                                FormatString(11) = "a10"
            PrintValue(12) = " ":                                   FormatString(12) = "a5"
            PrintValue(13) = "TAX":                                 FormatString(13) = "a10"
            PrintValue(14) = " ":                                   FormatString(14) = "a3"
            PrintValue(15) = " TAX":                                FormatString(15) = "a10"
            PrintValue(16) = " ":                                   FormatString(16) = "a5"
            PrintValue(17) = "TAX ":                                FormatString(17) = "a12"
            PrintValue(18) = " ":                                   FormatString(18) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "_"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
 
    End Select

End Sub

'=======================================   QTRLY TOTALS   ======================================

Public Sub QtrTotals(ByVal ReportList As String)
    frmProgress.lblMsg2 = "Printing Quarterly Final Totals . . . . ."
    frmProgress.Show
    If Ln > MaxLines Then
        FormFeed
        PageHeader ReportTitle, Msg1, "", ""
    End If
    SetFont 8, Equate.Portrait
    Ln = Ln + 1
        
    PrintValue(1) = String(118, "-"):                       FormatString(1) = "a118"
    PrintValue(2) = " ":                                    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
            
    Select Case ReportList
        Case "QtrlyFICAFWT"
            PrintValue(1) = "TOTALS":                               FormatString(1) = "a6"
            PrintValue(2) = "":                                     FormatString(2) = "a14"
            PrintValue(3) = Format(TotWageGross, "##,###,##0.00"):  FormatString(3) = "d8"
            PrintValue(4) = "":                                     FormatString(4) = "a14"
            PrintValue(5) = Format(TotWageMed, "##,###,##0.00"):    FormatString(5) = "d13"
            PrintValue(6) = "":                                     FormatString(6) = "a14"
            PrintValue(7) = Format(TotTaxSS, "##,###,##0.00"):      FormatString(7) = "d13"
            PrintValue(8) = "":                                     FormatString(8) = "a14"
            PrintValue(9) = Format(FinalFica, "##,###,##0.00"):     FormatString(9) = "d13"
            PrintValue(10) = "":                                    FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = " ":                                    FormatString(1) = "a34"
            PrintValue(2) = Format(TotWageSS, "##,###,##0.00"):     FormatString(2) = "d13"
            PrintValue(3) = " ":                                    FormatString(3) = "a14"
            PrintValue(4) = Format(TotTaxFed, "##,###,##0.00"):     FormatString(4) = "d13"
            PrintValue(5) = " ":                                    FormatString(5) = "a14"
            PrintValue(6) = Format(TotTaxMed, "##,###,##0.00"):     FormatString(6) = "d13"
            PrintValue(7) = " ":                                    FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "-"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
        Case "QtrlyStateCity"
            PrintValue(1) = "":                                     FormatString(1) = "a0"
            PrintValue(2) = "TOTALS":                               FormatString(2) = "a6"
            PrintValue(3) = " ":                                    FormatString(3) = "a38"
            PrintValue(4) = Format(TotWageState, "##,###,##0.00"):  FormatString(4) = "d13"
            PrintValue(5) = " ":                                    FormatString(5) = "a0"
            PrintValue(6) = Format(TotWageCity, "##,###,##0.00"):   FormatString(6) = "d13"
            PrintValue(7) = " ":                                    FormatString(7) = "a0"
            PrintValue(8) = Format(TotTaxState, "##,###,##0.00")
            FormatString(8) = "d13"
            PrintValue(9) = " ":                                    FormatString(9) = "a0"
            PrintValue(10) = Format(TotTaxCity, "##,###,##0.00"):   FormatString(10) = "d13"
            PrintValue(11) = " ":                                   FormatString(11) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "-"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
        
        Case "QtrlyFedUnemp"
            PrintValue(1) = "TOTALS":                               FormatString(1) = "a6"
            PrintValue(2) = " ":                                    FormatString(2) = "a37"
            PrintValue(3) = Format(TotQTDWageGross, "##,###,##0.00"):  FormatString(3) = "d13"
            PrintValue(4) = " ":                                    FormatString(4) = "a3"
            PrintValue(5) = Format(TotYTDWageGross, "##,###,##0.00"):  FormatString(5) = "d13"
            PrintValue(6) = " ":                                    FormatString(6) = "a3"
            PrintValue(7) = Format(TotWageFed, "##,###,##0.00"):    FormatString(7) = "d13"
            PrintValue(8) = " ":                                    FormatString(8) = "a3"
            PrintValue(9) = Format(TotWageState, "###,##0.00"):     FormatString(9) = "d10"
            PrintValue(10) = " ":                                   FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "-"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
       Case "QtrlyTipsTaxes"
            PrintValue(1) = "TOTALS":                               FormatString(1) = "a6"
            PrintValue(2) = "":                                     FormatString(2) = "a13"
            PrintValue(3) = Format(TotWageFIC, "##,###,##0.00"):    FormatString(3) = "d13"
            PrintValue(4) = "":                                     FormatString(4) = "a14"
            PrintValue(5) = Format(TotTipsFIC, "##,###,##0.00"):    FormatString(5) = "d13"
            PrintValue(6) = "":                                     FormatString(6) = "a14"
            PrintValue(7) = Format(TotTaxSS, "##,###,##0.00"):      FormatString(7) = "d13"
            PrintValue(8) = "":                                     FormatString(8) = "a14"
            PrintValue(9) = Format(TotTaxFed, "##,###,##0.00"):     FormatString(9) = "d13"
            PrintValue(10) = "":                                    FormatString(10) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = " ":                                    FormatString(1) = "a33"
            PrintValue(2) = Format(TotWageMed, "##,###,##0.00"):    FormatString(2) = "d13"
            PrintValue(3) = " ":                                    FormatString(3) = "a14"
            PrintValue(4) = Format(TotTipsMed, "##,###,##0.00"):    FormatString(4) = "d13"
            PrintValue(5) = " ":                                    FormatString(5) = "a14"
            PrintValue(6) = Format(TotTaxMed, "##,###,##0.00"):     FormatString(6) = "d13"
            PrintValue(7) = " ":                                    FormatString(7) = "~"
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = String(118, "-"):                       FormatString(1) = "a118"
            PrintValue(2) = " ":                                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 1
            
    End Select
    FormFeed
End Sub

'=======================================   QTRLY DEPARTMENT TOTALS   ======================================

Public Sub QtrDeptTotals(ByVal ReportList As String)
    trs.Sort = "DeptID"
    trs.MoveFirst

    frmProgress.lblMsg2 = "Printing Department Totals . . . . ."
    frmProgress.Show
    Ln = 0
    
    Select Case ReportList
        Case "QtrlyFICAFWT"
        
            Do
                If trs.EOF = True Then
                    Exit Do
                End If
                
                If Ln = 0 Or Ln > MaxLines Then
                 
                    If Ln Then FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    SetFont 8, Equate.Portrait
        
                    ' data header
                    Ln = Ln + 2
                    PrintCompanyHeader (ReportList)
        
                End If
                If PRDepartment.GetByID(trs!DeptID) Then
                    PrintValue(1) = PRDepartment.DepartmentNumber & " - " & PRDepartment.Name:
                                                                            FormatString(1) = "a18"
                Else
                    PrintValue(1) = "Department Not Found":                 FormatString(1) = "a18"
                End If
                DFicaTotal = DFicaTotal + trs!TaxSS + trs!TaxMed
                PrintValue(2) = " ":                                        FormatString(2) = "a2"
                PrintValue(3) = Format(trs!WageGross, "##,###,##0.00"):     FormatString(3) = "d13"
                PrintValue(4) = " ":                                        FormatString(4) = "a14"
                PrintValue(5) = Format(trs!WageMed, "##,###,##0.00"):       FormatString(5) = "d13"
                PrintValue(6) = " ":                                        FormatString(6) = "a14"
                PrintValue(7) = Format(trs!TaxSS, "##,###,##0.00"):         FormatString(7) = "d13"
                PrintValue(8) = " ":                                        FormatString(8) = "a14"
                PrintValue(9) = Format(DFicaTotal, "##,###,##0.00"):        FormatString(9) = "d13"
                PrintValue(10) = " ":                                       FormatString(10) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = " ":                                        FormatString(1) = "a34"
                PrintValue(2) = Format(trs!WageSS, "##,###,##0.00"):        FormatString(2) = "d13"
                PrintValue(3) = " ":                                        FormatString(3) = "a14"
                PrintValue(4) = Format(trs!TaxFed, "##,###,##0.00"):        FormatString(4) = "d13"
                PrintValue(5) = " ":                                        FormatString(5) = "a14"
                PrintValue(6) = Format(trs!TaxMed, "##,###,##0.00"):        FormatString(6) = "d13"
                PrintValue(7) = " ":                                        FormatString(7) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = String(118, "-"):                           FormatString(1) = "a118"
                PrintValue(2) = " ":                                        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                DFicaTotal = 0
                trs.MoveNext
            Loop
              
        Case "QtrlyStateCity"
            Do
                If trs.EOF = True Then
                    Exit Do
                End If
                
                If Ln = 0 Or Ln > MaxLines Then
                    If Ln Then FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    SetFont 8, Equate.Portrait
                    ' data header
                    Ln = Ln + 2
                    PrintCompanyHeader (ReportList)
                End If

                If PRDepartment.GetByID(trs!DeptID) Then
                    PrintValue(1) = PRDepartment.DepartmentNumber & " - " & PRDepartment.Name:
                                                                            FormatString(1) = "a18"
                Else
                    PrintValue(1) = "Department Not Found":                 FormatString(1) = "a18"
                End If
                PrintValue(2) = " ":                                        FormatString(2) = "a26"
                PrintValue(3) = Format(trs!WageState, "##,###,##0.00"):     FormatString(3) = "d13"
                PrintValue(4) = " ":                                        FormatString(4) = "a0"
                PrintValue(5) = Format(trs!WageCity, "##,###,##0.00"):      FormatString(5) = "d13"
                PrintValue(6) = " ":                                        FormatString(6) = "a0"
                PrintValue(7) = Format(trs!TaxState, "##,###,##0.00"):      FormatString(7) = "d13"
                PrintValue(8) = " ":                                        FormatString(8) = "a0"
                PrintValue(9) = Format(trs!TaxCity, "##,###,##0.00"):       FormatString(9) = "d13"
                PrintValue(10) = " ":                                       FormatString(10) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = String(118, "-"):                           FormatString(1) = "a118"
                PrintValue(2) = " ":                                        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                trs.MoveNext
            Loop
            
        Case "QtrlyFedUnemp"
            Do
                If trs.EOF = True Then
                    Exit Do
                End If
                
                If Ln = 0 Or Ln > MaxLines Then
                 
                    If Ln Then FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    SetFont 8, Equate.Portrait
        
                    ' data header
                    Ln = Ln + 2
                    PrintCompanyHeader (ReportList)
        
                End If
                
                If PRDepartment.GetByID(trs!DeptID) Then
                    PrintValue(1) = PRDepartment.DepartmentNumber & " - " & PRDepartment.Name:
                                                                            FormatString(1) = "a18"
                Else
                    PrintValue(1) = "Department Not Found":                 FormatString(1) = "a18"
                End If
                PrintValue(2) = " ":                                        FormatString(2) = "a25"
                PrintValue(3) = Format(trs!QTDWageGross, "##,###,##0.00"):  FormatString(3) = "d13"
                PrintValue(4) = " ":                                        FormatString(4) = "a3"
                PrintValue(5) = Format(trs!YTDWageGross, "##,###,##0.00"):  FormatString(5) = "d13"
                PrintValue(6) = " ":                                        FormatString(6) = "a3"
                PrintValue(7) = Format(trs!WageFed, "##,###,##0.00"):       FormatString(7) = "d13"
                PrintValue(8) = " ":                                        FormatString(8) = "a3"
                PrintValue(9) = Format(trs!WageState, "##,###,##0.00"):     FormatString(9) = "d13"
                PrintValue(10) = " ":                                       FormatString(10) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = String(118, "-"):                           FormatString(1) = "a118"
                PrintValue(2) = " ":                                        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                trs.MoveNext
            Loop
            
        Case "QtrlyTipsTaxes"
            Do
                If trs.EOF = True Then
                    Exit Do
                End If
                
                If Ln = 0 Or Ln > MaxLines Then
                 
                    If Ln Then FormFeed
                    PageHeader ReportTitle, Msg1, "", ""
                    SetFont 8, Equate.Portrait
        
                    ' data header
                    Ln = Ln + 2
                    PrintCompanyHeader (ReportList)
        
                End If
        
                If PRDepartment.GetByID(trs!DeptID) Then
                    PrintValue(1) = PRDepartment.DepartmentNumber & " - " & PRDepartment.Name:
                                                                            FormatString(1) = "a18"
                Else
                    PrintValue(1) = "Department Not Found":                 FormatString(1) = "a18"
                End If
                PrintValue(2) = " ":                                        FormatString(2) = "a1"
                PrintValue(3) = Format(trs!WageFIC, "##,###,##0.00"):       FormatString(3) = "d13"
                PrintValue(4) = " ":                                        FormatString(4) = "a14"
                PrintValue(5) = Format(trs!TipsFIC, "##,###,##0.00"):       FormatString(5) = "d13"
                PrintValue(6) = " ":                                        FormatString(6) = "a14"
                PrintValue(7) = Format(trs!TaxSS, "##,###,##0.00"):         FormatString(7) = "d13"
                PrintValue(8) = " ":                                        FormatString(8) = "a14"
                PrintValue(9) = Format(trs!TaxFed, "##,###,##0.00"):        FormatString(9) = "d13"
                PrintValue(10) = " ":                                       FormatString(10) = "~"
                FormatPrint
                Ln = Ln + 1
                              
                PrintValue(1) = " ":                                        FormatString(1) = "a33"
                PrintValue(2) = Format(trs!WageMed, "##,###,##0.00"):       FormatString(2) = "d13"
                PrintValue(3) = " ":                                        FormatString(3) = "a14"
                PrintValue(4) = Format(trs!TipsMed, "##,###,##0.00"):       FormatString(4) = "d13"
                PrintValue(5) = " ":                                        FormatString(5) = "a14"
                PrintValue(6) = Format(trs!TaxMed, "##,###,##0.00"):        FormatString(6) = "d13"
                PrintValue(7) = " ":                                        FormatString(7) = "~"
                FormatPrint
                Ln = Ln + 1
                
                PrintValue(1) = String(118, "-"):                           FormatString(1) = "a118"
                PrintValue(2) = " ":                                        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
                
                trs.MoveNext
            Loop
    End Select

    trs.Close
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================    FILE RATE LIST     ======================================

Public Sub RateList()
    frmRateList.Hide
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    
    PrtInit ("Port")    ' "Port" = Portrait
    
    ' set up SQL statement based upon order requested
    ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NO."
    If frmRateList.optNumber Then
        ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NO."
        SQLString = "SELECT * FROM PRCITY ORDER BY CityNumber"
    Else
        ReportTitle = "PAYROLL CITY RATE FILE LISTING BY CITY NAME"
        SQLString = "SELECT * FROM PRCITY ORDER BY CityName"
    End If
    
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If Not PRCity.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Payroll Rate File Listing"
        Exit Sub
    End If
    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, " ", "", ""
            SetFont 10, Equate.Portrait
            Ln = Ln + 2                ' Changed from Ln +1 to Ln + 2
            
            PrintValue(1) = "City Num":                         FormatString(1) = "a9"
            PrintValue(2) = " ":                                FormatString(2) = "a3"
            PrintValue(3) = "City Name":                        FormatString(3) = "a30"
            PrintValue(4) = " ":                                FormatString(4) = "a4"
            PrintValue(5) = "City Tax Rate":                    FormatString(5) = "a13"
            PrintValue(6) = " ":                                FormatString(6) = "~"
            FormatPrint
            Ln = Ln + 1
             
            PrintValue(1) = String(94, "-"):                       FormatString(1) = "a94"
            PrintValue(2) = " ":                                    FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1
        End If
    
        PrintValue(1) = PRCity.CityNumber:                      FormatString(1) = "n5"
        PrintValue(2) = " ":                                    FormatString(2) = "a7"
        PrintValue(3) = PRCity.CityName:                        FormatString(3) = "a30        "
        PrintValue(4) = " ":                                    FormatString(4) = "a0"
        PrintValue(5) = PRCity.CityRate:                        FormatString(5) = "d8"
        PrintValue(6) = " ":                                    FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 1
        
        If Not PRCity.GetNext Then
            Exit Do
        End If
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================   CITY TAX REPORT - by Employee     ======================================

Public Sub CityTaxRptEmployee()
Dim LastEmpID As Long

    trs.MoveFirst
    LastCityID = 0
    Ln = 0
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    PrtInit ("Port")
    ReportTitle = "CITY TAX REPORT BY EMPLOYEE"
    SetFont 10, Equate.Portrait

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
            
    Do

        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            CityTaxHeader (ReportTitle)
        End If
   
        If Not PREmployee.GetByID(trs!EmployeeID) Then
            MsgBox "Employee Info Not Found!!!", vbCritical, "Payroll Yearly City Tax Report"
            End
        End If
        
        If LastEmpID <> 0 And LastEmpID <> trs!EmployeeID Then

            PrintValue(1) = "       TOTALS "
            FormatString(1) = "a44"

            PrintValue(2) = SYTDGROSS
            FormatString(2) = "d10"

            PrintValue(3) = " "
            FormatString(3) = "a1"

            PrintValue(4) = SYTDTAX
            FormatString(4) = "d10"

            PrintValue(5) = " "
            FormatString(5) = "~"

            FormatPrint
            Ln = Ln + 1
            SMTDGross = 0
            SMTDTax = 0
            SQTDGross = 0
            SQTDTax = 0
            SYTDGROSS = 0
            SYTDTAX = 0

        End If
        frmProgress.lblMsg2 = "Employee: " & trs!EmployeeNumber & " - " & Trim(PREmployee.LFName)
        frmProgress.Show
        
        If LastEmpID = 0 Or LastEmpID <> trs!EmployeeID Then
            Ln = Ln + 1

               PrintValue(1) = trs!EmployeeNumber
               FormatString(1) = "a5"

               PrintValue(2) = " "
               FormatString(2) = "a2"

               PrintValue(3) = PREmployee.LFName
               FormatString(3) = "a30"

               PrintValue(4) = " "
               FormatString(4) = "a3"

               PrintValue(5) = " "
               FormatString(5) = "~"

               FormatPrint
               Ln = Ln + 1

        End If

        If Not PRCity.GetBySQL("Select * from PRCity where PRCity.Cityid = " & trs!CityID) Then
           CityName = trs!CityID & " Not Found!"
        End If

        PrintValue(1) = PRCity.CityNumber & "    " & PRCity.CityName & "            "
        FormatString(1) = "a24"
        
        PrintValue(2) = " "
        FormatString(2) = "a20"

        PrintValue(3) = trs!YTDGross
        FormatString(3) = "d13"
        
        SYTDGROSS = SYTDGROSS + trs!YTDGross
        TYTDGross = TYTDGross + trs!YTDGross

        PrintValue(4) = " "
        FormatString(4) = "a1"

        PrintValue(5) = trs!YTDTax
        FormatString(5) = "d13"
        
        SYTDTAX = SYTDTAX + trs!YTDTax
        TYTDTAX = TYTDTAX + trs!YTDTax

        PrintValue(6) = " "
        FormatString(6) = "~"

        FormatPrint

        Ln = Ln + 1

        LastEmpID = trs!EmployeeID
        trs.MoveNext

        If trs.EOF Then
           Exit Do
        End If

    Loop
    
    PrintValue(1) = "       TOTALS "
    FormatString(1) = "a24"

    PrintValue(2) = " "
    FormatString(2) = "a20"

    PrintValue(3) = SYTDGROSS
    FormatString(3) = "d10"

    PrintValue(4) = " "
    FormatString(4) = "a1"

    PrintValue(5) = SYTDTAX
    FormatString(5) = "d10"

    PrintValue(6) = " "
    FormatString(6) = "~"
    FormatPrint

    Ln = Ln + 2

    PrintValue(1) = "       GRAND TOTALS "
    FormatString(1) = "a24"

    PrintValue(2) = " "
    FormatString(2) = "a20"

    PrintValue(3) = TYTDGross
    FormatString(3) = "d10"

    PrintValue(4) = " "
    FormatString(4) = "a1"

    PrintValue(5) = TYTDTAX
    FormatString(5) = "d10"

    PrintValue(6) = " "
    FormatString(6) = "~"

    FormatPrint

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================   CITY TAX REPORT - by City     ======================================

Public Sub CityTaxRpt(ByVal RangeType As Byte, _
                            ByVal BatchNumbr As Long, _
                            ByVal PEDate As Long, _
                            ByVal CheckDt As Long, _
                            ByVal Startdate As Long, _
                            ByVal EndDate As Long)
Dim SYTDGROSS As Currency
Dim SYTDTAX As Currency
Dim TYTDGross As Currency
Dim TYTDTAX As Currency
Dim StartYM As Long
Dim EndYM As Long
Dim CityName As String
Dim CityNumber As Long
Dim LastCityID As Long
Dim LastCityName As String
Dim LastCityNumber As Long
Dim ReportTitle As String

    frmCityTaxRpt.Hide
    SetEquates
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
            
    trs.CursorLocation = adUseClient
    trs.Fields.Append "TempID", adDouble
    trs.Fields.Append "CityID", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    trs.Fields.Append "EmployeeNumber", adDouble
    trs.Fields.Append "LastName", adVarChar, 30, adFldIsNullable
    trs.Fields.Append "FirstName", adVarChar, 30, adFldIsNullable
    trs.Fields.Append "YTDGross", adCurrency
    trs.Fields.Append "YTDTax", adCurrency
    
    trs.Open , , adOpenDynamic, adLockOptimistic
           
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
           
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.BatchID = " & BatchNumbr
    Else
        SQLString = "SELECT * FROM PRDist WHERE PRDist.PEDate >= " & Startdate & _
        " AND PRDist.PEDate <= " & EndDate
    End If

    If Not PRDist.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "City Tax Report"
        Exit Sub
    End If
    
    If RangeType = PREquate.RangeTypeBatch Then
        Msg1 = "BATCH " & BatchNumbr & " - Period Ending: " & CDate(PEDate)
    Else
        Msg1 = "PERIOD ENDING DATE FROM: " & CDate(Startdate) & " TO: " & CDate(EndDate)
    End If
    
    Do
        If frmCityTaxRpt.optbyCity Then
            TempID = PRDist.CityID * 10 ^ 6 + PRDist.EmployeeID
        Else
            TempID = PRDist.EmployeeID * 10 ^ 6 + PRDist.CityID
        End If
        
        SQLString = "TempID = " & TempID
        trs.Find SQLString, 0, adSearchForward, 1
        
        If Not PREmployee.GetByID(PRDist.EmployeeID) Then
            MsgBox "Employee Info Not Found!!!", vbCritical, "Payroll Yearly City Tax Report"
            End
        End If
            
        If trs.EOF Then
            trs.AddNew
            trs!TempID = TempID
            trs!CityID = PRDist.CityID
            trs!EmployeeID = PRDist.EmployeeID
            trs!EmployeeNumber = PREmployee.EmployeeNumber
            trs!LastName = Trim(PREmployee.LastName)
            trs!FirstName = Trim(PREmployee.FirstName)
            trs!YTDGross = 0
            trs!YTDTax = 0
            trs.Update
        End If
        trs!YTDGross = trs!YTDGross + PRDist.CityWage
        trs!YTDTax = trs!YTDTax + PRDist.CityTax
        trs.Update
        
        If Not PRDist.GetNext Then Exit Do
    Loop

    trs.Sort = "TempID"
    
    If frmCityTaxRpt.optByEmployee Then
        CityTaxRptEmployee
        End
    End If
        
    trs.MoveFirst
    LastCityID = 0
    Ln = 0
    
    ReportTitle = "PAYROLL CITY TAX REPORT BY CITY"
    
    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            CityTaxHeader (ReportTitle)
        End If

        If LastCityID = 0 Then
            If PRCity.GetBySQL("Select * from PRCity where PRCity.CityID = " & trs!CityID) Then
                CityName = PRCity.CityName
                CityNumber = PRCity.CityNumber
            Else
                CityName = CityNumber & " Not Found!"
            End If
            
            PrintValue(1) = CityNumber
            FormatString(1) = "a5"
            
            PrintValue(2) = "   *** REPORT FOR CITY OF:  " & Trim(CityName) & "  ***"
            FormatString(2) = "a60"
            
            PrintValue(3) = " "
            FormatString(3) = "~"
            
            FormatPrint
            Ln = Ln + 2
        
        End If
        
        If LastCityID <> 0 And LastCityID <> trs!CityID Then
        
            Ln = Ln + 1
            
            PrintValue(1) = LastCityNumber & " - " & Trim(LastCityName) & " TOTALS"
            FormatString(1) = "a41"
            
            PrintValue(2) = " "
            FormatString(2) = "a3"
            
            PrintValue(3) = SYTDGROSS
            FormatString(3) = "d13"
            
            PrintValue(4) = " "
            FormatString(4) = "a2"
            
            PrintValue(5) = SYTDTAX
            FormatString(5) = "d13"
            
            PrintValue(6) = " "
            FormatString(6) = "~"
            
            FormatPrint
            Ln = Ln + 2
            
            SYTDGROSS = 0
            SYTDTAX = 0
            
            FormFeed
            CityTaxHeader (ReportTitle)
            
            If PRCity.GetBySQL("Select * from PRCity where PRCity.CityID = " & trs!CityID) Then
                CityName = PRCity.CityName
                CityNumber = PRCity.CityNumber
            Else
                CityName = CityNumber & " Not Found!"
            End If
            
            PrintValue(1) = CityNumber
            FormatString(1) = "a5"
            
            PrintValue(2) = "   *** REPORT FOR CITY OF:  " & Trim(CityName) & "  ***"
            FormatString(2) = "a60"
        
            PrintValue(3) = " "
            FormatString(3) = "~"
            
            FormatPrint
            Ln = Ln + 2
                            
        End If
        
        frmProgress.lblMsg2 = "Employee: " & trs!EmployeeNumber & " - " & Trim(trs!LastName) & ", " & Trim(trs!FirstName)
        frmProgress.Show
        
        PrintValue(1) = trs!EmployeeNumber
        FormatString(1) = "a6"
                      
        PrintValue(2) = " "
        FormatString(2) = "a1"
        
        PrintValue(3) = Trim(trs!LastName)
        FormatString(3) = "a20"
                      
        PrintValue(4) = " "
        FormatString(4) = "a1"
        
        PrintValue(5) = Trim(trs!FirstName)
        FormatString(5) = "a15"
                
        PrintValue(6) = " "
        FormatString(6) = "a1"

        PrintValue(7) = trs!YTDGross
        FormatString(7) = "d10"
        
        SYTDGROSS = SYTDGROSS + trs!YTDGross
        TYTDGross = TYTDGross + trs!YTDGross
        
        PrintValue(8) = " "
        FormatString(8) = "a2"
        
        PrintValue(9) = trs!YTDTax
        FormatString(9) = "d10"
        
        SYTDTAX = SYTDTAX + trs!YTDTax
        TYTDTAX = TYTDTAX + trs!YTDTax
        
        PrintValue(10) = " "
        FormatString(10) = "~"
        
        FormatPrint
        Ln = Ln + 1
              
        LastCityID = trs!CityID
        LastCityNumber = CityNumber
        LastCityName = CityName
        LastTempID = TempID
        
        trs.MoveNext
        
        If trs.EOF Then
            Exit Do
        End If
    Loop
    
    Ln = Ln + 1
    
    PrintValue(1) = LastCityNumber & " - " & Trim(LastCityName) & " TOTALS"
    FormatString(1) = "a41"
    
    PrintValue(2) = " "
    FormatString(2) = "a3"
    
    PrintValue(3) = SYTDGROSS
    FormatString(3) = "d13"
    
    PrintValue(4) = " "
    FormatString(4) = "a2"
    
    PrintValue(5) = SYTDTAX
    FormatString(5) = "d13"
    
    PrintValue(6) = " "
    FormatString(6) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "GRAND TOTALS"
    FormatString(1) = "a40"
    
    PrintValue(2) = " "
    FormatString(2) = "a4"
        
    PrintValue(3) = TYTDGross
    FormatString(3) = "d13"
    
    PrintValue(4) = " "
    FormatString(4) = "a2"
    
    PrintValue(5) = TYTDTAX
    FormatString(5) = "d13"
    
    PrintValue(6) = " "
    FormatString(6) = "~"
    
    FormatPrint

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

'=======================================   CITY TAX HEADER    ======================================

Public Sub CityTaxHeader(ReportTitle)
 
    PageHeader ReportTitle, Msg1, "", ""
    
    Ln = Ln + 1
    
    PrintValue(1) = "NUMBER"
    FormatString(1) = "a6"
    
    PrintValue(2) = " "
    FormatString(2) = "a1"
    
    PrintValue(3) = "EMPLOYEE NAME"
    FormatString(3) = "a30"
    
    PrintValue(4) = " "
    FormatString(4) = "a11"

    PrintValue(5) = "YTD GROSS"
    FormatString(5) = "a10"
    
    PrintValue(6) = " "
    FormatString(6) = "a8"
     
    PrintValue(7) = "YTD TAX"
    FormatString(7) = "a10"
    
    PrintValue(8) = " "
    FormatString(8) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "------------------------------------------------------------------------------------------------"
    FormatString(1) = "a96"
    
    PrintValue(2) = " "
    FormatString(2) = "~"
    
    FormatPrint
    Ln = Ln + 1
End Sub

'=======================================   WAGE REVIEW    ======================================

Public Sub WageReviewJournal()

    Ln = 0
    SetEquates
    NumEmployees = 0
    PrtInit ("Port")
    ReportTitle = "EMPLOYER'S REPORT OF WAGES - JOURNAL"
    SetFont 10, Equate.Portrait
    
    ' set up SQL statement based upon order requested
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If frmWageReview.cmbQtr = 1 Then
      QtrEnding = "QUARTER ENDING: 03/31/" & frmWageReview.cmbYear
    ElseIf frmWageReview.cmbQtr = 2 Then
      QtrEnding = "QUARTER ENDING: 06/30/" & frmWageReview.cmbYear
    ElseIf frmWageReview.cmbQtr = 3 Then
      QtrEnding = "QUARTER ENDING: 09/30/" & frmWageReview.cmbYear
    ElseIf frmWageReview.cmbQtr = 4 Then
      QtrEnding = "QUARTER ENDING: 12/31/" & frmWageReview.cmbYear
    End If
    
    Msg1 = QtrEnding
    
    frmWageReview.rs.MoveFirst
    If frmWageReview.optEmployee Then
        rs.Sort = "EmpID"
    Else
        rs.Sort = "SSN"
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
            PrintValue(2) = Format(frmWageReview.rs!SSN, "000-00-0000"):          FormatString(2) = "a11"
            NumEmployees = NumEmployees + 1
            PrintValue(3) = " ":                            FormatString(3) = "a2"
            PrintValue(4) = frmWageReview.rs!EmpName:       FormatString(4) = "a25"
                                   
            TotWageGross = TotWageGross + frmWageReview.rs!Gross
            PrintValue(5) = frmWageReview.rs!Gross:         FormatString(5) = "d13"
            PrintValue(6) = " ":                            FormatString(6) = "a3"
            PrintValue(7) = frmWageReview.rs!noweeks:       FormatString(7) = "a3"
            PrintValue(8) = " ":                            FormatString(8) = "~"
            FormatPrint
            Ln = Ln + 1
'cycle3:
        frmWageReview.rs.MoveNext
        If frmWageReview.rs.EOF Then Exit Do
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

Public Sub WageSupplemental()
Dim Remainder As Integer


PadNumber = 30

    
    SetEquates
    NumEmployees = 0
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 10, Equate.Portrait
    PrtTitle = frmWageReview.txtTitle
    frmWageReview.rs.MoveLast
    NumPages = frmWageReview.rs.RecordCount / 20
    frmWageReview.rs.MoveFirst
    CurrPg = 0
    LnCnt = 0
    Ln = 0
'    YUnits = 122
    
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.StateID) Then
        StateAbbrev = PRState.StateAbbrev
    End If
                           
    Do
        NumEmployees = NumEmployees + 1
        CustName = frmWageReview.rs!EmpName
   
        If Ln = 0 Or NumEmployees > 20 Then
            NumEmployees = 1
            Ln = Ln + 5
            LnCnt = 0
           
            Prvw.vsp.CurrentY = 1200      'row
            Prvw.vsp.CurrentX = 9550      'column
            Prvw.vsp.Text = qQuarter
        
            Prvw.vsp.CurrentY = 1200
            Prvw.vsp.CurrentX = 10200
            Prvw.vsp.Text = qYear
            
            Prvw.vsp.CurrentY = 1450
            Prvw.vsp.CurrentX = 800
            Prvw.vsp.Text = PRCompany.StateID
                        
            Prvw.vsp.CurrentY = 1850
            Prvw.vsp.CurrentX = 800
            Prvw.vsp.Text = PRCompany.Name

            Prvw.vsp.CurrentY = 2150
            Prvw.vsp.CurrentX = 800
            Prvw.vsp.Text = PRCompany.Address1
            
            Prvw.vsp.CurrentY = 2450
            Prvw.vsp.CurrentX = 800
            Prvw.vsp.Text = PRCompany.Address2
            
            Prvw.vsp.CurrentY = 2750
            Prvw.vsp.CurrentX = 800
            Prvw.vsp.Text = PRCompany.City & ", " & StateAbbrev & "  " & PRCompany.ZipCode
            
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Prvw.vsp.CurrentY = 3850
            Prvw.vsp.CurrentX = 900

        End If
        
        If NumEmployees = 1 Then
            Prvw.vsp.Text = Format(frmWageReview.rs!SSN, "000-00-0000")
            
            If Len(frmWageReview.rs!EmpName) < 30 Then
                CustLen = Len(frmWageReview.rs!EmpName)
                PadSpaces = PadNumber - CustLen
                PadString = Trim(frmWageReview.rs!EmpName) & Space(PadSpaces)
            End If
            
            Prvw.vsp.CurrentX = 3000
            Prvw.vsp.Text = PadString           '  Print Name

            TotWageGross = TotWageGross + frmWageReview.rs!Gross
            Prvw.vsp.CurrentX = 6780
            PadString = Format(frmWageReview.rs!Gross, "###,###,##0.00")
            AmtLen = Len(PadString)
            PadString = Space(DPadNumber - AmtLen) & PadString
            Prvw.vsp.Text = PadString           '  Print Amount
            
            Prvw.vsp.CurrentX = Prvw.vsp.CurrentX + 500
            Prvw.vsp.Text = frmWageReview.rs!noweeks
            Prvw.vsp.CurrentX = 0
            GoTo CycleIt
        Else
            Prvw.vsp.CurrentY = Prvw.vsp.CurrentY + 500
            Prvw.vsp.CurrentX = 900
        End If
        
        Prvw.vsp.Text = Format(frmWageReview.rs!SSN, "000-00-0000")

        Prvw.vsp.CurrentX = 3000
        
        If Len(frmWageReview.rs!EmpName) < 30 Then
            CustLen = Len(frmWageReview.rs!EmpName)
            PadSpaces = PadNumber - CustLen
            PadString = Trim(frmWageReview.rs!EmpName) & Space(PadSpaces)
        End If
        Prvw.vsp.Text = PadString                   '  Print Name

        TotWageGross = TotWageGross + frmWageReview.rs!Gross
        Prvw.vsp.CurrentX = 6780
        PadString = Format(frmWageReview.rs!Gross, "###,###,##0.00")
        AmtLen = Len(PadString)
        PadString = Space(DPadNumber - AmtLen) & PadString
        Prvw.vsp.Text = PadString                   '  Print Amount

        Prvw.vsp.CurrentX = Prvw.vsp.CurrentX + 500
        Prvw.vsp.Text = frmWageReview.rs!noweeks
CycleIt:
        If NumEmployees = 20 Then
            WageReviewTotals
            TotWageGross = 0
            FormFeed
        End If
        
        frmWageReview.rs.MoveNext
        If frmWageReview.rs.EOF Then Exit Do

    Loop
    WageReviewTotals

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

   
End Sub

Public Sub WageReviewTotals()

    Prvw.vsp.CurrentY = 13700
    Prvw.vsp.CurrentX = 6780

    PadString = Format(TotWageGross, "###,###,##0.00")
    AmtLen = Len(PadString)
    PadString = Space(DPadNumber - AmtLen) & PadString
    Prvw.vsp.Text = PadString                   '  Print Final Amount
    
    CurrPg = CurrPg + 1
    
    Prvw.vsp.CurrentY = 13660
    Prvw.vsp.CurrentX = 8900
    Prvw.vsp.Text = CurrPg
    
    Prvw.vsp.CurrentY = 13660
    Prvw.vsp.CurrentX = 9630
    Prvw.vsp.Text = NumPages
    
    Prvw.vsp.CurrentY = 14500
    Prvw.vsp.CurrentX = 5000
    Prvw.vsp.Text = PrtTitle

    Prvw.vsp.CurrentY = 14500
    Prvw.vsp.CurrentX = 7800
    Prvw.vsp.Text = Format(PrtDate, "mm/dd/yyyy")
    
End Sub

'=======================================   FORM 941     ======================================

Public Sub Form941APrint()
Dim VertSpace As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString As String
 
    CurrYear = Year(Now())
    Ln = 0
    SetEquates
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    SQLString = "SELECT * FROM PREmployee"
    
    rsInit SQLString, cn, rs

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

    PosPrint 9400, 3410, PadRight(Format(frm941Entry.Line1, FmtString), 13)
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
    PosPrint 5000, 3800, PadRight(Format(frm941Entry.Line17Mo1, FmtString), 13)
    PosPrint 5000, 4250, PadRight(Format(frm941Entry.Line17Mo2, FmtString), 13)
    PosPrint 5000, 4710, PadRight(Format(frm941Entry.Line17Mo3, FmtString), 13)
    PosPrint 5000, 5230, PadRight(Format(frm941Entry.Line17Total, FmtString), 13)
    PosPrint 1820, 5540, frm941Entry.AlphaCheckLine17c
    PosPrint 9200, 6500, frm941Entry.AlphaCheckLine18
        
    PosPrint 3900, 6950, frm941Entry.Line18Date
    PosPrint 9200, 7220, frm941Entry.AlphaCheckLine19
    PosPrint 880, 8200, frm941Entry.AlphaCheckPart4Yes
    PosPrint 3500, 8200, frm941Entry.Part4Name
    PosPrint 3500, 8650, Format(frm941Entry.Part4Phone, "###-###-####")
    PosPrint 10000, 8650, frm941Entry.Part4Pin
    PosPrint 900, 8900, frm941Entry.AlphaCheckPart4No
'    FormFeed
    PosPrint 3000, 10600, frm941Entry.Part5NameTitle
    PosPrint 3000, 11100, frm941Entry.Part5Date
    PosPrint 5100, 11100, Format(frm941Entry.Part5Phone, "###-###-####")
    PosPrint 3000, 12110, frm941Entry.Part5PrepName
    PosPrint 9000, 12500, Format(frm941Entry.Part5PrepPhone, "###-###-####")
    PosPrint 3000, 12980, frm941Entry.Part5Firm
    PosPrint 9000, 12980, frm941Entry.Part5EIN
    PosPrint 3000, 13470, frm941Entry.Part5Addr1
    PosPrint 9000, 13470, frm941Entry.Part5Zip
    PosPrint 3000, 13920, frm941Entry.Part5Addr2
    PosPrint 9000, 13920, frm941Entry.Part5SSN
    PosPrint 2660, 14450, frm941Entry.AlphaCheckPart5
    PosPrint 8180, 14450, frm941Entry.Part5PrepDate
    FormFeed
End Sub

Sub PrintFlexGridBuiltIn()
    fg.PrintGrid "My Grid"
End Sub

Public Sub Form941BPrint(ByVal VertPos As Long, ByRef rsB As ADODB.Recordset)

Dim VertSpace As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString As String

    SetEquates
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    SQLString = "SELECT * FROM PREmployee"
    
    rsInit SQLString, cn, rs

    VertSpace = 492
    FmtString = "##,###,##0.00"
    
    Col1X = 570
    Col2X = 2800
    Col3X = 5000
    Col4X = 7200

    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1a, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2a, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3a, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4a, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1b, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2b, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3b, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4b, FmtString), 13)
    PosPrint Col4X + 2300, VertPos, PadRight(Format(BMoTax, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1c, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2c, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3c, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4c, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1d, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2d, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3d, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4d, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1e, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2e, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3e, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4e, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1f, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2f, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3f, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4f, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1g, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2g, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3g, FmtString), 13)
    PosPrint Col4X, VertPos, PadRight(Format(rsB!cell4g, FmtString), 13)
    VertPos = VertPos + VertSpace
    
    PosPrint Col1X, VertPos, PadRight(Format(rsB!cell1h, FmtString), 13)
    PosPrint Col2X, VertPos, PadRight(Format(rsB!cell2h, FmtString), 13)
    PosPrint Col3X, VertPos, PadRight(Format(rsB!cell3h, FmtString), 13)

       
End Sub

Public Sub Form941BHdr()
FmtString = "##,###,##0.00"

    CurrYear = Year(Now())
    PosPrint 3380, 900, PRCompany.FederalID
    If frm941Entry.cmbQtr = 1 Then
        PosPrint 8450, 900, "X"
    ElseIf frm941Entry.cmbQtr = 2 Then
        PosPrint 8450, 1150, "X"
    ElseIf frm941Entry.cmbQtr = 3 Then
        PosPrint 8450, 1430, "X"
    ElseIf frm941Entry.cmbQtr = 4 Then
        PosPrint 8420, 1660, "X"
    End If
    
    PosPrint 3380, 1150, PRCompany.Name
    PosPrint 3380, 1430, CurrYear
    PosPrint 9600, 14300, PadRight(Format(TotTaxLiability, FmtString), 13)
    
End Sub

Sub PrintFlexGridOnVSPrinter()
    vp.StartDoc
    vp.RenderControl = fg.hWnd
    vp.EndDoc
End Sub

'=======================================   CHECK REGISTER    ======================================

Public Sub CheckRegister(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal CheckDt As Long, _
                         ByVal Startdate As Long, _
                         ByVal EndDate As Long)
Dim sqlstring1 As String
    
    frmCheckReg.Hide
    ReportTitle = "PAYROLL CHECK REGISTER"
    PrtInit ("Land")
    LandSW = 1
    SetFont 8, Equate.LandScape
    SetEquates
    PRTotal.CreateRS
           
    sqlstring1 = "SELECT PREmployee.*, PRHist.*, PRDepartment.* " & _
                 " FROM (PREmployee " & _
                 " INNER JOIN PRHist ON PREmployee.EmployeeID = PRHist.EmployeeID) " & _
                 " INNER JOIN PRDepartment ON PRDepartment.DepartmentID = PREmployee.DepartmentID "

    If RangeType = PREquate.RangeTypeBatch Then
        sqlstring1 = Trim(sqlstring1) & " WHERE PRHist.PEDate >= " & CLng(Startdate) & " AND " & _
                                      " PRHist.BatchID = " & BatchNumbr
        DedString = " PRHist.BatchID = " & BatchNumbr
                                      
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & CDate(CheckDt)
    Else
        sqlstring1 = Trim(sqlstring1) & " WHERE PRHist.PEDate >= " & CDate(Startdate) & " AND " & _
                                    " PRHist.PEDate <= " & CDate(EndDate)
        DedString = " PRHist.PEDate >= " & CDate(Startdate) & " AND " & _
                                    " PRHist.PEDate <= " & CDate(EndDate)
        Msg1 = "Date Range: " & CDate(Startdate) & " TO: " & CDate(EndDate)
    End If
        
    ' set up SQL statement based upon order checked
    If frmCheckReg.optCheckNo = True Then
        sqlstring1 = Trim(sqlstring1) & " ORDER BY PRHist.CheckNumber"
        ReportTitle = Trim(ReportTitle) & " BY CHECK NUMBER"
    ElseIf frmCheckReg.optEmpNo = True Then
        sqlstring1 = Trim(sqlstring1) & " ORDER BY PREmployee.EmployeeNumber"
        ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NUMBER"
    Else                                                          ' order by Employee Name
        sqlstring1 = Trim(sqlstring1) & " ORDER BY PREmployee.LastName, PREmployee.FirstName"
        ReportTitle = Trim(ReportTitle) & " BY EMPLOYEE NAME"
    End If

    rsInit sqlstring1, cn, rrs       ' rrs vars get assigned in rsInit

    If rrs.EOF = True And rrs.BOF = True Then
        MsgBox "No Data Found !!!", vbCritical, "Payroll Check Register"
        Prvw.vsp.EndDoc
        Exit Sub
    End If
    
    rrs.MoveFirst
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show

    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    If frmCheckReg.chkOEHrs Or frmCheckReg.chkOEAmt Or frmCheckReg.chkDed Then
        ChkRegGetHeaderData
    End If

    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            ChkRegFirstHeader
        End If

        If frmCheckReg.optSubTotal = True Then
            If LastEmpNumber <> 0 And rrs!EmployeeNumber <> LastEmpNumber Then
                GetEmpInfo (rrs!EmployeeNumber)
                EmpFlag = True
                ChkRegTotals
                PRTotal.Clear
                PRTotal.Save (Equate.RecPut)
                EmpFlag = False
                PRTotal.Clear
                RecCnt = 0

            ElseIf LastEmpNumber = rrs!EmployeeNumber Then
                TotalFlag = True
            End If
        Else
            If LastEmpNumber = rrs!EmployeeNumber Then
                TotalFlag = True
            End If
        End If
        
        RecCnt = RecCnt + 1

        '  RecType/IDNumber
        UpdateTotals 1, 999999998, 999999998, rrs![PREmployee.DepartmentID]
        UpdateTotals 2, rrs![PREmployee.DepartmentID], rrs!DepartmentNumber, rrs![PREmployee.DepartmentID]
        UpdateTotals 3, 999999999, 999999999, rrs![PREmployee.DepartmentID]        ' Update Grand Totals - 9's for recid so grand totals print last
      
      ' PRINT DETAIL   ##############################################################
      
        TotTaxes = rrs!SSTax + rrs!MedTax + rrs!FWTTax + rrs!SWTTax + rrs!CWTTax
        TOTHours = rrs!RegHours + rrs!OTHours + rrs!OEHours
        frmProgress.lblMsg2 = "Employee: " & rrs!EmployeeNumber & " - " & Trim(rrs!LastName) & ", " & (rrs!FirstName)
        frmProgress.Show
        
        PrintValue(1) = rrs!CheckNumber
        FormatString(1) = "n6"
        
        PrintValue(2) = " "
        FormatString(2) = "a1"
        
        PrintValue(3) = rrs!DepartmentNumber
        FormatString(3) = "n3"
                                 
        PrintValue(4) = " "
        FormatString(4) = "a1"
            
        PrintValue(5) = rrs!EmployeeNumber & "-" & Trim(rrs!LastName) & ", " & rrs!FirstName
        
        FormatString(5) = "a25"
        
        PrintValue(6) = " "
        FormatString(6) = "a2"
        
        PrintValue(7) = rrs!RegAmount
        FormatString(7) = "d9"
        
        PrintValue(8) = " "
        FormatString(8) = "a0"
        
        PrintValue(9) = rrs!OTAmount
        FormatString(9) = "d2"
        
        PrintValue(10) = " "
        FormatString(10) = "a0"

        PrintValue(11) = rrs!OEAmount
        FormatString(11) = "d8"
        
        PrintValue(12) = " "
        FormatString(12) = "a0"
        
        PrintValue(13) = rrs!Gross
        FormatString(13) = "d8"
        
        PrintValue(14) = " "
        FormatString(14) = "a0"
        
        PrintValue(15) = rrs!Deductions
        FormatString(15) = "d6"
        
        PrintValue(16) = " "
        FormatString(16) = "a0"
        
        PrintValue(17) = rrs!Net
        FormatString(17) = "d8"

        PrintValue(18) = " "
        FormatString(18) = "~"
        
        FormatPrint
        Ln = Ln + 1
                    
        PrintValue(1) = " "
        FormatString(1) = "a0"
      
        PrintValue(2) = rrs!PEDatef
        FormatString(2) = "a10"
        
        PrintValue(3) = " "
        FormatString(3) = "a0"
                 
        PrintValue(4) = rrs!SSTax
        FormatString(4) = "d6"
        
        PrintValue(5) = " "
        FormatString(5) = "a0"
        
        PrintValue(6) = rrs!MedTax
        FormatString(6) = "d6"
        
        PrintValue(7) = " "
        FormatString(7) = "a0"
        
        PrintValue(8) = rrs!FWTTax
        FormatString(8) = "d6"
        
        PrintValue(9) = " "
        FormatString(9) = "a0"
        
        PrintValue(10) = rrs!SWTTax
        FormatString(10) = "d6"
        
        PrintValue(11) = " "
        FormatString(11) = "a0"
        
        PrintValue(12) = rrs!CWTTax
        FormatString(12) = "d6"

        PrintValue(13) = TotTaxes
        FormatString(13) = "d6"
                    
        PrintValue(14) = " "
        FormatString(14) = "a0"
        
        PrintValue(15) = rrs!RegHours
        FormatString(15) = "d6"
        
        PrintValue(16) = " "
        FormatString(16) = "a0"
        
        PrintValue(17) = rrs!OTHours
        FormatString(17) = "d6"
        
        PrintValue(18) = " "
        FormatString(18) = "a0"
        
        PrintValue(19) = rrs!OEHours
        FormatString(19) = "d6"
        
        PrintValue(20) = " "
        FormatString(20) = "a0"
        
        PrintValue(21) = Format(TOTHours, "##0.00")
        FormatString(21) = "d6"
             
        PrintValue(22) = " "
        FormatString(22) = "~"
        
        FormatPrint
        Ln = Ln + 1

        ChkRegOEDed                 ''  ###   CHKREGOEDED
        
        PrintValue(1) = "______________________________________________________________________________________________________________________________________________________"
        FormatString(1) = "a154"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1

        LastEmpNumber = rrs!EmployeeNumber
        LastEmpLName = rrs!LastName
        LastEmpFName = rrs!FirstName
                  
        rrs.MoveNext
        
        If rrs.EOF Then
            trsDEDTot.Sort = "DeptID, Type,ItemID"
            ChkRegTotals
            PRTotal.Clear
            Exit Do
        End If
        
    Loop
'    trs.Close
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================   CHECK REGISTER - 1st header     ======================================

Public Sub ChkRegFirstHeader()
    Ln = Ln + 1
    PrintValue(1) = "CHK #"
    FormatString(1) = "a5"
    
    PrintValue(2) = " "
    FormatString(2) = "a2"
    
    PrintValue(3) = "DPT"
    FormatString(3) = "a4"
   
    PrintValue(4) = " "
    FormatString(4) = "a2"
    
    PrintValue(5) = "EMPLOYEE #/NAME"
    FormatString(5) = "a30"
    
    PrintValue(6) = " "
    FormatString(6) = "a1"
    
    PrintValue(7) = "REG PAY"
    FormatString(7) = "a7"
    
    PrintValue(8) = " "
    FormatString(8) = "a8"
    
    PrintValue(9) = "OT PAY"
    FormatString(9) = "a6"
    
    PrintValue(10) = " "
    FormatString(10) = "a7"
 
    PrintValue(11) = "OTH PAY"
    FormatString(11) = "a7"
    
    PrintValue(12) = " "
    FormatString(12) = "a7"
    
    PrintValue(13) = "GROSS PAY"
    FormatString(13) = "a9"
    
    PrintValue(14) = " "
    FormatString(14) = "a7"
    
    PrintValue(15) = "TOT DED"
    FormatString(15) = "a7"
    
    PrintValue(16) = " "
    FormatString(16) = "a7"
    
    PrintValue(17) = "NET PAY"
    FormatString(17) = "a7"
                                
    PrintValue(18) = " "
    FormatString(18) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " "
    FormatString(1) = "a0"
    
    PrintValue(2) = "PE DATE"
    FormatString(2) = "a8"
 
    PrintValue(3) = " "
    FormatString(3) = "a9"
    
    PrintValue(4) = "SS TAX"
    FormatString(4) = "a9"
    
    PrintValue(5) = " "
    FormatString(5) = "a4"
    
    PrintValue(6) = "MED TAX"
    FormatString(6) = "a9"
    
    PrintValue(7) = " "
    FormatString(7) = "a5"
    
    PrintValue(8) = "FWT TAX"
    FormatString(8) = "a9"
    
    PrintValue(9) = " "
    FormatString(9) = "a5"
    
    PrintValue(10) = "SWT TAX"
    FormatString(10) = "a9"
    
    PrintValue(11) = " "
    FormatString(11) = "a5"
    
    PrintValue(12) = "CWT TAX"
    FormatString(12) = "a9"
    
    PrintValue(13) = " "
    FormatString(13) = "a5"
 
    PrintValue(14) = "TOT TAXES"
    FormatString(14) = "a9"
    
    PrintValue(15) = " "
    FormatString(15) = "a7"
    
    PrintValue(16) = "REG HRS"
    FormatString(16) = "a7"
    
    PrintValue(17) = " "
    FormatString(17) = "a8"
    
    PrintValue(18) = "OT HRS"
    FormatString(18) = "a12"
    
    PrintValue(19) = " "
    FormatString(19) = "a1"
    
    PrintValue(20) = "OTH HRS"
    FormatString(20) = "a9"
    
    PrintValue(21) = " "
    FormatString(21) = "a5"
    
    PrintValue(22) = "TOT HRS"
    FormatString(22) = "a11"
    
    PrintValue(23) = " "
    FormatString(23) = "~"
    
    FormatPrint
    Ln = Ln + 1

    If frmCheckReg.chkOEHrs Or frmCheckReg.chkOEAmt Or frmCheckReg.chkDed Then
        ChkRegDedHeader
    End If
    
    PrintValue(1) = "______________________________________________________________________________________________________________________________________________________"
    FormatString(1) = "a154"
    
    PrintValue(2) = " "
    FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
End Sub

'=======================================   CHECK REGISTER - Get Header Data    ======================================

Public Sub ChkRegGetHeaderData()
    OEHours = 1
    OEAmount = 2
    DEDAmount = 3
    
    trsDED.CursorLocation = adUseClient

    trsDED.Fields.Append "Type", adInteger
    trsDED.Fields.Append "abbreviation", adVarChar, 13, adFldIsNullable
    trsDED.Fields.Append "Amount", adCurrency
    trsDED.Fields.Append "ItemID", adDouble
    trsDED.Fields.Append "LineNo", adInteger
    
    trsDED.Open , , adOpenDynamic, adLockOptimistic
    
    trsDEDTot.CursorLocation = adUseClient

    trsDEDTot.Fields.Append "Type", adInteger                   '  OEHours, OEAmount, DEDAmount
    trsDEDTot.Fields.Append "Amount", adCurrency
    trsDEDTot.Fields.Append "ItemID", adDouble
    trsDEDTot.Fields.Append "DeptID", adDouble
    trsDEDTot.Fields.Append "LineNo", adInteger
    
    trsDEDTot.Open , , adOpenDynamic, adLockOptimistic
    
    ' get the employer items
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeOE & " OR " & _
                " PRItem.ItemType = " & PREquate.ItemTypeDED & ")"
                
    If PRItem.GetBySQL(SQLString) Then
        Do
            trsDED.AddNew
            trsDEDTot.AddNew
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                trsDED!Type = OEAmount
                trsDED!LineNo = 2
                trsDEDTot!Type = OEAmount
                trsDEDTot!LineNo = 2
            Else
                trsDED!Type = DEDAmount
                trsDED!LineNo = 3
                trsDEDTot!Type = DEDAmount
                trsDEDTot!LineNo = 3
            End If

            trsDED!Abbreviation = PRItem.Abbreviation
            trsDED!ItemID = PRItem.ItemID
            trsDED.Update
            trsDEDTot!ItemID = PRItem.ItemID
            trsDEDTot.UpdateBatch
            
            ' add hour slot for EVERY other earning
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                trsDED.AddNew
                trsDEDTot.AddNew
                trsDED!Type = OEHours
                trsDED!LineNo = 1
                trsDED!Abbreviation = Trim(PRItem.Abbreviation) & " HR"
                trsDED!ItemID = PRItem.ItemID
                trsDED.Update
                trsDEDTot!Type = OEHours
                trsDEDTot!LineNo = 1
                trsDEDTot!ItemID = PRItem.ItemID
                trsDEDTot.UpdateBatch
            End If
                        
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If

    trsDED.Sort = "Type,ItemID"
    trsDEDTot.Sort = "Type,ItemID"

End Sub

'=======================================   CHECK REGISTER - Deduction Header     ======================================

Public Sub ChkRegDedHeader()
Dim LastAbbrev As String
    LastType = 0
    Colcount = 0
    PrtString = " "
    FormatString(1) = " "

    trsDED.MoveFirst
    
    Do
        '  PRINT HEADER - change in type or max number of columns
        Colcount = Colcount + 1
        If (LastType <> 0 And LastType <> trsDED!Type) Or Colcount = 12 Then
            PrtString = Space(15) & FormatString(Colcount)

            PrintValue(1) = PrtString
            FormatString(1) = "a200"

            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1

            Colcount = 1
            FormatString(1) = " "
            FormatString(Colcount) = " "

        End If

        LastType = trsDED!Type
'        Colcount = Colcount + 1
        
        FormatString(Colcount + 1) = FormatString(Colcount) & trsDED!Abbreviation & Space(5)
        LastAbbrev = trsDED!Abbreviation
        trsDED.MoveNext

        If trsDED.EOF Then
            Exit Do
        End If
    Loop
    
    If trsDED.EOF Then
        If Colcount > 0 Then
            PrtString = Space(15) & FormatString(Colcount) & LastAbbrev
            PrintValue(1) = PrtString
            FormatString(1) = "a200"

            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1
                
        End If
    End If

    
End Sub

'=======================================   CHECK REGISTER - OE/Deductions     ======================================

Public Sub ChkRegOEDed()
Dim PrtString As String
Dim LastCheck As Long
Dim LastEmployee As Long

WriteCt = 0

    SQLString = "SELECT * FROM PRHist WHERE " & Trim(DedString) & _
                " AND PRHist.histid = " & rrs!HistID

    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History!!!"
        End
    End If

    Do
        ' clear out temp record set
        trsDED.MoveFirst
        Do
            trsDED!Amount = 0
            trsDED.Update
            trsDED.MoveNext
            If trsDED.EOF Then Exit Do
        Loop

    ''''''''''''''''''''''''''''''     PROCESS OTHER EARNINGS   '''''''''''''''''''''''''''''''''''
        
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID & _
            " AND PRDist.DistType = " & PREquate.DistTypeItem

        If PRDist.GetBySQL(SQLString) Then
            Do
                If frmCheckReg.chkOEHrs Then
                    OtherEarningsHours
                End If

                If frmCheckReg.chkOEAmt Then
                    OtherEarningsAmount
                End If

            If Not PRDist.GetNext Then Exit Do
            Loop
        End If


    '''''''''''''''''''''''''''''''''''     PROCESS DEDUCTIONS     '''''''''''''''''''''''''''''''''''''''
        If frmCheckReg.chkDed Then
            ProcessDeductions
        End If
    '''''''''''''''''''''''''''''''''''        WRITE RECORD     '''''''''''''''''''''''''''''''''''''''
        
        ChkRegWrite

        LastCheck = PRHist.CheckNumber
        LastEmployee = PRDist.EmployeeID

        If Not PRHist.GetNext Then Exit Do

    Loop
    
End Sub

'=======================================   CHECK REGISTER - Other Earnings Hours    ======================================

Public Sub OtherEarningsHours()
    If PRDist.Hours <> 0 Then
        Flag = False
        trsDED.MoveFirst
        trsDEDTot.MoveFirst
        Do
            If trsDED!ItemID = PRDist.EmployerItemID Then
                trsDED!Amount = PRDist.Hours
                trsDEDTot!Amount = trsDEDTot!Amount + PRDist.Hours
                trsDEDTot!ItemID = trsDED!ItemID
                trsDEDTot!DeptID = PRDist.DepartmentID
                trsDEDTot!Type = trsDED!Type
                trsDEDTot.UpdateBatch
                Flag = True
                Exit Do
            End If
            trsDED.MoveNext
            trsDEDTot.MoveNext
            If trsDED.EOF Then Exit Do
        Loop
        If Not Flag Then
           MsgBox "Employer Item Not Found: " & PRHist.EmployeeID & vbCr & PRDist.DistID, vbCritical
        End If
    End If
        
End Sub

'=======================================   CHECK REGISTER - Other Earnings Amount    ======================================

Public Sub OtherEarningsAmount()
    If PRDist.Amount <> 0 Then
        Flag = False
        trsDED.MoveFirst
        trsDEDTot.MoveFirst

        Do
            If trsDED!ItemID = PRDist.EmployerItemID Then
                trsDED!Amount = PRDist.Amount
                trsDEDTot!Amount = trsDEDTot!Amount + PRDist.Amount
                trsDEDTot!ItemID = trsDED!ItemID
                trsDEDTot!DeptID = PRDist.DepartmentID
                trsDEDTot!Type = trsDED!Type
                trsDEDTot.UpdateBatch
                Flag = True
                Exit Do
            End If
            trsDED.MoveNext
            trsDEDTot.MoveNext
            If trsDED.EOF Then Exit Do
        Loop
        If Not Flag Then
            MsgBox "Employer Item Not Found: " & PRHist.EmployeeID & vbCr & PRDist.DistID, vbCritical
            End
        End If
    End If
End Sub

'=======================================   CHECK REGISTER - Process Deductions    ======================================

Public Sub ProcessDeductions()
    SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID & _
            " AND PRItemHist.ItemType <> " & PREquate.ItemTypeDirDepDed

    If PRItemHist.GetBySQL(SQLString) Then
        Do
            If PRItemHist.Amount <> 0 Then
                Flag = False
                trsDED.MoveFirst
                trsDEDTot.MoveFirst
                Do

                    If trsDED!ItemID = PRItemHist.EmployerItemID Then
                        trsDED!Amount = PRItemHist.Amount
                        trsDEDTot!Amount = trsDEDTot!Amount + PRItemHist.Amount
                        trsDEDTot!ItemID = trsDED!ItemID
                        trsDEDTot!DeptID = PRDist.DepartmentID
                        trsDEDTot!Type = trsDED!Type
                        trsDEDTot.UpdateBatch
                        Flag = True
                        Exit Do
                    End If
                    trsDED.MoveNext
                    trsDEDTot.MoveNext
                    If trsDED.EOF Then Exit Do
                Loop
                    If Not Flag Then
                        MsgBox "Employer Item Not Found: " & PRItemHist.EmployeeID & vbCr & PRItemHist.ItemHistID, vbCritical
                        End
                    End If
            End If
            If Not PRItemHist.GetNext Then Exit Do
        Loop
    End If
End Sub

'=======================================   CHECK REGISTER - Write   ======================================

Public Sub ChkRegWrite()
Dim SpaceNumber As Long
    
    LastType = 0
    Colcount = 0
    PrtString = " "
    FormatString(1) = ""
    SpaceNumber = 0
    OEHrsPrt = 0
    OEAmtPrt = 0
    DEDAmtPrt = 0
    trsDED.MoveFirst
    
    ' change in type or max number of columns
    Do
        If (LastType <> 0 And LastType <> trsDED!Type) Or Colcount = 12 Then
            If OEHrsPrt = 1 Or OEAmtPrt = 1 Or DEDAmtPrt = 1 Then

                PrtString = FormatString(Colcount)
    
                PrintValue(1) = PrtString
                FormatString(1) = "a200"
                
                PrintValue(2) = " "
                FormatString(2) = "~"
    
                FormatPrint
                Ln = Ln + 1
    
                Colcount = 0
                FormatString(1) = ""
                FormatString(Colcount) = ""
                
                WriteCt = 0
            End If
        End If
        
        LastType = trsDED!Type
        Colcount = Colcount + 1
        
        If Colcount = 1 Then
            If frmCheckReg.chkOEHrs And trsDED!Amount > 0 Then
                FormatString(Colcount + 1) = FormatString(Colcount) & "HOURS:     " & Space(6) & _
                                                Format(trsDED!Amount, "##,###,##0.00") & Space(11)
                OEHrsPrt = 1
            ElseIf frmCheckReg.chkOEAmt And trsDED!Amount > 0 Then
                FormatString(Colcount + 1) = FormatString(Colcount) & "OTHER EARN:" & Space(6) & _
                                                Format(trsDED!Amount, "##,###,##0.00") & Space(11)
                OEAmtPrt = 1
            ElseIf frmCheckReg.chkDed And trsDED!Amount > 0 Then
                FormatString(Colcount + 1) = FormatString(Colcount) & "DEDUCTIONS:" & Space(6) & _
                                                Format(trsDED!Amount, "##,###,##0.00") & Space(11)
                DEDAmtPrt = 1
            End If
        Else
            FormatString(Colcount + 1) = FormatString(Colcount) & Format(trsDED!Amount, "##,###,##0.00") & Space(11)
        End If

        trsDED.MoveNext

        If trsDED.EOF Then
            Exit Do
        End If
    Loop
    
    If frmCheckReg.chkOEHrs And OEHrsPrt = 1 Then
        PrtString = FormatString(Colcount)
        PrintValue(1) = PrtString
        FormatString(1) = "a200"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
    End If
    If frmCheckReg.chkOEAmt And OEAmtPrt = 1 Then
        PrtString = FormatString(Colcount)
        PrintValue(1) = PrtString
        FormatString(1) = "a200"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
    
        FormatPrint
        Ln = Ln + 1
    End If
    If frmCheckReg.chkDed And DEDAmtPrt = 1 Then
        PrtString = FormatString(Colcount)
        PrintValue(1) = PrtString
        FormatString(1) = "a200"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
    
        FormatPrint
        Ln = Ln + 1
    End If
    
End Sub

'=======================================   CHECK REGISTER - Totals   ======================================
Public Sub ChkRegTotals()

    ' Print Report Totals
    PRTotal.TSort       ' Sort is by "RecType" IDNumber
    PRTotal.FindFirst

    Do
        If Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            ChkRegFirstHeader
            FormatPrint
            Ln = Ln + 1
        End If

        TTotTaxes = PRTotal.SSTax + PRTotal.MedTax + PRTotal.StateTax + PRTotal.FWTTax + PRTotal.CityTax
        
        If PRTotal.IDNumber = 0 Then
            GoTo NoBreak
        End If
        
        If PRTotal.IDNumber = 999999998 Then
            
            PrintValue(2) = " "
            FormatString(2) = "~"
            
            FormatPrint
            Ln = Ln + 1
        
            If frmCheckReg.optSubTotal = True Then

                If RecCnt > 1 Then
                    GetEmpInfo (PREmployee.EmployeeID)
                    PrintValue(1) = "Emp Total " & LastEmpNumber & " - " & Trim(LastEmpLName) & ", " & Trim(LastEmpFName)
                Else
                    GoTo NoBreak
                End If
            Else
                GoTo NoBreak
            End If
            
        ElseIf TotalFlag = True Then
            If PRTotal.IDNumber = 999999999 Then
                PrintValue(1) = "GRAND TOTAL "
            ElseIf PRTotal.IDNumber <= 99 Then
                If Not PRDepartment.GetByID(PRTotal.DepartmentID) Then
                    MsgBox "Department Info Not Found!!!", vbCritical, "Check Register"
                    End
                End If
                
                PrintValue(1) = "DEPT " & PRTotal.IDNumber & " - " & Mid(PRDepartment.Name, 1, 8)
            End If
        Else
            GoTo NoBreak
        End If
        
        If PRTotal.IDNumber = 0 Then
            GoTo NoBreak
        End If
        
            FormatString(1) = "a38"
                       
            PrintValue(2) = " "
            FormatString(2) = "a0"
            
            PrintValue(3) = Format(PRTotal.RegAmount, "##,###,##0.00")
            FormatString(3) = "d13"
            
            PrintValue(4) = " "
            FormatString(4) = "a0"
            
            PrintValue(5) = Format(PRTotal.OTAmount, "##,###,##0.00")
            FormatString(5) = "d13"
            
            PrintValue(6) = " "
            FormatString(6) = "a0"
            
            PrintValue(7) = Format(PRTotal.OEAmount, "##,###,##0.00")
            FormatString(7) = "d13"
    
            PrintValue(8) = " "
            FormatString(8) = "a0"
            
            PrintValue(9) = Format(PRTotal.Gross, "##,###,##0.00")
            FormatString(9) = "d13"
                        
            PrintValue(10) = " "
            FormatString(10) = "a0"
            
'            PrintValue(11) = Format(PRTotal.Deductions, "##,###,##0.00")
'            FormatString(11) = "d13"
            
            PrintValue(11) = " "
            FormatString(11) = "a0"
             
            PrintValue(12) = Format(PRTotal.Net, "##,###,##0.00")
            FormatString(12) = "d13"
            
            PrintValue(13) = " "
            FormatString(13) = "~"
            
            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = " "
            FormatString(1) = "a10"
            
            PrintValue(2) = Format(PRTotal.SSTax, "##,###,##0.00")
            FormatString(2) = "d13"
                          
            PrintValue(3) = " "
            FormatString(3) = "a0"
            
            PrintValue(4) = Format(PRTotal.MedTax, "##,###,##0.00")
            FormatString(4) = "d13"
                                        
            PrintValue(5) = " "
            FormatString(5) = "a0"
            
            PrintValue(6) = Format(PRTotal.FWTTax, "##,###,##0.00")
            FormatString(6) = "d13"
            
            PrintValue(7) = " "
            FormatString(7) = "a0"
        
            PrintValue(8) = Format(PRTotal.StateTax, "##,###,##0.00")
            FormatString(8) = "d13"
                          
            PrintValue(9) = " "
            FormatString(9) = "a0"
        
            PrintValue(10) = Format(PRTotal.CityTax, "##,###,##0.00")
            FormatString(10) = "d13"
                                        
            PrintValue(11) = " "
            FormatString(11) = "a0"
            
            PrintValue(12) = Format(TTotTaxes, "###,###,##0.00")
            FormatString(12) = "d14"
            
            PrintValue(13) = " "
            FormatString(13) = "a0"
            
            PrintValue(14) = Format(PRTotal.RegHours, "#,###,##0.00")
            FormatString(14) = "d9"
            
            PrintValue(15) = " "
            FormatString(15) = "a0"
            
            PrintValue(16) = Format(PRTotal.OTHours, "#,###,##0.00")
            FormatString(16) = "d9"
                          
            PrintValue(17) = " "
            FormatString(17) = "a0"
            
            PrintValue(18) = Format(PRTotal.OEHours, "#,###,##0.00")
            FormatString(18) = "d9"
                                        
            PrintValue(19) = " "
            FormatString(19) = "a0"
         
            TTotHours = PRTotal.RegHours + PRTotal.OTHours + PRTotal.OEHours
            PrintValue(20) = Format(TTotHours, "##,###,##0.00")
            FormatString(20) = "d9"
            
            PrintValue(21) = " "
            FormatString(21) = "~"
            
            FormatPrint
            Ln = Ln + 1
            
            FindStr = "DeptID=" & CStr(PRTotal.DepartmentID)
            
            trsDEDTot.Find FindStr, 0, adSearchBackward, 1
'MsgBox trsDEDTot!Type & "  " & trsDEDTot!Amount & "  " & trsDEDTot!ItemID & vbCr & trsDEDTot!DeptID & "  " & trsDEDTot!LineNo

            If trsDEDTot.EOF Then
                GoTo NoBreak
            Else
                ChkRegOEDEDTotals
            End If
NoBreak:

            PrintValue(1) = "__________________________________________________________________________________________________________________________________________________________"
            FormatString(1) = "a157"
                 
            PrintValue(2) = " "
            FormatString(2) = "~"
                 
            FormatPrint
            Ln = Ln + 2
            
            PRTotal.Clear

            If EmpFlag = True Then
                Exit Sub
            End If
            
            If Not PRTotal.GetNext Then
                PRTotal.Clear
                Exit Do
            End If
        Loop

End Sub

'=======================================   CHECK REGISTER - OE/DED TOtals   ======================================
Public Sub ChkRegOEDEDTotals()

    LastType = 0
    Colcount = 0
    PrtString = " "
    FormatString(1) = ""
    SpaceNumber = 0
    trsDEDTot.MoveFirst
    ' change in type or max number of columns

        Do While trsDEDTot!DeptID = PRTotal.DepartmentID
            Colcount = Colcount + 1
            If Colcount = 1 Then
                If frmCheckReg.chkOEHrs And trsDEDTot!Type = OEHours Then
                    FormatString(Colcount + 1) = FormatString(Colcount) & "HOURS:     " & Space(6) & _
                                                    Format(trsDEDTot!Amount, "##,###,##0.00") & Space(11)
                ElseIf frmCheckReg.chkOEAmt And trsDEDTot!Type = OEAmount Then
                    FormatString(Colcount + 1) = FormatString(Colcount) & "OTHER EARN:" & Space(6) & _
                                                    Format(trsDEDTot!Amount, "##,###,##0.00") & Space(11)
                ElseIf frmCheckReg.chkDed And trsDEDTot!Type = DEDAmount Then
                    FormatString(Colcount + 1) = FormatString(Colcount) & "DEDUCTIONS:" & Space(6) & _
                                                    Format(trsDEDTot!Amount, "##,###,##0.00") & Space(11)
                End If
            Else
                FormatString(Colcount + 1) = FormatString(Colcount) & Format(trsDEDTot!Amount, "##,###,##0.00") & Space(11)
            End If
    
            trsDEDTot.MoveNext
    
            If trsDEDTot.EOF Then
                Exit Sub
            End If
        Loop
        If Trim(FormatString(Colcount)) <> "" Then
            PrtString = FormatString(Colcount)
            PrintValue(1) = PrtString
            FormatString(1) = "a200"
            
            PrintValue(2) = " "
            FormatString(2) = "~"
        
            FormatPrint
            Ln = Ln + 1
        End If

End Sub


Public Sub GetEmpInfo(ByVal EmpNo As Long)      '=========   GET EMPLOYEE INFO   =================

    If Not PREmployee.GetBySQL("SELECT * FROM PREmployee WHERE PREmployee.EmployeeNumber = " & rrs!EmployeeNumber) Then
        PREmployee.LastName = "None"
    End If

End Sub

'=======================================   CHECK REGISTER - Update Totals    ======================================

Private Sub UpdateTotals(ByVal RecType As Byte, ByVal RecId As Long, ByVal IDNumber As Long, ByVal Dept As Long)
    If Not PRTotal.tFind(RecType, RecId) Then
'        GetDeptByID rrs![PRemployee.DepartmentID]
        PRTotal.RecType = RecType
        PRTotal.RecId = RecId
        PRTotal.IDNumber = IDNumber
        PRTotal.RegHours = 0
        PRTotal.RegAmount = 0
        PRTotal.OTHours = 0
        PRTotal.OTAmount = 0
        PRTotal.OEHours = 0
        PRTotal.OEAmount = 0
        PRTotal.SSWage = 0
        PRTotal.SSTax = 0
        PRTotal.MEDWage = 0
        PRTotal.MedTax = 0
        PRTotal.FWTWage = 0
        PRTotal.FWTTax = 0
        PRTotal.StateWage = 0
        PRTotal.StateTax = 0
        PRTotal.CityWage = 0
        PRTotal.CityTax = 0
        PRTotal.Gross = 0
        PRTotal.Net = 0
        PRTotal.Save (Equate.RecAdd)
    End If

    PRTotal.DepartmentID = rrs![PREmployee.DepartmentID]
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
    PRTotal.StateWage = PRTotal.StateWage + rrs!SWTWage
    PRTotal.StateTax = PRTotal.StateTax + rrs!SWTTax
    PRTotal.CityWage = PRTotal.CityWage + rrs!CWTWage
    PRTotal.CityTax = PRTotal.CityTax + rrs!CWTTax
    PRTotal.Gross = PRTotal.Gross + rrs!Gross
    PRTotal.Net = PRTotal.Net + rrs!Net
    
    PRTotal.Save (Equate.RecPut)
End Sub

Private Sub TSort()         '=========   Total Sort   =================
    PRTotal.TSort
End Sub

'=======================================   DEPOSIT LISTING   ======================================

Public Sub DepositListing(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal CheckDt As Long, _
                         ByVal Startdate As Long, _
                         ByVal EndDate As Long)
Dim ReportTitle As String

    frmCheckReg.Hide
    ReportTitle = "PAYROLL EMPLOYEE DEPOSIT LISTING"

    Msg2 = "PERIOD ENDING DATE FROM: " & StartPEDate & " TO: " & EndPEDate
    msg3 = "CHECK DATE: " & TDBChkDate
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    SetEquates
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    Ln = Ln + 4
    
    PageHeader ReportTitle, Msg1, Msg2, msg3

    Ln = Ln + 3
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "FICA WITHHELD"
    FormatString(2) = "a41"
    
    PrintValue(3) = ""
    FormatString(3) = "a15"
                
    PrintValue(4) = DepFICAWH
    FormatString(4) = "d10"
    
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "MED WITHHELD"
    FormatString(2) = "a41"
    
    PrintValue(3) = ""
    FormatString(3) = "a15"

    PrintValue(4) = DepMEDWH
    FormatString(4) = "d10"
    
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "FICA MATCH"
    FormatString(2) = "a20"
    
    PrintValue(3) = ""
    FormatString(3) = "a2"
        
    PrintValue(4) = DepFICAAmt
    FormatString(4) = "d7"
    
    PrintValue(5) = ""
    FormatString(5) = "a1"
        
    PrintValue(6) = "x"
    FormatString(6) = "a1"
            
    PrintValue(7) = ""
    FormatString(7) = "a0"

    PrintValue(8) = DepFICAPct
    FormatString(8) = "d5"
    
    PrintValue(9) = "%"
    FormatString(9) = "a1"
    
    PrintValue(10) = ""
    FormatString(10) = "a3"
    
    PrintValue(11) = DepFICAMatch
    FormatString(11) = "d10"
                
    PrintValue(12) = " "
    FormatString(12) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " "
    FormatString(1) = "a5"

    PrintValue(2) = "MED MATCH"
    FormatString(2) = "a20"
        
    PrintValue(3) = ""
    FormatString(3) = "a2"
        
    PrintValue(4) = DepMedAmt
    FormatString(4) = "d7"
    
    PrintValue(5) = ""
    FormatString(5) = "a1"
        
    PrintValue(6) = "x"
    FormatString(6) = "a1"
            
    PrintValue(7) = ""
    FormatString(7) = "a0"

    PrintValue(8) = DepMedPct
    FormatString(8) = "d5"
    
    PrintValue(9) = "%"
    FormatString(9) = "a1"
    
    PrintValue(10) = ""
    FormatString(10) = "a3"
    
    PrintValue(11) = DepMEDMatch
    FormatString(11) = "d10"
                
    PrintValue(12) = " "
    FormatString(12) = "~"
    
    PrintValue(13) = " "
    FormatString(13) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " "
    FormatString(1) = "a5"
    
    PrintValue(2) = "FEDERAL TAX WITHHELD"
    FormatString(2) = "a41"
        
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepFedTaxWH
    FormatString(4) = "d10"
        
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "STATE TAX WITHHELD"
    FormatString(2) = "a41"
        
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepSTTaxWH
    FormatString(4) = "d10"
        
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " "
    FormatString(1) = "a5"

    PrintValue(2) = "CITY TAX WITHHELD"
    FormatString(2) = "a41"
                                            
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepCityTaxWH
    FormatString(4) = "d10"
    
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " "
    FormatString(1) = "a5"
    
    PrintValue(2) = "STATE UNEMPLOYMENT"
    FormatString(2) = "a20"
                                                
    PrintValue(3) = ""
    FormatString(3) = "a2"
        
    PrintValue(4) = DepSTUnempAmt
    FormatString(4) = "d7"
    
    PrintValue(5) = ""
    FormatString(5) = "a1"
        
    PrintValue(6) = "x"
    FormatString(6) = "a1"
            
    PrintValue(7) = ""
    FormatString(7) = "a0"

    PrintValue(8) = DepSTUnempPct
    FormatString(8) = "d5"
    
    PrintValue(9) = "%"
    FormatString(9) = "a1"
    
    PrintValue(10) = ""
    FormatString(10) = "a3"
    
    PrintValue(11) = DepSTUnempMatch
    FormatString(11) = "d10"
                
    PrintValue(12) = " "
    FormatString(12) = "~"
    
    PrintValue(13) = " "
    FormatString(13) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " "
    FormatString(1) = "a5"
    
    PrintValue(2) = "FEDERAL UNEMPLOYMENT"
    FormatString(2) = "a20"
        
    PrintValue(3) = ""
    FormatString(3) = "a2"
        
    PrintValue(4) = DepFedUnempAmt
    FormatString(4) = "d7"
    
    PrintValue(5) = ""
    FormatString(5) = "a1"
        
    PrintValue(6) = "x"
    FormatString(6) = "a1"
            
    PrintValue(7) = ""
    FormatString(7) = "a0"

    PrintValue(8) = DepFedUnempPct
    FormatString(8) = "d5"
    
    PrintValue(9) = "%"
    FormatString(9) = "a1"
    
    PrintValue(10) = ""
    FormatString(10) = "a3"
    
    PrintValue(11) = DepFedUnempMatch
    FormatString(11) = "d10"
                
    PrintValue(12) = " "
    FormatString(12) = "~"
    
    FormatPrint
    Ln = Ln + 4
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "FEDERAL DEPOSIT"
    FormatString(2) = "a41"
        
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepFedDep
    FormatString(4) = "d10"
        
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "TOTAL TAXES AND DEDUCTIONS TO BE ESCROWED"
    FormatString(2) = "a41"
        
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepTotEscrowed
    FormatString(4) = "d10"
        
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = ""
    FormatString(1) = "a67"
    
    PrintValue(2) = "_______"
    FormatString(2) = "a10"
        
    PrintValue(3) = " "
    FormatString(3) = "~"
    
    FormatPrint
    Ln = Ln + 4
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "NET PAYROLL CHECK AMOUNT"
    FormatString(2) = "a41"
        
    PrintValue(3) = ""
    FormatString(3) = "a15"
    
    PrintValue(4) = DepNetAmt
    FormatString(4) = "d10"
    
    PrintValue(5) = " "
    FormatString(5) = "~"
    
    FormatPrint
    Ln = Ln + 2
                
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "TOTAL GROSS PAY:"
    FormatString(2) = "a16"
        
    PrintValue(3) = TotalGrossPay
    FormatString(3) = "n3"
        
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = ""
    FormatString(1) = "a5"
    
    PrintValue(2) = "NUMBER OF RECORDS:"
    FormatString(2) = "a18"
        
    PrintValue(3) = NumberRecords
    FormatString(3) = "n3"
        
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================   DIRECT DEPOSIT REPORT   ======================================

Public Sub DirectDepositRpt(ByVal RangeType As Byte, _
                            ByVal BatchNumbr As Long, _
                            ByVal PEDate As Long, _
                            ByVal CheckDt As Long, _
                            ByVal Startdate As Long, _
                            ByVal EndDate As Long)
                            
Dim ReportTitle, LastABA As String
Dim LastEmployee As Long

    SetEquates
    
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "Batch", adInteger
    trs.Fields.Append "EmployeeNo", adInteger
    trs.Fields.Append "Routing", adVarChar, 14, adFldIsNullable
    trs.Fields.Append "AcctNo", adVarChar, 14, adFldIsNullable
    trs.Fields.Append "Name", adVarChar, 50, adFldIsNullable
    trs.Fields.Append "AcctType", adVarChar, 3, adFldIsNullable
'    trs.Fields.Append "Debit", adCurrency
    trs.Fields.Append "Credit", adCurrency
    trs.Fields.Append "BankName", adVarChar, 20, adFldIsNullable
    
    trs.Open , , adOpenDynamic, adLockOptimistic

    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 10, Equate.Portrait
    Ln = 0
        
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = "SELECT * FROM PRBatch WHERE PRBatch.BatchID = " & BatchNumbr
    Else
        SQLString = "SELECT * FROM PRBatch WHERE PRBatch.PEDate >= " & Startdate & _
        " AND PRBatch.PEDate <= " & EndDate
    End If

    rsInit SQLString, cn, rs
    
    If Not PRBatch.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Direct Deposit Report"
        Exit Sub
    End If
    
    Do
        SQLString = "SELECT * FROM PRItemHist " & _
                    "WHERE PRItemHist.batchid = " & PRBatch.BatchID & _
                    " AND PRItemhist.itemtype = " & PREquate.ItemTypeDirDepDed
                
        If PRItemHist.GetBySQL(SQLString) Then

            Do
                trs.AddNew
                trs!Batch = PRBatch.BatchID
                PREmployee.GetByID (PRItemHist.EmployeeID)   '  Get Employee Info
                trs!Employeeno = PREmployee.EmployeeNumber
                trs!Name = Trim(PREmployee.LastName) & ", " & Trim(PREmployee.FirstName)
                
                PRItem.GetByID (PRItemHist.ItemID)
                If Trim(PRItem.DirDepABA) = "" And PRItemHist.EmployeeID = LastEmployee Then
                    trs!Routing = LastABA
                Else
                    trs!Routing = Trim(PRItem.DirDepABA)
                End If
                trs!AcctNo = Trim(PRItem.DirDepAccount)
                If PRItem.DirDepType = PREquate.DirDepTypeChecking Then
                    trs!AcctType = "CHK"
                Else
                    trs!AcctType = "SVE"
                End If
                trs!Credit = PRItemHist.Amount
                trs!BankName = Trim(PRItem.DirDepBank)
                trs.Update
                LastEmployee = PRItemHist.EmployeeID
                LastABA = PRItem.DirDepABA
                If Not PRItemHist.GetNext Then Exit Do
            Loop
        End If

        If Not PRBatch.GetNext Then Exit Do
    Loop

    
'   Open output file and write first two lines
    If frmDirectDep.chkOutputFile Then
        WriteOneAndFive
    End If
    
'   Sort temporary recordset according to user sort selection

    If OrderType = 1 Then
        trs.Sort = "Batch,EmployeeNo"
    Else
        trs.Sort = "Batch,Name"
    End If
    
    ReportTitle = "CENTRALIZED DIRECT DEPOSIT REPORT"
    Msg1 = "LIVE RECORDS TRANSMITTED - CHECK DATE: " & PRBatch.CheckDate
    
    If RangeType = PREquate.RangeTypeBatch Then
        Msg2 = "BATCH " & BatchNumbr & " - Period Ending: " & CDate(PEDate)
    Else
        Msg2 = "PERIOD ENDING DATE FROM: " & CDate(Startdate) & " TO: " & CDate(EndDate)
    End If
    
    LineCt = 0
    trs.MoveFirst
    
    Do
        If trs.EOF = True Then
            Exit Do
        End If
        
        If Ln = 0 Or Ln > MaxLines Or trs!Batch <> LastBatch Then
            If trs!Batch <> LastBatch And LastBatch <> 0 Then
                PrintSubTotals
            End If
            
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, Msg2, ""
            SetFont 8, Equate.Portrait
            Ln = Ln + 2
        
            PrintValue(1) = "EMPLOYEE"
            FormatString(1) = "a8"
        
            PrintValue(2) = " "
            FormatString(2) = "a3"
            
            PrintValue(3) = "ROUTING"
            FormatString(3) = "a7"
            
            PrintValue(4) = ""
            FormatString(4) = "a5"
            
            PrintValue(5) = "ACCT"
            FormatString(5) = "a14"
            
            PrintValue(6) = " "
            FormatString(6) = "a31"
            
            PrintValue(7) = "ACCT"
            FormatString(7) = "a4"
            
            PrintValue(8) = ""
            FormatString(8) = "~"
            
            FormatPrint
            Ln = Ln + 1
                
            PrintValue(1) = " NUMBER"
            FormatString(1) = "a7"
        
            PrintValue(2) = " "
            FormatString(2) = "a2"
            
            PrintValue(3) = "AND TRANSIT"
            FormatString(3) = "a11"
            
            PrintValue(4) = ""
            FormatString(4) = "a2"
            
            PrintValue(5) = "NUMBER"
            FormatString(5) = "a14"
                
            PrintValue(6) = " "
            FormatString(6) = "a1"
            
            PrintValue(7) = "NAME"
            FormatString(7) = "a25"
                
            PrintValue(8) = " "
            FormatString(8) = "a6"
            
            PrintValue(9) = "TYPE"
            FormatString(9) = "a4"
                
            PrintValue(10) = " "
            FormatString(10) = "a6"
            
            PrintValue(11) = "DEBIT"
            FormatString(11) = "a5"
                
            PrintValue(12) = " "
            FormatString(12) = "a4"
            
            PrintValue(13) = "CREDIT"
            FormatString(13) = "a6"
            
            PrintValue(14) = " "
            FormatString(14) = "a2"
            
            PrintValue(15) = "BANK NAME"
            FormatString(15) = "a9"
            
            PrintValue(16) = " "
            FormatString(16) = "~"
                
            FormatPrint
            Ln = Ln + 1
                
            PrintValue(1) = "==============================================================================================================================="
            FormatString(1) = "a118"
            
            PrintValue(2) = " "
            FormatString(2) = "~"
            
            FormatPrint
            
            Ln = Ln + 1
            
        End If
        
'======================================================================================
'                                   DIRECT DEPOSIT DETAIL
'======================================================================================

        frmProgress.lblMsg2 = "Employee: " & trs!Employeeno & " - " & Trim(trs!Name)
        frmProgress.Show
        
        PrintValue(1) = trs!Employeeno
        FormatString(1) = "n6"
        
        PrintValue(2) = " "
        FormatString(2) = "a3"
        
        If IsNull(trs!Routing) Then
            trs!Routing = 0
        End If
        
        PrintValue(3) = trs!Routing
        FormatString(3) = "a11"
        
        PrintValue(4) = " "
        FormatString(4) = "a2"
        
        If IsNull(trs!AcctNo) Then
            trs!AcctNo = 0
        End If
        
        PrintValue(5) = trs!AcctNo
        FormatString(5) = "a14"
        
        PrintValue(6) = " "
        FormatString(6) = "a1"
        
        PrintValue(7) = trs!Name
        FormatString(7) = "a30"
        
        PrintValue(8) = " "
        FormatString(8) = "a2"
        
        If IsNull(trs!AcctType) Then
            trs!AcctType = 0
        End If
        PrintValue(9) = trs!AcctType
        FormatString(9) = "a3"
        
        PrintValue(10) = " "
        FormatString(10) = "a8"
        
        PrintValue(11) = trs!Credit
        FormatString(11) = "d8"
        TotCredAmt = TotCredAmt + trs!Credit
        SubCredAmt = SubCredAmt + trs!Credit
        
        PrintValue(12) = " "
        FormatString(12) = "a1"
        
        If IsNull(trs!BankName) Then
            trs!BankName = 0
        End If
        
        PrintValue(13) = trs!BankName
        FormatString(13) = "a20"
        
        PrintValue(14) = " "
        FormatString(14) = "~"
        
        FormatPrint
        Ln = Ln + 1
        LineCt = LineCt + 1
        SubLineCt = SubLineCt + 1
        
        LastBatch = trs!Batch
        If frmDirectDep.chkOutputFile Then
            WriteSixes     '  Write "6" detail lines of output file
        End If
        trs.MoveNext
    Loop
    DepositTotal = DepositTotal * 100
    
    If chkOutputFile Then
        If frmDirectDep.chkOutputFile Then
            WriteEmployerSix
        End If
        
    
        Write8sAnd9s
        PrintSubTotals
    End If
    Ln = Ln + 2
    PrintValue(1) = " "
    FormatString(1) = "a53"
    
    PrintValue(2) = "FINAL TOTAL CREDIT AMOUNT"
    FormatString(2) = "a27"
    
    PrintValue(3) = TotCredAmt
    FormatString(3) = "d8"
    
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " "
    FormatString(1) = "a59"
    
    PrintValue(2) = "NUMBER OF EMPLOYEES"
    FormatString(2) = "a26"
    
    PrintValue(3) = LineCt
    FormatString(3) = "n8"
    
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint

    trs.Close
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

'=======================================   DIRECT DEPOSIT REPORT   ======================================

Public Sub WriteEmployerSix()
Dim BAcct, RteNo  As Long
Dim CheckDigit, ChkDigNo As Byte
Dim FedNo As String

    ChkDigNo = Right(PRCompany.BankABA, 1)
    RteNo = PRCompany.BankABA
    FedID = Left(PRCompany.FederalID, 2)
    FedID = FedID & Mid(PRCompany.FederalID, 4, 8)
    
    RT1 = Left(RteNo, 1)
    RT2 = Mid(RteNo, 2, 1)
    RT3 = Mid(RteNo, 3, 1)
    RT4 = Mid(RteNo, 4, 1)
    RT5 = Mid(RteNo, 5, 1)
    RT6 = Mid(RteNo, 6, 1)
    RT7 = Mid(RteNo, 7, 1)
    RT8 = Mid(RteNo, 8, 1)
    
    Sum1 = RT1 * 3
    Sum2 = RT2 * 7
    Sum3 = RT3 * 1
    Sum4 = RT4 * 3
    Sum5 = RT5 * 7
    Sum6 = RT6 * 1
    Sum7 = RT7 * 3
    Sum8 = RT8 * 7
    
    TSUM = TSUM + Sum1 + Sum2 + Sum3 + Sum4 + Sum5 + Sum6 + Sum7 + Sum8
    
    diff = TSUM Mod 10
    
    If diff = 0 Then
        CheckDigit = 0
    Else
        CheckDigit = 10 - diff
    End If
    
    '  See if Calc of the Check Diget matches the 9th digit in the Routing Number
    If CheckDigit <> ChkDigNo Then
        MsgBox "Check Digit for Employer: " & Trim(PRCompany.Name) & " is not valid!!!", vbCritical, "Employee Lists and Labels"
        End
    End If
    
    BAcct = PRCompany.BankAccount
    
    If Len(BAcct) < 17 Then
        diff = 17 - Len(BAcct)
        AcctNum = BAcct & Space(diff)
    Else
        AcctNum = Left(BAcct, 17)
    End If
    
    BankName = PRCompany.BankName
    
    If Len(BankName) < 22 Then
        diff = 22 - Len(BankName)
        BankName = BankName & Space(diff)
    Else
        BankName = Left(BankName, 22)
    End If
    
    FedNo = FedID
    If Len(FedNo) < 15 Then
        diff = 15 - Len(FedNo)
        FedNo = FedNo & Space(diff)
    Else
        FedNo = Left(FedNo, 15)
    End If
    
    SeqNo = SeqNo + 1
    
'                 X = "6" & ChkSve & RteNo & CheckDigit & AcctNum & Format(trs!Credit * 100, "0000000000") & EmpNo & EName & Space(2) & "0" & Left(PRCompany.BankABA, 8) & Format(SeqNo, "0000000")
    x = "627" & Left(PRCompany.BankABA, 8) & CheckDigit & AcctNum & Format(DepositTotal, "0000000000") & FedNo & BankName & Space(2) & "0" & Left(PRCompany.BankABA, 8) & Format(SeqNo, "0000000")
    Print #TChannel, x  ' Output text.
    WriteCt = WriteCt + 1
    
End Sub

Public Sub PrintSubTotals()
    Ln = Ln + 1
    PrintValue(1) = " "
    FormatString(1) = "a51"
    
    PrintValue(2) = "BATCH " & LastBatch & " TOTAL CREDIT AMOUNT"
    FormatString(2) = "a29"
    
    PrintValue(3) = SubCredAmt
    FormatString(3) = "d8"
    
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " "
    FormatString(1) = "a59"
    
    PrintValue(2) = "NUMBER OF EMPLOYEES"
    FormatString(2) = "a26"
    
    PrintValue(3) = SubLineCt
    FormatString(3) = "n8"
    
    PrintValue(4) = " "
    FormatString(4) = "~"
    
    FormatPrint
    SubCredAmt = 0
    SubLineCt = 0

End Sub


Public Sub WriteOneAndFive()

Dim OutFile, FiveCoName As String
Dim CreateDate, CreateTime, lenCoName As Long
WriteCt = 0

    CreateDate = Format(Now, "yymmdd")
    CreateTime = Format(Now, "hhmm")
    CoName = PRCompany.Name
    BankName = PRCompany.BankName
    FiveCoName = PRCompany.Name
    FedID = Left(PRCompany.FederalID, 2)
    FedID = FedID & Mid(PRCompany.FederalID, 4, 8)
        
    If Len(CoName) < 23 Then
        diff = 23 - Len(CoName)
        CoName = CoName & Space(diff)
    Else
        CoName = Left(CoName, 23)
    End If
    
    If Len(BankName) < 23 Then
        diff = 23 - Len(BankName)
        BankName = BankName & Space(diff)
    Else
        BankName = Left(BankName, 23)
    End If
    
    If Len(CoName) < 16 Then
        diff = 23 - Len(CoName)
        FiveCoName = FiveCoName & Space(diff)
    Else
        FiveCoName = Left(CoName, 16)
    End If
    
    BatNo = Format(PRBatch.BatchID, "0000000")
    
    TChannel = FreeFile
                               
    Open "C:\balint\OutFile" For Output As #TChannel Len = 94
    
    '  File Header
 
    x = "101" & Space(1) & PRCompany.BankABA & Space(1) & FedID & CreateDate & CreateTime & "A094101" & BankName & CoName & Format(CheckDt, "yyyymmdd")
    Print #TChannel, x  ' Output text.
    WriteCt = WriteCt + 1
    
    x = "5200" & FiveCoName & "PAY ENDING: " & Format(PRItemHist.PEDate, "yyyymmdd") & Space(1) & FedID & "PPDPAYROLL" & Space(3) & Format(CheckDt, "yymmdd") & Format(CheckDt, "yymmdd") & Space(3) & "1" & Left(PRCompany.BankABA, 8) & BatNo
    Print #TChannel, x  ' Output text.
    WriteCt = WriteCt + 1

'Close #TChannel

End Sub

Public Sub WriteSixes()
Dim AcctNum, EName, EmpNo As String
Dim ChkSve, Sum1, Sum2, Sum3, Sum4, Sum5, Sum6, Sum7, Sum8, TSUM, CheckDigit As Integer
Dim RT1, RT2, RT3, RT4, RT5, RT6, RT7, RT8, diff, Wdiff, ChkDigNo As Long

    If trs!AcctType = "CHK" Then
        ChkSve = 22
    Else
        ChkSve = 32
    End If
    
    SeqNo = SeqNo + 1
     
    RteNo = Left(trs!Routing, 8)
    
    EEABA = EEABA + RteNo
    
    ChkDigNo = Right(trs!Routing, 1)
    
    RT1 = Left(RteNo, 1)
    RT2 = Mid(RteNo, 2, 1)
    RT3 = Mid(RteNo, 3, 1)
    RT4 = Mid(RteNo, 4, 1)
    RT5 = Mid(RteNo, 5, 1)
    RT6 = Mid(RteNo, 6, 1)
    RT7 = Mid(RteNo, 7, 1)
    RT8 = Mid(RteNo, 8, 1)
    
    Sum1 = RT1 * 3
    Sum2 = RT2 * 7
    Sum3 = RT3 * 1
    Sum4 = RT4 * 3
    Sum5 = RT5 * 7
    Sum6 = RT6 * 1
    Sum7 = RT7 * 3
    Sum8 = RT8 * 7
    
    TSUM = TSUM + Sum1 + Sum2 + Sum3 + Sum4 + Sum5 + Sum6 + Sum7 + Sum8
    
    diff = TSUM Mod 10
    
    If diff = 0 Then
        CheckDigit = 0
    Else
        CheckDigit = 10 - diff
    End If
    
    '  See if Calc of the Check Diget matches the 9th digit in the Routing Number
    If CheckDigit <> ChkDigNo Then
        MsgBox "Check Digit for Employee: " & trs!Employeeno & " - " & Trim(trs!Name) & " is not valid!!!", vbCritical, "Employee Lists and Labels"
        End
    End If
        
    If Len(trs!AcctNo) < 17 Then
        diff = 17 - Len(trs!AcctNo)
        AcctNum = trs!AcctNo & Space(diff)
    Else
        AcctNum = Left(trs!AcctNo, 17)
    End If
    
    If Len(trs!Employeeno) < 15 Then
        diff = 15 - Len(trs!Employeeno)
        EmpNo = trs!Employeeno & Space(diff)
    Else
        EmpNo = Left(trs!Employeeno, 15)
    End If
    
    If Len(trs!Name) < 22 Then
        diff = 22 - Len(trs!Name)
        EName = trs!Name & Space(diff)
    Else
        EName = Left(trs!Name, 22)
    End If

    x = "6" & ChkSve & RteNo & CheckDigit & AcctNum & Format(trs!Credit * 100, "0000000000") & EmpNo & EName & Space(2) & "0" & Left(PRCompany.BankABA, 8) & Format(SeqNo, "0000000")
    Print #TChannel, x  ' Output text.
    WriteCt = WriteCt + 1

    '  Accumulate Hash Total
    Hash = Hash + RteNo
    DepositTotal = DepositTotal + trs!Credit
    
End Sub


Public Sub Write8sAnd9s()

Dim HashString As String, BlockCt, diff, block As Long

    HashString = Right(Format(Hash, String(20, "0")), 10)
    
    diff = SeqNo Mod 10
    If diff = 0 Then
        BlockCt = SeqNo / 10
    Else
        BlockCt = (SeqNo / 10)
        block = CLng(BlockCt)
        block = block + 1
    End If
    
    If frmDirectDep.chkBalFile Then
        x = "8200" & Format(SeqNo, "000000") & Format(HashString, "0000000000") & Format(DepositTotal, "000000000000") & Format(DepositTotal, "000000000000") & Space(1) & FedID & Space(25) & Left(PRCompany.BankABA, 8) & BatNo
        Print #TChannel, x  ' Output text.
        WriteCt = WriteCt + 1
        x = "9000001" & Format(block, "000000") & Format(SeqNo, "00000000") & Format(HashString, "0000000000") & Format(DepositTotal, "000000000000") & Format(DepositTotal, "000000000000")
        Print #TChannel, x  ' Output text.
        WriteCt = WriteCt + 1
    Else
        x = "8200" & Format(SeqNo, "000000") & Format(HashString, "0000000000") & "000000000000" & Format(DepositTotal, "000000000000") & Space(1) & FedID & Space(25) & Left(PRCompany.BankABA, 8) & BatNo
        Print #TChannel, x  ' Output text.
        WriteCt = WriteCt + 1
        x = "9000001" & Format(block, "000000") & Format(SeqNo, "00000000") & Format(HashString, "0000000000") & "000000000000" & Format(DepositTotal, "000000000000")
        Print #TChannel, x  ' Output text.
        WriteCt = WriteCt + 1
    End If
    
    block = block * 10
    Wdiff = block - WriteCt
    
    If Wdiff <> 10 Then
        x = "9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"
        For i = 1 To Wdiff
            Print #TChannel, x  ' Output text.
        Next
    End If
    
End Sub


Public Sub CheckPrint()

Dim ReportTitle As String
Dim CheckBankName As String
Dim LineNo As Long
Dim ILineNo As Long
Dim CoName As String

    SetEquates

    LineNo = 25
    ILineNo = 25

    PrtInit ("Port")
    ReportTitle = "Print Checks "
            
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    
    CheckBankName = "XXXXXXXXXXXXXX"
    BankAddress = "XXXXXXXXXXXXXX"
    BankNumber = "XXXXXXXXXXXXX"
    
    SQLString = "SELECT * FROM PRHist WHERE PRHist.HistID = 1"
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "History Record Not Found !!!", vbCritical, "Print Checks"
        Exit Sub
    End If
    
    WrittenAmount = AmountInWords(PRHist.Net)

    If Not PREmployee.GetByID(PRHist.EmployeeID) Then
        MsgBox "Employee Not Found !!!", vbCritical, "Print Checks"
        Exit Sub
    End If
    
    If Not PRBatch.GetByID(PRHist.BatchID) Then
        MsgBox "Batch Not Found !!!", vbCritical, "Print Checks"
        Exit Sub
    End If

    frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.LFName)
    frmProgress.Show
        
    Prvw.vsp.FontBold = True             '  Turn on BOLD feature
    SetFont 13, Equate.Portrait
    Prt 3, 2, Trim(PRCompany.Name)
    
    SetFont 12, Equate.Portrait          '  Check Number and Company Name in  LARGE FONT and BOLD

    Prt 2, 70, PRHist.CheckNumber

    Prt 23, 1, Trim(PRCompany.Name)
    
    SetFont 10, Equate.Portrait         '  Set font to size 10
    
    Prt 14, 10, Trim(PREmployee.FLName)
    Prt 15, 10, Trim(PREmployee.Address1)
    
    If Trim(PREmployee.Address2) = "" Then
        Prt 16, 10, Trim(PREmployee.City) & "  " & Trim(PREmployee.ZipCode)
        Prt 17, 61, "___________________________________"
    Else
        Prt 16, 10, Trim(PREmployee.Address2)
        Prt 17, 10, Trim(PREmployee.City) & "  " & Trim(PREmployee.ZipCode)
        Prt 18, 61, "___________________________________"
    End If
    
    Prvw.vsp.FontBold = False           '  Turn off BOLD Feature
    
    Prt 4, 3, Trim(PRCompany.Address1)

    If Trim(PRCompany.Address2) = "" Then
        Prt 5, 3, Trim(PRCompany.City) & "  " & Trim(PRCompany.ZipCode)
    Else
        Prt 5, 3, Trim(PRCompany.Address2)
        Prt 6, 3, Trim(PRCompany.City) & "  " & Trim(PRCompany.ZipCode)
    End If
    
    Prt 9, 1, "PAY"
    Prt 9, 24, Trim(WrittenAmount)
    Prt 11, 40, PRHist.CheckNumber
    Prt 11, 55, Format(PRBatch.PEDate, "mm/dd/yyyy")
    Prt 11, 81, CheckAmount(PRHist.Net)
    
    Prt 23, 55, "CHK DATE: "
    Prt 23, 65, Format(PRBatch.CheckDate, "mm/dd/yyyy ")
    Prt 23, 78, "CHK #: "
    Prt 23, 85, PRHist.CheckNumber
          
    Prt 24, 1, "- - - - - - - - - - - CURRENT PD YR TO DATE       - - - - - - - - - - - - - "
    
    Prt 42, 50, Trim(PREmployee.FLName)
    Prt 42, 91, Trim(PREmployee.EmployeeNumber)
    
    If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
        MsgBox "Department Info Not Found!!!", vbCritical, "Check Print"
        End
    End If
        
    Prt 43, 50, "DPT: " & PRDepartment.DepartmentNumber
    Prt 43, 58, "RATE: " & Format(PRHist.RegRate, "##,###.#0")
    Prt 43, 79, "PE DATE: " & Format(PRHist.PEDate, "yyyymmdd")
    
     SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID & _
                 " AND (PRDist.ItemType = 3 OR PRDist.ItemType =  4)"

    If PRDist.GetBySQL(SQLString) Then
        Do
            PRItem.GetByID (PRDist.ItemID)
            Prt ILineNo, 1, Trim(PRItem.Title)
            Prt ILineNo, 18, CurrFormat(PRDist.Amount)
            Prt ILineNo, 30, CurrFormat(PRDist.Amount)
            ILineNo = ILineNo + 1
            If Not PRDist.GetNext Then Exit Do
        Loop
    End If

    SQLString = "SELECT * FROM PRItemHist " & _
                "WHERE PRItemHist.HistID = " & PRHist.HistID

    If PRItemHist.GetBySQL(SQLString) Then
        Prt LineNo, 50, "REG HRS"
        Prt LineNo, 66, CurrFormat(PRHist.RegHours)
        Prt LineNo, 83, CurrFormat(PRHist.RegHours)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "OVT HRS"
        Prt LineNo, 66, CurrFormat(PRHist.OTHours)
        Prt LineNo, 83, CurrFormat(PRHist.OTHours)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "OTH HRS"
        Prt LineNo, 66, CurrFormat(PRHist.OEHours)
        Prt LineNo, 83, CurrFormat(PRHist.OEHours)

        LineNo = LineNo + 1
        
        Prt LineNo, 50, "REG PAY"
        Prt LineNo, 66, CurrFormat(PRHist.RegAmount)
        Prt LineNo, 83, CurrFormat(PRHist.RegAmount)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "OVT PAY"
        Prt LineNo, 66, CurrFormat(PRHist.OTAmount)
        Prt LineNo, 83, CurrFormat(PRHist.OTAmount)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "OTH PAY"
        Prt LineNo, 66, CurrFormat(PRHist.OEAmount)
        Prt LineNo, 83, CurrFormat(PRHist.OEAmount)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "GROSS"
        Prt LineNo, 66, CurrFormat(PRHist.Gross)
        Prt LineNo, 83, CurrFormat(PRHist.Gross)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "DEDUCTIONS"
        Prt LineNo, 66, CurrFormat(PRHist.Deductions)
        Prt LineNo, 83, CurrFormat(PRHist.Deductions)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "FIC TAX"
        Prt LineNo, 66, CurrFormat(PRHist.SSTax)
        Prt LineNo, 83, CurrFormat(PRHist.SSTax)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "MED TAX"
        Prt LineNo, 66, CurrFormat(PRHist.MedTax)
        Prt LineNo, 83, CurrFormat(PRHist.MedTax)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "FWT TAX M02"
        Prt LineNo, 66, CurrFormat(PRHist.FWTTax)
        Prt LineNo, 83, CurrFormat(PRHist.FWTTax)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "SWT TAX M02"
        Prt LineNo, 66, CurrFormat(PRHist.SWTTax)
        Prt LineNo, 83, CurrFormat(PRHist.SWTTax)
        
        LineNo = LineNo + 1

        Prt LineNo, 50, "CWT TAX"
        Prt LineNo, 66, CurrFormat(PRHist.CWTTax)
        Prt LineNo, 83, CurrFormat(PRHist.CWTTax)
        
        LineNo = LineNo + 3

        Prt LineNo, 50, "NET PAY"
        Prt LineNo, 66, CurrFormat(PRHist.Net)
        Prt LineNo, 83, CurrFormat(PRHist.Net)
        LineNo = LineNo + 2

    End If

    SQLString = "SELECT * FROM PRItemHist " & _
                "WHERE PRItemHist.HistID = " & PRHist.HistID & _
                " AND PRItemHist.ItemType = 4"

    If PRItemHist.GetBySQL(SQLString) Then
        Do
            PRItem.GetByID (PRItemHist.ItemID)
            Prt ILineNo, 1, Trim(PRItem.Title)
            Prt ILineNo, 18, CurrFormat(PRItemHist.Amount)
            Prt ILineNo, 30, CurrFormat(PRItemHist.Amount)
            
            ILineNo = ILineNo + 1
            
            If Not PRItemHist.GetNext Then Exit Do
        Loop
    End If
    
'  Change font to small

    SetFont 8, Equate.Portrait
    
    Prvw.vsp.FontBold = True
    
    Prt 2, 75, Trim(CheckBankName)
    
    Prvw.vsp.FontBold = False
    
    Prt 3, 75, Trim(BankAddress)
    Prt 4, 75, BankNumber

    Prt 14, 1, "TO THE"
    Prt 15, 1, "ORDER"
    Prt 16, 2, "OF"
   

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub NewHireReport()
    frmNewHire.Hide
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
    
    PrtInit ("Port")    ' "Port" = Portrait
    
    ' set up SQL statement based upon order requested
    ReportTitle = ""
    
    SQLString = "SELECT * FROM PREmployee WHERE PREmployee.DateHired >= " & CLng(Startdate) & _
                " AND PREmployee.DateHired <= " & CLng(EndDate) & _
                " ORDER BY PREmployee.EmployeeNumber"
MsgBox SQLString
    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Found !!!", vbCritical, "State of Ohio New Hire Reporting Form 7048"
        Exit Sub
    End If
    
    Do

        Ln = Ln + 4
        PrintValue(1) = "STATE OF OHIO NEW HIRE REPORTING FORM 7048"
        FormatString(1) = "a90"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
        
        FormatPrint
        Ln = Ln + 4
        
        PrintValue(1) = "E M P L O Y E E   I N F O R M A T I O N"
        FormatString(1) = "a50"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
        
        FormatPrint
        Ln = Ln + 4
        
        PrintValue(1) = "SOCIAL SECURITY NUMBER:"
        FormatString(1) = "a26"
        
        PrintValue(2) = Format(PREmployee.SSN, "000-00-0000")
        FormatString(2) = "a11"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = "NAME:"
        FormatString(1) = "a26"
        
        PrintValue(2) = PREmployee.FLName
        FormatString(2) = "a30"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        If Trim(PREmployee.Address2) = "" Then
            PrintValue(1) = "ADDRESS:"
            FormatString(1) = "a26"
        Else
            PrintValue(1) = "ADDRESS 1:"
            FormatString(1) = "a26"
        End If
        
        PrintValue(2) = PREmployee.Address1
        FormatString(2) = "a30"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        If Trim(PREmployee.Address2) <> "" Then
            PrintValue(1) = "ADDRESS 2:"
            FormatString(1) = "a26"
            
            PrintValue(2) = PREmployee.Address2
            FormatString(2) = "a30"
            
            PrintValue(3) = " "
            FormatString(3) = "~"
            
            FormatPrint
            Ln = Ln + 2
        End If
        
        PrintValue(1) = "CITY/STATE/ZIP:"
        FormatString(1) = "a26"
        
        PrintValue(2) = Trim(PREmployee.City) & ", " & PREmployee.State & "  " & PREmployee.ZipCode
        FormatString(2) = "a90"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = "EMPLOYEE DATE OF HIRE:"
        FormatString(1) = "a26"
        
        PrintValue(2) = Format(PREmployee.DateHired, "mm/dd/yyyy")
        FormatString(2) = "a10"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = "DATE OF BIRTH:"
        FormatString(1) = "a26"
        
        If PREmployee.DateOfBirth = 0 Then
            PrintValue(2) = Format(PREmployee.DateOfBirth, "00/00/0000")
        Else
            PrintValue(2) = Format(PREmployee.DateOfBirth, "mm/dd/yyyy")
        End If
        FormatString(3) = "a10"
        
        PrintValue(4) = " "
        FormatString(4) = "~"
        
        FormatPrint
        Ln = Ln + 5
                 
        PrintValue(1) = "E M P L O Y E R   I N F O R M A T I O N"
        FormatString(1) = "a50"
        
        PrintValue(2) = " "
        FormatString(2) = "~"
        
        FormatPrint
        Ln = Ln + 4
        
        PrintValue(1) = "EMPLOYER FEDERAL EIN:"
        FormatString(1) = "a26"
        
        PrintValue(2) = PRCompany.FederalID
        FormatString(2) = "a11"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = "EMPLOYER NAME:"
        FormatString(1) = "a26"
        
        PrintValue(2) = PRCompany.Name
        FormatString(2) = "a30"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        If Trim(PRCompany.Address2) = "" Then
            PrintValue(1) = "ADDRESS:"
            FormatString(1) = "a26"
        Else
            PrintValue(1) = "ADDRESS 1:"
            FormatString(1) = "a26"
        End If
        
        PrintValue(2) = PRCompany.Address1
        FormatString(2) = "a30"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        If Trim(PRCompany.Address2) <> "" Then
            PrintValue(1) = "ADDRESS 2:"
            FormatString(1) = "a26"
        
            PrintValue(2) = PRCompany.Address2
            FormatString(2) = "a30"
        
            PrintValue(3) = " "
            FormatString(3) = "~"
        
            FormatPrint
            Ln = Ln + 2
        End If
        
        PrintValue(1) = "CITY/STATE/ZIP:"
        FormatString(1) = "a26"
        
        SQLString = "SELECT StateAbbrev from PRState where StateID = PRCompany.StateID"
        PrintValue(2) = Trim(PRCompany.City) & ", " & PRState.StateAbbrev & "  " & PRCompany.ZipCode
        FormatString(2) = "a60"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = "DATE: "
        FormatString(1) = "a26"
        
        PrintValue(2) = Format(Date, "mm/dd/yyyy")
        FormatString(2) = "a30"
        
        PrintValue(3) = " "
        FormatString(3) = "~"
        
        FormatPrint
        Ln = Ln + 2
                        
        If Not PREmployee.GetNext Then
            Exit Do
        End If
        
        FormFeed
            
    Loop
            
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub YECityTax()
Dim SYTDGROSS As Currency
Dim SYTDTAX As Currency
Dim TYTDGross As Currency
Dim TYTDTAX As Currency
Dim StartYM As Long
Dim EndYM As Long
Dim CityName As String
Dim CityNumber As Long
Dim LastCityID As Long
Dim LastCityName As String
Dim LastCityNumber As Long
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
                " AND PRDist.YEARMONTH <= " & EndYM

    If Not PRDist.GetBySQL(SQLString) Then
        MsgBox "Data Not Found!!!", vbCritical, "Payroll Yearly City Tax Report"
        End
    End If
    Do
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
               
        If LastCityID <> 0 And LastCityID <> trs!CityID Then
        
            Ln = Ln + 1
            
            PrintValue(1) = LastCityNumber & " - " & Trim(LastCityName) & " TOTALS"
                                                                FormatString(1) = "a41"
            PrintValue(3) = SYTDGROSS:                          FormatString(3) = "d13"
            PrintValue(5) = SYTDTAX:                            FormatString(5) = "d13"
            PrintValue(6) = " ":                                FormatString(6) = "~"
                        
            FormatPrint
            Ln = Ln + 2
            
            SYTDGROSS = 0
            SYTDTAX = 0
            
            FormFeed
            YECityHeader (ReportTitle)
                      
            FormatPrint
            Ln = Ln + 2
        End If
        
        If Not PREmployee.GetByID(trs!EmployeeID) Then
            MsgBox "Employee Info Not Found!!!", vbCritical, "Payroll Yearly City Tax Report"
            End
        End If
        
        frmProgress.lblMsg2 = "Employee: " & PREmployee.EmployeeNumber & " - " & Trim(PREmployee.LFName)
        frmProgress.Show
        
        PrintValue(1) = PREmployee.EmployeeNumber:              FormatString(1) = "a7"
        PrintValue(2) = Trim(PREmployee.LFName):                FormatString(2) = "a35"
        PrintValue(3) = Format(PREmployee.SSN, "000-00-0000"):  FormatString(3) = "a16"
        PrintValue(4) = Trim(PREmployee.Address1):              FormatString(4) = "a28"
        PrintValue(5) = trs!YTDGross:                           FormatString(5) = "d16"
        PrintValue(6) = trs!YTDTax:                             FormatString(6) = "d16"
        PrintValue(7) = " ":                                    FormatString(7) = "~"
        
        SYTDGROSS = SYTDGROSS + trs!YTDGross
        TYTDGross = TYTDGross + trs!YTDGross
        SYTDTAX = SYTDTAX + trs!YTDTax
        TYTDTAX = TYTDTAX + trs!YTDTax

        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                                    FormatString(1) = "a58"
                
        If Trim(PREmployee.Address2) <> "" Then
            PrintValue(2) = Trim(PREmployee.Address2):          FormatString(2) = "a30"
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
            Exit Do
        End If
    Loop
    Ln = Ln + 1
    
    PrintValue(1) = LastCityNumber & " - " & Trim(LastCityName) & " TOTALS"
                                                                FormatString(1) = "a86"
    PrintValue(2) = SYTDGROSS:                                  FormatString(2) = "d16"
    PrintValue(3) = SYTDTAX:                                    FormatString(3) = "d16"
    PrintValue(4) = " ":                                        FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "GRAND TOTALS":                             FormatString(1) = "a86"
    PrintValue(2) = TYTDGross:                                  FormatString(2) = "d16"
    PrintValue(3) = TYTDTAX:                                    FormatString(3) = "d16"
    PrintValue(4) = " ":                                        FormatString(4) = "~"
    
    FormatPrint

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub


Public Sub YECityHeader(ReportTitle)
 
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

Public Sub PayrollEntryHeader()
Dim LastAbbrev As String
    SetFont 8, Equate.Portrait
    
    PrintValue(1) = "Emp #"
    FormatString(1) = "a5"

    PrintValue(2) = " "
    FormatString(2) = "a2"
    
    PrintValue(3) = "NAME"
    FormatString(3) = "a25"
    
    PrintValue(4) = " "
    FormatString(4) = "a2"
    
    PrintValue(5) = "DPT"
    FormatString(5) = "a3"
        
    PrintValue(6) = " "
    FormatString(6) = "a3"
    
    PrintValue(7) = "RATE/SAL"
    FormatString(7) = "a8"
        
    PrintValue(8) = " "
    FormatString(8) = "a2"
    
    PrintValue(9) = "REG HRS"
    FormatString(9) = "a7"
        
    PrintValue(10) = " "
    FormatString(10) = "a2"
    
    PrintValue(11) = "OT HRS"
    FormatString(11) = "a7"
        
    PrintValue(12) = " "
    FormatString(12) = "a2"

    PrintValue(13) = " "
    FormatString(13) = "~"
            
    FormatPrint
    FormatString(1) = ""
    
    ChkRegDedHeader
    
    trsDED.MoveFirst
    SetFont 7, Equate.Portrait
    Do
        '  PRINT HEADER - change in type or max number of columns
        
        Colcount = Colcount + 1

        If (LastType <> 0 And LastType <> trsDED!Type) Or Colcount = 6 Or trsDED.EOF Then

            PrtString = Space(83) & FormatString(Colcount)
            PrintValue(1) = PrtString
            FormatString(1) = "a200"

            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1
            
            Colcount = 1
            FormatString(1) = " "
            FormatString(Colcount) = " "
            
        End If

        LastType = trsDED!Type
        
        FormatString(Colcount + 1) = FormatString(Colcount) & trsDED!Abbreviation & Space(3)
        LastAbbrev = trsDED!Abbreviation
        trsDED.MoveNext

        If trsDED.EOF Then
            Exit Do
        End If
    Loop
    If trsDED.EOF Then
        If Colcount > 0 Then
            PrtString = Space(83) & FormatString(Colcount) & LastAbbrev
            PrintValue(1) = PrtString
            FormatString(1) = "a200"

            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1
            
            PrintValue(1) = "________________________________________________________________________________________________________________________________________________"
            FormatString(1) = "a150"
            
            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1
                
        End If
    End If
    
End Sub

Public Sub PayrollDataEntry()                              ''''''''''''''''''''''''''''''''''''''''''
Dim LastLine, LineNumber, Colcount, LastType As Byte
Dim LastEmp As Integer
Dim LastPctAmt As Currency
LastLine = 0
LastEmp = 0
LastPctAmt = 0
Colcount = 0

    SetEquates
    
    PrtInit ("Port")
    ReportAbbreviation = "PAYROLL DATA ENTRY FORM"
    SetFont 8, Equate.Portrait

    frmProgress.lblMsg1 = "Printing " & ReportAbbreviation & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportAbbreviation)
    

    ' loop thru employees
    SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
  
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employees Were Found!!!", vbCritical, "Payroll Data Entry"
        End
    End If

    PageHeader ReportAbbreviation, Msg1, "", ""
    PrintCompanyHeader (ReportList)
    
    ChkRegGetHeaderData                 '  ##########   USED CHECK REGISTER PROCEDURE   ########
    
    Set trsded2 = trsDED.Clone
    
    trsDED.Sort = "Type,ItemID"
    trsded2.Sort = "Type,ItemID"
    PayrollEntryHeader
    
    Ln = Ln + 1
    LastLine = 0
    SetFont 8, Equate.Portrait
    Do

        FirstSw = True
        
        frmProgress.lblMsg2 = "Processing Employee No.:  " & PREmployee.EmployeeID
        frmProgress.Show
        
        trsded2.MoveFirst
            
        Do

            If Ln = 0 Or Ln > MaxLines Then
                If Ln Then FormFeed
                PageHeader ReportTitle, Msg1, "", ""
                PayrollEntryHeader
            End If
            
'            PayrollOEDed                 ''  ###   PAYROLLOEDED
'            PayrollDeductions

            If LastLine <> 0 And trsded2!LineNo <> LastLine Then
                WriteDed
                PrtLine = Space(255)
            End If
            If PREmployee.EmployeeID <> LastEmp Then
                Colcount = 0
            End If
            Colcount = Colcount + 1
            SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = " & PREmployee.EmployeeID & _
                " AND PRItem.EmployerItemID = " & trsded2!ItemID
            If PRItem.GetBySQL(SQLString) Then
                If PRItem.AmtPct > 0 Then
                    If Colcount > 1 Then
                        If PRItem.AmtPct < 100 Then
                            PrtLine = Trim(PrtLine) & "  " & Format(PRItem.AmtPct, " ###0.00") & " ."
                        Else
                            PrtLine = Trim(PrtLine) & "  " & Format(PRItem.AmtPct, "##0.00") & " ."
                        End If
                    Else
                        If PRItem.AmtPct < 100 Then
                            PrtLine = Trim(PrtLine) & ". " & Format(PRItem.AmtPct, " ###0.00") & " ."
                        Else
                            PrtLine = Trim(PrtLine) & ". " & Format(PRItem.AmtPct, "##0.00") & " ."
                        End If
                    End If
                Else
                    PrtLine = Trim(PrtLine) & "[___.__] ."
                End If
            Else
                PrtLine = Trim(PrtLine) & "[___.__] ."
            End If

            If FirstSw And Colcount = 5 Then
                WriteEmpLine
                WriteDed
                FirstSw = False
                PrtLine = Space(255)
                Colcount = 0
            ElseIf Colcount = 5 Then
                WriteDed
                PrtLine = Space(255)
                Colcount = 0
            End If

'            LastLine = trsded2!LineNo
            LastEmp = PREmployee.EmployeeID
            LastPctAmt = PRItem.AmtPct
            trsded2.MoveNext

            If trsded2.EOF Then
                If Colcount > 1 Then
                    If PRItem.AmtPct < 100 Then
                        PrtLine = Trim(PrtLine) & "  " & Format(LastPctAmt, " ###0.00") & " ."
                    Else
                        PrtLine = Trim(PrtLine) & "  " & Format(LastPctAmt, "##0.00") & " ."
                    End If
                Else
                    If PRItem.AmtPct < 100 Then
                        PrtLine = Trim(PrtLine) & ". " & Format(LastPctAmt, " ###0.00") & " ."
                    Else
                        PrtLine = Trim(PrtLine) & ". " & Format(LastPctAmt, "##0.00") & " ."
                    End If
                End If
                WriteDed
                Exit Do
            End If
        Loop
        WriteDed
'        trsded2.MoveNext

        If trsded2.EOF Then
            Exit Do
        End If
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub


Public Sub PayrollOEDed()
Dim PrtString As String
Dim LastCheck As Long
Dim LastEmployee As Long

WriteCt = 0

    SQLString = "SELECT * FROM PRHist WHERE PRHist.histid = " & rrs!HistID
  
    If Not PRHist.GetBySQL(SQLString) Then
       MsgBox "No History!!!"
       End
    End If

    Do
        ' clear out temp record set
        trsDED.MoveFirst
        Do
            trsDED!Amount = 0
            trsDED.Update
            trsDED.MoveNext
            If trsDED.EOF Then Exit Do
        Loop

    ''''''''''''''''''''''''''''''     PROCESS OTHER EARNINGS   '''''''''''''''''''''''''''''''''''
        
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID & _
            " AND PRDist.DistType = " & PREquate.DistTypeItem

        If PRDist.GetBySQL(SQLString) Then
            Do
                If frmPayEntry.chkOtherEarns Then
                    PayrollOEHours
                    PayrollOEAmount
                End If
                    
            If Not PRDist.GetNext Then Exit Do
            Loop
        End If


    '''''''''''''''''''''''''''''''''''     PROCESS DEDUCTIONS     '''''''''''''''''''''''''''''''''''''''
        If frmPayEntry.chkDeds Then
            SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID & _
                    " AND PRItemHist.ItemType <> " & PREquate.ItemTypeDirDepDed
            If PRItemHist.GetBySQL(SQLString) Then
                PayrollDeductions
            End If
        End If
    '''''''''''''''''''''''''''''''''''        WRITE RECORD     '''''''''''''''''''''''''''''''''''''''
        
'        WriteDed

        LastEmployee = PRDist.EmployeeID

        If Not PRHist.GetNext Then Exit Do

    Loop
    
End Sub

Public Sub PayrollOEHours()
    If LastLine <> 0 And trsDED!LineNo <> LastLine Then
        WriteDed
        PrtLine = Space(255)
    End If
    If rrs![PREmployee.EmployeeID] <> LastEmp Then
        Colcount = 0
    End If
    Colcount = Colcount + 1

    If PRItem.GetBySQL(SQLString) Then
        If PRItem.AmtPct > 0 Then
            If Colcount > 1 Then
                If PRItem.AmtPct < 100 Then
                    PrtLine = Trim(PrtLine) & "  " & Format(PRItem.AmtPct, " ###0.00") & " ."
                Else
                    PrtLine = Trim(PrtLine) & "  " & Format(PRItem.AmtPct, "##0.00") & " ."
                End If
            Else
                If PRItem.AmtPct < 100 Then
                    PrtLine = Trim(PrtLine) & ". " & Format(PRItem.AmtPct, " ###0.00") & " ."
                Else
                    PrtLine = Trim(PrtLine) & ". " & Format(PRItem.AmtPct, "##0.00") & " ."
                End If
            End If
        Else
            PrtLine = Trim(PrtLine) & "[___.__] ."
        End If
    Else
        PrtLine = Trim(PrtLine) & "[___.__] ."
    End If

        If FirstSw And Colcount = 5 Then
            WriteEmpLine
            WriteDed
            FirstSw = False
            PrtLine = Space(255)
            Colcount = 0
        ElseIf Colcount = 5 Then
            WriteDed
            PrtLine = Space(255)
            Colcount = 0
        End If

'            LastLine = trsDED!LineNo
        LastEmp = rrs![PREmployee.EmployeeID]
        LastPctAmt = PRItem.AmtPct
        trsDED.MoveNext
        
        If trsDED.EOF Then
            If Colcount > 1 Then
                If PRItem.AmtPct < 100 Then
                    PrtLine = Trim(PrtLine) & "  " & Format(LastPctAmt, " ###0.00") & " ."
                Else
                    PrtLine = Trim(PrtLine) & "  " & Format(LastPctAmt, "##0.00") & " ."
                End If
            Else
                If PRItem.AmtPct < 100 Then
                    PrtLine = Trim(PrtLine) & ". " & Format(LastPctAmt, " ###0.00") & " ."
                Else
                    PrtLine = Trim(PrtLine) & ". " & Format(LastPctAmt, "##0.00") & " ."
                End If
            End If
            WriteDed
        End If
 
'        WriteDed

        
End Sub

Public Sub PayrollOEAmount()
    If PRDist.Amount <> 0 Then
        Flag = False
        trsDED.MoveFirst

        Do
            If trsDED!ItemID = PRDist.EmployerItemID Then
                trsDED!Amount = PRDist.Amount
                Flag = True
                Exit Do
            End If
            trsDED.MoveNext
            If trsDED.EOF Then Exit Do
         Loop
            If Not Flag Then
                MsgBox "Employer Item Not Found: " & PRHist.EmployeeID & vbCr & PRDist.DistID, vbCritical
                End
            End If
    End If
End Sub

Public Sub PayrollDeductions()

    Do
        If PRItemHist.Amount <> 0 Then
            Flag = False
            trsDED.MoveFirst
            Do

                If trsDED!ItemID = PRItemHist.EmployerItemID Then
                    trsDED!Amount = PRItemHist.Amount
                    Flag = True
                    Exit Do
                End If
                trsDED.MoveNext
                If trsDED.EOF Then Exit Do
            Loop
                If Not Flag Then
                    MsgBox "Employer Item Not Found: " & PRItemHist.EmployeeID & vbCr & PRItemHist.ItemHistID, vbCritical
                    End
                End If
        End If
        If Not PRItemHist.GetNext Then Exit Do
    Loop

End Sub


Public Sub WriteEmpLine()

    If Ln > MaxLines Then
        If Ln Then FormFeed
        PayrollEntryHeader
        Ln = Ln + 1
    End If
    
    PrintValue(1) = PREmployee.EmployeeNumber
    FormatString(1) = "n6"

    PrintValue(2) = " "
    FormatString(2) = "a1"

    PrintValue(3) = Trim(PREmployee.FirstName) & " " & Trim(PREmployee.MidInit) & " " & Trim(PREmployee.LastName)
    FormatString(3) = "a25"

    PrintValue(4) = " "
    FormatString(4) = "a1"

    If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
        MsgBox "Department Info Not Found!!!", vbCritical, "Payroll Entry"
        End
    End If

    PrintValue(5) = PRDepartment.DepartmentNumber
    FormatString(5) = "n3"

    PrintValue(6) = " "
    FormatString(6) = "a4"

    PrintValue(7) = Format(PREmployee.SalaryAmount, "#,##0.00")
    FormatString(7) = "a8"

    PrintValue(8) = " "
    FormatString(8) = "a2"

    PrintValue(9) = "[___.__]"           ' Reg Hours
    FormatString(9) = "a8"

    PrintValue(10) = " "
    FormatString(10) = "a2"

    PrintValue(11) = "[___.__]"          ' OT Hours
    FormatString(11) = "a8"

    PrintValue(12) = " "
    FormatString(12) = "a1"
    
End Sub

Public Sub WriteDed()
    If Trim(PrtLine) <> "" Then
        If FirstSw = True Then

            PrintValue(13) = Trim(PrtLine)
            FormatString(13) = "a50"
        
            PrintValue(14) = " "
            FormatString(14) = "a1"
        
            PrintValue(15) = " "
            FormatString(15) = "~"
            
            FormatPrint
            Ln = Ln + 1
        Else

            PrintValue(1) = " "
            FormatString(1) = "a69"
            
            PrintValue(2) = Trim(PrtLine)
            FormatString(2) = "a50"
            
            PrintValue(3) = " "
            FormatString(3) = "a2"
        
            PrintValue(4) = " "
            FormatString(4) = "~"
            
            FormatPrint
            Ln = Ln + 1
        End If
    End If
    PrtLine = ""
    Colcount = 0

End Sub


Public Sub PayrollHeaderSetup()
''''''''''''''''''''''''''           NOT USED             ''''''''''''''''''''''''''''
Dim LineCt As Integer

    LastType = 0
    Colcount = 0
    PrtString = " "
    FormatString(1) = " "

    trsDED.MoveFirst
    
    Do
        '  PRINT HEADER - change in type or max number of columns
        
        If (LastType <> 0 And LastType <> trsDED!Type) Or Colcount = 5 Then
            PrtString = Space(15) & FormatString(Colcount)

            PrintValue(1) = PrtString
            FormatString(1) = "a200"

            PrintValue(2) = " "
            FormatString(2) = "~"

            FormatPrint
            Ln = Ln + 1

            Colcount = 0
            FormatString(1) = " "
            FormatString(Colcount) = " "

        End If

        LastType = trsDED!Type
        Colcount = Colcount + 1
        
        FormatString(Colcount + 1) = FormatString(Colcount) & trsDED!Abbreviation & Space(5)
        trsDED.MoveNext

        If trsDED.EOF Then
            Exit Do
        End If
    Loop
    
    PrtString = Space(15) & FormatString(Colcount)

    PrintValue(1) = PrtString
    FormatString(1) = "a200"

    PrintValue(2) = " "
    FormatString(2) = "~"

    FormatPrint
    Ln = Ln + 1
    
End Sub

Public Sub PayrollHeaderSpacing()
''''''''''''''''''''''''''           NOT USED  -  NOW           ''''''''''''''''''''''''''''
            
    AbbreviationCount = AbbreviationCount + 1
    VarLength = Len(TextAbbreviation)
    NumOfSpaces = 10 - VarLength
    ANumber = "a" & NumOfSpaces
    AbbreviationChars = "a" & VarLength

    PrintValue(1) = " "
    FormatString(1) = "a" & SpreadNumber

    PrintValue(2) = " "
    FormatString(2) = ANumber

    PrintValue(3) = TextAbbreviation
    FormatString(3) = AbbreviationChars
                            
    If AbbreviationCount = 5 Then
        FormatPrint
        Ln = Ln + 1
        AbbreviationCount = 0
        SpreadNumber = 83
    Else
        SpreadNumber = SpreadNumber + 10
        
        PrintValue(4) = " "
        FormatString(4) = "a" & SpreadNumber
    
        PrintValue(5) = " "
        FormatString(5) = "~"
        FormatPrint
    End If

End Sub


Public Sub CheckRecon(ByVal RangeType As Byte, _
                         ByVal BatchNumbr As Long, _
                         ByVal PEDate As Long, _
                         ByVal Startdate As Long, _
                         ByVal EndDate As Long)
Dim sqlstring1 As String
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

    SQLString = "Select * from PRHist"
 
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = Trim(SQLString) & " WHERE PRHist.PEDate >= " & CLng(Startdate) & " AND " & _
                                      " PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr
    Else
        SQLString = Trim(SQLString) & " WHERE PRHist.PEDate >= " & CLng(Startdate) & " AND " & _
                                    " PRHist.PEDate <= " & CLng(EndDate)
        Msg1 = "Date Range: " & CDate(Startdate) & " TO: " & CDate(EndDate)
    End If

    SQLString = Trim(SQLString) & " ORDER BY PRHist.CheckNumber"

    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History Found !!!", vbCritical, "Payroll Check Reconciliation"
        End
    End If

    Do
        If Ln = 0 Or Ln > MaxLines Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, "", ""
            Ln = Ln + 1
            
            PrintValue(1) = " ":                FormatString(1) = "a67"
            PrintValue(2) = "Check":            FormatString(2) = "a5"
                
            PrintValue(3) = " ":                FormatString(3) = "a8"
            PrintValue(4) = "Direct Deposit":   FormatString(4) = "a18"
                              
            PrintValue(5) = " ":                FormatString(5) = "a5"
            PrintValue(6) = " ":                FormatString(6) = "~"
                        
            FormatPrint
            Ln = Ln + 1
                        
            PrintValue(1) = " ":                FormatString(1) = "a0"
            PrintValue(2) = "Check No.":        FormatString(2) = "a9"
                        
            PrintValue(3) = " ":                FormatString(3) = "a2"
            PrintValue(4) = "P/E Date":         FormatString(4) = "a8"
                        
            PrintValue(5) = " ":                FormatString(5) = "a5"
            PrintValue(6) = "Payee":            FormatString(6) = "a5"
                        
            PrintValue(7) = " ":                FormatString(7) = "a38"
            PrintValue(8) = "Amount":           FormatString(8) = "a6"
                
            PrintValue(9) = " ":                FormatString(9) = "a11"
            PrintValue(10) = "Amount":          FormatString(10) = "a6"
            
            PrintValue(11) = " ":               FormatString(11) = "~"
            FormatPrint
                        
            PrintValue(1) = "______________________________________________________________________________________________________________"
            FormatString(1) = "a100"
            
            PrintValue(2) = " ":                FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 2
        End If

        sqlstring1 = "Date = " & PRHist.PEDate
        trs.Find sqlstring1, 0, adSearchForward, 1
        
        If trs.EOF Then
            trs.AddNew Array("Date", "Number", "Amount"), _
            Array(PRHist.PEDate, 0, 0)
            trs.UpdateBatch
        End If
        
        If Not PREmployee.GetBySQL("SELECT * FROM PREmployee WHERE PREmployee.EmployeeID = " & PRHist.EmployeeID) Then
            EmpName = "None"
        Else
            EmpName = PREmployee.FLName
        End If
        
        PrintValue(1) = PRHist.CheckNumber:         FormatString(1) = "n6"
        PrintValue(2) = " ":                        FormatString(2) = "a5"
                
        PrintValue(3) = Format(PRHist.PEDate, "yyyymmdd"):  FormatString(3) = "n8"
        PrintValue(4) = " ":                        FormatString(4) = "a5"
                
        PrintValue(5) = EmpName:                    FormatString(5) = "a30"
        PrintValue(6) = " ":                        FormatString(6) = "a5"
                
        PrintValue(7) = PRHist.Net:                 FormatString(7) = "d10"
        CheckAmt = CheckAmt + PRHist.Net
            
        PrintValue(8) = " ":                        FormatString(8) = "a5"
        PrintValue(9) = PRHist.DirectDeposit:       FormatString(9) = "d10"
        
        DepoAmt = DepoAmt + PRHist.DirectDeposit
        TotAmt = TotAmt + PRHist.Net + PRHist.DirectDeposit
        
        PrintValue(10) = " ":                       FormatString(10) = "~"
        FormatPrint
        Ln = Ln + 1
        
        trs!Amount = trs!Amount + PRHist.Net + PRHist.DirectDeposit
        trs!Number = trs!Number + 1
        
        trs.Update
        NoRecords = NoRecords + 1
        If Not PRHist.GetNext Then Exit Do
        
    Loop
    Ln = Ln + 1
    PrintValue(1) = " FINAL TOTAL:":                FormatString(1) = "a59"
    PrintValue(2) = CheckAmt:                       FormatString(2) = "d8"
            
    PrintValue(3) = " ":                            FormatString(3) = "a5"
    PrintValue(4) = DepoAmt:                        FormatString(4) = "d8"
                
    PrintValue(5) = " ":                            FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = "---------------  SUMMARY  ----------------------------"
    FormatString(1) = "a40"
    
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "Date ":                        FormatString(1) = "a9"
    PrintValue(2) = " ":                            FormatString(2) = "a5"
        
    PrintValue(3) = "# of Checks":                  FormatString(3) = "a15"
    PrintValue(4) = "    Amount ":                  FormatString(4) = "a12"
        
    PrintValue(5) = " ":                            FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = "------------------------------------------------------"
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
            
            PrintValue(1) = " ":                    FormatString(1) = "a67"
            PrintValue(2) = "Check":                FormatString(2) = "a5"
                            
            PrintValue(3) = " ":                    FormatString(3) = "a8"
            PrintValue(4) = "Direct Deposit":       FormatString(4) = "a18"
            
            PrintValue(5) = " ":                    FormatString(5) = "~"
            FormatPrint
            Ln = Ln + 1
                        
            PrintValue(1) = " ":                    FormatString(1) = "a0"
            PrintValue(2) = "Check No.":            FormatString(2) = "a9"
            
            PrintValue(3) = " ":                    FormatString(3) = "a2"
            PrintValue(4) = "P/E Date":             FormatString(4) = "a8"
            
            PrintValue(5) = " ":                    FormatString(5) = "a5"
            PrintValue(6) = "Payee":                FormatString(6) = "a5"

            PrintValue(7) = " ":                    FormatString(7) = "a38"
            PrintValue(8) = "Amount":               FormatString(8) = "a6"

            PrintValue(9) = " ":                    FormatString(9) = "a11"
            PrintValue(10) = "Amount":              FormatString(10) = "a6"

            PrintValue(11) = " ":                   FormatString(11) = "~"
            FormatPrint
                        
            PrintValue(1) = "______________________________________________________________________________________________________________"
            FormatString(1) = "a100"
            
            PrintValue(2) = " ":                    FormatString(2) = "~"
            FormatPrint
            Ln = Ln + 2
            
        End If
        
        PrintValue(1) = Format(trs!Date, "yyyymmdd"):   FormatString(1) = "n8"
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


Public Sub Form1099()

Dim LastYear As Long
Dim StartYRMO As Long
Dim EndYRMO As Long

Dim EmpCnt As Long
Dim Line3 As Currency
Dim Line4 As Currency
Dim Line7 As Currency
Dim Line16a As Currency
Dim Line17a As Long
Dim Line18a As Currency
Dim Line16b As Currency
Dim Line17b As Long
Dim Line18b As Currency

Line3 = 300
Line4 = 400
Line7 = 700
Line16 = 1600
Line18 = 1800
    
    PrtInit ("Port")
    SetEquates
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 10, Equate.Portrait
    Ln = 0
                   
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = " "
    End If
    
    SQLString = "SELECT * FROM PRemployee"

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Form 1099"
        Exit Sub
    End If
    
    Do
        If EmpCnt = 1 Then
            Ln = Ln + 9
        Else
            Ln = Ln + 6
        End If
                
        PrintValue(1) = " ":                        FormatString(1) = "a10"
        PrintValue(2) = PRCompany.Name:             FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a10"
        PrintValue(2) = PRCompany.Address1:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a10"
        PrintValue(2) = PRCompany.Address2:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1

        PrintValue(1) = " ":                        FormatString(1) = "a10"
        PrintValue(2) = Trim(PRCompany.City) & ", " & _
                        StateAbbrev & "  " & PRCompany.ZipCode
                                                    FormatString(2) = "a50"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 3
        
        PrintValue(1) = " ":                        FormatString(1) = "a48"
        PrintValue(2) = Line3:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a3"
        PrintValue(4) = Line4:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 4
        
        PrintValue(1) = " ":                        FormatString(1) = "a8"
        PrintValue(2) = PRCompany.FederalID:        FormatString(2) = "a10"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Format(PREmployee.SSN, "000-00-0000"):
                                                    FormatString(4) = "a11"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a8"
        PrintValue(2) = PREmployee.FLName:          FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a0"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a48"
        PrintValue(2) = Line7:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a8"
        PrintValue(2) = PREmployee.Address1:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a8"
        PrintValue(2) = PREmployee.Address2:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a8"
        PrintValue(2) = Trim(PREmployee.City) & ", " & _
                        PREmployee.State & "  " & _
                        PREmployee.ZipCode:
                                                    FormatString(2) = "a50"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 6
        
        PrintValue(1) = " ":                        FormatString(1) = "a48"
        PrintValue(2) = Line16a:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line17a:                    FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line18a:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a48"
        PrintValue(2) = Line16b:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line17b:                    FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line18b:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        
        EmpCnt = EmpCnt + 1
        If EmpCnt = 2 Then
            FormFeed
            EmpCnt = 0
        End If
                    
        If Not PREmployee.GetNext Then Exit Do
    Loop

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub Form1096()

Dim XXPhone As String
Dim Line3 As Long
Dim Line4 As Currency
Dim Line5 As Currency
Dim Misc1099 As String

Line3 = 3
Line4 = 400
Line5 = 500

XXPhone = "XXX-XXX-XXXX"
Misc1099 = "XXX"

    SetEquates
    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 11, Equate.Portrait
    Ln = 0
                   
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    End If

    Ln = Ln + 10
    
    PrintValue(1) = " ":                        FormatString(1) = "a8"
    PrintValue(2) = PRCompany.Name:             FormatString(2) = "a30"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 3
    
    PrintValue(1) = " ":                        FormatString(1) = "a8"
    PrintValue(2) = PRCompany.Address1:         FormatString(2) = "a30"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a8"
    PrintValue(2) = PRCompany.Address2:         FormatString(2) = "a30"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = " ":                        FormatString(1) = "a8"
    PrintValue(2) = Trim(PRCompany.City) & ", " & _
                    StateAbbrev & "  " & PRCompany.ZipCode
                                                FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a36"
    PrintValue(2) = Format(XXPhone, "000-000-0000"): FormatString(2) = "a12"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 4

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.FederalID:        FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "a21"
    PrintValue(4) = Line3:                      FormatString(4) = "n2"
    PrintValue(5) = " ":                        FormatString(5) = "a7"
    PrintValue(6) = Line4:                      FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "a6"
    PrintValue(8) = Line5:                      FormatString(8) = "d12"
    PrintValue(9) = " ":                        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 9
    
    PrintValue(1) = " ":                        FormatString(1) = "a6"
    PrintValue(2) = Misc1099:                   FormatString(2) = "a3"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub FormW3()
Dim Line1 As Currency
Dim Line2 As Currency
Dim Line3 As Currency
Dim Line4 As Currency
Dim Line5 As Currency
Dim Line6 As Currency
Dim Line7 As Currency
Dim Line8 As Currency
Dim Line9 As Currency
Dim Line10 As Currency
Dim Line11 As Currency
Dim Line12 As Currency
Dim Line13 As Currency
Dim Line14 As Currency
Dim Line15 As Currency
Dim Line16 As Currency
Dim Line17 As Currency
Dim Line18 As Currency
Dim Line19 As Currency
Dim Linec As Long
Dim h As String

Line1 = 100
Line2 = 200
Line3 = 300
Line4 = 400
Line5 = 500
Line6 = 600
Line7 = 700
Line8 = 800
Line9 = 900
Line10 = 1000
Line11 = 1100
Line12 = 1200
Line13 = 1300
Line14 = 1400
Line16 = 1600
Line17 = 1700
Line18 = 1800
Line19 = 1900
Linec = c
lineh = HH - HHHHHHH

    SetEquates
    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 11, Equate.Portrait
    Ln = 0
                   
    If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = PRState.StateAbbrev
    End If

    Ln = Ln + 7
    
    PrintValue(1) = " ":                        FormatString(1) = "a47"
    PrintValue(2) = Line1:                      FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Line2:                      FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a47"
    PrintValue(2) = Line3:                      FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Line4:                      FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a6"
    PrintValue(2) = Linec:                      FormatString(2) = "n2"
    PrintValue(3) = " ":                        FormatString(3) = "a39"
    PrintValue(4) = Line5:                      FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a8"
    PrintValue(6) = Line6:                      FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.FederalID:        FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "a30"
    PrintValue(4) = Line7:                      FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a8"
    PrintValue(6) = Line8:                      FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.Name:             FormatString(2) = "a30":
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Line9:                      FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a8"
    PrintValue(6) = Line10:                     FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 3

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.Address1:         FormatString(2) = "a30"
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Line11:                     FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a8"
    PrintValue(6) = Line12:                     FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.Address2:         FormatString(2) = "a30"

    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = Trim(PRCompany.City) & _
                    ", " & StateAbbrev:
                                                FormatString(2) = "a30"
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Line13:                     FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = PRCompany.ZipCode:          FormatString(2) = "a9"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " ":                        FormatString(1) = "a47"
    PrintValue(2) = Line14:                     FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = h:                          FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = StateAbbrev:                FormatString(2) = "a2"  '  Line 15
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = PRCompany.AddrStateID:      FormatString(4) = "n2"
    PrintValue(5) = " ":                        FormatString(5) = "a28"
    PrintValue(6) = Line16:                     FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "a8"
    PrintValue(8) = Line17:                     FormatString(8) = "d12"
    PrintValue(9) = " ":                        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a47"
    PrintValue(2) = Line18:                     FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Line19:                     FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub FormW2TwoPerPg()

Dim EmpCnt As Long

Dim Line1 As Currency
Dim Line2 As Currency
Dim Line3 As Currency
Dim Line4 As Currency
Dim Line5 As Currency
Dim Line6 As Currency
Dim Line7 As Currency
Dim Line8 As Currency
Dim Line9 As Currency
Dim Line10 As Currency
Dim Line11 As Currency
Dim Line12a1 As String
Dim Line12a2 As Currency
Dim Line12b1 As String
Dim Line12b2 As Currency
Dim Line12c1 As String
Dim Line12c2 As Currency
Dim Line12d1 As String
Dim Line12d2 As Currency
Dim Line13a As String
Dim Line13b As String
Dim Line13c As String
Dim Line14a As Currency
Dim Line14b As Currency
Dim Line14c As Currency
Dim Line14d As Currency
Dim Line15a1 As String
Dim Line15a2 As Long
Dim Line15b1 As String
Dim Line15b2 As Long
Dim Line16a As Long
Dim Line17a As Currency
Dim Line18a As Currency
Dim Line19a As Currency
Dim Line20a As String
Dim Line16b As Long
Dim Line17b As Currency
Dim Line18b As Currency
Dim Line19b As Currency
Dim Line20b As String
Dim LineD As Long

Line1 = 100
Line2 = 200
Line3 = 300
Line4 = 400
Line5 = 500
Line6 = 600
Line7 = 700
Line8 = 800
Line9 = 900
Line10 = 1000
Line11 = 1100
Line12a1 = 1
Line12a2 = 1201
Line13a = "X"
Line13b = "X"
Line13c = "X"
Line12b1 = 2
Line12b2 = 1202
Line14a = 1401
Line14b = 1402
Line14c = 1403
Line14d = 1404
Line12c1 = 3
Line12c2 = 1203
Line12d1 = 4
Line12d2 = 1204
Line16a = 1601
Line17a = 1701
Line18a = 1801
Line19a = 1901
Line20a = "ABCDEFG"
Line16b = 1602
Line17b = 1702
Line18b = 1802
Line19b = 1902
Line20b = "HIJKLMN"
LineD = 0           ' Control Number
Line15a1 = "OH"
Line15a2 = 34
Line15b1 = "PA"
Line15b2 = 42

    
    SetEquates
    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 10, Equate.Portrait
    Ln = 0
                   
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = " "
    Else
        StateAbbrev = PRState.StateAbbrev
    End If
    
    SQLString = "SELECT * FROM PRemployee"

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Form 1099"
        Exit Sub
    End If
    
    Do
        If EmpCnt = 1 Then
            Ln = Ln + 11
        Else
            Ln = Ln + 5
        End If
        
        PrintValue(1) = " ":                        FormatString(1) = "a28"
        PrintValue(2) = Format(PREmployee.SSN, "000-00-0000"):
                                                    FormatString(2) = "a11"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PRCompany.FederalID:        FormatString(2) = "a10"
        PrintValue(3) = " ":                        FormatString(3) = "a46"
        PrintValue(4) = Line1:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a7"
        PrintValue(6) = Line2:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PRCompany.Name:             FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a26"
        PrintValue(4) = Line3:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a7"
        PrintValue(6) = Line4:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PRCompany.Address1:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PRCompany.Address2:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a26"
        PrintValue(4) = Line5:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a7"
        PrintValue(6) = Line6:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = Trim(PRCompany.City) & ", " & _
                        StateAbbrev & "  " & PRCompany.ZipCode
                                                    FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
    
        LineD = LineD + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a63"
        PrintValue(2) = Line7:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a7"
        PrintValue(4) = Line8:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 3
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = LineD:                      FormatString(2) = "n4":
        PrintValue(3) = " ":                        FormatString(3) = "a52"
        PrintValue(4) = Line9:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a7"
        PrintValue(6) = Line10:                     FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PREmployee.FLName:          FormatString(2) = "a30":
        PrintValue(3) = " ":                        FormatString(3) = "a26"
        PrintValue(4) = Line11:                     FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line12a1:                   FormatString(6) = "a1"
        PrintValue(7) = " ":                        FormatString(7) = "a1"
        PrintValue(8) = Line12a2:                   FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PREmployee.Address1:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = PREmployee.Address2:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a23"
        PrintValue(4) = Line13a:                    FormatString(4) = "a1"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line13b:                    FormatString(6) = "a1"
        PrintValue(7) = " ":                        FormatString(7) = "a6"
        PrintValue(8) = Line13c:                    FormatString(8) = "a1"
        PrintValue(9) = " ":                        FormatString(9) = "a8"
        PrintValue(10) = Line12b1:                  FormatString(10) = "a1"
        PrintValue(11) = " ":                       FormatString(11) = "a1"
        PrintValue(12) = Line12b2:                  FormatString(12) = "d12"
        PrintValue(13) = " ":                       FormatString(13) = "~"
        FormatPrint
        Ln = Ln + 1
        
        
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = Trim(PREmployee.City) & ", " & _
                        PREmployee.State & "  " & _
                        PREmployee.ZipCode:
                                                    FormatString(2) = "a50"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a63"
        PrintValue(2) = Line14a:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a4"
        PrintValue(5) = Line12c1:                   FormatString(5) = "a1"
        PrintValue(6) = " ":                        FormatString(6) = "a1"
        PrintValue(7) = Line12c2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a63"
        PrintValue(2) = Line14b:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a63"
        PrintValue(2) = Line14c:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a4"
        PrintValue(5) = Line12d1:                   FormatString(5) = "a1"
        PrintValue(6) = " ":                        FormatString(6) = "a1"
        PrintValue(7) = Line12d2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a63"
        PrintValue(2) = Line14d:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "~"
        FormatPrint
        Ln = Ln + 2
                    
        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = Line15a1:                   FormatString(2) = "a2"
        PrintValue(3) = " ":                        FormatString(3) = "a8"
        PrintValue(4) = Line15a2:                   FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a13"
        PrintValue(6) = Line16a:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a0"
        PrintValue(8) = Line17a:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a2"
        PrintValue(10) = Line18a:                   FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a0"
        PrintValue(12) = Line19a:                   FormatString(12) = "d12"
        PrintValue(13) = " ":                       FormatString(13) = "a0"
        PrintValue(14) = Line20a:                   FormatString(14) = "a7"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        Ln = Ln + 2

        PrintValue(1) = " ":                        FormatString(1) = "a7"
        PrintValue(2) = Line15b1:                   FormatString(2) = "a2"
        PrintValue(3) = " ":                        FormatString(3) = "a8"
        PrintValue(4) = Line15b2:                   FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a13"
        PrintValue(6) = Line16b:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a0"
        PrintValue(8) = Line17b:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a2"
        PrintValue(10) = Line18b:                   FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a0"
        PrintValue(12) = Line19b:                   FormatString(12) = "d12"
        PrintValue(13) = " ":                       FormatString(13) = "a0"
        PrintValue(14) = Line20b:                   FormatString(14) = "a7"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        
        EmpCnt = EmpCnt + 1
        If EmpCnt = 2 Then
            FormFeed
            EmpCnt = 0
        End If
                    
        If Not PREmployee.GetNext Then Exit Do
        
    Loop
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub FormW2FourPerPg()

Dim EmpCnt As Long

Dim LineA As Long
Dim Line1 As Currency
Dim Line2 As Currency
Dim Line3 As Currency
Dim Line4 As Currency
Dim LineB As String

Dim Line5 As Currency
Dim Line6 As Currency
Dim Line7 As Currency
Dim Line8 As Currency
Dim Line9 As Currency
Dim Line10 As Currency
Dim Line11 As Currency
Dim Line12a1 As String
Dim Line12a2 As Currency
Dim Line12b1 As String
Dim Line12b2 As Currency
Dim Line12c1 As String
Dim Line12c2 As Currency
Dim Line12d1 As String
Dim Line12d2 As Currency
Dim Line13a As Currency
Dim Line13b As Currency
Dim Line13c As Currency
Dim Line14a As Currency
Dim Line14b As Currency
Dim Line14c As Currency
Dim Line14d As Currency
Dim Line15a1 As String
Dim Line15a2 As Long
Dim Line15b1 As String
Dim Line15b2 As Long
Dim Line16a As Currency
Dim Line17a As Currency
Dim Line18a As Currency
Dim Line19a As Currency
Dim Line20a As String
Dim Line16b As Currency
Dim Line17b As Currency
Dim Line18b As Currency
Dim Line19b As Currency
Dim Line20b As String
Dim LineD As Long

Line1 = 100
Line2 = 200
Line3 = 300
Line4 = 400
Line5 = 500
Line6 = 600
Line7 = 700
Line8 = 800
Line9 = 900
Line10 = 1000
Line11 = 1100
Line12a1 = 1
Line12a2 = 1200
Line13a = 1300
Line13b = 1310
Line13c = 1320
Line12b1 = 2
Line12b2 = 1202
Line12c1 = 3
Line12c2 = 1203
Line12d1 = 4
Line12d2 = 1204
Line13a = 1300
Line13b = 1310
Line13c = 1320
Line14a = 1401
Line14b = 1402
Line14c = 1403
Line14d = 1404
Line16a = 1601
Line17a = 1701
Line18a = 1801
Line19a = 1901
Line20a = "ABCDEFG"
Line16b = 1602
Line17b = 1702
Line18b = 1802
Line19b = 1902
Line20b = "HIJKLMN"
LineD = 0           ' Control Number
Line15a1 = "OH"
Line15a2 = 34
Line15b1 = "PA"
Line15b2 = 42

    
    SetEquates
    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 8, Equate.Portrait
    Ln = 0
                   
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = " "
    Else
        StateAbbrev = PRState.StateAbbrev
    End If
    
    SQLString = "SELECT * FROM PRemployee"

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Form 1099"
        Exit Sub
    End If
    
    Do
        If EmpCnt = 1 Then
            Ln = Ln + 5
        Else
            Ln = Ln + 4
        End If
                                                   
        PrintValue(1) = " ":                        FormatString(1) = "a25"
        PrintValue(2) = Line1:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a5"
        PrintValue(4) = Line2:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a29"
        PrintValue(6) = Line1:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a2"
        PrintValue(8) = Line2:                      FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = Format(PREmployee.SSN, "000-00-0000"):
                                                    FormatString(2) = "a11"

        PrintValue(3) = " ":                        FormatString(3) = "a52"
        PrintValue(4) = Format(PREmployee.SSN, "000-00-0000"):
                                                    FormatString(4) = "a11"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a25"
        PrintValue(2) = Line3:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a5"
        PrintValue(4) = Line4:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a29"
        PrintValue(6) = Line3:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a2"
        PrintValue(8) = Line4:                      FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a5"
        PrintValue(2) = PRCompany.FederalID:        FormatString(2) = "a10"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line5:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line6:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a10"
        PrintValue(8) = PRCompany.FederalID:        FormatString(8) = "a10"
        PrintValue(9) = " ":                        FormatString(9) = "a9"
        PrintValue(10) = Line5:                     FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a2"
        PrintValue(12) = Line6:                     FormatString(12) = "d12"
        PrintValue(13) = " ":                       FormatString(13) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PRCompany.Name:             FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PRCompany.Name:             FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PRCompany.Address1:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PRCompany.Address1:         FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PRCompany.Address2:         FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PRCompany.Address2:         FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = Trim(PRCompany.City) & ", " & _
                        StateAbbrev & "  " & PRCompany.ZipCode
                                                    FormatString(2) = "a50"
        PrintValue(3) = " ":                        FormatString(3) = "a14"
        PrintValue(4) = Trim(PRCompany.City) & ", " & _
                        StateAbbrev & "  " & PRCompany.ZipCode
                                                    FormatString(4) = "a50"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
    
        LineD = LineD + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = LineD:                      FormatString(2) = "n4"
        PrintValue(3) = " ":                        FormatString(3) = "a60"
        PrintValue(4) = LineD:                      FormatString(4) = "n4"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
                
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PREmployee.LFName:          FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PREmployee.LFName:          FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PREmployee.Address1:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PREmployee.Address1:        FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = PREmployee.Address2:        FormatString(2) = "a30"
        PrintValue(3) = " ":                        FormatString(3) = "a34"
        PrintValue(4) = PREmployee.Address2:        FormatString(4) = "a30"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a4"
        PrintValue(2) = Trim(PREmployee.City) & ", " & _
                        StateAbbrev & "  " & PREmployee.ZipCode
                                                    FormatString(2) = "a50"
        PrintValue(3) = " ":                        FormatString(3) = "a14"
        PrintValue(4) = Trim(PREmployee.City) & ", " & _
                        StateAbbrev & "  " & PREmployee.ZipCode
                                                    FormatString(4) = "a50"
        PrintValue(5) = " ":                        FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a5"
        PrintValue(2) = Line7:                      FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line8:                      FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = Line9:                      FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a9"
        PrintValue(8) = Line7:                      FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a6"
        PrintValue(10) = Line8:                     FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a2"
        PrintValue(12) = Line9:                     FormatString(12) = "d12"
        PrintValue(13) = " ":                       FormatString(13) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a5"
        PrintValue(2) = Line10:                     FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line11:                     FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a4"
        PrintValue(6) = Line12a1:                   FormatString(6) = "a1"
        PrintValue(7) = Line12a2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "a9"
        PrintValue(9) = Line10:                     FormatString(9) = "d12"
        PrintValue(10) = " ":                       FormatString(10) = "a6"
        PrintValue(11) = Line11:                    FormatString(11) = "d12"
        PrintValue(12) = " ":                       FormatString(12) = "a1"
        PrintValue(13) = Line12a1:                  FormatString(13) = "a1"
        PrintValue(14) = Line12a2:                  FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        Ln = Ln + 1
        
        PrintValue(1) = " ":                        FormatString(1) = "a1"
        PrintValue(2) = Line13a:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line14a:                    FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a4"
        PrintValue(6) = Line12b1:                   FormatString(6) = "a1"
        PrintValue(7) = Line12b2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "a4"
        PrintValue(9) = Line13a:                    FormatString(9) = "d12"
        PrintValue(10) = " ":                       FormatString(10) = "a11"
        PrintValue(11) = Line14a:                   FormatString(11) = "d12"
        PrintValue(12) = " ":                       FormatString(12) = "a1"
        PrintValue(13) = Line12b1:                  FormatString(13) = "a1"
        PrintValue(14) = Line12b2:                  FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a1"
        PrintValue(2) = Line13b:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line14b:                    FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a4"
        PrintValue(6) = Line12c1:                   FormatString(6) = "a1"
        PrintValue(7) = Line12c2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "a4"
        PrintValue(9) = Line13b:                    FormatString(9) = "d12"
        PrintValue(10) = " ":                       FormatString(10) = "a11"
        PrintValue(11) = Line14b:                   FormatString(11) = "d12"
        PrintValue(12) = " ":                       FormatString(12) = "a1"
        PrintValue(13) = Line12c1:                  FormatString(13) = "a1"
        PrintValue(14) = Line12c2:                  FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        Ln = Ln + 2
        
        PrintValue(1) = " ":                        FormatString(1) = "a1"
        PrintValue(2) = Line13c:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a10"
        PrintValue(4) = Line14c:                    FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a4"
        PrintValue(6) = Line12d1:                   FormatString(6) = "a1"
        PrintValue(7) = Line12d2:                   FormatString(7) = "d12"
        PrintValue(8) = " ":                        FormatString(8) = "a4"
        PrintValue(9) = Line13c:                    FormatString(9) = "d12"
        PrintValue(10) = " ":                       FormatString(10) = "a11"
        PrintValue(11) = Line14c:                   FormatString(11) = "d12"
        PrintValue(12) = " ":                       FormatString(12) = "a1"
        PrintValue(13) = Line12d1:                  FormatString(13) = "a1"
        PrintValue(14) = Line12d2:                  FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "~"
        FormatPrint
        Ln = Ln + 1
        
                    
        PrintValue(1) = " ":                        FormatString(1) = "a2"
        PrintValue(2) = Line15a1:                   FormatString(2) = "a2"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line15a2:                   FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a13"
        PrintValue(6) = Line16a:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a5"
        PrintValue(8) = Line17a:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a6"
        PrintValue(10) = Line15a1:                  FormatString(10) = "a2"
        PrintValue(11) = " ":                       FormatString(11) = "a6"
        PrintValue(12) = Line15a2:                  FormatString(12) = "n2"
        PrintValue(13) = " ":                       FormatString(13) = "a13"
        PrintValue(14) = Line16a:                   FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "a2"
        PrintValue(16) = Line17a:                   FormatString(16) = "d12"
        PrintValue(17) = " ":                       FormatString(17) = "~"
        FormatPrint
        Ln = Ln + 1

        PrintValue(1) = " ":                        FormatString(1) = "a2"
        PrintValue(2) = Line15b1:                   FormatString(2) = "a2"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line15b2:                   FormatString(4) = "n2"
        PrintValue(5) = " ":                        FormatString(5) = "a13"
        PrintValue(6) = Line16b:                    FormatString(6) = "d12"
        PrintValue(7) = " ":                        FormatString(7) = "a5"
        PrintValue(8) = Line17b:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a6"
        PrintValue(10) = Line15b1:                  FormatString(10) = "a2"
        PrintValue(11) = " ":                       FormatString(11) = "a6"
        PrintValue(12) = Line15b2:                  FormatString(12) = "n2"
        PrintValue(13) = " ":                       FormatString(13) = "a13"
        PrintValue(14) = Line16b:                   FormatString(14) = "d12"
        PrintValue(15) = " ":                       FormatString(15) = "a2"
        PrintValue(16) = Line17b:                   FormatString(16) = "d12"
        PrintValue(17) = " ":                       FormatString(17) = "~"
        FormatPrint
        Ln = Ln + 2
                
        PrintValue(1) = " ":                        FormatString(1) = "a5"
        PrintValue(2) = Line18a:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line19a:                    FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a10"
        PrintValue(6) = Line20a:                    FormatString(6) = "a7"
        PrintValue(7) = " ":                        FormatString(7) = "a11"
        PrintValue(8) = Line18a:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a6"
        PrintValue(10) = Line19a:                   FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a7"
        PrintValue(12) = Line20a:                   FormatString(12) = "a7"
        PrintValue(13) = " ":                       FormatString(13) = "~"
        FormatPrint
        Ln = Ln + 1

        PrintValue(1) = " ":                        FormatString(1) = "a5"
        PrintValue(2) = Line18b:                    FormatString(2) = "d12"
        PrintValue(3) = " ":                        FormatString(3) = "a6"
        PrintValue(4) = Line19b:                    FormatString(4) = "d12"
        PrintValue(5) = " ":                        FormatString(5) = "a10"
        PrintValue(6) = Line20b:                    FormatString(6) = "a7"
        PrintValue(7) = " ":                        FormatString(7) = "a11"
        PrintValue(8) = Line18b:                    FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a6"
        PrintValue(10) = Line19b:                   FormatString(10) = "d12"
        PrintValue(11) = " ":                       FormatString(11) = "a7"
        PrintValue(12) = Line20b:                   FormatString(12) = "a7"
        PrintValue(13) = " ":                       FormatString(13) = "~"
        FormatPrint
        Ln = Ln + 2
        EmpCnt = EmpCnt + 1
        If EmpCnt = 2 Then
            FormFeed
            EmpCnt = 0
        End If
                    
        If Not PREmployee.GetNext Then Exit Do
        
    Loop
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub W2File()
Dim W2File As String

' Dim Vars for "RA" Record
Dim FedID As Long
Dim PIN As String
Dim VendCode As Long
Dim ReSub As String
Dim WFID As String
Dim SoftCode As Long
Dim ZipCodeExt As String
Dim RAFSP As String
Dim RAFPC As String
Dim RACC As String
Dim SubName As String
Dim SubAddr As String
Dim SubDelivAddr As String
Dim SubCity As String
Dim SubState As String
Dim SubZip1 As String
Dim SubZip2 As String
Dim SubFSP As String
Dim SubFPC As String
Dim SubCC As String
Dim ContactName As String
Dim ContactPH As String
Dim ContactPHExt As String
Dim ContactEmail As String
Dim ContactFax As String
Dim ProbCode As Byte
Dim PrepCode As String

' Dim Vars for "RE" Record
Dim TaxYear As Long
Dim AgentCode As String
Dim EmpEin As Long
Dim AgentEIN As Long
Dim TermBusInd As Byte
Dim EstabNo As String
Dim OtherEIN As String
Dim REFSP As String
Dim REFPC As String
Dim RECC As String
Dim EmployCode As String
Dim TaxJurisCode As String
Dim RESickPayIndic As Byte

' Dim Vars for "RW" Records

Dim RWFSP As String
Dim RWFPC As String
Dim RWCC As String
Dim Wages As Currency
Dim FWT As Currency
Dim SSWage As Currency
Dim SSTax As Currency
Dim MEDWage As Currency
Dim MedTax As Currency
Dim SSTips As Currency
Dim AEIC As Currency
Dim DCB As Currency
Dim DC401K As Currency
Dim DC403B As Currency
Dim DC408K As Currency
Dim DC457B As Currency
Dim DC501C As Currency
Dim MilPay As Currency
Dim NQ457 As Currency
Dim HlthSvgs As Currency
Dim NQNot457 As Currency
Dim NonTaxCombat As Currency
Dim Premiums As Currency
Dim NonStatInc As Currency
Dim Def409A As Currency
Dim Roth401K As Currency
Dim Roth403B As Currency
Dim StatEmpIndic As Byte
Dim RetireIndic As Byte
Dim RWSickPayIndic As Byte

' Dim Vars for "RO" Records
Dim AllocTips As Currency
Dim UncollTaxTips As Currency
Dim MedSvgsAcct As Currency
Dim RetireAcct As Currency
Dim AdoptExp As Currency
Dim UncollSSRRTATax As Currency
Dim UncollMedTax As Currency
Dim Income409A As Currency
Dim PRWages As Currency
Dim PRComm As Currency
Dim PRAllow As Currency
Dim PRTips As Currency
Dim PRTotalInc As Currency
Dim PRTaxWH As Currency
Dim PRRetFundContr As Currency
Dim OtherIncTax As Currency
Dim OtherIncTaxWH As Currency

' Dim Vars for "RS" records
Dim OrgStCode As Long
Dim RSZipExt As String
Dim RSFSP As String
Dim RSFPC As String
Dim RSCC As String
Dim RptPer As Integer
Dim StQUnempWgs As Currency
Dim STQUnempTxWgs As Currency
Dim WksWorked As Long
Dim DtEmpl As Double
Dim DtSep As Double
Dim StEmplrAcct As String
Dim PostStCode As Integer
Dim OHTxWgs As Currency
Dim OHIncTxWH As Currency
Dim WgsTipsEtc As Currency
Dim TxType As String
Dim LocTxWgs As Currency
Dim LocIncTxWH As Currency
Dim SchDistNo As Integer

' Dim Vars for "RT" records
Dim TRWs As Long
Dim TWgs As Currency
Dim TFWT As Currency
Dim TSSWage As Currency
Dim TSSTax As Currency
Dim TMedWage As Currency
Dim TMedTax As Currency
Dim TSSTips As Currency
Dim TAEIC As Currency
Dim TDCB As Currency
Dim TDC401K As Currency
Dim TDC403B As Currency
Dim TDC408K As Currency
Dim TDC457B As Currency
Dim TDC501C As Currency
Dim TMilPay As Currency
Dim TNQ457 As Currency
Dim THlthSvgs As Currency
Dim TNQNot457 As Currency
Dim TNonTaxCombat As Currency
Dim TPremiums As Currency
Dim TSickTax As Currency
Dim TNonStatInc As Currency
Dim TDef409A As Currency
Dim TRoth401K As Currency
Dim TRoth403B As Currency

' Pop RA Vars
FedID = 340792940
PIN = "PIN"
VendCode = 1111
ReSub = 1
WFID = "WFID"
SoftCode = 98
ZipCodeExt = "1234"
RAFSP = "Foreign State Prov"
RAFPC = "Foreign Post Cd"
RACC = "CC"
SubName = "Submitter Name"
SubAddr = "Submitter Address"
SubDelivAddr = "Sub Delivery Address"
SubCity = "Sub City"
SubState = "SS"
SubZip1 = "SZip1"
SubZip2 = "Zip2"
SubFSP = "Sub For State Prov"
SubFPC = "Sub Post Code"
SubCC = "CC"
ContactName = "Contact Name"
ContactPH = "Contact Phone"
ContactPHExt = "PHExt"
ContactEmail = "Contact Email"
ContactFax = "Contact Fax"
ProbCode = 1
PrepCode = "A"

' Pop RE Vars
TaxYear = 2009
AgentCode = 2
EmpEin = 123456789
AgentEIN = 123456789
TermBusInd = 1
EstabNo = "1A2A"
OtherEIN = 888888888
REFSP = "Empl Foreign State Prov"
REFPC = "Emply Post Code"
RECC = "CC"
EmployCode = "Q"
TaxJurisCode = "G"
RESickPayIndic = 1

' Pop RW Vars

RWFSP = "Foreign State Province"
RWFPC = "For State Provi"
RWCC = "CC"
Wages = 1111111111
FWT = 222222222
SSWage = 333333333
SSTax = 444444444
MEDWage = 555555555
MedTax = 666666666
SSTips = 777777777
AEIC = 888888888
DCB = 999999999
DC401K = 1111111111
DC403B = 222222222
DC408K = 333333333
DC457B = 444444444
DC501C = 555555555
MilPay = 666666666
NQ457 = 777777777
HlthSvgs = 888888888
NQNot457 = 999999999
NonTaxCombat = 111111111
Premiums = 222222222
NonStatInc = 333333333
Def409A = 444444444
Roth401K = 555555555
Roth403B = 666666666
StatEmpIndic = 1
RetireIndic = 0
RWSickPayIndic = 0

' Pop RO Vars
AllocTips = 1111111111
UncollTaxTips = 222222222
MedSvgsAcct = 333333333
RetireAcct = 444444444
AdoptExp = 555555555
UncollSSRRTATax = 666666666
UncollMedTax = 777777777
Income409A = 888888888
PRWages = 999999999
PRComm = 1111111111
PRAllow = 222222222
PRTips = 333333333
PRTotalInc = 444444444
PRTaxWH = 555555555
PRRetFundContr = 666666666
OtherIncTax = 777777777
OtherIncTaxWH = 888888888

' Pop RS Vars
OrgStCode = 39
RSZipExt = "1234"
RSFSP = "Foreign State  Province"
RSFPC = "For State Provi"
RSCC = "CC"
RptPer = 32007
StQUnempWgs = 1111111111
STQUnempTxWgs = 222222222
WksWorked = 8
DtEmpl = 1312008
DtSep = 12312008
StEmplrAcct = 3201
PostStCode = 39
OHTxWgs = 333333333
OHIncTxWH = 222222222
WgsTipsEtc = 333333333
TxType = "E"
LocTxWgs = 1111111111
LocIncTxWH = 222222222
SchDistNo = 1234


'=============================================================================================

    SetEquates
    PrtInit ("Port")

    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
        
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    SetFont 8, Equate.Portrait
    Ln = 0
                   
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
        StateAbbrev = " "
    Else
        StateAbbrev = PRState.StateAbbrev
    End If
    
    TChannel = FreeFile
                               
    Open "C:\balint\W2File" For Output As #TChannel Len = 512
    
    ' Create Code "RA" - Submitter Record
    
    PrintValue(1) = "RA":                                           FormatString(1) = "a2"
    PrintValue(2) = FedID:                                          FormatString(2) = "n9"
    PrintValue(3) = PIN:                                            FormatString(3) = "a8"
    PrintValue(4) = VendCode:                                       FormatString(4) = "n4"
    PrintValue(5) = " ":                                            FormatString(5) = "a5"
    PrintValue(6) = ReSub:                                          FormatString(6) = "n1"
    PrintValue(7) = WFID:                                           FormatString(7) = "a6"
    PrintValue(8) = SoftCode:                                       FormatString(8) = "n2"
    PrintValue(9) = PRCompany.Name:                                 FormatString(9) = "a57"
    PrintValue(10) = PRCompany.Address1:                            FormatString(10) = "a22"
    PrintValue(11) = PRCompany.Address2:                            FormatString(11) = "a22"
    PrintValue(12) = PRCompany.City:                                FormatString(12) = "a22"
    PrintValue(13) = StateAbbrev:                                   FormatString(13) = "a2"
    PrintValue(14) = PRCompany.ZipCode:                             FormatString(14) = "a5"
    PrintValue(15) = ZipCodeExt:                                    FormatString(15) = "a4"
    PrintValue(16) = " ":                                           FormatString(16) = "a5"   ' blank
    PrintValue(17) = RAFSP:                                         FormatString(17) = "a23"
    PrintValue(18) = RAFPC:                                         FormatString(18) = "a15"
    PrintValue(19) = RACC:                                          FormatString(19) = "a2"
    PrintValue(20) = SubName:                                       FormatString(20) = "a57"
    PrintValue(21) = SubAddr:                                       FormatString(21) = "a22"
    PrintValue(22) = SubDelivAddr:                                  FormatString(22) = "a22"
    PrintValue(23) = SubCity:                                       FormatString(23) = "a22"
    PrintValue(24) = SubState:                                      FormatString(24) = "a2"
    PrintValue(25) = SubZip1:                                       FormatString(25) = "a5"
    PrintValue(26) = SubZip2:                                       FormatString(26) = "a4"
    PrintValue(27) = " ":                                           FormatString(27) = "a5"   ' blank
    PrintValue(28) = SubFSP:                                        FormatString(28) = "a23"
    PrintValue(29) = SubFPC:                                        FormatString(29) = "a15"
    PrintValue(30) = SubCC:                                         FormatString(30) = "a2"
    PrintValue(31) = ContactName:                                   FormatString(31) = "a27"
    PrintValue(32) = ContactPH:                                     FormatString(32) = "a15"
    PrintValue(33) = ContactPHExt:                                  FormatString(33) = "a5"
    PrintValue(34) = " ":                                           FormatString(34) = "a3"   ' blank
    PrintValue(35) = ContactEmail:                                  FormatString(35) = "a40"
    PrintValue(36) = " ":                                           FormatString(36) = "a3"   ' blank
    PrintValue(37) = ContactFax:                                    FormatString(37) = "a10"
    PrintValue(38) = ProbCode:                                      FormatString(38) = "n1"
    PrintValue(39) = PrepCode:                                      FormatString(39) = "a1"
    PrintValue(40) = " ":                                           FormatString(40) = "a12"  ' blank
    PrintValue(41) = " ":                                           FormatString(41) = "~"
    FormatPrint
 
    Print #TChannel, PrintString ' Output text.

'=============================================================================================
         
    ' Create Code "RE" - Submitter Record
    
    PrintValue(1) = "RE":                                           FormatString(1) = "a2"
    PrintValue(2) = TaxYear:                                        FormatString(2) = "n4"
    PrintValue(3) = AgentCode:                                      FormatString(3) = "n1"
    PrintValue(4) = EmpEin:                                         FormatString(4) = "n9"
    PrintValue(5) = AgentEIN:                                       FormatString(5) = "n9"
    PrintValue(6) = TermBusInd:                                     FormatString(6) = "n1"
    PrintValue(7) = EstabNo:                                        FormatString(7) = "a4"
    PrintValue(8) = OtherEIN:                                       FormatString(8) = "n9"
    PrintValue(9) = PRCompany.Name:                                 FormatString(9) = "a57"
    PrintValue(10) = PRCompany.Address1:                            FormatString(10) = "a22"
    PrintValue(11) = PRCompany.Address2:                            FormatString(11) = "a22"
    PrintValue(12) = PRCompany.City:                                FormatString(12) = "a22"
    PrintValue(13) = StateAbbrev:                                   FormatString(13) = "a2"
    PrintValue(14) = PRCompany.ZipCode:                             FormatString(14) = "a5"
    PrintValue(15) = ZipCodeExt:                                    FormatString(15) = "a4"
    PrintValue(16) = " ":                                           FormatString(16) = "a5"   ' Blank
    PrintValue(17) = REFSP:                                         FormatString(17) = "a23"
    PrintValue(18) = REFPC:                                         FormatString(18) = "a15"
    PrintValue(19) = RECC:                                          FormatString(19) = "a2"
    PrintValue(20) = EmployCode:                                    FormatString(20) = "a1"
    PrintValue(21) = TaxJurisCode:                                  FormatString(21) = "a1"
    PrintValue(22) = RESickPayIndic:                                FormatString(22) = "n1"
    PrintValue(23) = " ":                                           FormatString(23) = "a291"
    PrintValue(24) = " ":                                           FormatString(24) = "~"
    FormatPrint
    Print #TChannel, PrintString  ' Output text.

'=============================================================================================

    ' Create Code "RW" - Employee Wage Records
    SQLString = "SELECT * FROM PRemployee"

    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Data Found !!!", vbCritical, "Form 1099"
        Exit Sub
    End If

    Do
  
        PrintValue(1) = "RW":                                       FormatString(1) = "a2"
        PrintValue(2) = Format(PREmployee.SSN, "000000000"):        FormatString(2) = "r9"
        PrintValue(3) = PREmployee.FirstName:                       FormatString(3) = "a15"
        PrintValue(4) = PREmployee.MidInit:                         FormatString(4) = "a15"
        PrintValue(5) = PREmployee.LastName:                        FormatString(5) = "a20"
        PrintValue(6) = PREmployee.MidInit:                         FormatString(6) = "a4"
        PrintValue(7) = PREmployee.Address1:                        FormatString(7) = "a22"
        PrintValue(8) = PREmployee.Address2:                        FormatString(8) = "a22"
        PrintValue(9) = PREmployee.City:                            FormatString(9) = "a22"
        PrintValue(10) = PREmployee.State:                          FormatString(10) = "a2"
        PrintValue(11) = PREmployee.ZipCode:                        FormatString(11) = "a5"
        PrintValue(12) = ZipCodeExt:                                FormatString(12) = "a4"
        PrintValue(13) = " ":                                       FormatString(13) = "a5"   ' blank
        PrintValue(14) = RWFSP:                                     FormatString(14) = "a23"
        PrintValue(15) = RWFPC:                                     FormatString(15) = "a15"
        PrintValue(16) = RWCC:                                      FormatString(16) = "a2"
        PrintValue(17) = Format(Wages, "00000000000"):              FormatString(17) = "r11"
        PrintValue(18) = Format(FWT, "00000000000"):                FormatString(18) = "r11"
        PrintValue(19) = Format(SSWage, "00000000000"):             FormatString(19) = "r11"
        PrintValue(20) = Format(SSTax, "00000000000"):              FormatString(20) = "r11"
        PrintValue(21) = Format(MEDWage, "00000000000"):            FormatString(21) = "r11"
        PrintValue(22) = Format(MedTax, "00000000000"):             FormatString(22) = "r11"
        PrintValue(23) = Format(SSTips, "00000000000"):             FormatString(23) = "r11"
        PrintValue(24) = Format(AEIC, "00000000000"):               FormatString(24) = "r11"
        PrintValue(25) = Format(DCB, "00000000000"):                FormatString(25) = "r11"
        PrintValue(26) = Format(DC401K, "00000000000"):             FormatString(26) = "r11"
        PrintValue(27) = Format(DC403B, "00000000000"):             FormatString(27) = "r11"
        PrintValue(28) = Format(DC408K, "00000000000"):             FormatString(28) = "r11"
        PrintValue(29) = Format(DC457B, "00000000000"):             FormatString(29) = "r11"
        PrintValue(30) = Format(DC501C, "00000000000"):             FormatString(30) = "r11"
        PrintValue(31) = Format(MilPay, "00000000000"):             FormatString(31) = "r11"
        PrintValue(32) = Format(NQ457, "00000000000"):              FormatString(32) = "r11"
        PrintValue(33) = Format(HlthSvgs, "00000000000"):           FormatString(33) = "r11"
        PrintValue(34) = Format(NQNot457, "00000000000"):           FormatString(34) = "r11"
        PrintValue(35) = Format(NonTaxCombat, "00000000000"):       FormatString(35) = "r11"
        PrintValue(36) = " ":                                       FormatString(36) = "a11"
        PrintValue(37) = Format(Premiums, "00000000000"):           FormatString(37) = "r11"
        PrintValue(38) = Format(NonStatInc, "00000000000"):         FormatString(38) = "r11"
        PrintValue(39) = Format(Def409A, "00000000000"):            FormatString(39) = "r11"
        PrintValue(40) = Format(Roth401K, "00000000000"):           FormatString(40) = "r11"
        PrintValue(41) = Format(Roth403B, "00000000000"):           FormatString(41) = "r11"
        PrintValue(42) = " ":                                       FormatString(42) = "a23"
        PrintValue(43) = StatEmpIndic:                              FormatString(43) = "n1"
        PrintValue(44) = " ":                                       FormatString(44) = "a1"
        PrintValue(45) = RetireIndic:                               FormatString(45) = "n1"
        PrintValue(46) = RWSickPayIndic:                            FormatString(46) = "n1"
        PrintValue(47) = " ":                                       FormatString(47) = "a23"
        PrintValue(48) = " ":                                       FormatString(48) = "~"
        
        FormatPrint
        
        TRWs = TRWs + 1
        TWgs = TWgs + Wages
        TFWT = TFWT + FWT
        TSSWage = TSSWage + SSWage
        TSSTax = TSSTax + SSTax
        TMedWage = TMedWage + MEDWage
        TMedTax = TMedTax + MedTax
        TSSTips = TSSTips + SSTips
        TAEIC = TAEIC + AEIC
        TDCB = TDCB + DCB
        TDC401K = TDC401K + DC401K
        TDC403B = TDC403B + DC403B
        TDC408K = TDC408K + DC408K
        TDC457B = TDC457B + DC457B
        TDC501C = TDC501C + DC501C
        TMilPay = TMilPay + MilPay
        TNQ457 = TNQ457 + NQ457
        THlthSvgs = THlthSvgs + HlthSvgs
        TNQNot457 = TNQNot457 + NQNot457
        TNonTaxCombat = TNonTaxCombat + NonTaxCombat
        TPremiums = TPremiums + Premiums
        TNonStatInc = TNonStatInc + NonStatInc
        TDef409A = TDef409A + Def409A
        TRoth401K = TRoth401K + Roth401K
        TRoth403B = TRoth403B + Roth403B
        If RWSickPayIndic = 1 Then
            TSickTax = TSickTax + FWT
        End If
                
        Print #TChannel, PrintString  ' Output text
        
''=============================================================================================
'    ' Create Code "RO" - Employee Wage Records
'        PrintValue(1) = "RO":                                       FormatString(1) = "a2"
'        PrintValue(2) = " ":                                        FormatString(2) = "a9"
'        PrintValue(3) = Format(AllocTips, "00000000000"):           FormatString(3) = "r11"
'        PrintValue(4) = Format(UncollTaxTips, "00000000000"):       FormatString(4) = "r11"
'        PrintValue(5) = Format(MedSvgsAcct, "00000000000"):         FormatString(5) = "r11"
'        PrintValue(6) = Format(RetireAcct, "00000000000"):          FormatString(6) = "r11"
'        PrintValue(7) = Format(AdoptExp, "00000000000"):            FormatString(7) = "r11"
'        PrintValue(8) = Format(UncollSSRRTATax, "00000000000"):     FormatString(8) = "r11"
'        PrintValue(9) = Format(UncollMedTax, "00000000000"):        FormatString(9) = "r11"
'        PrintValue(10) = Format(Income409A, "00000000000"):         FormatString(10) = "r11"
'        PrintValue(11) = " ":                                       FormatString(11) = "a175"
'        PrintValue(12) = Format(PRWages, "00000000000"):            FormatString(12) = "r11"
'        PrintValue(13) = Format(PRComm, "00000000000"):             FormatString(13) = "r11"
'        PrintValue(14) = Format(PRAllow, "00000000000"):            FormatString(14) = "r11"
'        PrintValue(15) = Format(PRTips, "00000000000"):             FormatString(15) = "r11"
'        PrintValue(16) = Format(PRTotalInc, "00000000000"):         FormatString(16) = "r11"
'        PrintValue(17) = Format(PRTaxWH, "00000000000"):            FormatString(17) = "r11"
'        PrintValue(18) = Format(PRRetFundContr, "00000000000"):     FormatString(18) = "r11"
'        PrintValue(19) = " ":                                       FormatString(19) = "a11"
'        PrintValue(20) = Format(OtherIncTax, "00000000000"):        FormatString(20) = "r11"
'        PrintValue(21) = Format(OtherIncTaxWH, "00000000000"):      FormatString(21) = "r11"
'        PrintValue(22) = " ":                                       FormatString(22) = "a128"
'        PrintValue(23) = " ":                                       FormatString(23) = "~"
'
'        FormatPrint
'        Print #TChannel, PrintString  ' Output text
        
'=============================================================================================
        ' Create Code "RS" - Employee Supplemental Record

        PrintValue(1) = "RS":                                       FormatString(1) = "a2"
        PrintValue(2) = OrgStCode:                                  FormatString(2) = "n2"
        PrintValue(3) = " ":                                        FormatString(3) = "a5"
        PrintValue(4) = Format(PREmployee.SSN, "000000000"):        FormatString(4) = "r9"
        PrintValue(5) = PREmployee.FirstName:                       FormatString(5) = "a15"
        PrintValue(6) = PREmployee.MidInit:                         FormatString(6) = "a15"
        PrintValue(7) = PREmployee.LastName:                        FormatString(7) = "a20"
        PrintValue(8) = PREmployee.MidInit:                         FormatString(8) = "a4"
        PrintValue(9) = PREmployee.Address1:                        FormatString(9) = "a22"
        PrintValue(10) = PREmployee.Address2:                       FormatString(10) = "a22"
        PrintValue(11) = PREmployee.City:                           FormatString(11) = "a22"
        PrintValue(12) = PREmployee.State:                          FormatString(12) = "a2"
        PrintValue(13) = PREmployee.ZipCode:                        FormatString(13) = "a5"
        PrintValue(14) = RSZipExt:                                  FormatString(14) = "a4"
        PrintValue(15) = " ":                                       FormatString(15) = "a5"
        PrintValue(16) = RSFSP:                                     FormatString(16) = "a23"
        PrintValue(17) = RSFPC:                                     FormatString(17) = "a15"
        PrintValue(18) = RSCC:                                      FormatString(18) = "a2"
        PrintValue(19) = " ":                                       FormatString(19) = "a2"
        PrintValue(20) = Format(RptPer, "000000"):                  FormatString(20) = "r6"
        PrintValue(21) = Format(StQUnempWgs, "00000000000"):        FormatString(21) = "r11"
        PrintValue(22) = Format(STQUnempTxWgs, "00000000000"):      FormatString(22) = "r11"
        PrintValue(23) = WksWorked:                                 FormatString(23) = "a2"
        PrintValue(24) = Format(DtEmpl, "00000000"):                FormatString(24) = "r8"
        PrintValue(25) = Format(DtSep, "00000000"):                 FormatString(25) = "r8"
        PrintValue(26) = " ":                                       FormatString(26) = "a5"
        PrintValue(27) = StEmplrAcct:                               FormatString(27) = "n20"
        PrintValue(28) = " ":                                       FormatString(28) = "a6"
        PrintValue(29) = PostStCode:                                FormatString(29) = "a2"
        PrintValue(30) = Format(OHTxWgs, "00000000000"):            FormatString(30) = "r11"
        PrintValue(31) = Format(OHIncTxWH, "00000000000"):          FormatString(31) = "r11"
        PrintValue(32) = Format(WgsTipsEtc, "00000000000"):         FormatString(32) = "r10"
        PrintValue(33) = TxType:                                    FormatString(33) = "a1"
        PrintValue(34) = Format(LocTxWgs, "00000000000"):           FormatString(34) = "r11"
        PrintValue(35) = Format(LocIncTxWH, "00000000000"):         FormatString(35) = "r11"
        PrintValue(36) = Format(SchDistNo, "   0000"):              FormatString(36) = "r7"
        PrintValue(37) = " ":                                       FormatString(37) = "a75"
        PrintValue(38) = " ":                                       FormatString(38) = "a75"
        PrintValue(39) = " ":                                       FormatString(39) = "a25"
        PrintValue(40) = " ":                                       FormatString(40) = "~"

        FormatPrint
        Print #TChannel, PrintString  ' Output text
        If Not PREmployee.GetNext Then Exit Do
    Loop
       
'=============================================================================================
        ' Create Code "RT" - Totals of Amount fields in "RW" record

        PrintValue(1) = "RT":                                       FormatString(1) = "a2"
        PrintValue(2) = Format(TRWs, "0000000"):                    FormatString(2) = "r7"
        PrintValue(3) = Format(TWgs, "000000000000000"):            FormatString(3) = "r15"
        PrintValue(4) = Format(TFWT, "000000000000000"):            FormatString(4) = "r15"
        PrintValue(5) = Format(TSSWage, "000000000000000"):         FormatString(5) = "r15"
        PrintValue(6) = Format(TSSTax, "000000000000000"):          FormatString(6) = "r15"
        PrintValue(7) = Format(TMedWage, "000000000000000"):        FormatString(7) = "r15"
        PrintValue(8) = Format(TMedTax, "000000000000000"):         FormatString(8) = "r15"
        PrintValue(9) = Format(TSSTips, "000000000000000"):         FormatString(9) = "r15"
        PrintValue(10) = Format(TAEIC, "000000000000000"):          FormatString(10) = "r15"
        PrintValue(11) = Format(TDCB, "000000000000000"):           FormatString(11) = "r15"
        PrintValue(12) = Format(TDC401K, "000000000000000"):        FormatString(12) = "r15"
        PrintValue(13) = Format(TDC403B, "000000000000000"):        FormatString(13) = "r15"
        PrintValue(14) = Format(TDC408K, "000000000000000"):        FormatString(14) = "r15"
        PrintValue(15) = Format(TDC457B, "000000000000000"):        FormatString(15) = "r15"
        PrintValue(16) = Format(TDC501C, "000000000000000"):        FormatString(16) = "r15"
        PrintValue(17) = Format(TMilPay, "000000000000000"):        FormatString(17) = "r15"
        PrintValue(18) = Format(TNQ457, "000000000000000"):         FormatString(18) = "r15"
        PrintValue(19) = Format(THlthSvgs, "000000000000000"):      FormatString(19) = "r15"
        PrintValue(20) = Format(TNQNot457, "000000000000000"):      FormatString(20) = "r15"
        PrintValue(21) = Format(TNonTaxCombat, "000000000000000"):  FormatString(21) = "r15"
        PrintValue(22) = " ":                                       FormatString(22) = "a15"
        PrintValue(23) = Format(TPremiums, "000000000000000"):      FormatString(23) = "r15"
        PrintValue(24) = Format(TSickTax, "000000000000000"):       FormatString(24) = "r15"
        PrintValue(25) = Format(TNonStatInc, "000000000000000"):    FormatString(25) = "r15"
        PrintValue(26) = Format(TDef409A, "000000000000000"):       FormatString(26) = "r15"
        PrintValue(27) = Format(TRoth401K, "000000000000000"):      FormatString(27) = "r15"
        PrintValue(28) = Format(TRoth403B, "000000000000000"):      FormatString(28) = "r15"
        PrintValue(29) = " ":                                       FormatString(29) = "a113"
        PrintValue(30) = " ":                                       FormatString(30) = "~"
        
        FormatPrint
        Print #TChannel, PrintString  ' Output text
        W2ClearVars     ' Clear Total Variables
        
'=============================================================================================
        ' Create Code "RF" - Final record
        PrintValue(1) = "RF":                                       FormatString(1) = "a2"
        PrintValue(2) = " ":                                        FormatString(2) = "a5"
        PrintValue(3) = Format(TRWs, "000000000"):                  FormatString(3) = "r9"
        PrintValue(4) = " ":                                        FormatString(4) = "a496"
        PrintValue(5) = " ":                                        FormatString(5) = "~"
        
        FormatPrint
        Print #TChannel, PrintString  ' Output text
    End
    
End Sub

Public Sub W2ClearVars()
    TRWs = 0
    TWgs = 0
    TFWT = 0
    TSSWage = 0
    TSSTax = 0
    TMedWage = 0
    TMedTax = 0
    TSSTips = 0
    TAEIC = 0
    TDCB = 0
    TDC401K = 0
    TDC403B = 0
    TDC408K = 0
    TDC457B = 0
    TDC501C = 0
    TMilPay = 0
    TNQ457 = 0
    THlthSvgs = 0
    TNQNot457 = 0
    TNonTaxCombat = 0
    TPremiums = 0
    TSickTax = 0
    TNonStatInc = 0
    TDef409A = 0
    TRoth401K = 0
    TRoth403B = 0
        
End Sub


Public Sub ItemDetail(ByVal RangeType As Byte, _
                          ByVal BatchNumbr As Long, _
                          ByVal PEDate As Long, _
                          ByVal CheckDt As Long, _
                          ByVal Startdate As Long, _
                          ByVal EndDate As Long, _
                          ByVal OptDate As String)
Dim Item1ID, Item2ID, Item3ID, Item4ID, Item5ID, ctr, EECount, DTCount As Long
Dim SGross, SItem1, SItem2, SItem3, SItem4, SItem5 As Currency
Dim TGross, TItem1, TItem2, TItem3, TItem4, TItem5 As Currency

    SetEquates
    PrtInit ("Land")
    LandSW = 1
    SetFont 8, Equate.LandScape
    
    ReportTitle = "PAYROLL ITEM DETAIL LISTING"
    If frmItemDetail.optChkDate = True Then
        Msg2 = "ORDER BY CHECK DATE"
    Else
        Msg2 = "ORDER BY EMPLOYEE NUMBER"
    End If
    
    frmItemDetail.RSItem.MoveFirst
    ItemCount = 0
    Do
        If frmItemDetail.RSItem!Selected = True Then
            ItemCount = ItemCount + 1
            If ItemCount = 1 Then Item1ID = frmItemDetail.RSItem!ItemID
            If ItemCount = 2 Then Item2ID = frmItemDetail.RSItem!ItemID
            If ItemCount = 3 Then Item3ID = frmItemDetail.RSItem!ItemID
            If ItemCount = 4 Then Item4ID = frmItemDetail.RSItem!ItemID
            If ItemCount = 5 Then Item5ID = frmItemDetail.RSItem!ItemID
        End If
        frmItemDetail.RSItem.MoveNext
        If frmItemDetail.RSItem.EOF Then Exit Do
    Loop
    
    trs.CursorLocation = adUseClient
    trs.Fields.Append "EmpName", adChar, 30, adFldMayBeNull
    trs.Fields.Append "EmpID", adDouble:                            trs.Fields.Append "EmpNo", adDouble:
    trs.Fields.Append "HistID", adDouble:                           trs.Fields.Append "PEDate", adDouble:
    trs.Fields.Append "ChkDate", adDate:                            trs.Fields.Append "Gross", adCurrency:
    trs.Fields.Append "Item1Amount", adCurrency:                    trs.Fields.Append "Item2Amount", adCurrency
    trs.Fields.Append "Item3Amount", adCurrency:                    trs.Fields.Append "Item4Amount", adCurrency
    trs.Fields.Append "Item5Amount", adCurrency:                    trs.Fields.Append "YTDGross", adCurrency
      
    trs.Open , , adOpenDynamic, adLockOptimistic

    frmItemDetail.Hide
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    frmProgress.lblMsg1 = "Printing " & ReportTitle & " for: " & PRCompany.Name
    frmProgress.Show
            
    If RangeType = PREquate.RangeTypeBatch Then
        If Not PRBatch.GetByID(BatchNumbr) Then
            MsgBox "PR Batch Not Found: " & BatchNumbr, vbCritical
            GoBack
        End If
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.BatchID = " & BatchNumbr & _
        " AND ItemType = " & PREquate.ItemTypeDED
        Msg1 = "Batch: " & BatchNumbr & "  Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy") & _
               " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.CheckDate >= " & (Startdate) & " AND " & _
                                     " PRItemHist.CheckDate <= " & (EndDate) & _
                                     " AND ItemType = " & PREquate.ItemTypeDED
            Msg1 = "CHECK DATE RANGE: " & Format(Startdate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
                                    
        Else
            SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.PEDate >= " & (Startdate) & " AND " & _
                                    " PRItemHist.PEDate <= " & (EndDate) & _
                                    " AND ItemType = " & PREquate.ItemTypeDED
            Msg1 = "P/E DATE RANGE: " & Format(Startdate, "mm/dd/yyyy") & " TO: " & Format(EndDate, "mm/dd/yyyy")
        End If

    End If

    If Not PRItemHist.GetBySQL(SQLString) Then
        MsgBox "No P/R Item History Data Found!", vbExclamation
        GoBack
    End If
    
    Do
        ' has this employee been selected?
        frmItemDetail.RSEmp.Find "EmpID = " & PRItemHist.EmployeeID, 0, adSearchForward, 1

        If Not frmItemDetail.RSEmp.EOF And frmItemDetail.RSEmp!Selected = True Then
            SQLString = "HistID = " & PRItemHist.HistID
            trs.Find SQLString, 0, adSearchForward, 1
            If trs.EOF Then
                trs.AddNew
                trs!EmpID = PRItemHist.EmployeeID
                If Not PREmployee.GetByID(PRItemHist.EmployeeID) Then
                    MsgBox "Employee not found in Employee Master File!!!", vbCritical, "Item Detail Report"
                    GoBack
                End If
                trs!EmpNo = PREmployee.EmployeeNumber
                trs!EmpName = PREmployee.LFName
                trs!HistID = PRItemHist.HistID
                trs!ChkDate = PRItemHist.CheckDate
                trs!PEDate = PRItemHist.PEDate
                trs!Item1Amount = 0
                trs!Item2Amount = 0
                trs!Item3Amount = 0
                trs!Item4Amount = 0
                trs!Item5Amount = 0
                ctr = ctr + 1
                If Not PRHist.GetByID(PRItemHist.HistID) Then
                    trs!Gross = 0
                    trs!YTDGross = trs!YTDGross + PRHist.Gross
                Else
                    trs!Gross = PRHist.Gross
                    trs!YTDGross = trs!YTDGross + PRHist.Gross
                End If

                ItemCount = 0
                frmItemDetail.RSItem.MoveFirst
            End If
            
            If PRItemHist.EmployerItemID = Item1ID Then trs!Item1Amount = trs!Item1Amount + PRItemHist.Amount
            If PRItemHist.EmployerItemID = Item2ID Then trs!Item2Amount = trs!Item2Amount + PRItemHist.Amount
            If PRItemHist.EmployerItemID = Item3ID Then trs!Item3Amount = trs!Item3Amount + PRItemHist.Amount
            If PRItemHist.EmployerItemID = Item4ID Then trs!Item4Amount = trs!Item4Amount + PRItemHist.Amount
            If PRItemHist.EmployerItemID = Item5ID Then trs!Item5Amount = trs!Item5Amount + PRItemHist.Amount
        
            trs.Update
            
        End If
        If Not PRItemHist.GetNext Then Exit Do
    Loop

    ''''''''''''''''''''      PRINT REPORT DETAIL     ''''''''''''''''''''''''''''''''''''''
    If frmItemDetail.optChkDate = True Then
        trs.Sort = "chkdate"
    Else
        trs.Sort = "EmpNo"
    End If
    
    LastEmpID = 0
    
    If trs.RecordCount = 0 Then
        MsgBox "There are no Employees that fit your Criteria Selection", vbCritical, "Item Detail Report"
        GoBack
    End If
    
    trs.MoveFirst

    LastEmpNo = 0
    LastChkDate = 0
    ctr = 0
    DTCount = 0
    EECount = 0
    Do
        ctr = ctr + 1

        If Ln = 0 Or Ln > MaxLines - LineCount Then
            If Ln Then FormFeed
            PageHeader ReportTitle, Msg1, Msg2, ""
            ItemDetailHeader
        End If
        If PREmployee.GetByID(trs!EmpID) Then
        End If

        If frmItemDetail.optChkDate Then
            If LastChkDate <> "0" And LastChkDate <> trs!ChkDate And DTCount > 1 Then
                PrintValue(1) = "     Check Date: ":          FormatString(1) = "a17"
                PrintValue(2) = LastChkDate:                  FormatString(2) = "a46"
                PrintValue(3) = SGross:                       FormatString(3) = "d14"
                PrintValue(4) = SItem1:                       FormatString(4) = "d14"
                PrintValue(5) = SItem2:                       FormatString(5) = "d14"
                PrintValue(6) = SItem3:                       FormatString(6) = "d14"
                PrintValue(7) = SItem4:                       FormatString(7) = "d14"
                PrintValue(8) = SItem5:                       FormatString(8) = "d14"
                PrintValue(9) = " ":                          FormatString(9) = "~"
                FormatPrint
'                Ln = Ln + 2
                SGross = 0
                SItem1 = 0
                SItem2 = 0
                SItem3 = 0
                SItem4 = 0
                SItem5 = 0
                DTCount = 0
            End If
        Else
            If LastEmpNo <> 0 And LastEmpNo <> trs!EmpNo And EECount = 1 Then
'                Ln = Ln + 1
                PrintValue(1) = "     Employee: ":                          FormatString(1) = "a17"
                PrintValue(2) = LastEmpNo & " - " & LastEmpName: FormatString(2) = "a46"
                PrintValue(3) = SGross:                       FormatString(3) = "d14"
                PrintValue(4) = SItem1:                       FormatString(4) = "d14"
                PrintValue(5) = SItem2:                       FormatString(5) = "d14"
                PrintValue(6) = SItem3:                       FormatString(6) = "d14"
                PrintValue(7) = SItem4:                       FormatString(7) = "d14"
                PrintValue(8) = SItem5:                       FormatString(8) = "d14"
                PrintValue(9) = " ":                          FormatString(9) = "~"
                FormatPrint
                Ln = Ln + 1
                SGross = 0
                SItem1 = 0
                SItem2 = 0
                SItem3 = 0
                SItem4 = 0
                SItem5 = 0
                EECount = 0
            End If
        End If

        ' get the item
        PrintValue(1) = trs!EmpNo:                           FormatString(1) = "a7"
        PrintValue(2) = trs!EmpName:                         FormatString(2) = "a32"
        PrintValue(3) = Format(trs!PEDate, "mm/dd/yyyy"):    FormatString(3) = "a12"
        PrintValue(4) = Format(trs!ChkDate, "mm/dd/yyyy"):   FormatString(4) = "a12"
        PrintValue(5) = trs!Gross:                           FormatString(5) = "d14"
        PrintValue(6) = trs!Item1Amount:                     FormatString(6) = "d14"
        PrintValue(7) = trs!Item2Amount:                     FormatString(7) = "d14"
        PrintValue(8) = trs!Item3Amount:                     FormatString(8) = "d14"
        PrintValue(9) = trs!Item4Amount:                     FormatString(9) = "d14"
        PrintValue(10) = trs!Item5Amount:                    FormatString(10) = "d14"
        PrintValue(11) = " ":                                FormatString(11) = "~"
        FormatPrint
        Ln = Ln + 1
        SGross = SGross + trs!Gross
        TGross = TGross + trs!Gross
        SItem1 = SItem1 + trs!Item1Amount
        TItem1 = SItem1 + trs!Item1Amount
        SItem2 = SItem2 + trs!Item2Amount
        TItem2 = SItem2 + trs!Item2Amount
        SItem3 = SItem3 + trs!Item3Amount
        TItem3 = SItem3 + trs!Item3Amount
        SItem4 = SItem4 + trs!Item4Amount
        TItem4 = SItem4 + trs!Item4Amount
        SItem5 = SItem5 + trs!Item5Amount
        TItem5 = SItem5 + trs!Item5Amount
        DTCount = DTCount + 1
        
        If LastEmpNo = trs!EmpNo Then
            EECount = 1
        ElseIf LastEmpNo = 0 Then
            EECount = 0
            Ln = Ln + 1
            SGross = 0
            SItem1 = 0
            SItem2 = 0
            SItem3 = 0
            SItem4 = 0
            SItem5 = 0
        End If
        LastEmpNo = PREmployee.EmployeeNumber
        LastChkDate = trs!ChkDate
        LastEmpName = trs!EmpName

        
'MsgBox "  2  " & LastEmpNo & "  " & trs!EmpNo & "  " & EECount
        trs.MoveNext
        If trs.EOF Then Exit Do
    Loop
    
    If frmItemDetail.optChkDate Then
        PrintValue(1) = "     Check Date: ":                 FormatString(1) = "a17"
        PrintValue(2) = LastChkDate:                         FormatString(2) = "a46"
        PrintValue(3) = SGross:                              FormatString(3) = "d14"
        PrintValue(4) = SItem1:                              FormatString(4) = "d14"
        PrintValue(5) = SItem2:                              FormatString(5) = "d14"
        PrintValue(6) = SItem3:                              FormatString(6) = "d14"
        PrintValue(7) = SItem4:                              FormatString(7) = "d14"
        PrintValue(8) = SItem5:                              FormatString(8) = "d14"
        PrintValue(9) = " ":                                 FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 2
    Else
        PrintValue(1) = " ":                                 FormatString(1) = "a17"
        PrintValue(2) = LastEmpNo & " - " & LastEmpName: FormatString(2) = "a46"
        PrintValue(3) = SGross:                              FormatString(3) = "d14"
        PrintValue(4) = SItem1:                              FormatString(4) = "d14"
        PrintValue(5) = SItem2:                              FormatString(5) = "d14"
        PrintValue(6) = SItem3:                              FormatString(6) = "d14"
        PrintValue(7) = SItem4:                              FormatString(7) = "d14"
        PrintValue(8) = SItem5:                              FormatString(8) = "d14"
        PrintValue(9) = " ":                                 FormatString(9) = "~"
        FormatPrint
        Ln = Ln + 2
    End If
    PrintValue(1) = "GRAND TOTAL ":                          FormatString(1) = "a17"
    PrintValue(3) = TGross:                                  FormatString(3) = "d14"
    PrintValue(4) = TItem1:                                  FormatString(4) = "d14"
    PrintValue(5) = TItem2:                                  FormatString(5) = "d14"
    PrintValue(6) = TItem3:                                  FormatString(6) = "d14"
    PrintValue(7) = TItem4:                                  FormatString(7) = "d14"
    PrintValue(8) = TItem5:                                  FormatString(8) = "d14"
    PrintValue(9) = " ":                                     FormatString(9) = "~"
    FormatPrint
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Public Sub ItemDetailHeader()
Dim ColNumber As Long

    PrintValue(1) = "EMP NO.":                      FormatString(1) = "a7"
    PrintValue(2) = "EMPLOYEE NAME":                FormatString(2) = "a32"
    PrintValue(3) = "P/E DATE":                     FormatString(3) = "a12"
    PrintValue(4) = "CHK DATE":                     FormatString(4) = "a12"
    PrintValue(5) = "GROSS":                        FormatString(5) = "r13"

    ColNumber = 6
        
    frmItemDetail.RSItem.MoveFirst
    Do
        If frmItemDetail.RSItem!Selected = True Then
            ' get the item
            If PRItem.GetByID(frmItemDetail.RSItem!ItemID) Then
                PrintValue(ColNumber) = PRItem.Abbreviation
            ElseIf ColNumber = 11 Then
                Exit Do
            End If
            
            FormatString(ColNumber) = "r14"
            ColNumber = ColNumber + 1
            
        End If
        frmItemDetail.RSItem.MoveNext
        
    Loop Until frmItemDetail.RSItem.EOF

    PrintValue(ColNumber) = " ":                    FormatString(ColNumber) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(147, "="):               FormatString(1) = "a147"
    PrintValue(2) = " ":                            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
End Sub
