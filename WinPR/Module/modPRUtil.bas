Attribute VB_Name = "modPRUtil"
Option Explicit

Dim ASCIIChannel As Integer
Dim dType, X, TextName As String
Dim Ct, CT2, RCT, ImportCount As Long
Dim i, j, k As Integer
Dim DirDepDed1, DirDepDed2 As Byte
Dim CompanyID As Long
Dim mm, dd, yy As Long
Dim DfltCityID As Long
Dim DfltStateID As Long
Dim trs As New ADODB.Recordset
Dim OEBasis(10), DEDBasis(10) As Byte

Dim irs As New ADODB.Recordset
Dim CtyID As Long

' other tax used as separate city tax
Dim Tax6City, Tax7City, Tax8City, Tax9City, Tax0City As Long

' other tax used for SD tax
Dim SDTax6ID As Long
Dim SDTax7ID As Long
Dim SDTax8ID As Long
Dim SDTax9ID As Long
Dim SDTax0ID As Long

Dim DistFlag As Boolean
Public Sub aPRImport()
    
Dim X As String
    
    ' TO DO
    '
    ' SD Export - EE city includes state
    '             STMA - cram
    
    Tax6City = 0
    Tax7City = 0
    Tax8City = 0
    Tax9City = 0
    Tax0City = 0
        
    SDTax6ID = 0
    SDTax7ID = 0
    SDTax8ID = 0
    SDTax9ID = 0
    SDTax0ID = 0
    
    PRCompany.OpenRS
    PRDepartment.OpenRS
    PREmployee.OpenRS
    PRCity.OpenRS
    PRItem.OpenRS
    PRHist.OpenRS
    PRItemHist.OpenRS
    PRDist.OpenRS
    PRState.OpenRS
    PRBatch.OpenRS
    PRGLUpd.OpenRS
    
    ' get the PRCompany record
    If Not PRCompany.GetByID(User.LastPRCompany) Then
        MsgBox "Company record not found!", vbCritical
        End
    End If
    
'    TextName = "c:\balint\data\prx77701.txt"
'    TextName = "c:\balint\data\prx14001.txt"
'    TextName = "c:\balint\data\prx10901.txt"
    
    ' *** stuff it *** CUYFLS
    DfltCityID = 2
    
    Set Progress = New frmProgress
    Progress.Show
    Progress.Caption = "Windows PR SuperDOS Import"
    Progress.lblMsg1 = TextFileName
        
    ' create regular and overtime item types for every company
    PRItem.Clear
    PRItem.ItemType = PREquate.ItemTypeRegPay
    PRItem.EmployeeID = 0
    PRItem.Abbreviation = "RegPay"
    PRItem.Title = "Regular Pay"
    PRItem.Save (Equate.RecAdd)
    
    PRItem.Clear
    PRItem.ItemType = PREquate.ItemTypeOvtPay
    PRItem.EmployeeID = 0
    PRItem.Abbreviation = "OvtPay"
    PRItem.Title = "Overtime Pay"
    PRItem.Save (Equate.RecAdd)
    
    ' ???????????????????????????????
    ' get Ohio state record ????????
'    SQLString = "SELECT * FROM PRState WHERE StateAbbrev = OH"
'    If Not PRState.GetBySQL(SQLString) Then
'        MsgBox "PRState for Ohio not found", vbCritical
'        End
'    End If
    PRState.StateID = 36
    DfltStateID = 36
    
    DistFlag = False        ' set to true if a dist employer
    
    ASCIIChannel = FreeFile
    Open TextFileName For Input As ASCIIChannel
    
    RCT = 0
       
    Do
   
        Line Input #ASCIIChannel, X
      
        RCT = RCT + 1
      
        If Mid(X, 2, 3) = "END" Then Exit Do
      
        If RCT Mod 100 = 0 Then
            Progress.lblMsg2 = "Counting Records: " & CStr(RCT) & " " & TextName
            Progress.lblMsg2.Refresh
        End If
   
    Loop
    
    Ct = 0
   
    ' init the PRGLUpd file
    PRGLUpd.DeleteAll
    PRGLUpd.OpenRS
   
    Close #ASCIIChannel
    Open TextFileName For Input As ASCIIChannel
   
    Do
      
        Input #ASCIIChannel, dType

        Select Case dType
        
            Case "END"
                Exit Do
            
            ' Employer Records
            Case "ER1"
                ImportEmployer
            Case "ERDIRDEP"
                ImportERDirDep
            Case "EROTAX"
                ImportEROTax
            Case "ERTITLE"
                ImportERTitle
            Case "ERTYPE"
                ImportERType
            Case "ERGLACCT"
                ImportERGLAcct
            Case "ERFREQ"
                ' skip it - frequency codes
                ImportSkip 20
            Case "ERWC"
                ' skip it - work comp codes
                ImportSkip 20
            Case "ERCOM"
                ' skip it - comments
                ImportSkip 3
            Case "DPT"
                ImportDepartment
            
            ' Employee Records
            Case "EE1"
                ImportEmployee1
            Case "EE2"
                ImportEmployee2
            Case "EEDATE"
                ImportEEDate
            Case "EEOTHER"
                ImportEEOther
            Case "EECOM"
                ' skip it - comments
                ImportSkip 3  ' !!! 2 ???
                PREmployee.Save (Equate.RecAdd)
            
            Case "EEOE"
                ImportEEOE
            Case "EEDED"
                ImportEEDED
            Case "EEDIRDEP"
                ImportEEDirDep
            
            ' History and Distn
            Case "HIS"
                ImportHistory
            Case "DIS"
                DistFlag = True
                ImportDist
            
            Case "PRR"
                ImportPRRate
            
            Case "PNA1"
                ImportPRAcct 1
                
            Case "PNA2"
                ImportPRAcct 2
                
            Case "HIST"
                HistOnly
                
            Case Else
                MsgBox "Bad line type: " & dType, vbCritical
                End
        
        End Select

        ImportCount = ImportCount + 1
        If ImportCount Mod 100 = 1 Then
            Progress.lblMsg2 = "Importing record: " & Format(ImportCount, "###,##0") & " Of: " & Format(RCT, "###,##0")
            Progress.lblMsg2.Refresh
        End If

    Loop

    ' update PRItem.EmployerItemID for employees PRItem file
    EEItemUpdate
    
    ' ********
    '  ===> City & State Wage/Tax will NOT be used from PRHist !!!
    '       Only used to store total from PRDist
    ' ********
    
    ' select the default city
    SelectDfltCity
    
    ' assign the dflt cityID to the employees
    SetPREmployeeCityID
    
    ' assign the dflt CityID to the Employer
    PRCompany.CheckDays = frmStart.CheckDays
    PRCompany.DfltCityID = DfltCityID
    PRCompany.Save (Equate.RecPut)
    
    If DistFlag = True Then
        DistMatch   ' match dist to hist
        ' audit loop - compare wage and city tax
    Else
        SetPRDistCityID     ' update the dflt city id to PRDist
    End If
    
    ' assign city & state wages to PRDist
    ' assign state tax to PRDist
    TaxHistAssign
    
    ' create PRBatch from PRHist
    CreatePRBatch
    
    ' assign PREmployee.DepartmentID
    EEDeptAssign
    
    ' assign PRDist.EmployerItemID / PRItemHist.EmployerItemID
    ERIDAssign
    
    ' delete unassigned Dir Dep PRItem records
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDirDepDed
    rsInit SQLString, cn, irs
    If irs.RecordCount > 0 Then
        irs.MoveFirst
        Do
            X = Trim(irs!DirDepBank)
            If X = "" Then irs.Delete
            irs.MoveNext
            If irs.EOF Then Exit Do
        Loop
    End If
    
End Sub

Public Sub ImportState()

Dim Abbrev As String
Dim Name As String

    ' PRState.DeleteAll
    PRState.OpenRS
    
    Set Progress = New frmProgress
    Progress.Show
    Progress.Caption = "Windows PR State List Import"
    Progress.lblMsg1 = TextName
        
    ASCIIChannel = FreeFile
    On Error Resume Next
    TextName = "\Balint\Blank\StateList.csv"
    Open TextName For Input As ASCIIChannel
    If Err.Number <> 0 Then
        MsgBox "\Balint\Blank\StateList.csv Error: " & Err.Number & vbCr & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Do While Not EOF(ASCIIChannel)
           
        Input #ASCIIChannel, Abbrev, Name
        
        SQLString = "SELECT * FROM PRState WHERE StateAbbrev = '" & Abbrev & "'"
        If Not PRState.GetBySQL(SQLString) Then
            PRState.Clear
            PRState.StateAbbrev = Abbrev
            PRState.StateName = Trim(Name)
            PRState.Save (Equate.RecAdd)
        End If
    
        Progress.lblMsg1 = Abbrev & " " & Name
        Progress.Refresh
    
    Loop
    
    Close #ASCIIChannel

End Sub

Private Sub ImportPRRate()

    PRCity.Clear
    For i = 1 To 4
        Input #ASCIIChannel, X
        
        ' city number not filled in ????
        If i = 1 And X = "" Then
            ImportSkip 3
            Exit Sub
        End If
        
        If X <> "" Then
            Select Case i
                Case 1
                    ' add if it doesn't exist
                    SQLString = "SELECT * FROM PRCity WHERE CityNumber = " & X
                    If Not PRCity.GetBySQL(SQLString) Then
                        PRCity.Clear
                        PRCity.Save (Equate.RecAdd)
                    End If
                    
                    PRCity.CityNumber = CLng(X)
                Case 2
                    PRCity.CityName = X
                    PRCity.ShortName = X
                Case 3
                    PRCity.CityRate = CCur(X)
                Case 4
                    PRCity.CityRecipRate = CCur(X)
            End Select
        End If
    Next i
    
    PRCity.StateID = DfltStateID
    PRCity.Save (Equate.RecPut)

End Sub


Private Sub ImportEmployer()

Dim xID As Long

    PRCompany.Clear
    For i = 1 To 14
        Input #ASCIIChannel, X
        If X <> "" Then
            Select Case i
                Case 1
                    PRCompany.Name = Trim(X)
                Case 2
                    PRCompany.Address1 = X
                Case 3
                    PRCompany.Address2 = X
                Case 4
                    ' strip out "OHIO"
                    PRCompany.City = StripOhio(X)
                Case 5
                    PRCompany.AddrStateID = 36  ' default to OH
                Case 6
                    PRCompany.ZipCode = CLng(X)
                Case 7
                    xID = CLng(X)
                    If xID = 0 Then
                        PRCompany.StateID = ""
                    Else
                        PRCompany.StateID = Format(Int(xID / 10 ^ 6), "00") & "-" & _
                                            Format(xID Mod 10 ^ 6, "000000")
                    End If
                Case 8
                    xID = CLng(X)
                    If xID = 0 Then
                        PRCompany.FederalID = ""
                    Else
                        PRCompany.FederalID = Format(Int(xID / 10 ^ 7), "00") & "-" & _
                                            Format(xID Mod 10 ^ 7, "0000000")
                    End If
                Case 9
                    PRCompany.StateUnempPct = CCur(X)
                Case 10
                    PRCompany.FederalUnempPct = CCur(X)
                Case 11
                    PRCompany.DfltStateID = 36  ' default to OH
                Case 12
                    PRCompany.DfltMinWage = CCur(X)
                Case 13
                    PRCompany.DfltOTRate = CCur(X)
                Case 14
                    PRCompany.DfltRegHrs = CCur(X)
            End Select
        
        End If
    
    Next i
    
    PRCompany.FileName = dbName
    PRCompany.GLCompanyID = User.LastCompany
    PRCompany.FileName = dbName
    PRCompany.Save (Equate.RecPut)
    CompanyID = PRCompany.CompanyID
        
End Sub

Private Sub ImportERDirDep()

    If Not PRCompany.GetByID(CompanyID) Then
        MsgBox "Company record NF!", vbCritical
        End
    End If

    For i = 1 To 5
        
        Input #ASCIIChannel, X
        
        If X <> "" Then
            
            Select Case i
                
                Case 1      ' bank name
                    PRCompany.BankName = X
                
                Case 2      ' aba number
                    PRCompany.BankABA = X
                    
                Case 3      ' bank acct number
                    PRCompany.BankAccount = X
                    
                Case 4      ' first deduction number
                    DirDepDed1 = CByte(X)
                    
                Case 5      ' second deduction number
                    DirDepDed2 = CByte(X)
                    
            End Select
        
        End If
    
    Next i
    
    PRCompany.BankAddr1 = ""
    PRCompany.BankAddr2 = ""
    PRCompany.BankFraction = ""
    PRCompany.Save (Equate.RecPut)
    
End Sub

Private Sub ImportEROTax()

Dim OTXCode As String

    ' 5 other taxes
    For i = 1 To 5

        PRItem.Clear
        OTXCode = ""

        ' four fields
        For j = 1 To 4
            
            Input #ASCIIChannel, X
            
                Select Case j
                    
                    Case 1      ' other tax pct
                        
                        ' set up the PRItem record
                        PRItem.Clear
                        PRItem.EmployeeID = 0   ' signifies employer
                        PRItem.MaxPct = CCur(X)
                        
                    Case 2      ' other tax title
    
                        PRItem.Title = X
                        PRItem.Abbreviation = X
    
                    Case 3      ' other tax code
    
                        If X = "SD" Then
                            PRItem.ItemType = PREquate.ItemTypeSDTax
                        Else
                            PRItem.ItemType = PREquate.ItemTypeOtherTax
                        End If
                        
                        OTXCode = X
                        
                    Case 4      ' other tax max
                    
                        PRItem.MaxAmount = CCur(X)
                
                        ' save it ???
                        If PRItem.Title <> "" Then
                            PRItem.Active = 1
                            PRItem.Save (Equate.RecAdd)
                            ' SD tax ?
                            If PRItem.ItemType = PREquate.ItemTypeSDTax Then
                                If i = 1 Then SDTax6ID = PRItem.ItemID
                                If i = 2 Then SDTax7ID = PRItem.ItemID
                                If i = 3 Then SDTax8ID = PRItem.ItemID
                                If i = 4 Then SDTax9ID = PRItem.ItemID
                                If i = 5 Then SDTax0ID = PRItem.ItemID
                            End If
                        End If
                
                End Select
    
        Next j
        
        ' Eagl - other tax used for other cities
        If OTXCode = "CT" And PRItem.MaxAmount > 0 Then
            ' get the PRCityID
            SQLString = "SELECT * FROM PRCity WHERE CityNumber = " & Int(PRItem.MaxAmount)
            If PRCity.GetBySQL(SQLString) Then
                If i = 1 Then Tax6City = PRCity.CityID
                If i = 2 Then Tax7City = PRCity.CityID
                If i = 3 Then Tax8City = PRCity.CityID
                If i = 4 Then Tax9City = PRCity.CityID
                If i = 5 Then Tax0City = PRCity.CityID
            End If
        End If
        
    Next i
 
End Sub

Private Sub ImportERTitle()

    ' Order of import is IMPORTANT !!!
    ' this is the first of employer oe/ded setups
    ' 1 to 10 is other earnings / 11 to 20 is deductions

    For i = 1 To 20
        
        Input #ASCIIChannel, X
        
        ' skip Dir Deposit deductions
        If i = DirDepDed1 + 10 And DirDepDed1 <> 0 Then GoTo TTLNexti
        If i = DirDepDed2 + 10 And DirDepDed2 <> 0 Then GoTo TTLNexti
        
        If X = "" Then
            If i <= 10 Then X = "OE " & CStr(i) Else X = "DED " & CStr(i)
        End If
        
        PRItem.Clear
        
        If i <= 10 Then
            PRItem.ItemType = PREquate.ItemTypeOE
            PRItem.SDNumber = i
        Else
            PRItem.ItemType = PREquate.ItemTypeDED
            PRItem.SDNumber = i - 10
        End If
        
        PRItem.EmployeeID = 0
        PRItem.Title = X
        PRItem.Abbreviation = X
        
        PRItem.Save (Equate.RecAdd)
        
TTLNexti:
    Next i

End Sub

Private Sub ImportERType()

    ' update the type code in the PRItem file
    ' get by the SDNumber field

    For i = 1 To 20
        
        Input #ASCIIChannel, X
        
        If i <= 10 Then
            SQLString = "SELECT * FROM PRItem WHERE PRItem.SDNumber = " & CStr(i) & _
                        " AND PRItem.ItemType = " & PREquate.ItemTypeOE
        Else
            SQLString = "SELECT * FROM PRItem WHERE PRItem.SDNumber = " & CStr(i - 10) & _
                        " AND PRItem.ItemType = " & PREquate.ItemTypeDED
        End If
        
        If PRItem.GetBySQL(SQLString) Then
                    
            Select Case X
            
                Case "A"
                    PRItem.Basis = PREquate.BasisAmount
                Case "P"
                    PRItem.Basis = PREquate.BasisPercent
                Case "H"
                    PRItem.Basis = PREquate.BasisHourly
                Case Else
                    PRItem.Basis = PREquate.BasisAmount
            
            End Select
        
            PRItem.Active = 1
        
            PRItem.Save (Equate.RecPut)
    
            ' save for employee items assignment
            If i <= 10 Then
                OEBasis(i) = PRItem.Basis
            Else
                DEDBasis(i - 10) = PRItem.Basis
            End If
    
        End If
        
    Next i

End Sub

Private Sub ImportERGLAcct()

Dim iFlag As Boolean
Dim GLAcct As Long

    ' save company PRGLUpd records
    For i = 1 To 32
    
        Input #ASCIIChannel, X
        
        GLAcct = X
        
        If GLAcct <> 0 And Not (i >= 26 And i <= 30) Then
            PRGLUpd.Clear
            PRGLUpd.GLType = PREquate.GLTypeCompany
            PRGLUpd.RelatedID = 0
            PRGLUpd.GLAccountNum = GLAcct
            
            If i >= 1 And i <= 10 Then
                PRGLUpd.GLItemType = PREquate.GLItemTypeOE
                    
                ' get the PRItem record
                SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeOE & _
                            " AND SDNumber = " & i
                If PRItem.GetBySQL(SQLString) Then
                    iFlag = True
                    PRGLUpd.ItemID = PRItem.ItemID
                Else
                    iFlag = False
                End If
            
            
            ElseIf i >= 11 And i <= 20 Then
                PRGLUpd.GLItemType = PREquate.GLItemTypeDed
                
                ' get the PRItem record
                SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeDED & _
                            " AND SDNumber = " & i - 10
                If PRItem.GetBySQL(SQLString) Then
                    iFlag = True
                    PRGLUpd.ItemID = PRItem.ItemID
                Else
                    iFlag = False
                End If
            
            Else
                If i = 21 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSSTax
                If i = 22 Then PRGLUpd.GLItemType = PREquate.GLItemTypeMedTax
                If i = 23 Then PRGLUpd.GLItemType = PREquate.GLItemTypeFWTTax
                If i = 24 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSWTTax
                If i = 25 Then PRGLUpd.GLItemType = PREquate.GLItemTypeCWTTax
                If i = 31 Then PRGLUpd.GLItemType = PREquate.GLItemTypeGross
                If i = 32 Then PRGLUpd.GLItemType = PREquate.GLItemTypeNet
                iFlag = True
            End If
                        
            If iFlag Then PRGLUpd.Save (Equate.RecAdd)
        
        End If
    Next i

'    ' update the GL Acct # in the PRItem file
'    ' get by the SDNumber field
'
'    ' get the company record
'    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.CompanyID = " & CStr(CompanyID)
'    If Not PRCompany.GetBySQL(SQLString) Then
'        MsgBox "Company record not found ??? ", vbCritical
'        End
'    End If
'
'    For i = 1 To 32
'
'        Input #ASCIIChannel, X
'
'        If X = "" Then GoTo ERAcctNext
'
'        ' other earnings and deductions
'        If i >= 1 And i <= 20 Then
'
'            SQLString = "SELECT * FROM PRItem WHERE PRItem.SDNumber = " & CStr(i)
'            If PRItem.GetBySQL(SQLString) Then
'
'                PRItem.GLAccount = CLng(X)
'
'                PRItem.Save (Equate.RecPut)
'
'            End If
'
'        End If
'
'        ' skip 26 to 30 for other taxes
'        If i = 21 Then PRCompany.GLAcctSS = CLng(X)
'        If i = 22 Then PRCompany.GLAcctMED = CLng(X)
'        If i = 23 Then PRCompany.GLAcctFWT = CLng(X)
'        If i = 24 Then PRCompany.GLAcctSWT = CLng(X)
'        If i = 25 Then PRCompany.GLAcctCWT = CLng(X)
'        If i = 31 Then PRCompany.GLAcctGross = CLng(X)
'        If i = 32 Then PRCompany.GLAcctNet = CLng(X)
'
'ERAcctNext:
'    Next i
'
'    ' save the company record
'    PRCompany.Save (Equate.RecPut)

End Sub


Private Sub ImportDepartment()
    
    PRDepartment.Clear
    
    For i = 1 To 2
        Input #ASCIIChannel, X
        If X <> "" Then
            Select Case i
                Case 1
                    PRDepartment.DepartmentNumber = CLng(X)
                Case 2
                    PRDepartment.Name = X
            End Select
        End If
    Next i
    
    PRDepartment.Save (Equate.RecAdd)

End Sub

Private Sub ImportEmployee1()

    PREmployee.OpenRS
    PREmployee.Clear
    
    For i = 1 To 14
        Input #ASCIIChannel, X
        If X <> "" Then
            Select Case i
                Case 1
                    PREmployee.EmployeeNumber = CLng(X)
                Case 2
                    If Left(X, 1) = "," Then
                        PREmployee.LastName = Mid(X, 2, Len(X) - 1)
                    Else
                        PREmployee.LastName = X
                    End If
                Case 3
                    If Left(X, 1) = "," Then
                        PREmployee.FirstName = Mid(X, 2, Len(X) - 1)
                    Else
                        PREmployee.FirstName = X
                    End If
                Case 4
                    PREmployee.MidInit = X
                Case 5
                    PREmployee.Address1 = X
                Case 6
                    ' strip out "OHIO"
                    PREmployee.City = StripOhio(X)
                Case 7
                    PREmployee.State = X
                Case 8
                    PREmployee.ZipCode = CLng(X)
                Case 9
                    PREmployee.SSN = CLng(X)
                Case 10
                    PREmployee.DepartmentID = CLng(X)
                Case 11
                    PREmployee.SalaryAmount = CCur(X)
                    PREmployee.HourlyAmount = CCur(X)
                Case 12
                    PREmployee.Inactive = MakeBoo(X)
                Case 13
                    PREmployee.Salaried = MakeBoo(X)
                Case 14
                    PREmployee.FWTMarried = MakeBoo(X)
                    PREmployee.SWTMarried = MakeBoo(X)
            
            End Select
        
        End If
    
    Next i
    
End Sub

Private Sub ImportEmployee2()

Dim PPY, FWTEX, SWTEX As Integer

    For i = 1 To 19
        Input #ASCIIChannel, X
        If X <> "" Then
            Select Case i
                Case 1
                    PREmployee.NoSSTax = MakeBoo(X)
                    Case 2
                    PREmployee.NoMedTax = MakeBoo(X)
                Case 3
                    PREmployee.NoFedTax = MakeBoo(X)
                Case 4
                    PREmployee.NoStateTax = MakeBoo(X)
                Case 5
                    PREmployee.NoCityTax = MakeBoo(X)
                Case 6
                    ' no tx6
                Case 7
                    ' no tx7
                Case 8
                    ' no tx8
                Case 9
                    'no tx9
                Case 10
                    'no tx0
                Case 11
                    PREmployee.NoFedUnemp = MakeBoo(X)
                Case 12
                    PREmployee.NoStateUnemp = MakeBoo(X)
                Case 13
                    ' state number
                Case 14
                    ' city rate
                Case 15
                    If X = "26" Then
                        PREmployee.PaysPerYear = 26
                    ElseIf X = "24" Then
                        PREmployee.PaysPerYear = 24
                    ElseIf X = "12" Then
                        PREmployee.PaysPerYear = 12
                    Else
                        PREmployee.PaysPerYear = 52
                    End If
                    ' PREmployee.PaysPerYear = CInt(X)
                Case 16
                    FWTEX = CInt(X)
                Case 17
                    PREmployee.FWTExtraAmount = CCur(X)
                Case 18
                    SWTEX = CInt(X)
                Case 19
                    PREmployee.SWTExtraAmount = CCur(X)
            End Select
        End If
    Next i
    
    ' extra basis must be amount
    PREmployee.FWTExtraBasis = PREquate.BasisAmount
    PREmployee.SWTExtraBasis = PREquate.BasisAmount
    
    ' exemption number or percent?
    If FWTEX <= 20 Then
        PREmployee.FWTBasis = PREquate.BasisExemptions
        PREmployee.FWTAmount = FWTEX
    Else
        PREmployee.FWTBasis = PREquate.BasisPercent
        PREmployee.FWTAmount = FWTEX / 100
    End If
    
    If SWTEX <= 20 Then
        PREmployee.SWTBasis = PREquate.BasisExemptions
        PREmployee.SWTAmount = SWTEX
    Else
        PREmployee.SWTBasis = PREquate.BasisPercent
        PREmployee.SWTAmount = SWTEX / 100
    End If
    
End Sub

Private Sub ImportEEDate()
    
Dim EEDate As Date
    
    For i = 1 To 8
        
        Input #ASCIIChannel, X
        
        If X <> "" And X <> "0" Then
            
            On Error GoTo NextEEDate

            ' parse out the date mm/dd/yyyy except for birth date
            If i <= 7 Then
                mm = CLng(Mid(X, 1, 2))
                dd = CLng(Mid(X, 4, 2))
                yy = CLng(Mid(X, 7, 4))
            Else                ' assume yyyymmdd for birth date format
                mm = CLng(Mid(X, 5, 2))
                dd = CLng(Mid(X, 7, 2))
                yy = CLng(Mid(X, 1, 4))
            End If

            EEDate = DateSerial(yy, mm, dd)
            
            On Error GoTo 0

            Select Case i

                Case 1      ' Date Late Paid
                    PREmployee.DateLastPaid = EEDate
                Case 2
                    PREmployee.DateHired = EEDate
                Case 3
                    PREmployee.DateLastRaise = EEDate
                Case 4
                    PREmployee.DateLastReview = EEDate
                Case 5
                    PREmployee.DateLastLayoff = EEDate
                Case 6
                    PREmployee.DateLastRecall = EEDate
                Case 7
                    PREmployee.DateTerminated = EEDate
                Case 8
                    PREmployee.DateOfBirth = EEDate

            End Select

        End If
        
NextEEDate:
    Next i
    
End Sub
Private Sub ImportEEOther()

    For i = 1 To 10
        
        Input #ASCIIChannel, X

        If X <> "" Then

            Select Case i
                
                Case 1
                    ' term reason code
                Case 2
                    PREmployee.Sex = X
                Case 3
                    ' race code
                Case 4
                    If X = "0" Then
                        PREmployee.MaritalStatus = "S"
                    Else
                        PREmployee.MaritalStatus = "M"
                    End If
                Case 5
                    PREmployee.EducationLevel = CLng(X)
                Case 6
                    PREmployee.ShiftCode = CLng(X)
                Case 7
                    ' PREmployee.WorkCompNum = CLng(x)
                Case 8
                    ' phone number
                Case 9
                    ' 1099 emp
                Case 10
                    ' bank code ???
                    
            End Select
        
        End If
    
    Next i

End Sub
    
Private Sub ImportEEOE()
    
Dim OENumber As Byte
Dim ItemFlag As Boolean
Dim FieldNumber As Byte
Dim LastOE As Integer
    
    PRItem.Clear
    LastOE = 0
    
    For i = 1 To 180
        
        OENumber = Int((i - 1) / 18) + 1  ' OE numbers 1 to 10
        
        ' change in OE number - decide if to save
        If LastOE <> 0 And OENumber <> LastOE Then
            If ItemFlag = True And PRItem.Active = 1 Then
                PRItem.EmployeeID = PREmployee.EmployeeID
                PRItem.ItemType = PREquate.ItemTypeOE
                PRItem.SDNumber = LastOE
                PRItem.Basis = OEBasis(LastOE)
                
                PRItem.UseEmployer = 1
                
                PRItem.Save (Equate.RecAdd)
            End If
            ItemFlag = False
            PRItem.Clear
            FieldNumber = 0
        End If
        
        Input #ASCIIChannel, X

        FieldNumber = FieldNumber + 1

        ' flag to save it
        ItemFlag = True

        Select Case FieldNumber
        
            Case 1
                PRItem.AmtPct = CCur(X)
            Case 2
                PRItem.MaxAmount = CCur(X)
            Case 3
                If X = "X" Then PRItem.Active = 1
            Case 4
                If X = "X" Then PRItem.Tips = 1
            Case 5
                ' addl tips - skip it ....
            Case 6
                If X = "X" Then PRItem.NotInNet = 1
            Case 7
                If X = "X" Then PRItem.NoSSTax = 1
            Case 8
                If X = "X" Then PRItem.NoMedTax = 1
            Case 9
                If X = "X" Then PRItem.NoFWTTax = 1
            Case 10
                If X = "X" Then PRItem.NoSWTTax = 1
            Case 11
                If X = "X" Then PRItem.NoCWTTax = 1
            ' other tax tax flags are per employer
        End Select
        
NextEEOE:
        LastOE = OENumber
    
    Next i
            
    ' catch the last one
    If ItemFlag = True And PRItem.Active = 1 Then
        PRItem.EmployeeID = PREmployee.EmployeeID
        PRItem.ItemType = PREquate.ItemTypeOE
        PRItem.SDNumber = LastOE
        PRItem.Basis = OEBasis(LastOE)
        PRItem.UseEmployer = 1
        PRItem.Save (Equate.RecAdd)
    End If

End Sub
Private Sub ImportEEDED()
    
Dim DedNumber As Byte
Dim ItemFlag As Boolean
Dim FieldNumber As Byte
Dim LastDED As Integer
    
    PRItem.Clear
    LastDED = 0
    
    For i = 1 To 180
        
        DedNumber = Int((i - 1) / 18) + 1  ' DED numbers 1 to 10
        
        ' change in DED number - decide if to save
        If LastDED <> 0 And DedNumber <> LastDED Then
            If PRItem.Active = 1 Or LastDED = DirDepDed1 Or LastDED = DirDepDed2 Then
                PRItem.EmployeeID = PREmployee.EmployeeID
                If LastDED = DirDepDed1 Or LastDED = DirDepDed2 Then
                    PRItem.ItemType = PREquate.ItemTypeDirDepDed
                Else
                    PRItem.ItemType = PREquate.ItemTypeDED
                End If
                PRItem.SDNumber = LastDED
                PRItem.Basis = DEDBasis(LastDED)
                
                PRItem.UseEmployer = 1
                
                PRItem.Save (Equate.RecAdd)
            End If
            ItemFlag = False
            PRItem.Clear
            FieldNumber = 0
        End If
        
        Input #ASCIIChannel, X

        FieldNumber = FieldNumber + 1

        Select Case FieldNumber
        
            Case 1
                PRItem.AmtPct = CCur(X)
            Case 2
                PRItem.MaxAmount = CCur(X)
            Case 3
                If X = "X" Then PRItem.Active = 1
            Case 4
                If X = "X" Then PRItem.NoSSTax = 1
            Case 5
                If X = "X" Then PRItem.NoMedTax = 1
            Case 6
                If X = "X" Then PRItem.NoFWTTax = 1
            Case 7
                If X = "X" Then PRItem.NoSWTTax = 1
            Case 8
                If X = "X" Then PRItem.NoCWTTax = 1
            
            ' other tax flags in employer setup
            
            Case 14
                If X = "X" Then PRItem.NoFUNTax = 1
                If X = "X" Then PRItem.NoSUNTax = 1
        
        End Select
        
NextEEDED:
        LastDED = DedNumber
    
    Next i
            
    ' catch the last one
    If PRItem.Active = 1 Or LastDED = DirDepDed1 Or LastDED = DirDepDed2 Then
        PRItem.EmployeeID = PREmployee.EmployeeID
        
        If LastDED = DirDepDed1 Or LastDED = DirDepDed2 Then
            PRItem.ItemType = PREquate.ItemTypeDirDepDed
        Else
            PRItem.ItemType = PREquate.ItemTypeDED
        End If
        PRItem.SDNumber = LastDED
        PRItem.Basis = DEDBasis(LastDED)
        PRItem.UseEmployer = 1
        PRItem.Save (Equate.RecAdd)
    
    End If

End Sub

Private Sub ImportEEDirDep()

Dim DepAmount As Currency
Dim DedNumber As Byte
Dim DDCount As Byte

Dim BankName As String
Dim BankABA As String
Dim BankAccount As String
Dim BankType As String
Dim ItemID1, ItemID2 As Long

    ItemID1 = 0
    ItemID2 = 0

    For i = 1 To 2
                        
        For j = 1 To 4
            
            Input #ASCIIChannel, X
            
            If j = 1 Then BankName = X
            If j = 2 Then BankAccount = X
            If j = 3 Then BankABA = X
            If j = 4 Then BankType = X
            
        Next j
        
        If i = 1 Then
            DedNumber = DirDepDed1
        Else
            DedNumber = DirDepDed2
        End If
        
        ' find the PRItem record
        If Trim(BankName) <> "" And DedNumber <> 0 Then
                    
            SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID) & _
                         " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeDirDepDed) & _
                         " AND PRItem.SDNumber = " & CStr(DedNumber)
                         
            If Not PRItem.GetBySQL(SQLString) Then
                MsgBox "Dir Dep Item NF: EE# " & PREmployee.EmployeeNumber & vbCr & _
                       "Direct Deposit Setup will not be imported for this employee!", vbExclamation
            Else
                PRItem.DirDepABA = BankABA
                PRItem.DirDepAccount = BankAccount
                PRItem.DirDepBank = BankName
                PRItem.DirDepType = PREquate.DirDepTypeChecking
                If BankType = "SV" Or BankType = "sv" Then
                    PRItem.DirDepType = PREquate.DirDepTypeSavings
                End If
                
                PRItem.Active = 1
                
                PRItem.Save (Equate.RecPut)
            
                If i = 1 Then   ' use basis = net for the first dir dep type
                    PRItem.DirDepBasis = PREquate.DirDepBasisNet
                    ItemID1 = PRItem.ItemID
                Else
                    ItemID2 = PRItem.ItemID
                End If
            End If
        End If
        
    Next i
    
    ' fixed amount
    Input #ASCIIChannel, X
    
    ' no dir dep info inported - exit
    If ItemID1 = 0 Then Exit Sub
    
    DepAmount = CCur(X)
    
    If ItemID2 = 0 And DepAmount <> 0 Then
        If PRItem.GetByID(ItemID1) Then
            PRItem.DirDepBasis = PREquate.DirDepBasisAmt
            PRItem.DirDepAmtPct = DepAmount
            PRItem.Save (Equate.RecPut)
        End If
    ElseIf ItemID2 = 0 And DepAmount = 0 Then
        If PRItem.GetByID(ItemID1) Then
            PRItem.DirDepBasis = PREquate.DirDepBasisNet
            PRItem.Save (Equate.RecPut)
        End If
    ElseIf DepAmount <> 0 Then              ' ItemID2 is filled in
        If PRItem.GetByID(ItemID1) Then     ' use the amount for dir dep2
            PRItem.DirDepBasis = PREquate.DirDepBasisNet
            PRItem.DirDepAmtPct = 0
            PRItem.Save (Equate.RecPut)
        End If
        If PRItem.GetByID(ItemID2) Then
            PRItem.DirDepBasis = PREquate.DirDepBasisAmt
            PRItem.DirDepAmtPct = DepAmount
            PRItem.Save (Equate.RecPut)
        End If
    Else                                    ' ItemID2 is filled in
        If PRItem.GetByID(ItemID1) Then     ' amount not specified
            PRItem.DirDepBasis = PREquate.DirDepBasisNet
            PRItem.DirDepAmtPct = 0
            PRItem.Save (Equate.RecPut)
        End If
        If PRItem.GetByID(ItemID2) Then
            PRItem.DirDepBasis = PREquate.DirDepBasisAmt
            PRItem.DirDepAmtPct = 0
            PRItem.Save (Equate.RecPut)
        End If
    End If
    
End Sub

Private Sub ImportHistory()

Dim TotalOE As Currency
Dim TotalOEHours As Single
Dim TotalDed As Currency
Dim TotalDirDep As Currency

Dim OEHours(1 To 10) As Single
Dim OEAmount(1 To 10) As Currency
Dim TestDate As Date

Dim CWTTotal, CWTAmount, ERNTotal, ERNAmount As Currency
Dim SplitAmount, SWTTotal, SWTAmount As Currency
Dim RegDistID, DistID As Long
Dim OTXAmt As Currency

    ' *** SWT and CWT to be split among all PRDist records ***
    DistID = 0      ' storing the first PRDist item created
                    ' use for SWT and CWT rounding if necessary
    
    If ProgName <> "PRHIST" Then
        CtyID = 0
    End If
    
    PRHist.Clear
    
    For i = 1 To 53
        
        Input #ASCIIChannel, X

        If X = "" Or X = "0.00" Then GoTo NextHist

        If i = 1 Then PRHist.YearMonth = CLng(X)

        ' employee number - get EmployeeID
        If i = 2 Then
            SQLString = "SELECT * from PREmployee WHERE PREmployee.EmployeeNumber = " & CStr(X)
            If Not PREmployee.GetBySQL(SQLString) Then
                ' add a new employee
                PREmployee.Clear
                PREmployee.EmployeeNumber = X
                PREmployee.FirstName = "NEW"
                PREmployee.LastName = "EMPLOYEE"
                PREmployee.Save (Equate.RecAdd)
            End If
            PRHist.EmployeeID = PREmployee.EmployeeID
            PRHist.Save (Equate.RecAdd)     ' save it so a PRHistID is generated
        End If

        If i = 3 Then PRHist.CheckNumber = CLng(X)

        If i = 4 Then       ' PE Date yyyymmdd
            mm = Mid(X, 5, 2)
            dd = Mid(X, 7, 2)
            yy = Mid(X, 1, 4)
            PRHist.PEDate = DateSerial(yy, mm, dd)
        End If

        If i = 5 And X <> "0" Then       ' department number
            SQLString = "SELECT * FROM PRDepartment WHERE PRDepartment.DepartmentNumber = " & CStr(X)
            If Not PRDepartment.GetBySQL(SQLString) Then
            Else
                PRHist.DepartmentID = PRDepartment.DepartmentID
            End If
        End If

        If i = 6 Then PRHist.RegRate = CCur(X)
        ' i = 7 leave state number blank - assume Ohio for conversions
        If i = 8 Then PRHist.RegHours = CSng(X)
        If i = 9 Then PRHist.OTHours = CSng(X)

        If i >= 10 And i <= 19 Then
            TotalOEHours = TotalOEHours + CSng(X)
            OEHours(i - 9) = CSng(X)
        End If

        If i = 20 Then PRHist.RegAmount = CCur(X)
        If i = 21 Then PRHist.OTAmount = CCur(X)

        If i >= 22 And i <= 31 Then
            OEAmount(i - 21) = CCur(X)
        End If

        ' deductions - add to PRItemHist
        If i >= 32 And i <= 41 And X <> "0.00" Then

            ' find the PRItem record
            If DirDepDed1 = i - 31 Or DirDepDed2 = i - 31 Then
                SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeDirDepDed) & _
                            " AND PRItem.SDNumber = " & CStr(i - 31) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            Else
                SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeDED) & _
                            " AND PRItem.SDNumber = " & CStr(i - 31) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            End If
            
            ' not found - get the employer item and add an employee item
            If Not PRItem.GetBySQL(SQLString) Then
                
                SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 AND PRItem.SDNumber = " & CStr(i - 31) & _
                            " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeDED)
                
                If Not PRItem.GetBySQL(SQLString) Then
                    MsgBox "Employer PRItem not found - ded # " & CStr(i - 31) & " " & PREmployee.EmployeeNumber, vbCritical
                    End
                End If
                
                PRItem.UseEmployer = 1
                PRItem.EmployeeID = PREmployee.EmployeeID
                PRItem.EmployerItemID = PRItem.ItemID
                PRItem.Save (Equate.RecAdd)
            
            End If

            PRItemHist.Clear
            PRItemHist.EmployeeID = PRHist.EmployeeID
            PRItemHist.HistID = PRHist.HistID
            PRItemHist.DepartmentID = PRHist.DepartmentID
            PRItemHist.ItemID = PRItem.ItemID
            PRItemHist.EmployerItemID = PRItem.EmployerItemID
            PRItemHist.Hours = 0
            PRItemHist.Amount = CCur(X)
            PRItemHist.ManualAmount = 1
            PRItemHist.YearMonth = PRHist.YearMonth
            PRItemHist.PEDate = PRHist.PEDate

            If i - 31 = DirDepDed1 Or i - 31 = DirDepDed2 Then
                PRItemHist.ItemType = PREquate.ItemTypeDirDepDed
            Else
                PRItemHist.ItemType = PREquate.ItemTypeDED
            End If
            
            PRItemHist.Save (Equate.RecAdd)

        End If

        If i >= 22 And i <= 31 Then
            TotalOE = TotalOE + CCur(X)
        End If

        If i >= 32 And i <= 41 Then
            If i - 31 = DirDepDed1 Or i - 31 = DirDepDed2 Then
                TotalDirDep = TotalDirDep + CCur(X)
            Else
                TotalDed = TotalDed + CCur(X)
            End If
        End If

        If i = 42 Then PRHist.SSTax = CCur(X)
        If i = 43 Then PRHist.MedTax = CCur(X)
        If i = 44 Then PRHist.FWTTax = CCur(X)
        If i = 45 Then PRHist.SWTTax = CCur(X)
        If i = 46 Then PRHist.CWTTax = CCur(X)

        '  other tax as city tax?
        If i >= 47 And i <= 51 And PRHist.CWTTax = 0 And CtyID = 0 Then
            If i = 47 And Tax6City <> 0 Then
                PRHist.CWTTax = CCur(X)
                CtyID = Tax6City
            End If
            If i = 48 And Tax7City <> 0 Then
                PRHist.CWTTax = CCur(X)
                CtyID = Tax7City
            End If
            If i = 49 And Tax8City <> 0 Then
                PRHist.CWTTax = CCur(X)
                CtyID = Tax8City
            End If
            If i = 50 And Tax9City <> 0 Then
                PRHist.CWTTax = CCur(X)
                CtyID = Tax9City
            End If
            If i = 51 And Tax0City <> 0 Then
                PRHist.CWTTax = CCur(X)
                CtyID = Tax0City
            End If
        End If

        OTXAmt = CCur(X)
        If i = 47 And SDTax6ID <> 0 And OTXAmt <> 0 Then
            AddSDTax SDTax6ID, OTXAmt
        End If
        If i = 48 And SDTax7ID <> 0 And OTXAmt <> 0 Then
            AddSDTax SDTax7ID, OTXAmt
        End If
        If i = 49 And SDTax8ID <> 0 And OTXAmt <> 0 Then
            AddSDTax SDTax8ID, OTXAmt
        End If
        If i = 50 And SDTax9ID <> 0 And OTXAmt <> 0 Then
            AddSDTax SDTax9ID, OTXAmt
        End If
        If i = 51 And SDTax0ID <> 0 And OTXAmt <> 0 Then
            AddSDTax SDTax0ID, OTXAmt
        End If

        If i = 52 Then PRHist.Gross = CCur(X)
        If i = 53 Then PRHist.Net = CCur(X)

NextHist:
    
    Next i

    ' final updates
    PRHist.Deductions = TotalDed
    PRHist.OEAmount = TotalOE
    PRHist.OEHours = TotalOEHours
    PRHist.DirectDeposit = TotalDirDep

    ' set manual flags for all imports
    PRHist.ManualSSTax = 1
    PRHist.ManualMedTax = 1
    PRHist.ManualFWTTax = 1
    PRHist.StateID = DfltStateID

    PRHist.Save (Equate.RecPut)

    ' !!!!!!!!!!!!!!!!!!!!
    ' dont write to PRDist for dist companies
    ' !!!!!!!!!!!!!!!!!!!!
    If DistFlag = True Then Exit Sub

    ' init the variables
    CWTTotal = PRHist.CWTTax
    CWTAmount = 0
    SWTTotal = PRHist.SWTTax
    SWTAmount = 0
    ERNTotal = PRHist.Gross
    ERNAmount = 0

    ' write regular and overtime to PRDist
    PRDist.Clear
    
    ' ================================================
    ' if using other tax as city tax
    ' assign the cityid now
    ' dont overwrite later
    If CtyID <> 0 Then PRDist.CityID = CtyID
    ' ================================================
    
    PRDist.EmployeeID = PREmployee.EmployeeID
    PRDist.HistID = PRHist.HistID
    PRDist.DepartmentID = PRHist.DepartmentID
    PRDist.YearMonth = PRHist.YearMonth
    PRDist.PEDate = PRHist.PEDate
    PRDist.DistType = PREquate.DistTypeReg
    
    ' use the employer item id
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeRegPay) & _
                " AND PRItem.EmployeeID = 0"
    If Not PRItem.GetBySQL(SQLString) Then
        MsgBox "Regular pay item nf: ", vbCritical
        End
    End If
    PRDist.ItemID = PRItem.ItemID
    PRDist.EmployerItemID = PRItem.EmployerItemID
    PRDist.ItemType = PREquate.ItemTypeRegPay
    PRDist.Amount = PRHist.RegAmount
    PRDist.ManualAmount = 1
    PRDist.Hours = PRHist.RegHours
    PRDist.Rate = PRHist.RegRate
    
    PRDist.StateID = DfltStateID
    PRDist.HistFlag = 1
    
    ' split the state and city tax amounts
    CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
    PRDist.CityTax = CWTAmount
    CWTTotal = CWTTotal - CWTAmount
    PRDist.ManualCityTax = 1
    
    SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
    PRDist.StateTax = SWTAmount
    SWTTotal = SWTTotal - SWTAmount
    PRDist.ManualStateTax = 1
    
    PRDist.Save (Equate.RecAdd)
    
    RegDistID = PRDist.DistID
    If PRDist.Amount <> 0 Then DistID = PRDist.DistID
    
    ' Over Time
    If PRHist.OTAmount <> 0 Then
    
        PRDist.Clear
        
        ' ================================================
        ' if using other tax as city tax
        ' assign the cityid now
        ' dont overwrite later
        If CtyID <> 0 Then PRDist.CityID = CtyID
        ' ================================================
        
        PRDist.EmployeeID = PREmployee.EmployeeID
        PRDist.HistID = PRHist.HistID
        PRDist.DepartmentID = PRHist.DepartmentID
        PRDist.YearMonth = PRHist.YearMonth
        PRDist.PEDate = PRHist.PEDate
        PRDist.DistType = PREquate.DistTypeOT
        
        ' use the employer item id
        SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeOvtPay) & _
                    " AND PRItem.EmployeeID = 0"
        If Not PRItem.GetBySQL(SQLString) Then
            MsgBox "Overtime pay item nf: ", vbCritical
            End
        End If
        PRDist.ItemID = PRItem.ItemID
        PRDist.ItemType = PREquate.ItemTypeOvtPay
        PRDist.Amount = PRHist.OTAmount
        PRDist.ManualAmount = 1
        PRDist.Hours = PRHist.OTHours
        PRDist.Rate = PRHist.OTRate
        
        PRDist.StateID = DfltStateID
        PRDist.HistFlag = 1
        
        CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
        PRDist.CityTax = CWTAmount
        CWTTotal = CWTTotal - CWTAmount
        PRDist.ManualCityTax = 1
        
        SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
        PRDist.StateTax = SWTAmount
        SWTTotal = SWTTotal - SWTAmount
        PRDist.ManualStateTax = 1
        
        PRDist.Save (Equate.RecAdd)
    
        If DistID = 0 Then DistID = PRDist.DistID
    
    End If
    
    ' use the employer item id
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeRegPay) & _
                " AND PRItem.EmployeeID = 0"
    If Not PRItem.GetBySQL(SQLString) Then
        MsgBox "Regular pay item nf: ", vbCritical
        End
    End If
    PRDist.ItemID = PRItem.ItemID

    ' write other earnings to PRDist
    For i = 1 To 10
        If OEHours(i) <> 0 Or OEAmount(i) <> 0 Then
            
            PRDist.Clear
        
            ' ================================================
            ' if using other tax as city tax
            ' assign the cityid now
            ' dont overwrite later
            If CtyID <> 0 Then PRDist.CityID = CtyID
            ' ================================================

            PRDist.EmployeeID = PREmployee.EmployeeID
            PRDist.HistID = PRHist.HistID
            ' state id
            ' city id
            ' job id
            ' customer id
            PRDist.DepartmentID = PRHist.DepartmentID
            PRDist.YearMonth = PRHist.YearMonth
            PRDist.PEDate = PRHist.PEDate
            PRDist.CheckDate = PRHist.CheckDate
            PRDist.DistType = PREquate.DistTypeItem
            PRDist.ItemType = PREquate.ItemTypeOE
            
            SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeOE) & _
                        " AND PRItem.SDNumber = " & CStr(i) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            
            If Not PRItem.GetBySQL(SQLString) Then
                
                SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 AND PRItem.SDNumber = " & CStr(i) & _
                            " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeOE)
                
                If Not PRItem.GetBySQL(SQLString) Then
                    MsgBox "Employer PRItem not found - oe # " & CStr(i), vbCritical
                    End
                End If
                
                PRItem.EmployeeID = PREmployee.EmployeeID
                PRItem.EmployerItemID = PRItem.ItemID
                PRItem.ItemID = 0
                PRItem.UseEmployer = 1
                PRItem.Save (Equate.RecAdd)
            
            End If
            
            
            PRDist.ItemID = PRItem.ItemID
            
            PRDist.Hours = OEHours(i)
            If OEHours(i) = 0 Then
                PRDist.Rate = OEAmount(i)
            Else
                PRDist.Rate = OEAmount(i) / OEHours(i)
            End If
            PRDist.Amount = OEAmount(i)
            PRDist.ManualAmount = 1
            
            ' billing rate
            ' state wage
            ' state tax
            ' city wage
            ' city tax
            
            PRDist.HistFlag = 1
            PRDist.StateID = DfltStateID
        
            CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
            PRDist.CityTax = CWTAmount
            CWTTotal = CWTTotal - CWTAmount
            PRDist.ManualCityTax = 1
            
            SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
            PRDist.StateTax = SWTAmount
            SWTTotal = SWTTotal - SWTAmount
            PRDist.ManualStateTax = 1
            
            PRDist.Save (Equate.RecAdd)
            
            If DistID = 0 Then DistID = PRDist.DistID
        
        End If
    
    Next i

    ' if no earnging amounts - refer to the PRDist from Regular
    '    which is always added
    If DistID = 0 Then DistID = RegDistID

    ' rounding correction???
    If SWTTotal <> 0 Then
        If DistID = 0 Then
            MsgBox "No amounts ???" & vbCr & _
                   PREmployee.EmployeeNumber & vbCr & _
                   PREmployee.LFName & vbCr & _
                   PRHist.HistID & vbCr & _
                   PRHist.YearMonth, vbExclamation
            End
        End If
        If Not PRDist.GetByID(DistID) Then
            MsgBox "PRDist err: " & DistID, vbExclamation
            End
        End If
        PRDist.StateTax = PRDist.StateTax + SWTTotal
        PRDist.Save (Equate.RecPut)
    End If

    If CWTTotal <> 0 Then
        If DistID = 0 Then
            MsgBox "No amounts ???", vbExclamation
            End
        End If
        If Not PRDist.GetByID(DistID) Then
            MsgBox "PRDist err: " & DistID, vbExclamation
            End
        End If
        PRDist.CityTax = PRDist.CityTax + CWTTotal
        PRDist.Save (Equate.RecPut)
    End If

End Sub

Private Function AddSDTax(ByVal ItemID As Long, ByVal SDTaxAmt As Currency)

Dim SDPct As Currency

    ' get the employer item
    If PRItem.GetByID(ItemID) = False Then Exit Function
    SDPct = PRItem.MaxPct
    
    ' find the PRItem?
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & PRHist.EmployeeID & _
                " AND EmployerItemID = " & ItemID
    If Not PRItem.GetBySQL(SQLString) Then
        PRItem.Clear
        PRItem.EmployerItemID = ItemID
        PRItem.EmployeeID = PRHist.EmployeeID
        PRItem.ItemType = PREquate.ItemTypeSDTax
        PRItem.Basis = PREquate.BasisPercent
        PRItem.Active = 1
        PRItem.UseEmployer = 1
        PRItem.AmtPct = SDPct
        PRItem.Save (Equate.RecAdd)
    End If
    
    ' add the PRItemHist record
    PRItemHist.Clear
    PRItemHist.EmployeeID = PRHist.EmployeeID
    PRItemHist.HistID = PRHist.HistID
    PRItemHist.BatchID = PRHist.BatchID
    PRItemHist.ItemID = PRItem.ItemID
    PRItemHist.DepartmentID = PRHist.DepartmentID
    PRItemHist.EmployerItemID = PRItem.EmployerItemID
    PRItemHist.ItemType = PREquate.ItemTypeSDTax
    PRItemHist.YearMonth = PRHist.YearMonth
    PRItemHist.CheckDate = PRHist.CheckDate
    PRItemHist.PEDate = PRHist.PEDate
    PRItemHist.Hours = 0
    PRItemHist.Rate = 0
    PRItemHist.Amount = SDTaxAmt
    PRItemHist.Save (Equate.RecAdd)

End Function

Private Sub ImportDist()

Dim TypeCode As Byte
Dim IType As Byte
Dim SDNum As Byte

    PRDist.Clear

    For i = 1 To 12
        
        Input #ASCIIChannel, X

        ' type code
        If i = 1 Then
            
            TypeCode = CByte(X)
                        
            If TypeCode < 1 Or TypeCode > 12 Then
                MsgBox "Bad Type Code: " & CStr(TypeCode), vbCritical
                End
            End If
                        
            If TypeCode = 1 Then IType = PREquate.ItemTypeRegPay
            If TypeCode = 2 Then IType = PREquate.ItemTypeOvtPay
            
            If TypeCode >= 3 And TypeCode <= 12 Then
                IType = PREquate.ItemTypeOE
                SDNum = TypeCode - 2
            End If
        
            PRDist.ItemType = IType
        
        End If

        If i = 2 Then PRDist.YearMonth = CLng(X)
        
        If i = 3 Then       ' employee number
            
            SQLString = "SELECT * FROM PREmployee WHERE EmployeeNumber = " & X
            If Not PREmployee.GetBySQL(SQLString) Then
                MsgBox "Employee # " & X & " from Dist " & PRDist.YearMonth & " NF", vbCritical
                End
            End If
        
            PRDist.EmployeeID = PREmployee.EmployeeID
        
        End If
        
        If i = 4 Then       ' city number
        
            SQLString = "SELECT * FROM PRCity WHERE CityNumber = " & X
            If Not PRCity.GetBySQL(SQLString) Then
                MsgBox "City # " & X & " from Dist " & PRDist.YearMonth & " NF", vbCritical
                End
            End If
            
            PRDist.CityID = PRCity.CityID
            
        End If
        
        If i = 5 Then PRDist.CityTax = CCur(X)
        
        If i = 6 Then
            
            mm = Mid(X, 5, 2)
            dd = Mid(X, 7, 2)
            yy = Mid(X, 1, 4)
            PRDist.PEDate = DateSerial(yy, mm, dd)
            
        End If
        
        If i = 7 Then PRDist.Hours = CSng(X)
        ' 8 = multiplier
        If i = 9 Then PRDist.Rate = CCur(X)
        If i = 10 Then PRDist.Amount = CCur(X)
        
        If i = 11 And X <> "0" Then
            SQLString = "SELECT * FROM PRDepartment WHERE DepartmentNumber = " & X
            If Not PRDepartment.GetBySQL(SQLString) Then
            Else
                PRDist.DepartmentID = PRDepartment.DepartmentID
            End If
        End If
        
        ' 12 = job number
        
    Next i
        
    ' get the PRItem
    If TypeCode = 1 Or TypeCode = 2 Then    ' Reg / Ovt
        
        SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(IType)
        If Not PRItem.GetBySQL(SQLString) Then
            MsgBox "Reg or OT Pay item NF", vbCritical
            End
        End If
    
    Else
        
        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & CStr(PREmployee.EmployeeID) & _
                    " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeOE) & _
                    " AND PRItem.SDNumber = " & CStr(SDNum)
        
        If Not PRItem.GetBySQL(SQLString) Then
            
            ' get from employer
            SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0" & _
                        " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeOE) & _
                        " AND PRItem.SDNumber = " & CStr(SDNum)
            If Not PRItem.GetBySQL(SQLString) Then
                MsgBox "Employer OE# " & CStr(SDNum) & " PRItem NF", vbCritical
                End
            End If
        End If
    End If
    
    If TypeCode = 1 Then PRDist.DistType = PREquate.DistTypeReg
    If TypeCode = 2 Then PRDist.DistType = PREquate.DistTypeOT
    If TypeCode > 2 Then PRDist.DistType = PREquate.DistTypeItem
    
    PRDist.StateID = PRState.StateID
    PRDist.ItemID = PRItem.ItemID
    
    PRDist.ManualAmount = 1
    PRDist.ManualCityTax = 1
    PRDist.ManualStateTax = 1
    
    PRDist.Save (Equate.RecAdd)

End Sub

Private Sub EEItemUpdate()
    
' Assign the EmployerItemID for Employee PRItem records
    
Dim IID As Long
    
    Progress.lblMsg2 = "Now Updating Employee Item Records"
    Progress.lblMsg2.Refresh
    
    ' set up temp record set
    trs.CursorLocation = adUseClient
   
    trs.Fields.Append "ItemType", adInteger
    trs.Fields.Append "SDNumber", adDouble
    trs.Fields.Append "ItemID", adDouble
    
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    ' load up Employer Items
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0"
    If Not PRItem.GetBySQL(SQLString) Then
        MsgBox "No Employer items found!", vbCritical
        End
    End If
    
    Do
    
        IID = PRItem.ItemType * 100 + PRItem.SDNumber
    
        trs.AddNew
        trs.Fields("SDNumber") = IID
        trs.Fields("ItemID") = PRItem.ItemID
        trs.Update
        
        If Not PRItem.GetNext Then Exit Do
        
    Loop
    
    ' load up Employee Items
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID <> 0"
    If Not PRItem.GetBySQL(SQLString) Then
        Exit Sub
    End If
    
    Do
    
        ' standard other earnings and deductions
        If PRItem.ItemType <> PREquate.ItemTypeDirDepDed And PRItem.ItemType <> PREquate.ItemTypeSDTax Then
    
            IID = PRItem.ItemType * 100 + PRItem.SDNumber
            SQLString = "SDNumber = " & CStr(IID)
            trs.Find SQLString, 0, adSearchForward, 1
            
            If trs.EOF Then
                MsgBox "Employer Item NF: " & PRItem.ItemType & " " & PRItem.SDNumber
                End
            End If
            
            PRItem.EmployerItemID = trs!ItemID
            PRItem.Save (Equate.RecPut)
    
        End If
    
        If Not PRItem.GetNext Then Exit Do
    
    Loop

    Set trs = Nothing

End Sub

Private Sub DistMatch()

Dim Gross As Currency

    SQLString = "SELECT * FROM PREmployee"
    If Not PREmployee.GetBySQL(SQLString) Then Exit Sub        ' no employees ???
    
    Do
    
        ' get the history records
        SQLString = "SELECT * FROM PRHist WHERE EmployeeID = " & CStr(PREmployee.EmployeeID) & _
                    " ORDER BY PRHist.HistID"
        If Not PRHist.GetBySQL(SQLString) Then
            GoTo NextEmployee
        End If
        
        Do
            
            Gross = 0
            
            SQLString = "SELECT * FROM PRDist WHERE PRDist.EmployeeID = " & CStr(PREmployee.EmployeeID) & _
                        " AND PRDist.YearMonth = " & CStr(PRHist.YearMonth) & _
                        " AND PRDist.PEDate = " & CLng(PRHist.PEDate) & _
                        " AND PRDist.HistID = 0 ORDER BY PRDist.DistID"
                        
            If Not PRDist.GetBySQL(SQLString) Then
                GoTo NextHist
            End If
            
            Do
            
                PRDist.HistID = PRHist.HistID
                PRDist.Save (Equate.RecPut)
            
                Gross = Gross + PRDist.Amount
                If Gross >= PRHist.Gross Then
                    If Gross > PRHist.Gross Then
                        MsgBox "Warning: Hist/Dist Match for EmployeeID: " & PREmployee.EmployeeID & _
                               " PE Date: " & PRHist.PEDate & vbCr & _
                               " Dist Gross: " & Gross & vbCr & _
                               " Hist Gross: " & PRHist.Gross, vbExclamation
                    End If
                    
                    Exit Do
                End If
            
                If Not PRDist.GetNext Then Exit Do
            
            Loop
            
NextHist:
            If Not PRHist.GetNext Then Exit Do
        
        Loop

NextEmployee:
        If Not PREmployee.GetNext Then Exit Do
    
    Loop

End Sub

Private Sub SelectDfltCity()
    Do
        MsgBox "Please select the default for city withholding", vbInformation, "Windows PR Import"
        ModeSelect = True
        frmPRCity.Show vbModal
        DfltCityID = frmPRCity.SelectedCityID
        If DfltCityID <> 0 Then Exit Do
    Loop
    Unload frmPRCity
End Sub

Private Sub SetPRDistCityID()
    
    SQLString = "SELECT * FROM PRCity WHERE CityID = " & CStr(DfltCityID)
    If Not PRCity.GetBySQL(SQLString) Then
        MsgBox "City NF: ???"
        End
    End If
    Progress.lblMsg2 = "Now updating History with City: " & PRCity.CityName
    Progress.lblMsg2.Refresh
            
    SQLString = "SELECT * FROM PRDist WHERE CityID = 0"
    If Not PRDist.GetBySQL(SQLString) Then Exit Sub
    Do
        PRDist.CityID = DfltCityID
        PRDist.Save (Equate.RecPut)
        If Not PRDist.GetNext Then Exit Do
    Loop

End Sub

Private Sub SetPREmployeeCityID()

    SQLString = "SELECT * FROM PRCity WHERE CityID = " & CStr(DfltCityID)
    If Not PRCity.GetBySQL(SQLString) Then
        MsgBox "City NF: ???"
        End
    End If
    Progress.lblMsg2 = "Now updating Employees with City: " & PRCity.CityName
    Progress.lblMsg2.Refresh
            
    SQLString = "SELECT * FROM PREmployee"
    If Not PREmployee.GetBySQL(SQLString) Then Exit Sub
    Do
        PREmployee.DefaultCityID = DfltCityID
        PREmployee.Save (Equate.RecPut)
        If Not PREmployee.GetNext Then Exit Do
    Loop

End Sub
Private Sub CreatePRBatch()

Dim LastYearMonth As Long
Dim LastPED As Date
Dim RecCount As Long
Dim YM, yy, mm, dd As Integer

    Progress.lblMsg2 = "Now creating Batch Records"
    Progress.lblMsg2.Refresh

    ' create PRBatch from PRHist
    ' one record per PE Date
            
    LastYearMonth = 0
    LastPED = 0
    RecCount = 0
        
    SQLString = "SELECT * FROM PRHist ORDER BY YearMonth, PEDate"
    If Not PRHist.GetBySQL(SQLString) Then Exit Sub
    
    Do
    
        ' create the PRBatch Record
        If LastYearMonth = 0 Or LastPED = 0 Or LastYearMonth <> PRHist.YearMonth Or LastPED <> PRHist.PEDate Then
        
            ' update the fields
            If LastYearMonth <> 0 Then
                PRBatch.RecCount = RecCount
                PRBatch.Save (Equate.RecPut)
                RecCount = 0
            End If
            
            PRBatch.Clear
            PRBatch.UserID = User.ID
            PRBatch.CreateDate = Now()
            PRBatch.PEDate = PRHist.PEDate
            PRBatch.YearMonth = PRHist.YearMonth
            
            ' must be in same YearMonth
            yy = Year(PRBatch.PEDate)
            mm = Month(PRBatch.PEDate)
            dd = Day(PRBatch.PEDate) + frmStart.CheckDays
            PRBatch.CheckDate = DateSerial(yy, mm, dd)
            YM = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
            If YM <> PRBatch.YearMonth Then
                yy = Int(PRBatch.YearMonth) / 100
                mm = PRBatch.YearMonth Mod 100
                PRBatch.CheckDate = DateSerial(yy, mm, 1)
            End If
            
            PRBatch.Save (Equate.RecAdd)
        
        End If
        
        LastYearMonth = PRHist.YearMonth
        LastPED = PRHist.PEDate
        RecCount = RecCount + 1
    
        ' update the prhist record
        PRHist.BatchID = PRBatch.BatchID
        PRHist.CheckDate = PRBatch.CheckDate
        PRHist.Save (Equate.RecPut)
    
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    ' update the last batch record
    PRBatch.RecCount = RecCount
    PRBatch.Save (Equate.RecPut)

    Progress.lblMsg2 = "Now Updating Batch Numbers to Distribution"
    Progress.lblMsg2.Refresh

    ' update the BatchID to PRDist
    SQLString = "SELECT * FROM PRHist"
    If Not PRHist.GetBySQL(SQLString) Then Exit Sub
    
    Do
    
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & CStr(PRHist.HistID)
        If PRDist.GetBySQL(SQLString) Then
        
            Do
            
                PRDist.BatchID = PRHist.BatchID
                PRDist.CheckDate = PRHist.CheckDate
                PRDist.Save (Equate.RecPut)
                
                If Not PRDist.GetNext Then Exit Do
                
            Loop
            
        End If
        
        ' update PRItemHist also
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID
        If PRItemHist.GetBySQL(SQLString) Then
        
            Do
            
                PRItemHist.BatchID = PRHist.BatchID
                PRItemHist.CheckDate = PRHist.CheckDate
                PRItemHist.Save (Equate.RecPut)
                
                If Not PRItemHist.GetNext Then Exit Do
                
            Loop
            
        End If
        
        If Not PRHist.GetNext Then Exit Do
        
    Loop
    
    
End Sub

Private Sub EEDeptAssign()

    ' PREmployee.DepartmentID is set to the department number on import
    '    switch it to the DepartmentID

    SQLString = "SELECT * FROM PREmployee WHERE DepartmentID <> 0"
    
    If Not PREmployee.GetBySQL(SQLString) Then Exit Sub
    
    Do
    
        SQLString = "SELECT * FROM PRDepartment WHERE DepartmentNumber = " & CStr(PREmployee.DepartmentID)
                
        If Not PRDepartment.GetBySQL(SQLString) Then       ' add it
        
            PRDepartment.Clear
            PRDepartment.DepartmentNumber = PREmployee.DepartmentID
            PRDepartment.Name = "Dept " & CStr(PREmployee.DepartmentID)
            PRDepartment.Save (Equate.RecAdd)
            
        End If
        
        PREmployee.DepartmentID = PRDepartment.DepartmentID
        PREmployee.Save (Equate.RecPut)
        
        If Not PREmployee.GetNext Then Exit Do
        
    Loop
    
End Sub

Private Sub ImportSkip(NumFields As Integer)

Dim iSkip As Integer

    For iSkip = 1 To NumFields
        Input #ASCIIChannel, X
    Next iSkip

End Sub

Private Function MakeBoo(ByVal InputString As String) As Byte

    If CByte(InputString) <> 0 Or InputString = "X" Then
        MakeBoo = 1
    Else
        MakeBoo = 0
    End If

End Function

Public Sub FWTImport()
    
Dim FWTID As String
Dim FWTYear, Yr As Long
Dim FWTMonth As Byte
Dim MS As String
Dim LowAmount As Currency
Dim HiAmount As Currency
Dim AddAmount As Currency
Dim Pct As Currency
Dim PctBasis As Currency
Dim FWTAmt As Long
    
    Progress.lblMsg2 = "Now importing federal tax tables"
    Progress.lblMsg2.Refresh

    PRFWTTable.DeleteAll
    PRFWTTable.OpenRS
        
    ASCIIChannel = FreeFile
    
    On Error Resume Next
    TextName = "\Balint\Blank\FWTTable.csv"
    Open TextName For Input As ASCIIChannel
    If Err.Number <> 0 Then
        MsgBox "\Balint\Blank\FWTTable.csv Error: " & Err.Number & vbCr & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' get rid of header line
    Input #ASCIIChannel, FWTID, FWTYear, FWTMonth, MS, LowAmount, HiAmount, AddAmount, Pct, PctBasis
    
    RCT = 0
       
    Do While Not EOF(ASCIIChannel)
   
        Input #ASCIIChannel, FWTID, FWTYear, FWTMonth, MS, LowAmount, HiAmount, AddAmount, Pct, PctBasis
      
        If Trim(MS) = "" Then Exit Do
      
        PRFWTTable.Clear
        
        ' FWT or Ohio SWT?
        If FWTID = "FWT" Then
            PRFWTTable.StateID = 0
        Else
            SQLString = "SELECT * FROM PRState WHERE PRState.StateAbbrev = " & "'" & Trim(FWTID) & "'"
            If Not PRState.GetBySQL(SQLString) Then
                MsgBox "State NF: " & FWTID, vbCritical
                End
            End If
            PRFWTTable.StateID = PRState.StateID
        End If
      
        PRFWTTable.TaxYear = FWTYear
        PRFWTTable.TaxMonth = FWTMonth
        
        If MS = "M" Then
            PRFWTTable.msMarried = 1
            PRFWTTable.msSingle = 0
        ElseIf MS = "S" Then
            PRFWTTable.msMarried = 0
            PRFWTTable.msSingle = 1
        ElseIf MS = "X" Then                ' Ohio not per marital status
            PRFWTTable.msMarried = 0
            PRFWTTable.msSingle = 0
        Else
            MsgBox "Marital Status ??? " & MS, vbCritical
            End
        End If
        
        PRFWTTable.LowAmount = LowAmount
        PRFWTTable.HiAmount = HiAmount
        PRFWTTable.Amount = AddAmount
        PRFWTTable.Percent = Pct
        PRFWTTable.ExcessBase = PctBasis
        
        PRFWTTable.Save (Equate.RecAdd)
        
    Loop

End Sub

Private Sub ImportPRAcct(ByVal LineNum As Byte)

Dim Dpt As Integer
Dim pFlag As Boolean
Dim iFlag As Boolean
Dim GLAcct As Long

    ' get the department number
    Input #ASCIIChannel, X
    Dpt = CInt(X)
    
    ' get the department record
    ' if invalid dept# or find fails - still need to read the fields in
    ' form the import file
    pFlag = True
    If Dpt < 10 Then pFlag = False
    If Dpt > 99 Then pFlag = False
    
    SQLString = "SELECT * FROM PRDepartment WHERE DepartmentNumber = " & X
    If Not PRDepartment.GetBySQL(SQLString) Then pFlag = False
    
    For i = 1 To 20
        
        Input #ASCIIChannel, X
        GLAcct = CLng(X)
        
        If pFlag And GLAcct <> 0 Then
        
            PRGLUpd.Clear
            PRGLUpd.GLType = PREquate.GLTypeDept
            PRGLUpd.RelatedID = PRDepartment.DepartmentID
            PRGLUpd.GLAccountNum = GLAcct
            
            If LineNum = 1 Then
                If i >= 1 And i <= 10 Then      ' Other Earnings
                    
                    ' get the PRItem record
                    SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeOE & _
                                " AND SDNumber = " & i
                    If PRItem.GetBySQL(SQLString) Then
                        iFlag = True
                    Else
                        iFlag = False
                    End If
                    
                    PRGLUpd.GLItemType = PREquate.GLItemTypeOE
                    PRGLUpd.ItemID = PRItem.ItemID  ' ***
                
                Else                        ' deductions
                    
                    ' get the PRItem record
                    SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeDED & _
                                " AND SDNumber = " & i - 10
                    If PRItem.GetBySQL(SQLString) Then
                        iFlag = True
                    Else
                        iFlag = False
                    End If
                    
                    PRGLUpd.GLItemType = PREquate.GLItemTypeDed
                    PRGLUpd.ItemID = PRItem.ItemID ' ***
                
                End If
            Else
                If i = 1 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSSTax
                If i = 2 Then PRGLUpd.GLItemType = PREquate.GLItemTypeMedTax
                If i = 3 Then PRGLUpd.GLItemType = PREquate.GLItemTypeFWTTax
                If i = 4 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSWTTax
                If i = 5 Then PRGLUpd.GLItemType = PREquate.GLItemTypeCWTTax
                If i = 11 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSUN
                If i = 12 Then PRGLUpd.GLItemType = PREquate.GLItemTypeFUN
                If i = 13 Then PRGLUpd.GLItemType = PREquate.GLItemTypeWkcExp
                If i = 14 Then PRGLUpd.GLItemType = PREquate.GLItemTypeGross
                If i = 15 Then PRGLUpd.GLItemType = PREquate.GLItemTypeNet
                If i = 16 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSSExp
                If i = 17 Then PRGLUpd.GLItemType = PREquate.GLItemTypeMEDExp
                If i = 18 Then PRGLUpd.GLItemType = PREquate.GLItemTypeSUNExp
                If i = 19 Then PRGLUpd.GLItemType = PREquate.GLItemTypeFUNExp
                If i = 20 Then PRGLUpd.GLItemType = PREquate.GLItemTypeWkcExp
                PRGLUpd.ItemID = 0
                If i >= 6 And i <= 10 Then      ' skip other taxes from SuperDOS
                    iFlag = False
                Else
                    iFlag = True
                End If
            End If
                
            If iFlag Then PRGLUpd.Save (Equate.RecAdd)
        
        End If
            
    Next i

End Sub

Public Sub TaxMaxImport()

Dim TaxYear As Integer
Dim TaxAmt As Currency
Dim TaxType As String
Dim TaxEquate As Byte
Dim TaxState As String
Dim TaxDesc As String
Dim AddFlag As Boolean
    
    Progress.lblMsg2 = "Now Importing tax tables"
    Progress.lblMsg2.Refresh
    
    ASCIIChannel = FreeFile
    TextName = "\Balint\Blank\TaxMax.csv"
    On Error Resume Next
    Open TextName For Input As ASCIIChannel
    If Err.Number <> 0 Then
        MsgBox "\Balint\Blank\TaxMax.csv Error # " & Err.Number & vbCr & Err.Description, vbExclamation
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' get rid of header line
    Input #ASCIIChannel, TaxYear, TaxDesc, TaxAmt
    
    RCT = 0
       
    Do While Not EOF(ASCIIChannel)
   
        Input #ASCIIChannel, TaxYear, TaxType, TaxAmt
      
        TaxEquate = 0
        TaxState = ""
        If TaxType = "FICA" Then
            TaxEquate = PREquate.GlobalTypeSSMax
            TaxDesc = "SS Max"
        ElseIf TaxType = "ALLW" Then
            TaxEquate = PREquate.GlobalTypeFWTAllow
            TaxDesc = "FWT Allow"
        ElseIf TaxType = "FDUN" Then
            TaxEquate = PREquate.GlobalTypeFUNMax
            TaxDesc = "FUN Max"
        ElseIf Mid(TaxType, 3, 2) = "UN" Then
            TaxEquate = PREquate.GlobalTypeSUNMax
            TaxState = Mid(TaxType, 1, 2)
        ElseIf TaxType = "SSP" Then
            TaxEquate = PREquate.GlobalTypeSSPct
            TaxDesc = "SS Pct"
        ElseIf TaxType = "MEDP" Then
            TaxEquate = PREquate.GlobalTypeMEDPct
            TaxDesc = "MED Pct"
        ElseIf TaxType = "OHALLW" Then
            TaxEquate = PREquate.GLobalTypeOHAllow
            TaxDesc = "OH Allow"
        End If

        ' update to PRGlobal
        If TaxEquate <> 0 Then
        
            If TaxState = "" Then
                SQLString = "SELECT * FROM PRGlobal WHERE PRGlobal.TypeCode = " & TaxEquate & _
                            " AND PRGlobal.Year = " & TaxYear
                If Not PRGlobal.GetBySQL(SQLString) Then
                    PRGlobal.TypeCode = TaxEquate
                    PRGlobal.Year = TaxYear
                    PRGlobal.Save (Equate.RecAdd)
                End If
                PRGlobal.Amount = TaxAmt
                PRGlobal.Description = TaxDesc
                PRGlobal.Save (Equate.RecPut)
            Else
                AddFlag = True
                SQLString = "SELECT * FROM PRGlobal WHERE PRGlobal.TypeCode = " & TaxEquate & _
                            " AND PRGLobal.Year = " & TaxYear
                If PRGlobal.GetBySQL(SQLString) Then
                    Do
                        If Mid(PRGlobal.Description, 1, 2) = TaxState Then
                            AddFlag = False
                            Exit Do
                        End If
                        If Not PRGlobal.GetNext Then Exit Do
                    Loop
                End If
                If AddFlag = True Then
                    PRGlobal.TypeCode = TaxEquate
                    PRGlobal.Year = TaxYear
                    PRGlobal.Description = Trim(TaxState) & " Unemp Max"
                    PRGlobal.Save (Equate.RecAdd)
                End If
                PRGlobal.Amount = TaxAmt
                PRGlobal.Save (Equate.RecPut)
            End If
        End If

    Loop

End Sub




Private Sub TaxHistAssign()

Dim p1, p2, GrossWage As Currency
Dim SSWage, MEDWage, FWTWage, SWTWage, CWTWage, SWTGross, CWTGross As Currency
Dim FUNWage, SUNWage As Currency
Dim SSDed, MEDDed, FWTDed, SWTDed, CWTDed As Currency
Dim SWTTax, CWTTax As Currency
Dim LastYear, TaxYear As Long
Dim SSMax, FUNMax, SUNMax As Currency
Dim YTDSSWage, YTDFUNWage, YTDSUNWage As Currency
Dim CWTDedDist, SWTDedDist, SWTDist As Currency

    Progress.lblMsg2 = "Now Calculating Tax History"
    Progress.lblMsg2.Refresh

    ' **** different state UN max ******

    ' assign the CWT wage and SWT wage to PRDist
    ' split up SWT to PRDist proportionally
    ' calculate all taxable wages
        
    SQLString = "SELECT * FROM PREmployee"
    If Not PREmployee.GetBySQL(SQLString) Then Exit Sub    ' ???
    
    Do
    
        SQLString = "SELECT * FROM PRHist WHERE PRHist.EmployeeID = " & PREmployee.EmployeeID & _
                    " ORDER BY PRHIST.YearMonth, PRHist.PEDate"
                    
        LastYear = 0
                    
        If PRHist.GetBySQL(SQLString) Then
    
            Do
            
                ' change in year
                TaxYear = Int(PRHist.YearMonth / 100)
                If LastYear = 0 Or TaxYear <> LastYear Then
                
                    ' clear variables
                    YTDSSWage = 0
                    YTDFUNWage = 0
                    YTDSUNWage = 0
                    
                    ' get tax maximums
                    
                    SSMax = PRGlobal.GetAmount(PREquate.GlobalTypeSSMax, TaxYear)
                    FUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, TaxYear)
                    SUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeSUNMax, TaxYear)
                    
                End If
                LastYear = TaxYear
            
                ' loop thru the earnings
                SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
                If PRDist.GetBySQL(SQLString) Then
                
                    Do
                    
                        If PRDist.DistType = PREquate.DistTypeItem Then
                                
                            ' other earning - get the PRItem record
                            SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & PRDist.ItemID
                            If Not PRItem.GetBySQL(SQLString) Then
                                MsgBox "PRItem NF: " & PRDist.ItemID, vbCritical
                                End
                            End If
                            
                            If PRItem.NoFWTTax = 0 Then FWTWage = FWTWage + PRDist.Amount
                            If PRItem.NoMedTax = 0 Then MEDWage = MEDWage + PRDist.Amount
                            If PRItem.NoSSTax = 0 Then SSWage = SSWage + PRDist.Amount
                            
                            ' ************ write the earnings to PRDist if taxable
                            
                            If PRItem.NoSWTTax = 0 Then
                                SWTGross = SWTGross + PRDist.Amount
                                SWTWage = SWTWage + PRDist.Amount
                                PRDist.StateWage = PRDist.Amount
                            End If
                            
                            If PRItem.NoCWTTax = 0 Then
                                CWTGross = CWTGross + PRDist.Amount
                                CWTWage = CWTWage + PRDist.Amount
                                PRDist.CityWage = PRDist.Amount
                            End If
                            
                            If PRItem.NoSUNTax = 0 Then
                                SUNWage = SUNWage + PRDist.Amount
                                PRDist.SUNWage = PRDist.Amount
                            End If
                            
                            ' ***************************************
                            
                            If PRItem.NoFUNTax = 0 Then FUNWage = FUNWage + PRDist.Amount
                        
                        Else            ' regualar or overtime
                            
                            SSWage = SSWage + PRDist.Amount
                            MEDWage = MEDWage + PRDist.Amount
                            FWTWage = FWTWage + PRDist.Amount
                            SWTWage = SWTWage + PRDist.Amount
                            CWTWage = CWTWage + PRDist.Amount
                            SUNWage = SUNWage + PRDist.Amount
                            FUNWage = FUNWage + PRDist.Amount
                            
                            SWTGross = SWTGross + PRDist.Amount
                            CWTGross = CWTGross + PRDist.Amount
                        
                            ' ***** write the taxable wage
                            PRDist.StateWage = PRDist.Amount
                            PRDist.CityWage = PRDist.Amount
                        
                        End If
                        
                        ' save it
                        PRDist.Save (Equate.RecPut)
                        
                        GrossWage = GrossWage + PRDist.Amount
                        
                        If Not PRDist.GetNext Then Exit Do
                    
                    Loop
                
                End If
        
                ' loop thru the deductions
                SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID
                If PRItemHist.GetBySQL(SQLString) Then
                
                    Do
                    
                        SQLString = "SELECT * FROM PRITem WHERE PRItem.ItemID = " & PRItemHist.ItemID
                        If Not PRItem.GetBySQL(SQLString) Then
                            MsgBox "PRItem NF: " & PRItemHist.ItemID, vbCritical
                            End
                        End If
                        
                        If PRItem.NoSSTax = 1 Then SSWage = SSWage - PRItemHist.Amount
                        If PRItem.NoMedTax = 1 Then MEDWage = MEDWage - PRItemHist.Amount
                        If PRItem.NoFWTTax = 1 Then FWTWage = FWTWage - PRItemHist.Amount
                        If PRItem.NoSWTTax = 1 Then SWTWage = SWTWage - PRItemHist.Amount
                        If PRItem.NoCWTTax = 1 Then CWTWage = CWTWage - PRItemHist.Amount
                        If PRItem.NoFUNTax = 1 Then FUNWage = FUNWage - PRItemHist.Amount
                        If PRItem.NoSUNTax = 1 Then SUNWage = SUNWage - PRItemHist.Amount
                        
                        If PRItem.NoSWTTax = 1 Then SWTDed = SWTDed + PRItemHist.Amount
                        If PRItem.NoCWTTax = 1 Then CWTDed = CWTDed + PRItemHist.Amount
        
                        If Not PRItemHist.GetNext Then Exit Do
                        
                    Loop
                
                End If
        
                PRHist.SSWage = CalcWage(SSWage, YTDSSWage, SSMax)
                YTDSSWage = YTDSSWage + PRHist.SSWage
                
                PRHist.MEDWage = MEDWage
                PRHist.FWTWage = FWTWage
                
                PRHist.FUNWage = CalcWage(FUNWage, YTDFUNWage, FUNMax)
                YTDFUNWage = YTDFUNWage + PRHist.FUNWage
        
                PRHist.Save (Equate.RecPut)
        
                SWTTax = PRHist.SWTTax
        
                SWTDedDist = 0
                SWTDist = 0
                CWTDedDist = 0
        
                ' loop back thru the PRDist to split up the SWT and CWT deductions
                SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
                If PRDist.GetBySQL(SQLString) Then
                    
                    Do
                    
                        If PRDist.StateWage <> 0 Then
                            If SWTGross <> 0 Then
                                p1 = PRDist.Amount / SWTGross * SWTDed
                                p2 = PRDist.Amount / SWTGross * PRHist.SWTTax
                            Else
                                p1 = 0
                                p2 = 0
                            End If
                            PRDist.StateWage = PRDist.StateWage - p1
                            PRDist.StateTax = p2
                            SWTDedDist = SWTDedDist + p1
                            SWTDist = SWTDist + p2
                        End If
                        
                        If PRDist.CityWage <> 0 Then
                            If CWTGross <> 0 Then
                                p1 = PRDist.CityWage / CWTGross * CWTDed
                            Else
                                p1 = 0
                            End If
                            PRDist.CityWage = PRDist.CityWage - p1
                            CWTDedDist = CWTDedDist + p1
                        End If
                    
                        PRDist.Save (Equate.RecPut)
        
                        If Not PRDist.GetNext Then Exit Do
        
                    Loop
        
                End If
                
                ' clean up rounding?
                If SWTDedDist <> SWTDed Then
                    SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
                    If PRDist.GetBySQL(SQLString) Then
                        Do
                            If PRDist.StateWage <> 0 Then
                                PRDist.StateWage = PRDist.StateWage + SWTDed - SWTDedDist
                                PRDist.Save (Equate.RecPut)
                                Exit Do
                            End If
                            If Not PRDist.GetNext Then Exit Do
                        Loop
                    End If
                End If
                
                If SWTDist <> PRHist.SWTTax Then
                    
                    SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
                    If PRDist.GetBySQL(SQLString) Then
                        Do
                            If PRDist.StateWage <> 0 Then
                                PRDist.StateTax = PRDist.StateTax + PRHist.SWTTax - SWTDist
                                PRDist.Save (Equate.RecPut)
                                Exit Do
                            End If
                            If Not PRDist.GetNext Then Exit Do
                        Loop
                    End If
                
                End If
                
                If CWTDedDist <> CWTDed Then
                    SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
                    If PRDist.GetBySQL(SQLString) Then
                        Do
                            If PRDist.StateWage <> 0 Then
                                PRDist.StateWage = PRDist.StateWage + CWTDed - CWTDedDist
                                PRDist.Save (Equate.RecPut)
                                Exit Do
                            End If
                            If Not PRDist.GetNext Then Exit Do
                        Loop
                    End If
                End If
                
                ' clear the variables
                GrossWage = 0
                
                SSWage = 0
                MEDWage = 0
                FWTWage = 0
                SWTWage = 0
                CWTWage = 0
                
                SWTGross = 0
                CWTGross = 0
                
                FUNWage = 0
                SUNWage = 0
                
                SWTDed = 0
                CWTDed = 0
                
                If Not PRHist.GetNext Then Exit Do
    
            Loop

        End If
        
        If Not PREmployee.GetNext Then Exit Do

    Loop

End Sub

Private Function CalcWage(ByVal PayAmount As Currency, _
                          ByVal YTDAmount As Currency, _
                          ByVal MaxAmount As Currency) As Currency
                           
    If YTDAmount >= MaxAmount Then
        CalcWage = 0
    ElseIf YTDAmount + PayAmount <= MaxAmount Then
        CalcWage = PayAmount
    Else
        CalcWage = MaxAmount - YTDAmount
    End If
    
End Function

Private Sub ERIDAssign()
    
    Progress.lblMsg2 = "Now Update History Employer Item IDs"
    Progress.lblMsg2.Refresh

    SQLString = "SELECT * FROM PRDist"
    If PRDist.GetBySQL(SQLString) Then
        Do
            If Not PRItem.GetByID(PRDist.ItemID) Then
                MsgBox "ItemID NF: " & PRDist.ItemID, vbCritical
                End
            End If
            PRDist.EmployerItemID = PRItem.EmployerItemID
            PRDist.Save (Equate.RecPut)
            If Not PRDist.GetNext Then Exit Do
        Loop
    End If
    
    SQLString = "SELECT * FROM PRItemHist"
    If PRItemHist.GetBySQL(SQLString) Then
        Do
            If Not PRItem.GetByID(PRItemHist.ItemID) Then
                MsgBox "ItemID NF: " & PRItemHist.ItemID, vbCritical
                End
            End If
            PRItemHist.EmployerItemID = PRItem.EmployerItemID
            PRItemHist.Save (Equate.RecPut)
            If Not PRItemHist.GetNext Then Exit Do
        Loop
    End If

End Sub

Public Sub GlobalInit()

Dim i As Integer
Dim GlobalString As String
Dim TypeCode As Byte

    ' add if DNE
    For i = 1 To 25
        
        TypeCode = PREquate.GlobalTypeW2Box12
        If i = 1 Then GlobalString = "(A) UNCOLLECTED SS TAX ON TIPS"
        If i = 2 Then GlobalString = "(B) UNCOLLECTED MED TAX ON TIPS"
        If i = 3 Then GlobalString = "(C) COST OF GROUP TERM LIFE INS > $50000"
        If i = 4 Then GlobalString = "(D) SECTION 401K"
        If i = 5 Then GlobalString = "(E) SECTION 403B"
        If i = 6 Then GlobalString = "(F) SECTION 408K 6"
        If i = 7 Then GlobalString = "(G) SECTION 457B"
        If i = 8 Then GlobalString = "(H) SECTION 501 C 18 D"
        If i = 9 Then GlobalString = "(J) NON TAXABLE SICK PAY"
        If i = 10 Then GlobalString = "(K) GOLDEN PARACHUTE PAYMENTS"
        If i = 11 Then GlobalString = "(L) NON TAXABLE REIMBURSEMENTS"
        If i = 12 Then GlobalString = "(M) UNCOLL. SS TAX OF TERM LIFE"
        If i = 13 Then GlobalString = "(N) UNCOLL. MED TAX OF TERM LIFE"
        If i = 14 Then GlobalString = "(P) MOVING EXPENSE REIMB."
        If i = 15 Then GlobalString = "(Q) MILITARY BASIC QUARTERS"
        If i = 16 Then GlobalString = "(R) EMPLOYER CONTR. TO MSA"
        If i = 17 Then GlobalString = "(S) EMPLOYEE SAL REDUCTION 408P"
        If i = 18 Then GlobalString = "(T) ADOPTION BENEFITS"
        If i = 19 Then GlobalString = "(V) EXERCISE OF NON-STAT OPTIONS"
        If i = 20 Then GlobalString = "(W) EMP CONTR TO HEALTH SVGS"
        If i = 21 Then GlobalString = "(Y) DEFERRALS SEC 409A"
        If i = 22 Then GlobalString = "(Z) INCOME SEC 409A"
        If i = 23 Then GlobalString = "(1) AA ROTH 401K CONTRIBUTIONS"
        If i = 24 Then GlobalString = "(2) BB ROTH 403B CONTRIBUTIONS"
        If i = 25 Then
            GlobalString = "SEC. 125"
            TypeCode = PREquate.GlobalTypeW2Box14
        End If
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & TypeCode & _
                    " AND Description = '" & GlobalString & "'"
        If Not PRGlobal.GetBySQL(SQLString) Then
            PRGlobal.Clear
            PRGlobal.TypeCode = TypeCode
            PRGlobal.Description = GlobalString
            PRGlobal.Save (Equate.RecAdd)
        End If
    Next i

End Sub

Private Sub HistOnly()
    
Dim CompanyName, Date1, Date2 As String
Dim Answer As Variant
    
    ' show file being imported
    Input #ASCIIChannel, CompanyName, Date1, Date2
    
    Answer = MsgBox("OK to import Payroll History for: " & vbCr & _
                     CompanyName & vbCr & vbCr & _
                     "From: " & Date1 & " to: " & Date2, vbQuestion + vbOKCancel)
               
    If Answer = vbCancel Then End
    
    CtyID = PRCompany.DfltCityID
    
    Do
      
        Input #ASCIIChannel, dType

        Select Case dType
        
            Case "HIS"
                ImportHistory
            
            Case "END"
                Exit Do
    
        End Select
    
        ImportCount = ImportCount + 1
        If ImportCount Mod 100 = 1 Then
            Progress.lblMsg2 = "Importing record: " & Format(ImportCount, "###,##0") & " Of: " & Format(RCT, "###,##0")
            Progress.lblMsg2.Refresh
        End If
    
    Loop

    ' assign city & state wages to PRDist
    ' assign state tax to PRDist
    TaxHistAssign
    
    ' create PRBatch from PRHist
    CreatePRBatch

    ' assign PRDist.EmployerItemID / PRItemHist.EmployerItemID
    ERIDAssign
    
    ' delete unassigned Dir Dep PRItem records
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDirDepDed
    rsInit SQLString, cn, irs
    If irs.RecordCount > 0 Then
        irs.MoveFirst
        Do
            X = Trim(irs!DirDepBank)
            If X = "" Then irs.Delete
            irs.MoveNext
            If irs.EOF Then Exit Do
        Loop
    End If

    GoBack

End Sub


