Attribute VB_Name = "modPRGlobal"
Public SQLString As String

Public rc4Key As String

Public SessMgr As New QBSessionManager

Public OpenTab As Byte
Public Equate As cEquate
Public PREquate As cPREquate
Public PRBatch As cPRBatch
Public PRCompany As cPRCompany
Public PRDepartment As cPRDepartment
Public PREmployee As cPREmployee
Public PRItem As cPRItem
Public PRHist As cPRHist
Public PRItemHist As cPRItemHist
Public PRDist As cPRDist
Public PRFWTTable As cPRFWTTable
Public PRGlobal As cPRGlobal
Public PRGLUpd As cPRGLUpd

Public GLAccount As cGLAccount
Public GLBatch As cGLBatch
Public GLCompany As cGLCompany
Public GLHistory As cGLHistory
Public GLJournal As cGLJournal
Public GLPrint As cGLPrint

Public PRTotal As cPRTotal

Public JCCustomer As cJCCustomer
Public JCJob As cJCJob
Public PRTimeSheet As cPRTimeSheet
Public QBAccount As cQBAccount
Public QBUpdate As cQBUpdate

Public PRW2 As cPRW2
Public PRW2City As cPRW2City
Public PRW2State As cPRW2State

Public PRCounty As cPRCounty

Public Notes As cNotes

Public User As cGLUser

Public dbName As String

Public PRCity As cPRCity
Public PRState As cPRState

Public EmpID As Long

Public MsgResponse As Integer
Public Response As Boolean
Public CompFlag As Boolean
Public Pg As Integer

Public ErrMsg As String
Public ErrMessage As String
Public Ln As Integer
Public Lines As Integer
Public Columns As Integer
Public MaxLines As Integer
Public FontName As String
Public FontSize As Integer
Public Prvw As frmPreview
Public Progress As frmProgress

Public Condensed As Boolean
Public CharSize As Double
Dim CPI As Integer

Public XUnits As Integer
Public YUnits As Integer

Public PgNum As Long
Public FormatString(50) As String
Public PrintValue(50) As Variant
Public PS As String
Public FValue As Integer

Public ii As Long
Public jj As Long
Public kk As Long
Public xx As String
Public x As Variant

' command line variables
Public CompanyID As Long
Public dbPwd As String
Public ProgName As String
Public SysFile As String
Public UserID As Long
Public PRBatchID As Long
Public BackName As String
Public MenuName As String
Public Period As Long

Public TaskID As Long

Public FormType As Byte        ' 1=add / 2=edit / 3=delete

Public uDB As XArrayDB

Public PrintString As String
Public CCol As Long

Public TextChannel2 As Integer      ' used to bypass auto export in FormatPrint
Public TextChannel As Integer
Public TextFileName As String

' for lists and labels
Public NoLabels As Integer
Public LabelString(4, 4) As String
Public Label2String(5, 5) As String

Public FilterSw As Byte

Public DfltCityID As Long
Public DfltStateID As Long
Public DfltJobID As Long

Public DisConn As Boolean
Public SelID As Long
Public ChangeFlag As Boolean
Public ModeSelect As Boolean
Public RangeType As Byte
Public InitFlag As Boolean
Public PEDate, CheckDate, StartDate, EndDate As Date

Public jbFlag As Boolean

' *** from PRGGlobal ***
Public TxtDate As Date
Public PrtDate As Date
Public PrtTitle As String
Public GrossPay As Currency
Public NetPay As Currency
Public TotalFlag As Boolean
Public TGrossPay As Currency
Public TTotTaxes As Currency
Public TNetPay As Currency
Public TTotHours As Single
Public rrs As ADODB.Recordset

Public Dpts As New ADODB.Recordset
Public BatchNumbr As Long
Public CheckDt As Date
Public ConvNumMon As Integer
Public txtDisplay As String
' *** from PRGGlobal ***

Public Msg1, Msg2, msg3, Msg4 As String
Public CurrDate As Date
Public CurrYear As Long
Public qQuarter As Byte
Public NumEmployees As Long

Public ZipString As String
Public HorzNudge, VertNudge, Nudge As Byte
Public OptDate As String
Public Ct, Recs As Long
Public RecCt As Long
Public CountFormat As String

Public PrvwReturn As Boolean

Public LandSw As Byte

Public QBOK As Boolean
Public QBFileName, QBFedID, QBCompanyName As String

Public Ver941 As String

Public BalintFolder As String

Public SEH_Flag As Boolean

Public Sub SetEquates()
   
    rc4Key = "B@lint19742101!@#$%^&*"
   
    ' format for record count progress displays
    CountFormat = "#,###,##0"
   
    Nudge = 45              ' 1440 / 32 = 45  nudge by 1/32"
    Equate.PgTwips = 11520 ' 8" x 1440 twips/inch = 11520

    Equate.RecAdd = True
    Equate.RecPut = False

    Equate.FormAdd = 1
    Equate.FormEdit = 2
    Equate.FormDel = 3
    
    Equate.Portrait = 1
    Equate.LandScape = 2
   
    PREquate.RecTypeAdjust = 1
    PREquate.RecTypeDepartment = 2
    PREquate.RecTypeEERate = 3
    PREquate.RecTypeEmployee = 4
    PREquate.RecTypeEmployer = 5
    PREquate.RecTypeHistory = 6
    PREquate.RecTypeItem = 7
    PREquate.RecTypeList = 8
    PREquate.RecTypePhone = 9
   
    PREquate.ItemTypeRegPay = 1
    PREquate.ItemTypeOvtPay = 2
    PREquate.ItemTypeOE = 3
    PREquate.ItemTypeDED = 4
    PREquate.ItemTypeSDTax = 5
    PREquate.ItemTypeOtherTax = 6
    PREquate.ItemTypeWorkComp = 7
    PREquate.ItemTypeShiftDiff = 8
    PREquate.ItemType1099Pay = 9
    PREquate.ItemTypeDirDepDed = 10
    PREquate.ItemTypeSick = 11
    PREquate.ItemTypeVacation = 12
    PREquate.ItemTypeNotInNet = 13
    PREquate.ItemTypeEIC = 14
    PREquate.ItemType3PSickPay = 15
    PREquate.ItemTypeRegTax = 16
   
    PREquate.qbItem_RegPay = 1
    PREquate.qbItem_OvtPay = 2
    PREquate.qbItem_OE = 3
    PREquate.qbItem_SSTax = 4
    PREquate.qbItem_MEDTax = 5
    PREquate.qbItem_FWTTax = 6
    PREquate.qbItem_SWTTax = 7
    PREquate.qbItem_CWTTax = 8
    PREquate.qbItem_DED = 9
    PREquate.qbItem_Net = 10
    PREquate.qbItem_ERSS = 11
    PREquate.qbItem_ERMED = 12
    PREquate.qbItem_ERFUN = 13
    PREquate.qbItem_ERSUN = 14
    PREquate.qbItem_ERWkc = 15
    PREquate.qbItem_AccrueSS = 16
    PREquate.qbItem_AccrueMED = 17
    PREquate.qbItem_AccrueFUN = 18
    PREquate.qbItem_AccrueSUN = 19
    PREquate.qbItem_AccrueWkc = 20
   
    PREquate.qbJobStatus_Awarded = 0
    PREquate.qbJobStatus_Closed = 1
    PREquate.qbJobStatus_InProgress = 2
    PREquate.qbJobStatus_None = 3
    PREquate.qbJobStatus_NotAwarded = 4
    PREquate.qbJobStatus_Pending = 5
   
    PREquate.DistTypeReg = 1
    PREquate.DistTypeOT = 2
    PREquate.DistTypeItem = 3
   
    PREquate.BasisAmount = 1
    PREquate.BasisHourly = 2
    PREquate.BasisPercent = 3
    PREquate.BasisNet = 4
    PREquate.BasisExemptions = 5
   
    PREquate.DirDepTypeChecking = 1
    PREquate.DirDepTypeSavings = 2
   
    PREquate.DirDepBasisAmt = 1
    PREquate.DirDepBasisPct = 2
    PREquate.DirDepBasisNet = 3
    
    PREquate.RangeTypeBatch = 1
    PREquate.RangeTypeMonths = 2
    PREquate.RangeTypePEDate = 3
    
    ' numbering out of sequence ... OK
    PREquate.CheckTypeBlankStock = 255
    PREquate.CheckTypePrePrintedA = 2
    PREquate.CheckTypePrePrintedB = 1
    PREquate.CheckTypePrePrintedC = 3
    
    ' ==========================================================
    ' used for GetPRAmount - no duplicates!!!
   
    PREquate.WageGross = 1
    PREquate.WageFed = 2
    PREquate.WageSS = 3
    PREquate.WageMed = 4
    PREquate.WageState = 5
    PREquate.WageCity = 6
    PREquate.WageFUN = 7
    PREquate.WageSUN = 8
    
    PREquate.TaxFed = 9
    PREquate.TaxSS = 10
    PREquate.TaxMed = 11
    PREquate.TaxState = 12
    PREquate.TaxCity = 13
    
    PREquate.TipsSS = 14
    PREquate.TipsMed = 15
    
    ' ==========================================================
    
    PREquate.EEListsPhone = 1
    PREquate.EEListsDptRate = 2
    PREquate.EEListsPRNotes = 3

    PREquate.GlobalTypeRaceCode = 1
    PREquate.GlobalTypeEducationLevel = 2
    PREquate.GlobalTypeContact = 3
    PREquate.GlobalTypeShiftCode = 4
    PREquate.GlobalTypeTerminationCode = 5
    
    PREquate.GlobalTypeSSMax = 6
    PREquate.GlobalTypeSSPct = 7
    PREquate.GlobalTypeMEDPct = 8
    PREquate.GlobalTypeFWTAllow = 9
    PREquate.GLobalTypeOHAllow = 10
    PREquate.GlobalTypeFUNMax = 11
    
    ' *******************
    PREquate.GlobalTypeFUNPct = 34
    
    PREquate.GlobalTypeSUNMax = 12
    
    PREquate.GlobalTypeW2Box12 = 13
    PREquate.GlobalTypeW2Box14 = 14
    PREquate.GlobalTypePaysPerYear = 15
        
    PREquate.GlobalTypeEICMaxAdv = 16
    PREquate.GlobalTypeEICMaxWage = 17
        
    PREquate.GlobalTypeWkcCat = 18
        
    PREquate.GlobalTypeNudge = 19
        
    PREquate.GlobalType941Part4 = 20
    PREquate.GlobalType941Part5 = 21
    PREquate.GlobalType941PaidPrep = 22
        
    PREquate.GlobalTypeBUCForm = 23
        
    PREquate.GlobalTypePRCheckPrefix = 24
        
    PREquate.GlobalTypeOHSDTaxAllow = 25
        
    PREquate.GlobalTypeW3A = 26
    PREquate.GlobalTypeW3B = 27
    PREquate.GlobalTypeW3C = 28
    PREquate.GlobalTypeW3D = 29
    PREquate.GlobalTypeW3E = 30
    
    PREquate.GlobalTypeQBPRChk = 31
    PREquate.GlobalTypeDeductBasis = 32
    
    PREquate.GlobalTypeDirDepFolder = 33
    
    
    ' **** >>>> FUN Pct is 34 <<<< ****
    
    PREquate.GlobalTypeQB_File = 35     ' *** not used ***
    PREquate.GlobalTypeQB_EE_Wage = 36
    PREquate.GlobalTypeQB_EE_Tax = 37
    PREquate.GlobalTypeQB_EE_Ded = 38
    PREquate.GlobalTypeQB_ER_Tax = 39
    
    PREquate.GlobalTypeQB_Register = 40
    
    PREquate.GlobalTypeScreenDefault = 41
    
    PREquate.GlobalTypeQBInv = 42
    
    ' link TimeSheet Week Ended to PRBatch
    PREquate.GlobalTypePRBatchWE = 43
    
    PREquate.GlobalTypePWCraft = 44
    PREquate.GlobalTypePWCounty = 45
    PREquate.GlobalTypePWUnion = 46
    PREquate.GlobalTypePWWage = 47
    
    PREquate.GlobalTypeQBPayFED = 48
    PREquate.GlobalTypeQBPayState = 49
    PREquate.GlobalTypeQBPayCity = 50
    PREquate.GlobalTypeQBPayWKC = 51
    PREquate.GlobalTypeQBPayFUN = 52
    PREquate.GlobalTypeQBPaySUN = 53
    PREquate.GlobalTypeQBPayItem = 54
    PREquate.GlobalTypeQBPayDeductItem = 56
    PREquate.GlobalTypeQBPayNetPay = 57
    PREquate.GlobalTypeQBPayGrossPay = 58
    PREquate.GlobalTypeQBPayCompany = 60

    ' Var1 = Description
    ' Var2 = CompanyID
    PREquate.GlobalTypeCompanyOption = 55
        
    PREquate.GlobalTypeGLFFSched = 59
    PREquate.GlobalTypeGLFFColumn = 61
    PREquate.GlobalTypeGLFFSetup = 62
    
    PREquate.GlobalTypeUnionEmployee = 63
    PREquate.GlobalTypeUnionJob = 64
    
    PREquate.GlobalTypeOtherStateID = 65
    
    PREquate.GlobalTypeMEDAddPct = 66
    PREquate.GlobalTypeMEDAddAmt = 67
    
    PREquate.GlobalTypeOHMultiplier = 68
    
    PREquate.EICTypeSingle = 1
    PREquate.EICTypeMarriedWith = 2
    PREquate.EICTypeMarriedWO = 3
    
    PREquate.x1099Reg = 1
    PREquate.x1099Inc = 2
    
    ' ****************************************
    ' ** GL Update
    
    PREquate.GLTypeEmployee = 1
    PREquate.GLTypeDept = 2
    PREquate.GLTypeCompany = 3

    PREquate.GLItemTypeOE = 1
    PREquate.GLItemTypeDed = 2
    PREquate.GLItemTypeSSTax = 3
    PREquate.GLItemTypeSSMatch = 4
    PREquate.GLItemTypeMedTax = 5
    PREquate.GLItemTypeMedMatch = 6
    PREquate.GLItemTypeFWTTax = 7
    PREquate.GLItemTypeSWTTax = 8
    PREquate.GLItemTypeCWTTax = 9
    PREquate.GLItemTypeFUN = 10
    PREquate.GLItemTypeSUN = 11
    PREquate.GLItemTypeGross = 12
    PREquate.GLItemTypeNet = 13
    PREquate.GLItemTypeSSExp = 14
    PREquate.GLItemTypeMEDExp = 15
    PREquate.GLItemTypeFUNExp = 16
    PREquate.GLItemTypeSUNExp = 17
    PREquate.GLItemTypeWkcExp = 18
    PREquate.GLItemTypeRegPay = 19
    PREquate.GLItemTypeOvtPay = 20
    PREquate.GLItemTypeWkcTax = 21
    
    ' ** GL Update
    ' ****************************************
    
    ' default sort order for the company
    PREquate.SortOrderNumber = 0
    PREquate.SortOrderName = 1
    PREquate.SortOrderDeptNumber = 2
    PREquate.SortOrderDeptName = 3
    
    PREquate.ActionAdd = 1
    PREquate.ActionEdit = 2
    PREquate.ActionDelete = 3
    
    PREquate.QBError_NoFileOpen = -2147220458
    PREquate.QBError_QBBeginSession = -2147220472
    
    PREquate.PeriodTypeMonth = 1
    PREquate.PeriodTypeQuarter = 2
    PREquate.PeriodTypeYear = 3
    PREquate.PeriodTypePay = 4
    
    PREquate.PayTypeCheck = 1
    PREquate.PayTypeEFT = 2
    
    Equate.NoteTypeEE = 1
    Equate.NoteTypeER = 2
    
    PrvwReturn = False
    
End Sub

Public Function GetCmd(ByVal CmdLine As String, ByVal Argument As String, ByVal StrNum As String) As Variant

' return xxxx - Argument=xxxx

Dim i As Long
Dim cmd As String
Dim C1 As String
    
    StrNum = LCase(StrNum)
    If StrNum <> "str" And StrNum <> "num" Then
        MsgBox "StrNum not assigned !"
        GetCmd = ""
        Exit Function
    End If
    
    If StrNum = "str" Then
        GetCmd = ""
    Else
        GetCmd = 0
    End If

    ' bad value traps
    If IsNull(CmdLine) Then Exit Function
    If IsNull(Argument) Then Exit Function
    If CmdLine = "" Then Exit Function
    If Argument = "" Then Exit Function

    ' ignore case for argument type but keep it for the return string
    cmd = LCase(CmdLine)
    Argument = LCase(Argument)

    ' search for Argument=xxxxx
    i = InStr(1, cmd, Argument, vbTextCompare)
    If i = 0 Then Exit Function
    
    ' now look for the "=" sign
    If Mid(CmdLine, i + Len(Argument), 1) <> "=" Then Exit Function
    
    ' append to make return string until a space or end of line
    i = i + Len(Argument) + 1
    Do
        If i > Len(CmdLine) Then Exit Do
        C1 = Mid(CmdLine, i, 1)
        If C1 = " " Then Exit Do
        GetCmd = GetCmd & C1
        i = i + 1
    Loop

End Function

Public Function Centered(ByVal Strg As String, _
                         ByVal FldWidth As Integer) As String
                         
    If (FldWidth - Len(Strg)) < 0 Then
       Centered = Strg
    Else
       Centered = Space((FldWidth - Len(Strg)) / 2) & Strg
    End If
    
End Function


Public Sub prt(ByVal Line As Byte, ByVal Col As Byte, ByVal Str As String)

Dim pi As Integer

       Prvw.vsp.CurrentX = (XUnits * Col) + 200 + (Nudge * HorzNudge)
       Prvw.vsp.CurrentY = (YUnits * Line) + (Nudge * VertNudge)
       
       Prvw.vsp.text = Str
       
'       Ln = Ln + 1

'       Printer.CurrentX = (Col * 120) - 240 + (hadj * 45)
'       Printer.CurrentY = ((Line - 1) * 240) - 960 + (vadj * 45)
    
End Sub

' col as actual twip value
Public Sub PrtCenter(ByVal Line As Byte, ByVal Str As String)

Dim pi As Integer

'       Prvw.vsp.CurrentX = ((Equate.PgTwips - Prvw.vsp.TextWidth(Str)) / 2) + 200
'       Prvw.vsp.CurrentY = 240 * Line
'       Prvw.vsp.Text = Str
'
''       Ln = Ln + 1
'
'       ' clear the print values
'       For pi = 1 To 40
'           PrintValue(pi) = ""
'       Next pi
'
''       Printer.CurrentX = (Col * 120) - 240 + (hadj * 45)
''       Printer.CurrentY = ((Line - 1) * 240) - 960 + (vadj * 45)
            
    i = ((Columns - Len(Str)) / 2)
    PrintValue(1) = " ":            FormatString(1) = "a" & i
    PrintValue(2) = Str:     FormatString(2) = "a" & Len(Str)
    PrintValue(3) = " ":            FormatString(3) = "~"
    FormatPrint
    
End Sub
Public Sub FormFeed()

   Prvw.vsp.NewPage
   Ln = 0

End Sub

Public Sub LineFeed(ByVal Lines As Integer)
    
Dim lx As Integer
    
    PrintValue(1) = " "
    FormatString(1) = "a1"
    PrintValue(2) = " "
    FormatString(2) = "~"

    For lx = 1 To Lines
        FormatPrint
    Next lx

End Sub


Public Sub PrtInit(ByVal PortLand As String)
   
   Set Prvw = New frmPreview
   
   Prvw.vsp.Preview = True
   
   Prvw.vsp.PhysicalPage = False
'   Prvw.vsp.PhysicalPage = True
   
   Prvw.vsp.MarginRight = 0
   Prvw.vsp.MarginLeft = 0
   Prvw.vsp.MarginBottom = 0
   Prvw.vsp.MarginTop = 0
   
   If PortLand = "Port" Then
      Prvw.vsp.Orientation = orPortrait
   Else
      Prvw.vsp.Orientation = orLandscape
   End If
   
' -------- from glstmt
'   If GLPrint.PrintBIB = Equate.PrtISOnly And GLPrint.WidePrint = False Then
'      Prvw.vsp.Orientation = orLandscape
'   End If
' --------
   
   Prvw.vsp.Font.Name = "Courier New"
   
'   ' bold for Richlak - make an option ....
'   If GLPrint.Output = "Bold" Then
'      Prvw.vsp.Font.Bold = True
'   Else
'      Prvw.vsp.Font.Bold = False
'   End If
    
   Prvw.Height = Screen.Height
   Prvw.Width = Prvw.vsp.Width + 1000
   Prvw.Top = 0
   Prvw.Left = 0

   Prvw.vsp.Height = Screen.Height - 500
   Prvw.vsp.Left = 500
   Prvw.vsp.Top = 0
 
'   Prvw.vsp.ExportFormat = vpxPagedHTML
'   Prvw.vsp.ExportFile = "c:\asend\htmltest.html"
 
   Prvw.vsp.StartDoc

   YUnits = 240 - 15       ' = 1440 / 6

'   YUnits = 240 - 25

End Sub
Public Sub SetFont(ByVal FSize As Integer, ByVal PortLand As Byte)
   
   Prvw.vsp.FontSize = FSize
   
   Select Case FSize
      Case 6
         XUnits = 72   ' 1440 / 20
         CPI = 20
      Case 7
         XUnits = 80   ' 1440 / 18
         CPI = 18
      Case 8
         XUnits = 96   ' 1440 / 15
         CPI = 15
      Case 9
         XUnits = 110  ' 1440 / 13
         CPI = 13
      Case 10
         XUnits = 120  ' 1440 / 12
         CPI = 12
      Case 11
         XUnits = 131  ' 1440 / 11
         CPI = 11
      Case 12
         XUnits = 144  ' 1440 / 10
         CPI = 10
      Case 13
         XUnits = 180   ' 1440 / 8   for Report of wages supplemental
         CPI = 8
      Case Else
         MsgBox "Incorrect font size selected: " & FSize
         End
   End Select
   
   ' store the width of one character in twips
   ' using a monospaced font - will be the same size no matter what the letter
   CharSize = Prvw.vsp.TextWidth("W")
   
   If PortLand = Equate.Portrait Then
      Prvw.vsp.Orientation = orPortrait
      MaxLines = 63
      Columns = 8 * CPI
   Else
      Prvw.vsp.Orientation = orLandscape
      MaxLines = 49
      Columns = 11 * CPI
   End If
   
End Sub

Public Sub FormatPrint()

Dim pi As Integer
Dim Lngth As Integer
Dim PCol As Integer
Dim RightString As String
Dim l As Integer

Dim DWidth As Integer
Dim DExp As Integer
Dim CCount As Integer

Dim Quote As String
Dim Comma As String
Dim TextString As String

Dim TxtX As String
Dim StrFormat As String
Dim CommaFlag As Boolean

    ' init variables for csv output
    Quote = """"
    Comma = ","

   pi = 1
   PS = ""
   PCol = 1
   
   If IsNull(CCol) Then
      PrintString = ""
   ElseIf CCol = 0 Then
      PrintString = ""
   Else
      PrintString = Space(CCol)
   End If

   TextString = ""

   Do
                    
      Lngth = Len(FormatString(pi))
      
      FValue = CInt(Mid(FormatString(pi), 2, Lngth - 1))
      
      Select Case Mid(FormatString(pi), 1, 1)
         
         Case "x"
         
            PrintString = PrintString & Space(FValue)
         
         Case "t"
            
            Lngth = Len(PrintString)
            If FValue <= Lngth Then
               PrintString = Mid(PrintString, 1, FValue - 1)
            Else
               PrintString = PrintString & Space(FValue - Lngth - 1)
            End If
            
         Case "a"
            
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = ""
  
            l = Len(PrintValue(pi))
            If l < FValue Then
               PrintValue(pi) = PrintValue(pi) & Space(FValue - l)
            End If
            
            PrintString = PrintString & Mid(PrintValue(pi), 1, FValue)
            ' StrFormat = String(FValue, "@")
            ' PrintString = PrintString & Format(PrintValue(pi), StrFormat)
            TextString = TextString & Quote & Mid(PrintValue(pi), 1, FValue) & Quote & Comma
            
         Case "r"       ' right justified
         
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = ""
  
            If FValue > Len(PrintValue(pi)) Then
               PrintString = PrintString & Space(FValue - Len(PrintValue(pi))) & PrintValue(pi)
               TextString = TextString & Quote & Space(FValue - Len(PrintValue(pi))) & PrintValue(pi) & Quote & Comma
            Else
               PrintString = PrintString & Mid(PrintValue(pi), 1, FValue)
               TextString = TextString & Quote & Mid(PrintValue(pi), 1, FValue) & Quote & Comma
            End If
         
         Case "d"
            
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = 0
  
            ' limits / unassigned
            If FValue <= 0 Or FValue >= 14 Then FValue = 14
            If FValue <= 5 Then FValue = 5
            
            DWidth = 5         ' 0.00-
            DExp = 1
            CCount = 0
            
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
               If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
            Loop
            DWidth = DWidth + CCount       ' add in the comma count
               
            If DWidth > FValue Then         ' take out the commas
                CommaFlag = False
                DWidth = FValue
            Else
                CommaFlag = True
            End If
            
            If CommaFlag Then       ' put in commas if there is enough room
                Select Case FValue
                    Case 5
                        x = Format(Abs(PrintValue(pi)), "0.00")
                    Case 6
                        x = Format(Abs(PrintValue(pi)), "#0.00")
                    Case 7
                        x = Format(Abs(PrintValue(pi)), "##0.00")
                    Case 8
                        x = Format(Abs(PrintValue(pi)), "###0.00")
                    Case 9
                        x = Format(Abs(PrintValue(pi)), "#,##0.00")
                    Case 10
                        x = Format(Abs(PrintValue(pi)), "##,###0.00")
                    Case 11
                        x = Format(Abs(PrintValue(pi)), "###,##0.00")
                    Case 12
                        x = Format(Abs(PrintValue(pi)), "####,##0.00")
                    Case 13
                        x = Format(Abs(PrintValue(pi)), "#,###,##0.00")
                    Case 14
                        x = Format(Abs(PrintValue(pi)), "##,###,##0.00")
                    Case Else
                        MsgBox "Invalid Format Value: " & FValue, vbCritical
                        End
                End Select
            Else
                y = String(FValue - 5, "#") & "0.00"
                x = Format(Abs(PrintValue(pi)), y)
            End If
                
            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0.00")    ' leading minus sign for text output
            Else
               x = x & Space(1)
               TxtX = x
            End If
               
            If DWidth <= 14 Then
                PrintString = PrintString & Space(FValue - DWidth) & x
                TextString = TextString & Quote & Space(FValue - DWidth) & TxtX & Quote & Comma
            Else
                PrintString = PrintString & x
                TextString = TextString & Quote & TxtX & Quote & Comma
            End If
         
'         Case "d"
'
'            DWidth = 5         ' 0.00-
'            DExp = 1
'            CCount = 0
'            Do
'               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
'               DExp = DExp + 1
'               DWidth = DWidth + 1
'               If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
'            Loop
'            DWidth = DWidth + CCount       ' add in the comma count
'
'            x = Format(Abs(PrintValue(pi)), "##,###,##0.00")
'            If PrintValue(pi) < 0 Then
'               x = x & "-"
'               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0.00")    ' leading minus sign for text output
'            Else
'               x = x & Space(1)
'               TxtX = x
'            End If
'
'            If DWidth <= 14 Then
'                PrintString = PrintString & Space(14 - DWidth) & x
'                TextString = TextString & Quote & Space(14 - DWidth) & TxtX & Quote & Comma
'            Else
'                PrintString = PrintString & x
'                TextString = TextString & Quote & TxtX & Quote & Comma
'            End If
         
         Case "p"
               
            DWidth = 4  ' 0.0-
            DExp = 1
            
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = 0
  
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
            Loop
            
            x = Format(Abs(PrintValue(pi)), "##0.0")

            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
            Else
               x = x & Space(1)
               TxtX = x
            End If
            
            If DWidth <= 6 Then
               PrintString = PrintString & Space(6 - DWidth) & x
               TextString = TextString & Quote & Space(6 - DWidth) & TxtX & Quote & Comma
            Else
               PrintString = PrintString & x
               TextString = TextString & Quote & TxtX & Quote & Comma
            End If
            
            PCol = PCol + FValue
            
         Case "q"     ' same as p except blank if zero
               
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = 0
  
            If PrintValue(pi) <> 0 Then
            
                x = Format(Abs(PrintValue(pi)), "##0.0")

                If PrintValue(pi) < 0 Then
                   x = x & "-"
                   TxtX = "-" & Format(Abs(PrintValue(pi)), "##0.0")
                Else
                   x = x & Space(1)
                   TxtX = x
                End If
            
                x = Format(Abs(PrintValue(pi)), "##0.0")
            
                DWidth = 4  ' 0.0-
                DExp = 1
                Do
                  If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
                  DExp = DExp + 1
                  DWidth = DWidth + 1
                Loop
            
                PrintString = PrintString & Space(6 - DWidth) & x
                
                TextString = TextString & Quote & Space(6 - DWidth) & TxtX & Quote & Comma
            
            Else
            
                PrintString = PrintString & Space(FValue)
                TextString = TextString & Quote & Quote & Comma
            
            End If
         
         Case "i"
               
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = 0
  
            DWidth = 2         ' 0-
            DExp = 1
            CCount = 0
            Do
               If Abs(PrintValue(pi)) < 10 ^ DExp Then Exit Do
               DExp = DExp + 1
               DWidth = DWidth + 1
               If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
            Loop
            DWidth = DWidth + CCount       ' add in the comma count
               
            x = Format(Abs(PrintValue(pi)), "##,###,##0")

            If PrintValue(pi) < 0 Then
               x = x & "-"
               TxtX = "-" & Format(Abs(PrintValue(pi)), "##,###,##0")
            Else
               x = x & Space(1)
               TxtX = x
            End If
            
            PrintString = PrintString & Space(14 - DWidth) & x
            
            TextString = TextString & Comma & Space(14 - DWidth) & TxtX & Quote & Comma
            
         Case "n"
               
            ' 2017-01-14 !!!
            If IsNull(PrintValue(pi)) Then PrintValue(pi) = 0
  
            x = Format(Abs(PrintValue(pi)), "########0")
                        
            PrintString = PrintString & Space(FValue - Len(x)) & x
            TextString = TextString & Quote & Space(FValue - Len(x)) & x & Quote & Comma
            
         Case Else
            
            MsgBox "Bad Format: " & FormatString(pi)
      
      End Select
      
      pi = pi + 1
   
   Loop Until FormatString(pi) = "~"
       
   ' print the string
   Prvw.vsp.CurrentY = Ln * YUnits + (Nudge * VertNudge)
   Prvw.vsp.CurrentX = XUnits * 2 + (Nudge * HorzNudge)
   
   Prvw.vsp.CurrentY = (Ln * YUnits) + (Nudge * VertNudge)
   Prvw.vsp.CurrentX = (XUnits * 2) + (Nudge * HorzNudge)
   
   Prvw.vsp.text = PrintString
       
   ' clear the print values
   For pi = 1 To 40
       PrintValue(pi) = ""
   Next pi

   ' output to text if necessary
   If TextChannel <> 0 Then
      Print #TextChannel, TextString
   End If

End Sub
Public Sub PosPrint(ByVal CurrX As Long, ByVal CurrY As Long, ByVal PrintString As String)
    
    Prvw.vsp.CurrentX = CurrX + HorzNudge * Nudge
    Prvw.vsp.CurrentY = CurrY + VertNudge * Nudge
    
    Prvw.vsp.text = PrintString
End Sub
Public Function PadRight(ByVal InString As String, ByVal Length As Long) As String
    If InString = "" Then
        PadRight = ""
    ElseIf Len(InString) > Length Then      ' too long - cut it off
        PadRight = Mid(InString, 1, Length)
    Else
        PadRight = Space(Length - Len(InString)) & InString
    End If
End Function

Public Sub GoBack()
   
    ' return call if given
    If BackName <> "" Then
        
        x = BackName & " UserID=" & UserID & _
            " dbPwd=" & dbPwd & _
            " OpenTab=" & OpenTab & _
            " MenuName=" & MenuName
            
        If BalintFolder <> "" Then
            x = x & " BalintFolder=" & BalintFolder
        End If
        TaskID = Shell(x, vbMaximizedFocus)
    End If

    End

End Sub

Public Function GetPRAmount(ByVal EmployeeID As Long, _
                            ByVal FieldID As Long, _
                            ByVal StartYear As Long, _
                            ByVal EndYear As Long, _
                            ByVal StartMonth As Long, _
                            ByVal EndMonth As Long, _
                            ByVal StartPED As Date, _
                            ByVal EndPED As Date) As Currency
                            
                            
Dim NW As Variant

    NW = Now
    
    GetPRAmount = Hour(NW) * 100 + Minute(NW) + Second(NW) / 100
                            
End Function

Public Sub SetGrid(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid, Optional KeepGrid As Byte)
    
    If KeepGrid = 0 Then
        gfg.Clear
    End If
    
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '

    gfg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    gfg.TabBehavior = flexTabCells                       ' tab moves between cells
    gfg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    ' gfg.HighLight = flexHighlightNever                   ' don't select ranges

End Sub
Public Sub SetGridFree(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    gfg.FixedCols = 0                   ' see all cols selected by SQL
    gfg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    gfg.DataMode = flexDMFree         ' Recordset cursor is maintained by grid
    gfg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set gfg.DataSource = grs.DataSource '
    gfg.DataMember = grs.DataMember     '

    gfg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    gfg.TabBehavior = flexTabCells                       ' tab moves between cells
    gfg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    ' gfg.HighLight = flexHighlightNever                   ' don't select ranges

End Sub

Public Sub AddAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
    
    grs.AddNew          ' Add to the recordset
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid

End Sub
Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid)
' Public Sub DelAdo(ByRef grs As ADODB.Recordset, ByRef gfg As VSFlexGrid, ByVal Number As Long)
    
    grs.Delete
    grs.Update          ' Record (save to file)
    grs.MoveLast        ' Move to the last record in the record set
    gfg.DataRefresh     ' Update the grid data
    gfg.Col = 0         ' Go to the first column
    gfg.SetFocus        ' Move from add button to grid
End Sub

Public Sub SetAdo(ByRef gcn As ADODB.Connection, ByRef grs As ADODB.Recordset, ByVal SQL As String)
    ' Common behavior for Recordsets
    Set grs = New ADODB.Recordset       ' set the recordset
    grs.LockType = adLockOptimistic     '
    grs.CursorType = adOpenDynamic      '
    grs.Source = SQL                    '
    Set grs.ActiveConnection = gcn      ' connection set previous to call
    grs.Open                            ' start the record
End Sub


Public Sub tdbTextSet(ByRef tdbTXT As TDBText, Optional tdbLen As Integer)

    tdbTXT.Key.Clear = ""       ' no key to clear field
    tdbTXT.FormatMode = dbiIncludeFormat
    tdbTXT.Format = "A9#@"
    tdbTXT.text = ""
    tdbTXT.Key.Clear = "{F2}"
    If tdbLen <> 0 Then tdbTXT.MaxLength = tdbLen
    
End Sub

Public Sub tdbAmountSet(ByRef tdbAmt As TDBNumber)

    tdbAmt.Format = "##,###,##0.00;(##,###,##0.00)"
    tdbAmt.DisplayFormat = "##,###,##0.00;(##,###,##0.00);0.00"
    tdbAmt.HighlightText = True
    tdbAmt.Key.Clear = ""
    tdbAmt.MinValue = -99999999.99
    tdbAmt.MaxValue = 99999999.99
    tdbAmt.Value = 0

End Sub

Public Sub tdbIntegerSet(ByRef tdbAmt As TDBNumber)

    tdbAmt.Format = "##,###,##0;(##,###,##0)"
    tdbAmt.DisplayFormat = "##,###,##0;(##,###,##0);0"
    tdbAmt.HighlightText = True
    tdbAmt.Key.Clear = ""
    tdbAmt.MinValue = -99999999
    tdbAmt.MaxValue = 99999999
    tdbAmt.Value = 0

End Sub

Public Sub tdbDateSet(ByRef tdbDate As tdbDate, ByVal DateValue As Date)

    tdbDate.Format = "mm/dd/yyyy"
    tdbDate.DisplayFormat = "mm/dd/yyyy"
    tdbDate.ErrorBeep = True
    tdbDate.DropDown.Enabled = True
    tdbDate.DropDown.Visible = dbiShowAlways
    tdbDate.DropDown.Position = dbiDropPosInside

    tdbDate.text = ""
    If IsNull(DateValue) Then
        tdbDate.text = ""
    ElseIf DateValue = 0 Then
        tdbDate.text = ""
    Else
        On Error Resume Next
        tdbDate.Value = DateValue
        On Error GoTo 0
    End If

End Sub

Public Function ItemName(ByVal ItemType As Byte) As String

    ItemName = "? Item Type NF: " & ItemType

    If ItemType = PREquate.ItemTypeRegPay Then ItemName = "Regular Pay"
    If ItemType = PREquate.ItemTypeOvtPay Then ItemName = "Overtime Pay"
    If ItemType = PREquate.ItemTypeOE Then ItemName = "Other Earning"
    If ItemType = PREquate.ItemTypeDED Then ItemName = "Deduction"
    If ItemType = PREquate.ItemTypeSDTax Then ItemName = "SD Tax"
    If ItemType = PREquate.ItemTypeOtherTax Then ItemName = "Other Tax"
    If ItemType = PREquate.ItemTypeWorkComp Then ItemName = "Wrkmn Comp"
    If ItemType = PREquate.ItemTypeShiftDiff Then ItemName = "Shift Diff"
    If ItemType = PREquate.ItemType1099Pay Then ItemName = "1099 Pay"
    If ItemType = PREquate.ItemTypeDirDepDed Then ItemName = "Direct Deposit"
    If ItemType = PREquate.ItemTypeSick Then ItemName = "Sick Pay"
    If ItemType = PREquate.ItemTypeVacation Then ItemName = "Vacation Pay"
    If ItemType = PREquate.ItemTypeNotInNet Then ItemName = "Not In Net"
    If ItemType = PREquate.ItemTypeRegTax Then ItemName = "Regular Tax"
    If ItemType = PREquate.ItemTypeEIC Then ItemName = "EIC Adv"
    If ItemType = PREquate.ItemType3PSickPay Then ItemName = "3rd Party Sick Pay"
    
End Function

Public Sub rsSave(rsDIS As ADODB.Recordset, cnn As ADODB.Connection)

    ' update a disconnected record set
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = cnn
    
    ' reopen the disconnected record set
    rs.Open rsDIS, cnn
    
    rs.UpdateBatch

End Sub

'Public Function cmbCheck(ByRef cmb As TDBCombo, ByRef xdb As XArrayDB) As Boolean
'
'Dim cmbString As String
'
'    ' make sure a tdb combo entry is in the list
'    cmbString = cmb
'
'    If cmbString = "" Then
'        cmbCheck = False
'    Else
'        If xdb.Find(0, 1, cmbString, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING) < 0 Then
'            cmbCheck = False
'        Else
'            cmbCheck = True
'        End If
'    End If
'
'End Function

Public Function CurrFormat(ByVal Amount As Currency) As String
            
Dim pi As Integer
Dim Lngth As Integer
Dim PCol As Integer
Dim RightString As String
Dim l As Integer

Dim DWidth As Integer
Dim DExp As Integer
Dim CCount As Integer

Dim Quote As String
Dim Comma As String
Dim TextString As String

Dim TxtX As String
                
    DWidth = 5         ' 0.00-
    DExp = 1
    CCount = 0
    Do
                
        If Abs(Amount) < 10 ^ DExp Then Exit Do
        DExp = DExp + 1
        DWidth = DWidth + 1
        If DExp Mod 3 = 1 And DExp <> 1 Then CCount = CCount + 1 ' comma count
    Loop
    DWidth = DWidth + CCount       ' add in the comma count
               
    x = Format(Abs(Amount), "##,###,##0.00")
    If Round(Amount, 2) < 0 Then
        x = x & "-"
        TxtX = "-" & Format(Abs(Amount), "##,###,##0.00")    ' leading minus sign for text output
    Else
        x = x & Space(1)
        TxtX = x
    End If
           
    If DWidth <= 14 Then
        CurrFormat = Space(14 - DWidth) & x
    Else
        CurrFormat = x
    End If

End Function

Public Function AmountInWords(ByVal Amount As Currency, ByVal Cents As Boolean) As String

    ' Cents = true - "99 CENTS"
    ' Cents = false - "99/100" - used for pre-printed checks

Dim Amt1 As Long         ' millions
Dim Amt2 As Long         ' hundreds of thousands
Dim Amt3 As Long         ' singles
Dim Amt4 As Long         ' cents
Dim MaxLen, i As Long
Dim x As String

    MaxLen = 77
 
    ' must be less than $100 million
    If Amount > 99999999.99 Then
        AmountInWords = ""
        Exit Function
    End If

    AmountInWords = ""
    
    i = Int(Amount)
    
    Amt1 = Int(i / 10 ^ 6)                  ' millions
    Amt2 = Int(i / 10 ^ 3) Mod 1000         ' hundreds of thousands
    Amt3 = i Mod 1000                       ' singles
    ' Amt4 = Amount * 100 Mod 100             ' cents
    Amt4 = (Amount - Int(Amount)) * 100
    
    ' millions
    If Amt1 <> 0 Then
        AmountInWords = AIW(Amt1, "")
        AmountInWords = Trim(AmountInWords) & " MILLION"
    End If
    
    ' thousands
    If Amt2 <> 0 Then
        If Amt2 > 99 Then
            AmountInWords = AIW(Int(Amt2 / 100), Trim(AmountInWords)) & " HUNDRED"
        End If
        AmountInWords = " " & AIW(Amt2 Mod 100, Trim(AmountInWords))
        AmountInWords = Trim(AmountInWords) & " THOUSAND"
    End If
    
    ' hundreds
    If Amt3 > 99 Then
        AmountInWords = " " & AIW(Int(Amt3 / 100), AmountInWords)
        AmountInWords = Trim(AmountInWords) & " HUNDRED"
    End If
    
    ' singles
    If Amt3 Mod 100 <> 0 Then
        AmountInWords = " " & AIW(Amt3 Mod 100, AmountInWords)
    End If
    
    ' "dollars"
    If Cents = True Then
        If Amount < 1 Then
            AmountInWords = Trim(AmountInWords) & " NO DOLLARS AND"
        Else
            If Amt3 = 1 And Amount < 2 Then
                AmountInWords = Trim(AmountInWords) & " DOLLAR AND"
            Else
                AmountInWords = Trim(AmountInWords) & " DOLLARS AND"
            End If
        End If
    Else
        If Amount < 1 Then
            AmountInWords = Trim(AmountInWords) & " NO DOLLARS AND"
        Else
            AmountInWords = Trim(AmountInWords) & " AND"
        End If
    End If
    
    ' cents
    If Cents = True Then
        If Amt4 = 0 Then
            AmountInWords = Trim(AmountInWords) & " NO CENTS"
        Else
            AmountInWords = Trim(AmountInWords) & " " & Format(Amt4, "00")
            If Amt4 = 1 Then
                AmountInWords = Trim(AmountInWords) & " CENT"
            Else
                AmountInWords = Trim(AmountInWords) & " CENTS"
            End If
        End If
    Else
        AmountInWords = AmountInWords & " " & Format(Amt4, "00") & "/100"
    End If
    
    ' string too long ???
    If Len(AmountInWords) > MaxLen Then
        
        AmountInWords = ""
        
        If Amt1 <> 0 Then AmountInWords = Amt1 & " MILLION"
        If Amt2 <> 0 Then AmountInWords = Trim(AmountInWords) & " " & Amt2 & " THOUSAND "
        If Amt3 <> 0 Then AmountInWords = Trim(AmountInWords) & " " & Amt3 & " "
        
        If Amount < 1 Then
            AmountInWords = "NO DOLLARS AND"
        Else
            If Amt3 = 1 And Amount < 2 Then
                AmountInWords = Trim(AmountInWords) & " DOLLAR AND"
            Else
                AmountInWords = Trim(AmountInWords) & " DOLLARS AND"
            End If
        End If
        
        If Amt4 = 0 Then
            AmountInWords = Trim(AmountInWords) & " NO CENTS"
        Else
            AmountInWords = Trim(AmountInWords) & " " & Format(Amt4, "00")
            If Amt4 = 1 Then
                AmountInWords = Trim(AmountInWords) & " CENT"
            Else
                AmountInWords = Trim(AmountInWords) & " CENTS"
            End If
        End If
    
    End If
 
    ' right justify it
    AmountInWords = Space(77 - Len(Trim(AmountInWords))) & AmountInWords

End Function

Private Function AIW(ByVal Amount As Integer, ByVal InString As String) As String

    AIW = Trim(InString)
    
    Do

        If Amount = 1 Then AIW = Trim(AIW) & " ONE"
        If Amount = 2 Then AIW = Trim(AIW) & " TWO"
        If Amount = 3 Then AIW = Trim(AIW) & " THREE"
        If Amount = 4 Then AIW = Trim(AIW) & " FOUR"
        If Amount = 5 Then AIW = Trim(AIW) & " FIVE"
        If Amount = 6 Then AIW = Trim(AIW) & " SIX"
        If Amount = 7 Then AIW = Trim(AIW) & " SEVEN"
        If Amount = 8 Then AIW = Trim(AIW) & " EIGHT"
        If Amount = 9 Then AIW = Trim(AIW) & " NINE"
        If Amount = 10 Then AIW = Trim(AIW) & " TEN"
        If Amount = 11 Then AIW = Trim(AIW) & " ELEVEN"
        If Amount = 12 Then AIW = Trim(AIW) & " TWELVE"
        If Amount = 13 Then AIW = Trim(AIW) & " THIRTEEN"
        If Amount = 14 Then AIW = Trim(AIW) & " FOURTEEN"
        If Amount = 15 Then AIW = Trim(AIW) & " FIFTEEN"
        If Amount = 16 Then AIW = Trim(AIW) & " SIXTEEN"
        If Amount = 17 Then AIW = Trim(AIW) & " SEVENTEEN"
        If Amount = 18 Then AIW = Trim(AIW) & " EIGHTEEN"
        If Amount = 19 Then AIW = Trim(AIW) & " NINETEEN"
        If Amount >= 20 And Amount <= 29 Then AIW = Trim(AIW) & " TWENTY"
        If Amount >= 30 And Amount <= 39 Then AIW = Trim(AIW) & " THIRTY"
        If Amount >= 40 And Amount <= 49 Then AIW = Trim(AIW) & " FORTY"
        If Amount >= 50 And Amount <= 59 Then AIW = Trim(AIW) & " FIFTY"
        If Amount >= 60 And Amount <= 69 Then AIW = Trim(AIW) & " SIXTY"
        If Amount >= 70 And Amount <= 79 Then AIW = Trim(AIW) & " SEVENTY"
        If Amount >= 80 And Amount <= 89 Then AIW = Trim(AIW) & " EIGHTY"
        If Amount >= 90 And Amount <= 99 Then AIW = Trim(AIW) & " NINETY"

        If Amount >= 0 And Amount <= 19 Then Exit Do
        If Amount >= 20 And Amount <= 99 Then Amount = Amount Mod 10
        If Not (Amount >= 1 And Amount <= 9) Then Exit Do
    
    Loop

End Function

Public Function CheckAmount(ByVal Amount As Currency) As String

    If Amount > 99999999.99 Then
        CheckAmount = Format(Amount, "###########0.00")
        Exit Function
    End If

    x = Format(Amount, "$##,###,##0.00")
    CheckAmount = String(15 - Len(Trim(x)), "*") & Trim(x)

End Function

Public Function StripOhio(ByVal x As String) As String

    ' take "OHIO" out of a string

Dim pos As Long

    If IsNull(x) Then
        StripOhio = ""
        Exit Function
    End If

    pos = InStr(1, x, "OHIO", vbTextCompare)
    
    If pos = 0 Then
        StripOhio = x
    Else
        StripOhio = Mid(x, 1, pos - 1) & Mid(x, pos + 4, 99)
    End If

End Function

Public Sub cmbPPYSet(ByRef cmb As ComboBox, ByVal DfltValue As Byte)

    With cmb
        
        .AddItem "12"
        .AddItem "24"
        .AddItem "26"
        .AddItem "52"
        
        ' init to current value
        Select Case DfltValue
            Case 0
                .ListIndex = 3
            Case 12
                .ListIndex = 0
            Case 24
                .ListIndex = 1
            Case 26
                .ListIndex = 2
            Case 52
                .ListIndex = 3
            Case Else
                .ListIndex = 3
        End Select
    
    End With

End Sub

Public Function nNull(ByVal InVal As Variant) As Variant

    nNull = 0
    
    ' 140118 - add err catch
    On Error Resume Next
    
    If IsNull(InVal) Then
        nNull = 0
    ElseIf InVal = "" Then
        nNull = 0
    Else
        nNull = InVal
    End If

    On Error GoTo 0

End Function

Public Function MonthName(ByVal MonthNum) As String

    Select Case MonthNum
        Case 1
            MonthName = "JANUARY"
        Case 2
            MonthName = "FEBRUARY"
        Case 3
            MonthName = "MARCH"
        Case 4
            MonthName = "APRIL"
        Case 5
            MonthName = "MAY"
        Case 6
            MonthName = "JUNE"
        Case 7
            MonthName = "JULY"
        Case 8
            MonthName = "AUGUST"
        Case 9
            MonthName = "SEPTEMBER"
        Case 10
            MonthName = "OCTOBER"
        Case 11
            MonthName = "NOVEMBER"
        Case 12
            MonthName = "DECEMBER"
        Case Else
            MonthName = ""
        End Select

End Function

Public Function GetNumMon(ByVal MonthAbbrev As String) As Integer

    MonthAbbrev = StrConv(MonthAbbrev, vbLowerCase)

    Select Case MonthAbbrev
        Case "jan"
            GetNumMon = 1
        Case "feb"
            GetNumMon = 2
        Case "mar"
            GetNumMon = 3
        Case "apr"
            GetNumMon = 4
        Case "may"
            GetNumMon = 5
        Case "jun"
            GetNumMon = 6
        Case "jul"
            GetNumMon = 7
        Case "aug"
            GetNumMon = 8
        Case "sep"
            GetNumMon = 9
        Case "oct"
            GetNumMon = 10
        Case "nov"
            GetNumMon = 11
        Case "dec"
            GetNumMon = 12
        Case Else
            GetNumMon = 0
    End Select

End Function

Public Function GetMonthAbbrev(ByVal Mth As Byte) As String

    GetMonthAbbrev = ""
    If IsNull(Mth) Then Exit Function
    Select Case Mth
        Case 1:     GetMonthAbbrev = "JAN"
        Case 2:     GetMonthAbbrev = "FEB"
        Case 3:     GetMonthAbbrev = "MAR"
        Case 4:     GetMonthAbbrev = "APR"
        Case 5:     GetMonthAbbrev = "MAY"
        Case 6:     GetMonthAbbrev = "JUN"
        Case 7:     GetMonthAbbrev = "JUL"
        Case 8:     GetMonthAbbrev = "AUG"
        Case 9:     GetMonthAbbrev = "SEP"
        Case 10:    GetMonthAbbrev = "OCT"
        Case 11:    GetMonthAbbrev = "NOV"
        Case 12:    GetMonthAbbrev = "DEC"
    End Select

End Function

Public Function PadString(ByVal InString As String, ByVal StrLen As Integer, Optional ByVal Justify As String) As String

    ' left justified by default
    If Len(InString) > StrLen Then
        PadString = Mid(InString, 1, StrLen)
    Else
        If Justify = "R" Then
            PadString = Space(StrLen - Len(InString)) & InString
        Else
            PadString = InString & Space(StrLen - Len(InString))
        End If
    End If

End Function

Public Function OutNumber(ByVal InNumber As Long, ByVal StrLen As Integer) As String

Dim nString As String

    If InNumber < 0 Then
        MsgBox "OutNumber can not process a negative value! " & InNumber, vbCritical
        End
    End If

    nString = CStr(InNumber)
    
    If Len(nString) > StrLen Then
        MsgBox "OutNumber can not process an excess value! " & InNumber, vbCritical
        End
    End If

    OutNumber = String(StrLen - Len(CStr(InNumber)), "0") & CStr(InNumber)
    
End Function

Public Function ABACheckDigit(ByVal InString As String, Optional Verify As Byte) As Byte

Dim CheckSum As Long
Dim Weight As Byte
    
    ABACheckDigit = 99

    If Len(InString) <> 9 Then
        If Verify = 1 Then
            MsgBox "Invalid length: " & Len(InString) & " " & InString, vbExclamation
        End If
        Exit Function
    End If
    
    ' must be all numeric
    For i = 1 To 8
        
        If InStr("0123456789", Mid(InString, i, 1)) = 0 Then
            If Verify = 1 Then
                MsgBox "Invalid ABA - Numbers only! " & InString, vbCritical
            End If
            Exit Function
        End If
    
        If i = 1 Or i = 4 Or i = 7 Then
            Weight = 3
        ElseIf i = 2 Or i = 5 Or i = 8 Then
            Weight = 7
        Else
            Weight = 1
        End If
        
        CheckSum = CheckSum + Weight * CByte(Mid(InString, i, 1))
        
    Next i
    
    ABACheckDigit = CheckSum Mod 10
    If ABACheckDigit > 0 Then ABACheckDigit = 10 - ABACheckDigit

    ' compare to given value
    If ABACheckDigit <> CByte(Right(InString, 1)) Then
        If Verify = 1 Then
            MsgBox "Invalid Check Digit!", vbExclamation
        End If
        ABACheckDigit = 99
        Exit Function
    End If

End Function

Public Sub GetNudge(ByVal UserID As Long, _
                    ByVal ReportName As String)
                    
    SQLString = "SELECT * FROM PRGlobal WHERE " & _
                "TypeCode = " & PREquate.GlobalTypeNudge & " AND " & _
                "Description = '" & ReportName & "' AND " & _
                "UserID = " & UserID
    
    If PRGlobal.GetBySQL(SQLString) Then
        HorzNudge = PRGlobal.Var1
        VertNudge = PRGlobal.Var2
    Else
        HorzNudge = 0
        VertNudge = 0
    End If
 
End Sub

Public Sub SaveNudge(ByVal UserID As Long, _
                     ByVal ReportName As String)
                         
    SQLString = "SELECT * FROM PRGlobal WHERE " & _
                "TypeCode = " & PREquate.GlobalTypeNudge & " AND " & _
                "Description = '" & ReportName & "' AND " & _
                "UserID = " & UserID
    
    If PRGlobal.GetBySQL(SQLString) Then
        PRGlobal.Var1 = nNull(HorzNudge)
        PRGlobal.Var2 = nNull(VertNudge)
        PRGlobal.Save (Equate.RecPut)
    Else
        If HorzNudge = 0 And VertNudge = 0 Then     ' don't create if value is zero
            Exit Sub
        Else
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeNudge
            PRGlobal.Description = ReportName
            PRGlobal.Var1 = nNull(HorzNudge)
            PRGlobal.Var2 = nNull(VertNudge)
            PRGlobal.UserID = User.ID
            PRGlobal.Save (Equate.RecAdd)
        End If
    End If

End Sub

Public Sub SetNudge(ByRef tdbNum As TDBNumber)
    tdbIntegerSet tdbNum
    With tdbNum
        .Spin = dbiShowAlways
        .MinValue = -255
        .MaxValue = 255
    End With
End Sub

Public Function cmbYrQtrSet(ByRef cmbYr As ComboBox, ByRef cmbQtr As ComboBox) As Boolean
Dim yrs As ADODB.Recordset
Dim i, j, k As Integer

    SQLString = "SELECT DISTINCT YearMonth FROM PRHist ORDER BY YearMonth DESC"
    rsInit SQLString, cn, yrs
    If yrs.RecordCount = 0 Then
        MsgBox "No Payroll History Data Found!!", vbExclamation
        cmbYrQtrSet = False
        Exit Function
    End If
    
    cmbYrQtrSet = True
    
    yrs.MoveFirst
    cmbYr.AddItem Int(yrs!YearMonth / 100)
    Do
        yrs.MoveNext
        If yrs.EOF Then Exit Do
        k = 0
        j = cmbYr.ListCount
        For i = 0 To j - 1
            cmbYr.ListIndex = i
            If cmbYr.text = Int(yrs!YearMonth / 100) Then
                k = 1
                Exit For
            End If
        Next i
        If k = 0 Then
            cmbYr.AddItem (Int(yrs!YearMonth / 100))
        End If
    Loop
    cmbYr.ListIndex = 0

    cmbQtr.AddItem "1"
    cmbQtr.AddItem "2"
    cmbQtr.AddItem "3"
    cmbQtr.AddItem "4"
    
    ' select the default qtr
    Select Case Month(Now())
        Case 1
            cmbQtr.ListIndex = 3    ' Q4
            If cmbYr.ListCount > 1 Then cmbYr.ListIndex = 1
        Case 2 To 4
            cmbQtr.ListIndex = 0    ' Q1
        Case 5 To 7
            cmbQtr.ListIndex = 1    ' Q2
        Case Else
            cmbQtr.ListIndex = 2    ' Q3
    End Select

End Function

Public Sub cmbPoint(ByRef cmb As ComboBox, ByVal ID As Long)

Dim li As Long

    With cmb
        If .ListCount = 0 Then Exit Sub
        .ListIndex = 0
        For li = 0 To .ListCount - 1
            If .ItemData(li) = ID Then
                .ListIndex = li
                Exit For
            End If
        Next li
    End With

End Sub

Public Function AmtMax(ByVal CurrAmt As Currency, _
                       ByVal YTDAmt As Currency, _
                       ByVal MaxAmt As Currency) As Currency

    If YTDAmt + CurrAmt <= MaxAmt Then
        AmtMax = CurrAmt
    ElseIf YTDAmt >= MaxAmt Then
        AmtMax = 0
    Else
        AmtMax = MaxAmt - YTDAmt
    End If
    
End Function

Public Function SuperRound(ByVal Hrs As Currency, ByVal Rate As Currency) As Currency

    ' simulate SuperDOS rounding - to three places then round to two

Dim p1 As Currency

    If Hrs = 0 Or Rate = 0 Then
        SuperRound = 0
        Exit Function
    End If
    
    p1 = Round(Hrs * Rate, 3)
    p1 = p1 + 0.005
    SuperRound = (Int(p1 * 10 ^ 2)) / 10 ^ 2
    
End Function
Public Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim fString As String
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
           
    TableExists = False
           
    Do Until frs.EOF = True
                  
        If frs!Table_Name = TableName Then
            TableExists = True
            Exit Do
        End If
        
       frs.MoveNext
   
   Loop

End Function

Public Function TextSet(ByVal InString As String)
    If IsNull(InString) Then
        TextSet = ""
    Else
        TextSet = Trim(InString)
    End If
End Function

Public Sub PageHeader(Optional ByVal ReportName As String, _
                       Optional ByVal Msg1 As String, _
                       Optional ByVal Msg2 As String, _
                       Optional ByVal msg3 As String, _
                       Optional ByVal SkipLines As Byte, _
                       Optional ByVal UseGLName As Boolean)
                       
Dim SideCols As Integer
Dim HdrName As String
                       
    Ln = SkipLines
    Pg = Pg + 1
   
    If UseGLName = True Then
        HdrName = GLCompany.Name
    Else
        HdrName = PRCompany.Name
    End If
   
    ' 29 characters for fixed left and right portion of first header line
    '    1             8       1   8                    10         1
    ' first line - system date & time / company name / page #
    x = Trim(HdrName)
    y = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
    z = "Page: " & Format(Pg, "####")
   
    If Len(x) > Columns - 39 Then
       x = Mid(Trim(HdrName), 1, Columns - 39)
    End If
           
    If LandSw = 1 Then
        i = ((Columns - Len(x)) / 2) - 29           ' i = 49
        w = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss") & _
            Space(i) & x
        i = Columns - Len(w) - 30
        w = w & Space(i) & "Page: " & Format(Pg, "###0")
    Else
        i = ((Columns - Len(x)) / 2) - 19
        w = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss") & _
            Space(i) & x
        i = Columns - Len(w) - 10
        w = w & Space(i) & "Page: " & Format(Pg, "###0")

    End If
    
    SideCols = (Columns - Len(x)) / 2
    
    w = "  " & Format(Now(), "mm/dd/yy") & " " & Format(Now(), "hh:mm:ss")
    w = Trim(w) & Space(SideCols - 19) & Trim(x)
    w = Trim(w) & Space(SideCols - 10) & "Page: " & Format(Pg, "###0")
    
    PrintValue(1) = w
    FormatString(1) = "a" & Columns
    
    PrintValue(2) = " "
    FormatString(2) = "~"
    FormatPrint
    
    ' PrtCenter Ln, w
    
    Ln = Ln + 1
   
    If ReportName <> "" Then
        PrtCenter 0, ReportName
        Ln = Ln + 1
    End If
           
'    If QtrEnding <> "" Then
'       PrtCenter Ln, QtrEnding
'       Ln = Ln + 1
'    End If
   
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

Public Function ParseString(ByVal InString As String, ByVal SepString As String) As ADODB.Recordset

Dim rs As New ADODB.Recordset
Dim i, j As Long
Dim x, y As String

    Set ParseString = New ADODB.Recordset
    ParseString.CursorLocation = adUseClient
    ParseString.Fields.Append "ListValue", adVarChar, 255, adFldIsNullable
    ParseString.Open , , adOpenDynamic, adLockOptimistic
    
    If IsNull(InString) Then Exit Function
    If InString = "" Then Exit Function
    
    j = Len(Trim(InString))
    x = ""
    y = ""
    For i = 1 To j
        If Mid(InString, i, 1) = SepString Then
            ParseString.AddNew
            ParseString!listvalue = x
            ParseString.Update
            x = ""
            i = i + 1
            If i > j Then
                Exit For
            End If
        End If
        x = Trim(x) & Mid(InString, i, 1)
    Next i
    If x <> "" Then
        ParseString.AddNew
        ParseString!listvalue = x
        ParseString.Update
    End If

End Function

Public Sub TestPattern()
    
    For i = 1 To MaxLines
        x = ""
        For j = 1 To Columns
            If i Mod 2 = 1 Then
                x = Trim(x) & j Mod 10
            Else
                If j Mod 10 = 0 Then
                    x = x & Int(j / 10)
                Else
                    x = x & " "
                End If
            End If
        Next j
        PrintValue(1) = x
        FormatString(1) = "a" & Columns
        PrintValue(2) = " "
        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
    Next i

End Sub

Public Function LoadQBDrop(ByVal SQLString As String) As String

    If QBAccount.GetBySQL(SQLString) = False Then
        LoadQBDrop = ""
        Exit Function
    End If
    
    Do
        LoadQBDrop = Trim(LoadQBDrop) & "|#" & QBAccount.QBID & ";" & QBAccount.Name
        If QBAccount.GetNext = False Then Exit Do
    Loop

End Function

Public Function SplitCalc(ByVal BasisAmt As Currency, _
                           ByVal BasisTotal As Currency, _
                           ByVal DistAmt As Currency) As Currency

    If BasisTotal = 0 Then
        SplitCalc = 0
    ElseIf BasisAmt = BasisTotal Then
        SplitCalc = DistAmt
    Else
        SplitCalc = Round(BasisAmt / BasisTotal * DistAmt, 2)
    End If

End Function

Public Function GetFileName(ByVal Str As String) As String
    
Dim i, j As Long
Dim x As String
    
    GetFileName = ""
    
    If IsNull(Str) Then Exit Function
    If Str = "" Then Exit Function
    
    x = Trim(Str)
    j = Len(x)
    For i = j To 1 Step -1
        If Mid(x, i, 1) = "\" Then Exit For
        GetFileName = Mid(x, i, 1) & GetFileName
    Next i

End Function

Public Function PRBilling() As Boolean

    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                " AND Var1 = 'YES'" & _
                " AND Var2 = '" & PRCompany.GLCompanyID & "'"
    PRBilling = PRGlobal.GetBySQL(SQLString)

End Function

Public Sub rsDelAll(ByRef rs As ADODB.Recordset)

    Do
        If rs.RecordCount = 0 Then Exit Sub
        rs.MoveFirst
        rs.Delete
    Loop

End Sub

Public Sub Delay(ByVal Secs As Integer)

Dim n As Variant
Dim s1 As Long
Dim s2 As Long

    ' delay a specified number of seconds - what if spans a day?
    n = Now()
    s1 = (n - Int(n)) * 60 * 60 * 24   ' number of seconds since midnight
    
    Do
       n = Now()
       s2 = (n - Int(n)) * 60 * 60 * 24   ' number of seconds since midnight
       If s2 - s1 >= Secs Then Exit Do
    Loop
 
End Sub

Public Function mdbName(ByVal Str As String) As String

Dim mdbI, mdbJ, mdbK As Long

    mdbName = ""
    If Str = "" Then Exit Function
    If InStr(1, Str, "\", vbTextCompare) = 0 Then Exit Function
    
    mdbK = Len(Str)
    For mdbI = mdbK To 1 Step -1
        If Mid(Str, mdbI, 1) = "\" Then
            Exit For
        End If
    Next mdbI
    If mdbI = 0 Then Exit Function
    mdbName = Trim(Mid(Str, mdbI + 1, mdbK))

End Function

Public Function MaxLen(ByVal Str As String, ByVal Ln As Integer) As String

    If IsNull(Str) Then
        MaxLen = ""
    Else
        MaxLen = Trim(Str)
        If Len(Str) > Ln Then
            MaxLen = Mid(Str, 1, Ln)
        End If
    End If

End Function

Public Function SlashSplit(ByVal SString As String, sSide As Integer) As String
    ' divide string by slash - sSide = 1 - left of slash / sSide = 2 - right of slash
Dim sPos As Integer

    If IsNull(SString) = True Or Len(Trim(SString)) <= 1 Then
        SlashSplit = ""
        Exit Function
    End If

    sPos = InStr(1, SString, "/")
    
    If sSide = 1 Then
        If sPos = 0 Then
            SlashSplit = SString
        Else
            SlashSplit = Mid(SString, 1, sPos - 1)
        End If
    Else
        If sPos = 0 Then
            SlashSplit = ""
        Else
            SlashSplit = Mid(SString, sPos + 1)
        End If
    End If

End Function

Public Function DateSplit(ByVal sDate As Date) As String
    If IsNull(sDate) = True Then
        DateSplit = ""
    Else
        DateSplit = Format(Month(sDate), "00") & "  " & Format(Day(sDate), "00") & "  " & Year(sDate)
    End If
End Function

Public Function FNameOnly(ByVal fnm As String) As String

    Dim pos As Integer
    pos = InStrRev(fnm, "\")
    If pos = -1 Then
        FNameOnly = fnm
    Else
        FNameOnly = Right(fnm, Len(fnm) - pos)
    End If

End Function

Public Function DollarAndCents(ByVal Amt As Currency) As String
    ' 2018-09-29 - needs rounded?  - returned "8  100" for 9.00 ..
    '   corrected double negative error
    Amt = nNull(Amt)

    Dim dAmt, cAmt As Integer
    Dim AmtSign As String
        
    AmtSign = "+"
    If Amt < 0 Then
        AmtSign = "-"
    End If
    
    Amt = Abs(Amt)
    
    dAmt = Int(Amt)
    cAmt = (Amt - Int(Amt)) * 100
    
    If AmtSign = "+" Then
        DollarAndCents = Format(dAmt, "##,###,##0") & "  " & Format(cAmt, "00")
    Else
        ' 2018-09-29
        ' DollarAndCents = "-" & Format(-dAmt, "##,###,##0") & "  " & Format(cAmt, "00")
        DollarAndCents = "-" & Format(dAmt, "##,###,##0") & "  " & Format(cAmt, "00")
    End If

End Function

