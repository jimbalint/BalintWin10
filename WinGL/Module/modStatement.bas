Attribute VB_Name = "modStatement"
Option Explicit

Dim SignMode As Integer
Dim SignTemp As Integer
Dim DollarSign As String
Dim lngBalSht As Long
Dim GTotal(20) As Currency
Dim BSTotal(4) As Currency
Dim GLDate(20) As String

Dim frmGLPring As New frmGLPrint
Dim flg As Boolean
Dim DBName As String
Dim CompanyID As Long
Dim act As Long
Dim CYAmt As Currency
Dim CYSum As Currency
Dim PYAmt As Currency
Dim PYSum As Currency
Dim PFormat As String

Dim PctBaseCYAmt As Currency
Dim PctBasePYAmt As Currency
Dim PctBaseCYSum As Currency
Dim PctBasePYSum As Currency

Dim PctValCYAmt As Currency
Dim PctValPYAmt As Currency
Dim PctValCYSum As Currency
Dim PctValPYSum As Currency

Dim AmtString As String
Dim PrtString As String
Dim xCurr As Currency

Dim DateString As String
Dim DW As String

Dim PrtTab As Byte

Dim BalFlg1 As Long        ' sp%[1]
Dim BalFlg2 As Long        ' sp%[2]
Dim GFlag As Byte          ' sp%

Dim BSFlag As Integer         ' BS

Dim AcctD As String
Dim Mths As Integer

Dim P1 As Long
Dim AcctFlag As Boolean

Dim LoCons As Long
Dim HiCons As Long

Dim LoBranch As Integer
Dim HiBranch As Integer

Dim BranchTotal1 As Currency   ' Y[1]
Dim BranchTotal2 As Currency   ' Y[2]

Dim PLFlag As Boolean      ' P1

Dim StartPD As Byte
Dim FiscalYear As Integer
Dim EndPd As Byte
Dim MathFlag As Boolean
Dim PctAcct As Long
Dim PrevAcct As Long
Dim OutCH As Integer
Dim LastAcct As Long
Dim GLFName As String
Dim GLID As Long
Dim BatchNum As Long
Dim CompanyOption As String

Public Sub GLStatement(Optional StmtCount As Byte)
   
   ' store if comparative
   ' if not - separate record set for prev year not needed
   ' from GLAccount.GetRecordSets
   If GLPrint.RegCmp = Equate.Comp Then
      CompFlag = True
      HAdj = 1.5
   Else
      CompFlag = False
      HAdj = 0
   End If
   
    ' Company Option
    ' special formatting for Richlak
    ' GLPrint / "R"
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                " AND Description = 'GLPrint'"
    If PRGlobal.GetBySQL(SQLString) Then
        CompanyOption = PRGlobal.Var1
    Else
        CompanyOption = ""
    End If
   
   ' convert calendar month to period #
   StartPD = GetPeriod(GLCompany.NumberPds, GLPrint.BeginDate Mod 100, GLCompany.FirstPeriod)
   EndPd = GetPeriod(GLCompany.NumberPds, GLPrint.EndDate Mod 100, GLCompany.FirstPeriod)
   
   LoBranch = GLPrint.LowBranchAcct
   HiBranch = GLPrint.HiBranchAcct
   
   LoCons = GLPrint.LowConsAcct
   HiCons = GLPrint.HiConsAcct
   
   SignMode = 1   ' Currently on:  1=BalSht  2=IncStmt
   lngBalSht = 0 ' balance sheet flag
   
   PLFlag = False      ' set to true when first P record is encountered
   
   BalFlg1 = 0       ' sp%[1]
   BalFlg2 = 0       ' sp%[2]
   GFlag = 0         ' sp%
   BSFlag = 0        ' BS
   
   ' call routine from global
   GetDates GLID
   
   ' set up date strings
   DateSetup
       
    ' initialize the print screen - 12 point/portrait for balance sheet
    If StmtCount <= 1 Then          ' from the menu or free format first time through
        PrtInit "Port"
   
        If GLPrint.PrintBIB = Equate.PrtBSOnly Or GLPrint.PrintBIB = Equate.PrtBoth Then
        ' 10 ???
              
            ' balance sheet not comparative
            If Not CompFlag Then
                SetFont 11, Equate.Portrait
            Else
                If GLPrint.WidePrint = True Then
                    SetFont 8, Equate.Portrait             ' compressed - portrait
                Else
                    SetFont 9, Equate.LandScape            ' not compressed - landscape
                End If
            End If
    
        Else
            If GLPrint.RegCmp = Equate.Comp Then
                If GLPrint.WidePrint = True Then
                    SetFont 8, Equate.Portrait             ' compressed - portrait
                Else
                    SetFont 9, Equate.LandScape            ' not compressed - landscape
                End If
            Else
                SetFont 11, Equate.Portrait
            End If
        End If

    Else
    
        Prvw.vsp.Orientation = orPortrait
   
        If GLPrint.PrintBIB = Equate.PrtBSOnly Or GLPrint.PrintBIB = Equate.PrtBoth Then
        ' 10 ???
              
            ' balance sheet not comparative
            If Not CompFlag Then
                Prvw.vsp.Font.Size = 11
            Else
                If GLPrint.WidePrint = True Then
                    Prvw.vsp.FontSize = 8           ' compressed - portrait
                Else
                    Prvw.vsp.Orientation = orLandscape
                    Prvw.vsp.FontSize = 9         ' not compressed - landscape
                End If
            End If
    
        Else
            If GLPrint.RegCmp = Equate.Comp Then
                If GLPrint.WidePrint = True Then
                    Prvw.vsp.FontSize = 8           ' compressed - portrait
                Else
                    Prvw.vsp.Orientation = orLandscape
                    Prvw.vsp.FontSize = 9          ' not compressed - landscape
                End If
            Else
                Prvw.vsp.FontSize = 11
            End If
        End If
    
        Prvw.vsp.NewPage
    
    End If

   ' 1050
   If GLPrint.RegBraCon = Equate.Branch Then
      LoCons = LoBranch
      HiCons = LoBranch
   End If
   
GLFS:
   
   GLFS
 
    ' 3900 / 3920
    If GLPrint.PrintBIB <> Equate.PrtBSOnly Or (GLPrint.RegBraCon = Equate.Regular And BSFlag = 0) = False Then
        ' >>>> form feed below if printing multiple branches
        ' FormFeed
    End If
   
   ' 4040
   LoCons = LoCons + 1
   HiCons = LoCons
   
   BSFlag = 0
   BranchTotal1 = BranchTotal1 + GTotal(5)
   BranchTotal2 = BranchTotal2 + GTotal(10)
   Ln = 0
   
   ' 4050    - Print next branch
   If LoCons <= HiBranch And GLPrint.RegBraCon = Equate.Branch Then
   
      ' clear variables - 1090
      BSFlag = 0            ' BS
      BalFlg1 = 0           ' sp%[1]
      BalFlg2 = 0           ' sp%[2]
      lngBalSht = 0         ' B1
      BranchTotal1 = 0      ' Y[1]
      BranchTotal2 = 0      ' Y[2]
      SignMode = 1          ' L1
      
      For ii = 1 To 20
         GTotal(ii) = 0           ' G[ ]
         BSTotal(ii / 5) = 0      ' Z[ ]
      Next ii
   
      ' >>>>
      FormFeed
   
      GoTo GLFS
   
   End If
   
   If GLPrint.RegBraCon = Equate.Branch Then
      
      Call Prt(Ln, 1, "")
      Call Prt(Ln, 1, "")
      Call Prt(Ln, 1, "")
      Ln = Ln + 3
      
      FormatString(1) = "a42"
      FormatString(2) = "t49"
      FormatString(3) = "d11"
      FormatString(4) = "t65"
      FormatString(5) = "d11"
      FormatString(6) = "~"
      
      PrintValue(1) = "BRANCH GRAND TOTALS"
      PrintValue(3) = BranchTotal1
      PrintValue(5) = BranchTotal2
   
   End If
   
   frmProgress.MousePointer = vbArrow
   frmProgress.Hide
  
End Sub

Private Sub GLFS()

Dim g(20) As Currency
Dim Z(5) As Currency
Dim flg As Boolean
Dim SP As Integer
Dim Acct As Long


   AcctFlag = False

   ' opens company record and assigns date variables
   GetDates GLID
   
   ' set variables
   P1 = 2147483647       ' set to highest possible value for a long
   Ln = 0
   
   ' 1060
   If GLCompany.SubDigits <= 0 Then GLCompany.SubDigits = 1
   LoCons = LoCons Mod 10 ^ GLCompany.SubDigits
   HiCons = HiCons Mod 10 ^ GLCompany.SubDigits
   
   ' get record set of all accounts in range
   ' terminate if none found
   
   ' open the progress window
   frmProgress.lblMsg1 = "Printing Statement for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering GLAccount Info ... " & LoCons
   frmProgress.Show
   
   ' ========================= Get Account Record Sets ==================================
   If GLAccount.GetRecordSets(Year(CurrYrFYEnd), Year(CurrYrFYEnd) - 1) = False Then
      End
   End If
   ' ====================================================================================
   
   ' skip to first record if not default low gl account #
   If GLPrint.LowAccount <> 0 And GLPrint.LowAccount <> 1 Then
      Acct = GLPrint.LowAccount
      Do
         If GLAccount.GetAccount(Acct) Then Exit Do
         Acct = Acct + 1
         If Acct > GLPrint.HiAccount Then      ' No accts found ????
            MsgBox "No accounts found for range: " & GLPrint.LowAccount & " to: " & GLPrint.HiAccount, vbExclamation + vbOKOnly, _
                   "Statement Print"
            Exit Sub
         End If
      Loop
   End If
   
   Do
      
      ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
      ' xxx check the GLAccount record  filters, skips etc. xxx
      ' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
      
      If GLAccount.Account < GLPrint.LowAccount Then GoTo Cycle
      If GLAccount.Account > GLPrint.HiAccount Then Exit Do
      
      ' 1390 - sp%=0
      GFlag = 0
      PrtTab = GLAccount.PrintTab
      
      ' special line from Richlak - skip GLFG header records
      If GLAccount.AcctType = "H" And GLAccount.TotalLevel > 0 Then GoTo Cycle
      
      ' balance sheet only selected 1440
      If GLPrint.PrintBIB = Equate.PrtBSOnly And GLAccount.Account > GLCompany.FirstPAcct Then Exit Do
      
      ' first p record encountered - 1450 --> 3220
      If P1 = 2147483647 And GLAccount.AcctType = "P" Then
         PRec
         GoTo Cycle
      End If
      
      ' income statement only selected 1460
      If GLPrint.PrintBIB = Equate.PrtISOnly And GLAccount.Account < GLCompany.FirstPAcct Then GoTo Cycle
      
      ' branch or consolidated selected - check range 1470 & 1480
      If GLPrint.RegBraCon = Equate.Branch Or GLPrint.RegBraCon = Equate.Consol Then
         If GLAccount.Account Mod 10 ^ GLCompany.SubDigits < LoCons Then GoTo Cycle
         If GLAccount.Account Mod 10 ^ GLCompany.SubDigits > HiCons Then GoTo Cycle
      End If
      
      ' all stmts/sched  ... 1500-1520 - last filter
      flg = False
      If GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules Then flg = True                         ' all statements
      If GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements Then flg = True                        ' all schedules
      
      If GLPrint.RegBraCon = Equate.Branch And GLAccount.BranchAcct Then flg = True  ' branch acct
      If GLPrint.RegBraCon = Equate.Consol And GLAccount.ConsAcct Then flg = True    ' cons acct
      
      If flg = False Then
      
         If InStr("N0TM", GLAccount.AcctType) = 0 Then GoTo Cycle
      
         ' >>>>> 1540 let sp%=1 <<<<<<<
         GFlag = 1
      
      End If
      
      GetDesc
      
      ' 1910 0NMT
      
      ' 1920
      If InStr("0N", GLAccount.AcctType) <> 0 Then Type0NM
      
      ' 1930
      If GLAccount.AcctType = "M" And GLPrint.UseMathRec = True Then Type0NM
      
      ' 1940
      If GLAccount.AcctType = "T" Then PrintT
      
      ' 1950
      If GLPrint.RegBraCon = Equate.Consol And GLAccount.Account Mod 10 ^ GLCompany.SubDigits <> HiCons Then
         GoTo Cycle
      End If
      
      ' 1960
      If GLAccount.AcctType = "C" Then TotalClear
      
      ' 1970
      If GLAccount.AcctType = "B" Then BalSht
      
      ' 1980
      If GLAccount.AcctType = "U" Then UnderLn
      
      ' 1990
      If GLAccount.AcctType = "P" Then PRec
      
      ' 2000
      If GLAccount.AcctType = "." Then Percent
      
Cycle:
   
      ' status display update
      act = act + 1
      If act Mod 10 = 0 Then
         If GLPrint.RegBraCon = Equate.Branch Then
            frmProgress.lblMsg2 = "On Account #: " & GLAccount.Account & " Branch #: " & LoCons
         Else
            frmProgress.lblMsg2 = "On Account#: " & GLAccount.Account
         End If
         frmProgress.Refresh
      End If
   
      ' get next glaccount record
      LastAcct = GLAccount.Account
      If GLAccount.GetNext = False Then Exit Do
      
'      If OutCH <> 0 Then
'         Print #OutCH, LastAcct; "  "; GLAccount.Account; " "; P1; " "; LoCons; " "; LoBranch
'      End If
   
   Loop
      
End Sub

Private Sub DateSetup()       ' 0660
   
Dim DString As String
Dim Months As String
Dim Digit(13) As String
Dim fmt As String
Dim ii As Double

   fmt = "mmmm dd, yyyy"
   
   Digit(0) = " Zero"
   Digit(1) = " One"
   Digit(2) = " Two"
   Digit(3) = " Three"
   Digit(4) = " Four"
   Digit(5) = " Five"
   Digit(6) = " Six"
   Digit(7) = " Seven"
   Digit(8) = " Eight"
   Digit(9) = " Nine"
   Digit(10) = " Ten"
   Digit(11) = " Eleven"
   Digit(12) = " Twelve"
   Digit(13) = " Thirteen"
   
   GLDate(0) = Format(CurrYrPdEnd, fmt)
   
   
   If EndPd = 1 Then
      GLDate(1) = CStr(Digit(EndPd)) & " Month Ended " & Format(CurrYrPdEnd, fmt)
   Else
      GLDate(1) = CStr(Digit(EndPd)) & " Months Ended " & Format(CurrYrPdEnd, fmt)
   End If
   
   GLDate(2) = Format(CurrYrCurrPdBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
   GLDate(3) = Format(CurrYrFYBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
   If EndPd = 1 Then
      GLDate(4) = CStr(Digit(EndPd)) & " Month Ended " & Format(PrevYrPdEnd, fmt)
   Else
      GLDate(4) = CStr(Digit(EndPd)) & " Months Ended " & Format(PrevYrPdEnd, fmt)
   End If
   
   GLDate(5) = Format(PrevYrPDBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
   GLDate(6) = Format(PrevYrFYBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
   GLDate(7) = Format(Now, fmt)
   
   GLDate(8) = CStr(EndPd * 4) & " Weeks Ended " & Format(CurrYrPdEnd, fmt)
   
   GLDate(9) = CStr(EndPd * 4) & " Weeks Ended " & Format(PrevYrPdEnd, fmt)
   
   X = Trim(Format(CurrYrPdEnd, fmt))
   GLDate(10) = Space(18 - Len(X)) & X
   
   GLDate(11) = Space(18 - Len(Format(CurrYrCurrPdBeg, fmt))) & Format(CurrYrCurrPdBeg, fmt)
   
   GLDate(12) = Space(18 - Len(Format(CurrYrFYBeg, fmt))) & Format(CurrYrFYBeg, fmt)
   
   GLDate(13) = Space(18 - Len(Format(PrevYrPdEnd, fmt))) & Format(PrevYrPdEnd, fmt)
   
   GLDate(14) = Space(18 - Len(Format(PrevYrPDBeg, fmt))) & Format(PrevYrPDBeg, fmt)
   
   GLDate(15) = Space(18 - Len(Format(PrevYrFYBeg, fmt))) & Format(PrevYrFYBeg, fmt)
   
   If EndPd - StartPD + 1 = 1 Then
      DString = CStr(Digit(EndPd - StartPD + 1)) & " Month"
   Else
      DString = CStr(Digit(EndPd - StartPD + 1)) & " Months"
   End If
   X = Trim(DString)
   GLDate(16) = Space(18 - Len(X)) & X

   If EndPd = 1 Then
      DString = CStr(Digit(EndPd)) & " Month"
   Else
      DString = CStr(Digit(EndPd)) & " Months"
   End If
   X = Trim(DString)
   GLDate(17) = Space(18 - Len(X)) & X
   
   GLDate(18) = " And " & Year(CurrYrFYEnd)
   
   GLDate(19) = " And " & Year(CurrYrFYEnd) - 1
   
   ' date type 1   - ex. pd 7 - 9
   Mths = EndPd - StartPD + 1
   GLDate(20) = CStr(Digit(Mths)) & " Month"                              ' Three Month
   If Mths > 1 Then GLDate(20) = GLDate(20) & "s"                          ' Three Months
   GLDate(20) = GLDate(20) & " And" & CStr(Digit(EndPd)) & " Month"       ' Three Months And Nine Month
   If EndPd > 1 Then GLDate(20) = GLDate(20) & "s"                         ' Three Months And Nine Months
   GLDate(20) = GLDate(20) & " Ended " & Format(CurrYrPdEnd, fmt)
   
   ' shift to upper case 0980
   If GLPrint.LowerCaseDate = False Then
      For ii = 0 To 20
          GLDate(ii) = UCase(GLDate(ii))
      Next ii
   End If

End Sub

Private Sub GetDesc()        ' 1690
   
'   AcctD = AcctDes(GLAccount.Account)
    
    AcctD = GLAccount.FullDesc
   
   ' 1840
   If GLPrint.PrtAcctNum = True And GLAccount.AcctType = "0" Then
      AcctD = CStr(GLAccount.Account) & " " & AcctD
   End If
   
   ' 1850
   If GLPrint.PrtAcctNum = True And GLAccount.AcctType = "T" And GLAccount.BSColumn = 1 Then
      AcctD = CStr(GLAccount.Account) & " " & AcctD
   End If
   
   ' 1860
   If GLPrint.RegBraCon = Equate.Branch And GLAccount.Account Mod 10 ^ GLCompany.SubDigits = HiCons Then
      BSFlag = BSFlag + 1
   End If
   
   ' 1870
   If GLPrint.RegBraCon = Equate.Consol And GLAccount.Account Mod 10 ^ GLCompany.SubDigits <> HiCons Then
      Exit Sub       ' goto 1910
   End If
   
   If GLAccount.DescNumber = 1 And GLAccount.AcctType = "H" Then
      AcctD = GLCompany.Name
   End If
   
   ' 1880
   If GLAccount.AcctType = "D" Then DateFormat
   
   ' 1890
   If GLAccount.AcctType = "H" Then HeaderFormat
   
   ' 1900
   If InStr("ALIE", GLAccount.AcctType) <> 0 Then ALIE
   
End Sub


Private Sub DateFormat()     ' 2020

   If GLAccount.Account > P1 Then
      AcctD = GLDate(2)
   Else
      AcctD = GLDate(1)
   End If
   
'   If GLAccount.TotalLevel <= 17 Then AcctD = AcctDes(GLAccount.Account) & GLDate(GLAccount.TotalLevel)

   ' 2070
   If GLAccount.TotalLevel <= 17 Then AcctD = GLAccount.FullDesc & GLDate(GLAccount.TotalLevel)
   
   ' new format !!! For the ### months and ### months ended MMMMM, dd, yyyy
   If GLAccount.TotalLevel = 20 Then AcctD = GLAccount.FullDesc & GLDate(20)
   
   HeaderFormat
   
End Sub

Private Sub HeaderFormat()       ' 2080 - center string

Dim ss1 As Integer
Dim ss2 As Integer
Dim PFlg As Boolean

Dim sLen As Integer
    
Dim x1 As String
Dim x2 As String
    
    ' supp cur pd  2110
    ' acutal start character
    sLen = Len(LastH) + GLAccount.PrintTab - 1
    sLen = Len(LastH) + GLAccount.PrintTab - 1
   
    ' ???? this works for R&C ????
    ' If GLPrint.SupprCP = True And GLAccount.Account > P1 And GLAccount.PrintTab <> 0 Then
    
    ' ???? use this for Richlak ???
    If GLPrint.SupprCP = True And GLAccount.Account > P1 And GLAccount.PrintTab >= 20 And GLAccount.PrintTab <= 56 Then
      
        ss1 = sLen
        ss2 = sLen + Len(AcctD) - 1
        
        x1 = Space(sLen) & AcctD & Space(200)
        x1 = Space(sLen) & AcctD & Space(100)
 

        ' x1 = Left(x1, 34) & Space(22) & Mid(x1, 57, 22) & Space(22) & Right(x1, Len(x1) - 104)
            
        
        ' richlak 06/28/10
        If CompanyOption = "RL" Then
            x2 = Left(x1, 33) & Space(23) & Mid(x1, 56, 22) & Space(22) & Mid(x1, 102, 23)
        Else
            ' R&C 2/22/2010
            ' x2 = Left(x1, 34) & Space(22) & Mid(x1, 57, 22) & Space(22) & Mid(x1, 102, 23)
            x2 = Left(x1, 34) & Space(22) & Mid(x1, 57, 22) & Space(23) & Mid(x1, 103, 22)
        End If
        
        If ss1 = 0 Or ss2 <= ss1 Then
            AcctD = ""
        Else
            AcctD = Mid(x2, ss1, ss2 - ss1)
            AcctD = Mid(x2, ss1 + 1, Len(AcctD) + 2)
        End If
    
    End If
      
' x1 = "|" & AcctD
' x2 = "x"
'
'        If sLen >= 36 And sLen <= 57 Then                   ' starts in 1st column
' x2 = "a "
'            If Len(AcctD) > 22 Then
'                AcctD = Right(AcctD, 58 - sLen)
'            Else
'                ' AcctD = Space(Len(AcctD))
'                GLAccount.PrintTab = GLAccount.PrintTab + Len(AcctD)
'            End If
'
'        ElseIf sLen >= 82 And sLen <= 103 Then              ' starts in 3rd column
' x2 = "b "
'
'            If Len(AcctD) > 22 Then
'                AcctD = Right(AcctD, 104 - sLen)
'            Else
'                ' AcctD = Space(Len(AcctD))
'                GLAccount.PrintTab = GLAccount.PrintTab + Len(AcctD)
'            End If
'
'        ElseIf sLen < 36 And sLen + Len(AcctD) >= 36 Then   ' starts before 1st column goes into 1st column
' x2 = "c "
'            If Len(AcctD) > 22 Then
'                AcctD = Left(AcctD, 36 - sLen)
'            Else
'                ' AcctD = Space(Len(AcctD))
'                GLAccount.PrintTab = GLAccount.PrintTab + Len(AcctD)
'            End If
'
'
'        ElseIf sLen < 82 And sLen + Len(AcctD) >= 82 Then   ' starts before 3rd column goes into 3rd column
' x2 = "d "
'            If Len(AcctD) > 22 Then
'                AcctD = Left(AcctD, 82 - sLen)
'            Else
'                ' AcctD = Space(Len(AcctD))
'                GLAccount.PrintTab = GLAccount.PrintTab + Len(AcctD)
'            End If
'
'        End If
'


' If GLAccount.Account >= 4890 And GLAccount.Account <= 5000 Then
'   MsgBox x2 & GLAccount.Account & vbCr & x1 & vbCr & Len(x1) & vbCr & AcctD & vbCr & sLen
' End If

'        ' blank out 1st column?
'        If sLen >= 36 And sLen <= 57 Then
'            ss1 = 35 - Len(LastH) - GLAccount.PrintTab
'            If ss1 < 1 Then ss1 = 1
'
'            ss2 = 57 - Len(LastH) - GLAccount.PrintTab
'            If ss2 > Len(AcctD) Then ss2 = Len(AcctD)
'
'            AcctD = Mid(AcctD, 1, ss1 - 1) & Space(ss2 - ss1 + 1) & Mid(AcctD, ss2 + 1)
'
'        End If
'
'        If sLen >= 80 And sLen <= 103 Then
'            ss1 = 80 - Len(LastH) - GLAccount.PrintTab
'            If ss1 < 1 Then ss1 = 1
'
'            ss2 = 103 - Len(LastH) - GLAccount.PrintTab
'            If ss2 > Len(AcctD) Then ss2 = Len(AcctD)
'
'            AcctD = Mid(AcctD, 1, ss1 - 1) & Space(ss2 - ss1 + 1) & Mid(AcctD, ss2 + 1)
'
'        End If
        
'
'
'      ' 2088
'      If Len(AcctD) + GLAccount.PrintTab >= 81 Then
'
'         If 81 - GLAccount.PrintTab < 1 Then
'            ss1 = 1
'         Else
'            ss1 = 81 - GLAccount.PrintTab
'         End If
'
'         If 103 - GLAccount.PrintTab > Len(AcctD) Then
'            ss2 = Len(AcctD)
'         Else
'            ss2 = 103 - GLAccount.PrintTab
'         End If
'
'         AcctD = Mid(AcctD, 1, ss1 - 1) & Space(ss2 - ss1 + 1) & Mid(AcctD, ss2 + 1)
'
'      End If
'      ' 2088
      
'   End If
   
   ' >>> 2110 - keep within 80 columns if ....
   Ln = Ln + 1
   
' 0505
'   AcctD = GLAccount.Account & " " & AcctDes(GLAccount.Account)
   
 If OutCH <> 0 Then
    Print #OutCH, GLAccount.Account; " "; GLAccount.PrintTab; " "; CCol; " "; Len(AcctD); " "; AcctD
 End If
   
   If GLAccount.PrintTab = 0 Then
       
       FormatString(1) = "a124"
       PrintValue(1) = Centered(Trim(AcctD), Columns)
       FormatString(2) = "~"
       FormatPrint
       
   Else
       
       ' takes the place of line 2110
       '    don't print wide if not comparative or budget
       PFlg = False
       If GLPrint.RegCmp = Equate.Comp Then PFlg = True
       If GLPrint.RegBraCon = Equate.Budget Then PFlg = True

' suzy
'       If GLAccount.Account < 900 Then PFlg = True ' ??? 2110 ac[1]
       
       If GLAccount.PrintTab + CCol + Len(Trim(AcctD)) <= 82 Then PFlg = True
       
'       If PFlg Then Prt Ln, CCol, Space(GLAccount.PrintTab) & AcctD
       
       If PFlg Then
          FormatString(1) = "x" & GLAccount.PrintTab
          PrintValue(1) = ""
          X = "a" & Len(AcctD)
          FormatString(2) = X
          
          PrintValue(2) = AcctD
          FormatString(3) = "~"
          FormatPrint
       End If
       
   End If

   LnFeeds

End Sub

Private Sub LnFeeds()            ' 2130

   ' >>> coding for landscape/legal
   ' 0505
   If GLAccount.LineFeeds = 255 Then
      Ln = Ln - 1
      CCol = Len(PrintString) + 1  ' + 1 ??? - date formats ???
      
      LastH = LastH & Space(GLAccount.PrintTab) & AcctD
      
      Exit Sub
   Else
      CCol = 0
      
      LastH = ""
   End If
   
   Lines = GLAccount.LineFeeds
   
   If Lines >= 40 And Ln < Lines Then
      Lines = Lines - Ln
   End If
   
   ' 0505
   If Ln > MaxLines Then
      FormFeed
   End If
   
   For ii = 1 To Lines - GFlag
       Ln = Ln + 1
       Prt Ln, 1, ""
       If Ln > MaxLines Then
          FormFeed
       End If
   
   Next ii
   
' MsgBox "Acct#: " & GLAccount.Account & vbCrLf & _
'        "Ln: " & Ln & vbCrLf & _
'        "Lines: " & GLAccount.LineFeeds & " " & Lines & vbCrLf & _
'        AcctD
   
End Sub


Private Sub UnderLn()      ' 3100
   
Dim u As String
   
    AcctD = String(17, " ") & String(14, "=")
   
    If GLAccount.TotalLevel >= 1 And GLAccount.TotalLevel <= 3 Then
        AcctD = String(17, " ") & String(14, "-")
    End If
   
    If GLAccount.TotalLevel <> 1 And GLAccount.BSColumn <> 1 Then
        ii = 48
    Else
        ii = 32
    End If
    
    If GLAccount.Account >= GLCompany.FirstPAcct Then   ' 3150 - Income Stmt
      
        If GLAccount.TotalLevel >= 1 And GLAccount.TotalLevel <= 3 Then
            u = "-"
        Else
            u = "="
        End If
      
        Let AcctD = String(13, u) & String(2, " ") & String(5, u) & _
                    String(2, " ") & String(13, u) & String(2, " ") & String(5, u)

        ii = 36
      
        If GLPrint.SupprCP Then
            ' R&C - 02/22/2010
            ' AcctD = Space(22) & Mid(AcctD, 23, 14)
        End If
      
    End If
   
    If GLPrint.RegCmp = Equate.Comp Then     ' duplicate for comparatives 3200
        ' AcctD = RTrim(AcctD) & String(5, " ") & AcctD
        AcctD = Trim(AcctD) & String(5, " ") & AcctD
    End If

'   If GLAccount.PrintTab = 0 Then GLAccount.PrintTab = ii
    GLAccount.PrintTab = ii

   HeaderFormat

End Sub

' section header for Assets/Liabilities/Income/Expense
Private Sub ALIE()         ' 2230
   
   If InStr(1, "AE", GLAccount.AcctType) Then SignMode = 1
   If InStr(1, "IL", GLAccount.AcctType) Then SignMode = -1
   
   If GLAccount.PrintTab = 0 Then
      PrtTab = 1
   Else
      PrtTab = GLAccount.PrintTab
   End If
   
   Ln = Ln + 1
   
   FormatString(1) = "x" & PrtTab
   
   X = "a" & Len(AcctD)
   FormatString(2) = X
   PrintValue(2) = AcctD
   
   FormatString(3) = "~"
     
   FormatPrint
   
   LnFeeds
   
End Sub

Private Sub BalSht()       ' 3030  Type "B"
   
   ' first time through - set the flag and exit
   If lngBalSht = 0 Then
      lngBalSht = GLAccount.Account
      Exit Sub
   End If
   
   FormFeed
   
End Sub

Private Sub Type0NM()      ' 2280
   
   If GLAccount.AcctType = "M" Then TotalClear     ' 2290
   
   AddAmounts           ' 2300
   
   ' exit if consolidated 2310
   If GLPrint.RegBraCon = Equate.Consol And GLAccount.Account Mod 10 ^ GLCompany.SubDigits <> HiCons Then
      Exit Sub
   End If
   
   ' add the totals - 2320
   For ii = 1 To 5
      GTotal(ii) = GTotal(ii) + CYAmt
      GTotal(ii + 5) = GTotal(ii + 5) + CYSum
      GTotal(ii + 10) = GTotal(ii + 10) + PYAmt
      GTotal(ii + 15) = GTotal(ii + 15) + PYSum
   
      ' 2340 - 2350
      If GLAccount.AcctType = "M" And ii = GLAccount.TotalLevel Then
         MathFlag = True
         Exit Sub
      End If
   
   Next ii
   
'   ' rem
'   MsgBox "Acct#: " & GLAccount.Account & vbCrLf & _
'          "Tl 1 : " & GTotal(1) & vbCrLf & _
'          "Tl 6 : " & GTotal(6)
   
   ' 2370 set the flag
   If GFlag = 0 Then
      BalFlg1 = BalFlg1 + 1
   Else
      BalFlg2 = BalFlg2 + 1
   End If
   
   ' 2380
   If GLPrint.RegCmp = Equate.NonComp And CYAmt = 0 And CYSum = 0 Then
      If GFlag = 0 Then
         BalFlg1 = BalFlg1 - 1
      Else
         BalFlg2 = BalFlg2 - 1
      End If
   End If
      
   '2390
   If GLPrint.RegCmp = Equate.Comp And CYAmt = 0 And CYSum = 0 And PYAmt = 0 And PYSum = 0 Then
      If GFlag = 0 Then
         BalFlg1 = BalFlg1 - 1
      Else
         BalFlg2 = BalFlg2 - 1
      End If
   End If
   
   ' 2400
   If GFlag = 1 Then Exit Sub
   
   ' 2410
   If PrtTab = 0 Then PrtTab = 3
   
   Print0NT
   
End Sub

Private Sub Print0NT()     ' 2420 / 2540 (p&l)
       
    ' print zero balance ? 2430
    If GLPrint.PrtZeroBal = False Then
        If CYAmt = 0 And CYSum = 0 And PYAmt = 0 And PYSum = 0 Then Exit Sub
    End If
   
    ' not comparatives 2440
    If GLPrint.PrtZeroBal = False And GLPrint.RegCmp = Equate.NonComp Then
        If CYAmt = 0 And CYSum = 0 Then Exit Sub
    End If
   
    ' round dollars 2450
    If GLPrint.RoundDollars = True Then
        CYAmt = Round(CYAmt)
        CYSum = Round(CYSum)
        PYAmt = Round(PYAmt)
        PYSum = Round(PYSum)
    End If
   
    ' dollar sign 2460
    If GLAccount.DollarSign = True Then
        DollarSign = "$"
    Else
        DollarSign = " "
    End If
   
    SignTemp = SignMode
      
    ' 2480 - make negative
    If GLAccount.AcctType = "T" And GLAccount.TotalLevel = 5 And _
        GLAccount.Account > P1 Then SignTemp = -1
      
    ' 2490 - sign reversals
    If GLPrint.StaSch = Equate.Stmt And GLAccount.SignRevStmt = True Then
        SignTemp = -SignTemp
    End If
   
    If GLPrint.StaSch = Equate.Sched And GLAccount.SignRevSched = True Then
        SignTemp = -SignTemp
    End If
   
    ' 2550
    If PctValCYAmt = 0 And _
        PctValCYSum = 0 And _
        PctValPYAmt = 0 And _
        PctValPYSum = 0 Then
      
        PFormat = "q6"
    Else
        PFormat = "p6"
    End If
   
   
    ' 2510
    If GLAccount.Account > P1 Then
   
        ' 2555
        If GLPrint.RegBraCon = Equate.Budget And _
            GLAccount.AcctType = "T" And _
            GLAccount.TotalLevel = 5 Then
          
            PctValCYAmt = PctValCYAmt * SignTemp
            PctValCYSum = PctValCYSum * SignTemp
            PctValPYAmt = PctValPYAmt * SignTemp
            PctValPYSum = PctValPYSum * SignTemp
          
        End If
   
        ' 2560
        If GLPrint.SupprCP = True Then
            GLAccount.BSColumn = 2
        End If
       
        ' 2565 - 2600
        If GLAccount.BSColumn = 0 Or _
            GLAccount.BSColumn = 3 Or _
            GLAccount.BSColumn = 4 Then
       
            FormatString(1) = "x" & CStr(PrtTab)
            FormatString(2) = "a30"
            FormatString(3) = "t36"
          
            If GLPrint.RoundDollars = True Then
                FormatString(4) = "a1"
                FormatString(5) = "i14"
            Else
                FormatString(4) = "a1"
                FormatString(5) = "d14"
            End If
          
            FormatString(6) = "x1"
            FormatString(7) = PFormat
          
            FormatString(8) = FormatString(4)
            FormatString(9) = FormatString(5)
          
            FormatString(10) = FormatString(6)
            FormatString(11) = FormatString(7)
          
            If GLPrint.RegCmp = Equate.Comp Then     ' comparatives
                FormatString(12) = "x2"
                FormatString(13) = "a1"
             
                ' fix for nearest dollar comparatives
                FormatString(14) = FormatString(5)
                ' FormatString(14) = "d14"
             
                FormatString(15) = "x1"
                FormatString(16) = PFormat
                FormatString(17) = "a1"
             
                ' fix for nearest dollar comparatives
                FormatString(18) = FormatString(5)
                ' FormatString(18) = "d14"
             
                FormatString(19) = "x1"
                FormatString(20) = PFormat
                FormatString(21) = "~"
             
                PrintValue(13) = DollarSign
                PrintValue(14) = PYAmt * SignTemp
                PrintValue(16) = PctValPYAmt
                PrintValue(17) = DollarSign
                PrintValue(18) = PYSum * SignTemp
                PrintValue(20) = PctValPYSum
             
            Else
                FormatString(12) = "~"
            End If
                  
            PrintValue(2) = AcctD
            PrintValue(4) = DollarSign
            PrintValue(8) = DollarSign
                  
            Select Case GLAccount.BSColumn
            Case 0
                PrintValue(5) = CYAmt * SignTemp
                PrintValue(7) = Round(PctValCYAmt, 1)
                PrintValue(9) = CYSum * SignTemp
                PrintValue(11) = Round(PctValCYSum, 1)
            Case 3
                PrintValue(5) = CYAmt * SignTemp
                PrintValue(7) = Round(PctValCYAmt, 1)
                PrintValue(9) = PrintValue(5)
                PrintValue(11) = PrintValue(7)
            Case 4
                PrintValue(5) = CYSum * SignTemp
                PrintValue(7) = Round(PctValCYSum, 1)
                PrintValue(9) = PrintValue(5)
                PrintValue(11) = PrintValue(7)
            End Select
          
        Else   ' bc%=1 or 2
       
            FormatString(1) = "x" & CStr(PrtTab)
            FormatString(2) = "a30"
            FormatString(3) = "t36"
          
            If GLAccount.BSColumn = 1 Then
             
                If GLPrint.RoundDollars = True Then
                    FormatString(4) = "a1"
                    FormatString(5) = "i14"
                Else
                    FormatString(4) = "a1"
                    FormatString(5) = "d14"
                End If
                FormatString(6) = "x1"
                FormatString(7) = PFormat
                FormatString(8) = "x22"
                FormatString(9) = "~"
             
                PrintValue(2) = AcctD
                PrintValue(4) = DollarSign
                PrintValue(5) = CYAmt * SignTemp
                PrintValue(7) = PctValCYAmt
             
            Else    ' bc%=2 --> Suppress CP
          
                FormatString(4) = "x22"
          
                If GLPrint.RoundDollars = True Then
                    FormatString(5) = "a1"
                    FormatString(6) = "i14"
                Else
                    FormatString(5) = "a1"
                    FormatString(6) = "d14"
                End If
             
                FormatString(7) = "x1"
                FormatString(8) = PFormat
             
                PrintValue(2) = AcctD
                PrintValue(5) = DollarSign
             
                PrintValue(6) = CYSum * SignTemp
                PrintValue(8) = PctValCYSum
          
                ' comparatives ???
                If GLPrint.RegCmp <> Equate.Comp Then
                    FormatString(9) = "~"
                Else    ' format for prior year
                
                    FormatString(9) = "x24"
                    PrintValue(10) = DollarSign
                    
                    If GLPrint.RoundDollars = True Then
                        FormatString(10) = "a1"
                        FormatString(11) = "i14"
                    Else
                        FormatString(10) = "a1"
                        FormatString(11) = "d14"
                    End If
             
                    FormatString(12) = "x1"
                    FormatString(13) = PFormat
             
                    PrintValue(11) = PYSum * SignTemp
                    PrintValue(13) = PctValPYSum
          
                    FormatString(14) = "~"
          
                End If
            
            End If
       
        End If
       
    Else  ' GLAccount.Account <= P1 --> Balance Sheet
      
        ' 2515 / 2520
        If (GLAccount.AcctType = "T" And GLAccount.BSColumn <> 1) Or _
            (GLAccount.AcctType = "0" And GLAccount.BSColumn = 2) Then
         
            FormatString(1) = "x" & CStr(PrtTab)
            FormatString(2) = "a42"
            FormatString(3) = "t49"
            FormatString(4) = "x16"
            If GLPrint.RoundDollars = True Then
                FormatString(5) = "a1"
                FormatString(6) = "i14"
            Else
                FormatString(5) = "a1"
                FormatString(6) = "d14"
            End If
         
            ' 2525
            PrintValue(2) = AcctD
            PrintValue(5) = DollarSign
            PrintValue(6) = CYSum * SignTemp
         
            If GLPrint.RegCmp = Equate.Comp Then
            
                FormatString(7) = "x16"
                FormatString(8) = FormatString(5)
                FormatString(9) = FormatString(6)
                FormatString(10) = "~"
         
                PrintValue(8) = DollarSign
                PrintValue(9) = PYSum * SignTemp
         
            Else
            
                FormatString(7) = "~"
         
            End If
         
        Else     ' 2525 II
      
            FormatString(1) = "x" & CStr(PrtTab)
            FormatString(2) = "a42"
            FormatString(3) = "t49"
         
            If GLPrint.RoundDollars = True Then
                FormatString(4) = "a1"
                FormatString(5) = "i14"
            Else
                FormatString(4) = "a1"
                FormatString(5) = "d14"
            End If
         
            PrintValue(2) = AcctD
            PrintValue(4) = DollarSign
            PrintValue(5) = CYSum * SignTemp
         
            If GLPrint.RegCmp = Equate.Comp Then
            
                FormatString(6) = "x16"
                FormatString(7) = FormatString(4)
                FormatString(8) = FormatString(5)
                FormatString(9) = "~"
         
                PrintValue(7) = DollarSign
                PrintValue(8) = PYSum * SignTemp
         
            Else
            
                FormatString(6) = "~"
         
            End If
         
        End If
    End If
   
    Ln = Ln + 1
    FormatPrint
   
    LnFeeds
   
End Sub

Private Sub PrintT()     ' 2650
   
Dim TFlag As Boolean
   
   ' 2660 exit for consolidated
   If GLPrint.RegBraCon = Equate.Consol And GLAccount.Account Mod 10 ^ GLCompany.SubDigits <> HiCons Then
      Exit Sub
   End If
      
   AddAmounts     ' 2670 gosub 3490
   
   ' 2680
   If GLAccount.PrintTab = 0 Then GLAccount.PrintTab = 5
   
   ' 2690
   If GFlag Then
      TotalClear
      Exit Sub
   End If
   
   ' 2700
   Print0NT       ' gosub 2420
   
   ' 2710 - 2720
   ' 09-10-07 ????????
'   If BalFlg1 And BalFlg2 Then
'      Ln = Ln + 1
'      Prt Ln, 1, "Error in Addition"
'   End If
   
   FormatString(1) = "a25"
   FormatString(2) = "i10"
   FormatString(3) = "a10"
   FormatString(4) = "d12"
   FormatString(5) = "a10"
   FormatString(6) = "d12"
   FormatString(7) = "~"
   
   PrintValue(3) = "Upd Tot ="
   PrintValue(5) = "Acc Tot ="
   
   ' 2740 check CurrYr/CurrPd math
   TFlag = False
   If GLPrint.RoundDollars = True Then
      If CYAmt <> Round(GTotal(GLAccount.TotalLevel)) Then TFlag = True
   Else
      If CYAmt <> GTotal(GLAccount.TotalLevel) Then TFlag = True
   End If
   
   If TFlag = True Then
      
      PrintValue(1) = "error in CYR CUR PD AC#: "
      PrintValue(2) = GLAccount.Account
      PrintValue(4) = CYAmt
      PrintValue(6) = GTotal(GLAccount.TotalLevel)
      
'      Ln = Ln + 1
'
' error suppr
'      FormatPrint

   End If
   
   ' 2790 check CurrYr/YTD math
   TFlag = False
   If GLPrint.RoundDollars = True Then
      If CYSum <> Round(GTotal(GLAccount.TotalLevel + 5)) Then TFlag = True
   Else
      If CYSum <> GTotal(GLAccount.TotalLevel + 5) Then TFlag = True
   End If
   
   If TFlag = True Then
      PrintValue(1) = "Error in CYR Y.T.D. AC#: "
      PrintValue(2) = GLAccount.Account
      PrintValue(4) = CYSum
      PrintValue(6) = GTotal(GLAccount.TotalLevel + 5)
      
      
' error suppr
'
'      Ln = Ln + 1
'      FormatPrint
   
   End If
   
   ' 2850
   If GLPrint.RegCmp = Equate.Comp Then
      TotalClear
      Exit Sub
   End If
   
   ' 2840 check LastYr/CurrPd math
   TFlag = False
   If GLPrint.RoundDollars = True Then
      If PYAmt <> Round(GTotal(GLAccount.TotalLevel + 10)) Then TFlag = True
   Else
      If PYAmt <> GTotal(GLAccount.TotalLevel + 10) Then TFlag = True
   End If
   
' ???
'   If TFlag = True Then
'      PrintValue(1) = "Error in LYR Cur Pd AC#: "
'      PrintValue(2) = GLAccount.Account
'      PrintValue(4) = PYAmt
'      PrintValue(6) = GTotal(GLAccount.TotalLevel + 10)
'
'      Ln = Ln + 1
'      FormatPrint
'
'   End If
   
   ' 2900 check LastYr/YTD math
   TFlag = False
   If GLPrint.RoundDollars = True Then
      If PYSum <> Round(GTotal(GLAccount.TotalLevel + 15)) Then TFlag = True
   Else
      If PYSum <> GTotal(GLAccount.TotalLevel + 15) Then TFlag = True
   End If
   
' ???
'   If TFlag = True Then
'      PrintValue(1) = "Error in LYR Y.T.D. AC#: "
'      PrintValue(2) = GLAccount.Account
'      PrintValue(4) = PYSum
'      PrintValue(6) = GTotal(GLAccount.TotalLevel + 15)
'
'      Ln = Ln + 1
'      FormatPrint
'
'   End If
   
   TotalClear
   
End Sub


Private Function AmtFormat(ByVal xAmt As Currency, _
                           ByVal RevSign As Boolean) As String
   
   If RevSign Then xAmt = -xAmt
   
   If GLAccount.DollarSign Then
      AmtFormat = "$"
   Else
      AmtFormat = " "
   End If
   
   
End Function


Private Sub TotalClear()        ' 2950
   For ii = 1 To GLAccount.TotalLevel
      
      ' don't clear level 5 on the income statement   2970
      If GLAccount.AcctType <> "C" And _
         GLAccount.Account > GLCompany.FirstPAcct And _
         ii = 5 Then GoTo Cycle01
      
      ' hold balance sheet totals    2980
      If GLAccount.AcctType = "T" And _
         GLAccount.Account < GLCompany.FirstPAcct And _
         ii = 5 Then
         BSTotal(1) = BSTotal(1) + CYAmt
         BSTotal(2) = BSTotal(2) + CYSum
         BSTotal(3) = BSTotal(3) + PYAmt
         BSTotal(4) = BSTotal(4) + PYSum
      End If
      
      GTotal(ii) = 0
      GTotal(ii + 5) = 0
      GTotal(ii + 10) = 0
      GTotal(ii + 15) = 0
      
      BalFlg1 = 0
      BalFlg2 = 0
      
Cycle01:
   Next ii

End Sub

Private Sub AddAmounts()       ' 3490

   If GLPrint.RegBraCon <> Equate.Consol Then     ' clear amts if not consolidated  3500
      CYAmt = 0
      CYSum = 0
      PYAmt = 0
      PYSum = 0
   End If
   
   ' clear for consolidated  3510
   If GLAccount.Account Mod 10 ^ GLCompany.SubDigits = LoCons Then
      CYAmt = 0
      CYSum = 0
      PYAmt = 0
      PYSum = 0
   End If
   
   If GLPrint.RegBraCon <> Equate.Budget Then      ' not budget
      
      CYSum = CYSum + GLAccount.GetCurrAmount(1, EndPd)
      
      PYSum = PYSum + GLAccount.GetPrevAmount(1, EndPd)
      
      CYAmt = CYAmt + GLAccount.GetCurrAmount(StartPD, EndPd)
      
      PYAmt = PYAmt + GLAccount.GetPrevAmount(StartPD, EndPd)
      
   Else                                ' budget
      
      CYSum = CYSum - GLAccount.GetBudget(1, EndPd)
      
      PYSum = PYSum - GLAccount.GetBudget(1, EndPd)
                                 
      CYAmt = CYAmt + GLAccount.GetCurrAmount(StartPD, EndPd)
      
      PYAmt = PYAmt + GLAccount.GetCurrAmount(StartPD, EndPd)
      
      
   End If
   
   ' 3580 - 3590
   If GLPrint.RegBraCon = Equate.Regular And MathFlag = False Then
   Else
      If GLAccount.AcctType = "T" Then
         CYAmt = GTotal(GLAccount.TotalLevel)
         CYSum = GTotal(GLAccount.TotalLevel + 5)
         PYAmt = GTotal(GLAccount.TotalLevel + 10)
         PYSum = GTotal(GLAccount.TotalLevel + 15)
         MathFlag = False
      End If
   End If
       
   ' 3610
   If GLAccount.Account < GLCompany.FirstPAcct Then Exit Sub
   
   ' Div 3620 & 3760
   If GLPrint.RegBraCon <> Equate.Budget Then     ' regular
      PctValCYAmt = Abs(Div0(CYAmt, PctBaseCYAmt))
      PctValCYSum = Abs(Div0(CYSum, PctBaseCYSum))
      PctValPYAmt = Abs(Div0(PYAmt, PctBasePYAmt))
      PctValPYSum = Abs(Div0(PYSum, PctBasePYSum))
   Else                               ' budget
      PctValCYAmt = Div0(CYAmt, PctBaseCYAmt)
      PctValCYSum = Div0(CYAmt - CYSum, CYSum)
      PctValPYAmt = Div0(PYAmt, PctBaseCYSum)
      PctValPYSum = Div0(PYAmt - PYSum, PYSum)
   End If
   
End Sub

Public Sub Percent()    ' 3350
   
   PrevAcct = GLAccount.Account
   
   If GLAccount.Description = "" Then
      PctAcct = 0
   Else
      PctAcct = CLng(GLAccount.Description)
   End If
   
   If PctAcct = 0 Or GLAccount.GetAccount(PctAcct) = False Then
      
      ' dont' trap the error
      PctBaseCYSum = 0
      PctBasePYSum = 0
      PctBaseCYAmt = 0
      PctBasePYAmt = 0
      
      If GLAccount.GetAccount(PrevAcct) = True Then
      Else
         MsgBox "PrevAcct NF: " & PrevAcct
      End If
      
      Exit Sub
      
      ' ============================
      
      MsgBox GLAccount.Account & " Percent base account not found: " & PctAcct, vbCritical
      End
   Else
   
   End If
   
   PctBaseCYSum = GLAccount.GetCurrAmount(1, EndPd)
   
   PctBasePYSum = GLAccount.GetPrevAmount(1, EndPd)
                              
   PctBaseCYAmt = GLAccount.GetCurrAmount(StartPD, EndPd)
   
   PctBasePYAmt = GLAccount.GetPrevAmount(StartPD, EndPd)
   
   If GLAccount.GetAccount(PrevAcct) = True Then
   Else
      MsgBox "PrevAcct NF: " & PrevAcct
   End If
   
   
End Sub


Private Sub PRec()        ' 3220
   
   If P1 <> 2147483647 Then   ' not first P record
      
      If Ln > 0 Then ' ??? avoid blank first sheet for cons inc stmt
         FormFeed
      End If
      
   Else
      
      If GLPrint.RegCmp = Equate.NonComp Then     ' not comparative
         jj = 2
      Else
         jj = 4
      End If
      
      kk = 0
      
      For ii = 1 To jj
          If Round(BSTotal(ii), 2) <> 0# Then kk = 1
      Next ii
      
      If kk = 1 Then     ' balance error
         Ln = Ln + 1
         Prt Ln, 1, "Error - Balance Sheet does not balance !!!"
      End If
      
      P1 = GLAccount.Account
      PLFlag = True                  ' p1
      
      ' if not income statment only
'      If GLPrint.PrintBIB <> Equate.PrtISOnly Then LnFeeds
      If GLPrint.PrintBIB <> Equate.PrtISOnly Then FormFeed
      
      ' change font size
      If GLPrint.RegCmp = Equate.Comp Then
         If GLPrint.WidePrint = True Then
            SetFont 7, Equate.Portrait             ' compressed - portrait
         Else
            SetFont 8, Equate.LandScape            ' not compressed - landscape
         End If
      Else
'         Prvw.vsp.FontSize = 8       ' non comparative
      End If
      
   End If

End Sub

Private Sub Branches()        ' 4030

End Sub




