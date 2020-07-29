Attribute VB_Name = "modGLPrint"
Option Explicit

Dim w, X, Y, Z As String
Dim I, J, K As Integer
Dim c As Currency
Dim SString As String

Dim Pg As Integer
Dim ContFlg, PgBrk As Boolean
Dim Dgt As Integer
Dim RBC As String

Dim rs As New ADODB.Recordset

' Dtl GL

Dim Amount(11) As Currency
Dim PB As Currency
Dim YB As Currency
Dim Pg2 As Integer
   
Dim GB As Long   ' GLAccount.Date1
Dim GE As Long   ' GLAccount.Date2
   
Dim Mo, Yr As Integer

Dim GLHAmt As Currency
Dim GLHRef, GLHDes As String
Dim GLHPd, GLHJS As Integer

Dim RecCount As Long

Dim trs As New ADODB.Recordset
Dim LastAcct As Long
Dim CAcct As Long
Dim PAcct As Long
Dim LastPd As Long
Dim HFlg As Boolean
Dim PFlg As Boolean
Dim BFlg As Boolean     ' BL%
Dim jFlg As Boolean
Dim CY1 As Currency  ' CY[M1%]
Dim FirstFlg As Boolean
Dim glMsg(3) As String
Dim ct As Long

Public Sub ChartOfAccts(ByVal LoAcct As Long, _
                        ByVal HiAcct As Long, _
                        ByVal RegCons As String, _
                        ByVal Digits As Byte)
                        
   SetEquates
   PrtInit ("Port")
   SetFont 7, Equate.Portrait
                             
   If RegCons = "Reg" Then
      Prvw.Caption = GLCompany.Name & " - Chart of Accounts"
   Else
      Prvw.Caption = GLCompany.Name & " - CONSOLIDATED Chart of Accounts"
   End If
   
   frmProgress.lblMsg1 = "Printing Chart Of Accounts for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering account information ... "
   frmProgress.Show
                             
   ct = 0
                             
   If Not GLAccount.GetAcctRecSet(LoAcct, HiAcct) Then
      MsgBox "No Account records found !", vbOKOnly + vbCritical
      frmPreview.Hide
      Exit Sub
   End If
   
   Do
   
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On account " & GLAccount.Account
         frmProgress.lblMsg2.Refresh
      End If
   
      ' filters
      If InStr(1, "AEIL0", GLAccount.AcctType, vbTextCompare) = 0 Then GoTo CCycle
   
      ' consolidated filter
      If RegCons <> "Reg" Then
         If GLAccount.Account Mod 10 ^ Digits <> GLCompany.HiConsolidated Then GoTo CCycle
      End If
   
      If Ln = 0 Or Ln > MaxLines Then
         
         If Ln Then FormFeed
         
         If RegCons = "Reg" Then
            PageHeader "Chart Of Accounts", " ", "", ""
         Else
            PageHeader "CONSOLIDATED Chart Of Accounts", " ", "", ""
         End If
      
         ' data header
         Ln = Ln + 1
         
         PrintValue(1) = "Account Number"
         FormatString(1) = "r15"
         
         PrintValue(2) = " "
         FormatString(2) = "a3"
         
         PrintValue(3) = "D e s c r i p t i o n"
         FormatString(3) = "a30"
         
         PrintValue(4) = " "
         FormatString(4) = "~"
         
         FormatPrint
         
         Ln = Ln + 1
         
         PrintValue(1) = String(15, "-")
         FormatString(1) = "r15"
         
         PrintValue(2) = " "
         FormatString(2) = "t17"
         
         PrintValue(3) = String(110, "-")
         FormatString(3) = "a110"
         
         PrintValue(4) = " "
         FormatString(4) = "~"
         
         FormatPrint
      
      End If
   
      Ln = Ln + 1
   
      PrintValue(1) = " "
      FormatString(1) = "a6"
   
      If GLAccount.AcctType = "0" Then
         If RegCons = "Reg" Then
            PrintValue(2) = GLAccount.Account
         Else
            PrintValue(2) = Int(GLAccount.Account / 10 ^ Digits)
         End If
         FormatString(2) = "n9"
      Else
         PrintValue(2) = " "
         FormatString(2) = "a9"
      End If
      
      PrintValue(3) = ""
      FormatString(3) = "x3"
      
      PrintValue(4) = GLAccount.FullDesc
      FormatString(4) = "a120"
           
      PrintValue(5) = ""
      FormatString(5) = "~"
         
      FormatPrint
   
CCycle:
      If Not GLAccount.GetNextAcct Then Exit Do
   
   Loop
   
End Sub

Public Sub PrintDesc(ByVal LoAcct As Long, _
                     ByVal HiAcct As Long)
                     
                     
Dim GLDescription As New cGLDescription
                     
   SetEquates
   PrtInit ("Port")
   SetFont 7, Equate.Portrait
                             
   Ln = 0
                             
   Prvw.Caption = GLCompany.Name & " - Print Description File"
                     
   frmProgress.lblMsg1 = "Printing description list for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering description items ... "
   frmProgress.Show
   
   ct = 0
                     
   GLDescription.OpenRS
   
   Do
      
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
         frmProgress.lblMsg2.Refresh
      End If
      
      If LoAcct <> 0 And GLDescription.Number < LoAcct Then GoTo DCycle
      If HiAcct <> 0 And GLDescription.Number > HiAcct Then GoTo DCycle
   
      If Ln = 0 Or Ln > MaxLines Then
         
         If Ln Then FormFeed
         
         PageHeader "Description File", " ", "", ""
      
         ' data header
         Ln = Ln + 1
         
         PrintValue(1) = "Number"
         FormatString(1) = "r9"
         
         PrintValue(2) = " "
         FormatString(2) = "a3"
         
         PrintValue(3) = "D e s c r i p t i o n"
         FormatString(3) = "a30"
         
         PrintValue(4) = " "
         FormatString(4) = "~"
         
         FormatPrint
         
         Ln = Ln + 1
         
         PrintValue(1) = String(9, "-")
         FormatString(1) = "r9"
         
         PrintValue(2) = " "
         FormatString(2) = "t13"
         
         PrintValue(3) = String(110, "-")
         FormatString(3) = "a110"
            
         PrintValue(4) = " "
         FormatString(4) = "~"
         
         FormatPrint
      
      End If
   
      Ln = Ln + 1
   
      PrintValue(1) = GLDescription.Number
      FormatString(1) = "n9"
      
      PrintValue(2) = " "
      FormatString(2) = "t13"
      
      PrintValue(3) = GLDescription.Description
      FormatString(3) = "a110"
      
      PrintValue(4) = " "
      FormatString(4) = "~"
      
      FormatPrint
      
DCycle:
      If Not GLDescription.GetNext Then Exit Do
   Loop
                     
                     
End Sub


Public Sub PrintGLAccount(ByVal FiscalYear As Integer, _
                          ByVal StartPD As Integer, _
                          ByVal EndPd As Integer, _
                          ByVal LoAcct As Long, _
                          ByVal HiAcct As Long, _
                          ByVal LoMain As Long, _
                          ByVal HiMain As Long, _
                          ByVal LoBranch As Integer, _
                          ByVal HiBranch As Integer, _
                          ByVal Digits As Integer)
                           
   SetEquates
   PrtInit ("Port")
   SetFont 10, Equate.Portrait
                             
   Prvw.Caption = GLCompany.Name & " - Print GLAccount File"
                             
   frmProgress.lblMsg1 = "Printing GLAccount file for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering Account information ... "
   frmProgress.Show
   ct = 0
                             
   If Not GLAccount.GetAcctRecSet(LoAcct, HiAcct) Then
      frmPreview.Hide
      Exit Sub
   End If
   
   Do
   
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
         frmProgress.lblMsg2.Refresh
      End If
   
'      If HiMain <> 0 Then
'         If Int(GLAccount.Account / 10 ^ Digits) < LoMain Then GoTo GCycle
'         If Int(GLAccount.Account / 10 ^ Digits) > HiMain Then GoTo GCycle
'      End If
   
      If HiBranch <> 0 Then
         If GLAccount.Account Mod 10 ^ Digits < LoBranch Then GoTo GCycle
         If GLAccount.Account Mod 10 ^ Digits > HiBranch Then GoTo GCycle
      End If
   
      If Ln = 0 Or Ln > MaxLines Then GLAcctHdr FiscalYear, StartPD, EndPd
         
      Ln = Ln + 1
   
      PrintValue(1) = GLAccount.Account
      FormatString(1) = "n9"
   
      If GLAccount.DescNumber <> 0 Then
         PrintValue(3) = "," & GLAccount.DescNumber & " " & GLAccount.FullDesc
      Else
         PrintValue(3) = GLAccount.Description
      End If
      
      PrintValue(5) = " " & GLAccount.AcctType
      
      PrintValue(7) = " " & GLAccount.TotalLevel
      
      PrintValue(9) = " " & GLAccount.PrintTab
      
      PrintValue(11) = GLAccount.LineFeeds
      
      PrintValue(13) = " " & GLAccount.BSColumn
      
      If GLAccount.AllStatements Then
         X = "X"
      Else
         X = "-"
      End If
      
      If GLAccount.AllSchedules Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.BranchAcct Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.ConsAcct Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.TotalOnLedger Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.DollarSign Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.SignRevStmt Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      If GLAccount.SignRevSched Then
         X = X & "X"
      Else
         X = X & "-"
      End If
      
      PrintValue(15) = X
      
      PrintValue(17) = GLAccount.Date1
      
      PrintValue(19) = GLAccount.Date2
   
      PrintValue(21) = GLAmount.GetAmount(GLAccount.Account, FiscalYear, StartPD, EndPd)
      FormatString(21) = "d14"
   
      PrintValue(22) = " "
      FormatString(22) = "~"
   
      FormatPrint

GCycle:
      If Not GLAccount.GetNextAcct Then Exit Do
   
   Loop

End Sub


Private Sub GLAcctHdr(ByVal FiscalYear As Integer, ByVal StartPD As Byte, ByVal EndPd As Byte)
         
      If Ln Then FormFeed
      
      PageHeader "GLAccount File Listing", "Amounts For:", _
                 "Year: " & FiscalYear & " Periods: " & StartPD & " to: " & EndPd, ""
   
      ' data header
      Ln = Ln + 1
      
      PrintValue(1) = "Account"
      FormatString(1) = "r9"
      
      PrintValue(2) = " "
      FormatString(2) = "x1"
      
      PrintValue(3) = " "               ' desc
      FormatString(3) = "a20"
      
      PrintValue(4) = " "
      FormatString(4) = "a1"
      
      PrintValue(5) = " "               ' type
      FormatString(5) = "a4"
      
      PrintValue(6) = " "
      FormatString(6) = "a1"
      
      PrintValue(7) = "TOT"
      FormatString(7) = "a3"
      
      PrintValue(8) = " "
      FormatString(8) = "a1"
      
      PrintValue(9) = "PRT"
      FormatString(9) = "a3"
      
      PrintValue(10) = " "
      FormatString(10) = "a1"
      
      PrintValue(11) = "LN"
      FormatString(11) = "r3"
      
      PrintValue(12) = " "
      FormatString(12) = "a1"
      
      PrintValue(13) = "BS"
      FormatString(13) = "r3"
      
      PrintValue(14) = " "
      FormatString(14) = "a1"
       
      PrintValue(15) = "SSBCT$RR"
      FormatString(15) = "a8"
      
      PrintValue(16) = " "
      FormatString(16) = "a1"
       
      PrintValue(17) = "  Last"
      FormatString(17) = "a8"
      
      PrintValue(18) = " "
      FormatString(18) = "a1"
       
      PrintValue(19) = "  Curr"
      FormatString(19) = "a8"
  
      PrintValue(20) = " "
      FormatString(20) = "a1"
  
      PrintValue(21) = " "
      FormatString(21) = "a1"
      
      PrintValue(22) = " "
      FormatString(22) = "~"
  
      FormatPrint
      
      Ln = Ln + 1
      
      PrintValue(1) = "Number"
      FormatString(1) = "r9"
      
      PrintValue(2) = " "
      FormatString(2) = "a1"
      
      PrintValue(3) = "Description"               ' desc
      FormatString(3) = "a20"
      
      PrintValue(4) = " "
      FormatString(4) = "a1"
      
      PrintValue(5) = "Type"               ' type
      FormatString(5) = "a4"
      
      PrintValue(6) = " "
      FormatString(6) = "a1"
      
      PrintValue(7) = "LVL"
      FormatString(7) = "a3"
      
      PrintValue(8) = " "
      FormatString(8) = "a1"
      
      PrintValue(9) = "TAB"
      FormatString(9) = "a3"
      
      PrintValue(10) = " "
      FormatString(10) = "a1"
      
      PrintValue(11) = "FDS"
      FormatString(11) = "a3"
      
      PrintValue(12) = " "
      FormatString(12) = "a1"
      
      PrintValue(13) = "COL"
      FormatString(13) = "a3"
      
      PrintValue(14) = " "
      FormatString(14) = "a1"
       
      PrintValue(15) = "TCROL$12"
      FormatString(15) = "a8"
      
      PrintValue(16) = " "
      FormatString(16) = "a1"
       
      PrintValue(17) = "  Date"
      FormatString(17) = "a8"
      
      PrintValue(18) = " "
      FormatString(18) = "a1"
       
      PrintValue(19) = "  Date"
      FormatString(19) = "a8"
      
      PrintValue(20) = " "
      FormatString(20) = "a1"
      
      PrintValue(21) = "     Amount"
      FormatString(21) = "a14"
      
      PrintValue(22) = " "
      FormatString(22) = "~"
      
      FormatPrint
      Ln = Ln + 1
      
      X = String(30, "-")
      
      PrintValue(1) = X
      FormatString(1) = "r9"
      
      PrintValue(2) = " "
      FormatString(2) = "a1"
      
      PrintValue(3) = X               ' desc
      FormatString(3) = "a20"
      
      PrintValue(4) = " "
      FormatString(4) = "a1"
      
      PrintValue(5) = X               ' type
      FormatString(5) = "a4"
      
      PrintValue(6) = " "
      FormatString(6) = "a1"
      
      PrintValue(7) = X
      FormatString(7) = "a3"
      
      PrintValue(8) = " "
      FormatString(8) = "a1"
      
      PrintValue(9) = X
      FormatString(9) = "a3"
      
      PrintValue(10) = " "
      FormatString(10) = "a1"
      
      PrintValue(11) = X
      FormatString(11) = "a3"
      
      PrintValue(12) = " "
      FormatString(12) = "a1"
      
      PrintValue(13) = X
      FormatString(13) = "a3"
      
      PrintValue(14) = " "
      FormatString(14) = "a1"
       
      PrintValue(15) = X
      FormatString(15) = "a8"
      
      PrintValue(16) = " "
      FormatString(16) = "a1"
       
      PrintValue(17) = X
      FormatString(17) = "a8"
      
      PrintValue(18) = " "
      FormatString(18) = "a1"
       
      PrintValue(19) = X
      FormatString(19) = "a8"
      
      PrintValue(20) = " "
      FormatString(20) = "a1"
      
      PrintValue(21) = X
      FormatString(21) = "a14"
      
      PrintValue(22) = " "
      FormatString(22) = "~"
      
      FormatPrint
      
End Sub

Public Sub DetailGL(ByVal RegBraCon As String, _
                    ByVal FiscalYear As Integer, _
                    ByVal StartPD As Integer, _
                    ByVal EndPd As Integer, _
                    ByVal LoAcct As Long, _
                    ByVal HiAcct As Long, _
                    ByVal LoCons As Long, _
                    ByVal HiCons As Long, _
                    ByVal LoBranch As Integer, _
                    ByVal HiBranch As Integer, _
                    ByVal Digits As Integer, _
                    ByVal PgBreak As Boolean, _
                    ByVal PrtZero As Boolean, _
                    ByVal CompanyID As Long)


   PgBrk = PgBreak
   ContFlg = False
   Dgt = Digits
   RBC = RegBraCon

   ' init variables
   LastAcct = 0
   LastPd = 0
   PFlg = False
   
   ' set up temp record set
   trs.CursorLocation = adUseClient
   
   trs.Fields.Append "Account", adDouble
   trs.Fields.Append "BaseNum", adDouble
   trs.Fields.Append "Period", adInteger
   trs.Fields.Append "Amount", adCurrency
   trs.Fields.Append "JS", adInteger
   trs.Fields.Append "ID", adDouble
   trs.Fields.Append "Reference", adVarChar, 20, adFldIsNullable
   trs.Fields.Append "Description", adVarChar, 20, adFldIsNullable
   trs.Fields.Append "Type", adVarChar, 1, adFldIsNullable
   trs.Fields.Append "PostDate", adDate
   
   trs.Open , , adOpenDynamic, adLockOptimistic
   
   SetEquates
   PrtInit ("Port")
   SetFont 9, Equate.Portrait
                             
   Prvw.Caption = GLCompany.Name & " - Print Detail GL"
                             
   ' progress form init
   frmProgress.Show
   frmProgress.lblMsg1 = GLCompany.Name & " - Print Detail GL"
   frmProgress.lblMsg2 = "Gathering account and detail information .... "
   frmProgress.Refresh
                             
   ' header strings
   GetDates (CompanyID)
   glMsg(1) = "Fiscal Year: " & FiscalYear & " " & _
              " Period #: " & StartPD & " To: # " & EndPd & "   -   " & _
              Format(CurrYrCurrPdBeg, "mm/dd/yyyy") & " To: " & _
              Format(CurrYrPdEnd, "mm/dd/yyyy")
   I = 2
   
   If LoAcct <> 0 Then
      glMsg(2) = "Account Number From: " & LoAcct & " To: " & HiAcct
      I = 3
   End If
   
   If RegBraCon <> "Reg" Then
      glMsg(I) = ""
      If LoCons <> 0 Then
         glMsg(I) = "Cons From: " & LoCons & " To: " & HiCons
      End If
      If LoBranch <> 0 Then
         glMsg(I) = glMsg(I) & "  Branch From: " & LoBranch & " To: " & HiBranch
      End If
   End If
   
   If Not GLAccount.GetRecordSetsNoBudget(FiscalYear, FiscalYear) Then
      frmPreview.Hide
      Exit Sub
   End If
   
   frmProgress.lblMsg2 = "Opening History ...."
   frmProgress.Refresh
   
   SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
               " AND Period >= " & StartPD & " AND Period <= " & EndPd & " AND HisType <> 'B'"
   
   If Not GLHistory.GetByString(SString) Then
      MsgBox "No History Found !!!", vbOKOnly + vbCritical, "Detail GL Print"
      GLAccount.CloseRS
      GLDescription.CloseRS
      Exit Sub
   End If

   ' loop thru the history record set and make sort string in temp rs
   Do
      
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "Creating Sort File: " & _
                               GLHistory.FiscalYear & " " & _
                               GLHistory.Period & " " & _
                               "Records Added: " & Format(ct, "##,###,##0")
         frmProgress.Refresh
      
      End If
      
      If LoAcct <> 0 Then
         If GLHistory.Account < LoAcct Then GoTo NextHist
         If GLHistory.Account > HiAcct Then GoTo NextHist
      End If
      
      ' takes care of filter lines 1800 and 1810
      
      If RegBraCon = "Bra" Then
         If GLHistory.Account Mod 10 ^ Digits < LoBranch Then GoTo NextHist
         If GLHistory.Account Mod 10 ^ Digits > HiBranch Then GoTo NextHist
      End If
      
      ' filter by glh type
' >>>>>>>>>>><<<<<<<<<<<<
'      If InStr(1, "NORTW", GLHistory.HisType, vbTextCompare) Then GoTo NextHist
' >>>>>>>>>>><<<<<<<<<<<<
       
      trs.AddNew Array("Account", _
                       "BaseNum", _
                       "Period", _
                       "Amount", _
                       "JS", _
                       "ID", _
                       "Reference", _
                       "Description", _
                       "PostDate"), _
                 Array(GLHistory.Account, _
                       Int(GLHistory.Account / 10 ^ Digits), _
                       GLHistory.Period, _
                       GLHistory.Amount, _
                       GLHistory.JournalSource, _
                       GLHistory.ID, _
                       GLHistory.Reference, _
                       GLHistory.Description, _
                       GLHistory.PostDate)

      RecCount = RecCount + 1

NextHist:
      If Not GLHistory.GetNext Then Exit Do
   
   Loop
   
   GLHistory.CloseRS
   
   frmProgress.lblMsg2 = "Now Sorting ... "
   frmProgress.Refresh
   
   If RegBraCon <> "Bra" Then
      trs.Sort = "Account,Period,JS,PostDate"
   Else
      trs.Sort = "BaseNum,Reference,Period,JS,PostDate"
   End If
   
   ' 1440
   If RegBraCon = "Bra" Then
      LoCons = LoBranch
      HiCons = LoBranch
      BFlg = False
   End If
   Ln = 0
   Pg = 0
   
   ' 1460 - Branch
NextBranch:
   
   LastAcct = 0
   Amount(11) = 0
   PB = 0
   YB = 0
   FirstFlg = True
   
   ' 1480
   If RegBraCon <> "Bra" Then
      Ln = 0
      Pg = 0
   End If
   
   ' 1670
   PFlg = False
   Ln = 0
   
   ' 1680
   If RegBraCon <> "Bra" Or (RegBraCon = "Bra" And Not BFlg) Then
      DtlGLHeader
   End If
   
   ct = 0
   
   ' loop thru GLAccount
   Do
      
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         If RegBraCon <> "Bra" Then
            frmProgress.lblMsg2 = "Now Printing: " & GLAccount.Account
         Else
            frmProgress.lblMsg2 = "Now Printing: " & GLAccount.Account & " Branch: " & LoCons
         End If
         frmProgress.Refresh
      End If
      
      ' filters 1780
      If LoAcct <> 0 Then
         If GLAccount.Account < LoAcct Then GoTo NextAcct
         If GLAccount.Account > HiAcct Then GoTo NextAcct
      End If
      
      If RegBraCon = "Bra" And GLAccount.Account Mod 10 ^ Digits < LoCons Then GoTo NextAcct
      If RegBraCon = "Bra" And GLAccount.Account Mod 10 ^ Digits > HiCons Then GoTo NextAcct
      
      ' 1820
      If RegBraCon = "Con" And LastAcct <> 0 Then
         If Int(GLAccount.Account / 10 ^ Digits) <> Int(LastAcct / 10 ^ Digits) Then
            GoTo FirstSort
         End If
      End If
               
      ' 1860 - type filter
      If InStr(1, "0NPT", GLAccount.AcctType, vbTextCompare) = 0 Then GoTo NextAcct
      
      ' assign variables from GLAccount
      GB = GLAccount.Date1
      GE = GLAccount.Date2
      
      CY1 = GLAccount.GetCurrAmount(StartPD, StartPD)
      LastAcct = GLAccount.Account
      
      ' 1910 ....
      If GLAccount.AcctType = "P" Then
         If RegBraCon <> "Bra" And PFlg = False Then DtlGLHeader
         PFlg = True
         GoTo NextAcct
      ElseIf GLAccount.AcctType = "T" Then
         If RegBraCon = "Con" Then GoTo NextAcct
         If Not GLAccount.TotalOnLedger Then GoTo NextAcct
      End If
      
      ' 2110
      If RegBraCon <> "Con" Then
         PB = 0
         YB = 0
      End If
      
      ' 2120 prev bal
      If StartPD <> 1 Then
         PB = GLAccount.GetCurrAmount(1, StartPD - 1)
      End If
      
      ' 2160
      Amount(3) = PB
      Amount(6) = PB
      Amount(10) = PB
      
      ' 2170 YTD Bal
      YB = GLAccount.GetCurrAmount(1, EndPd)
      
      ' ~ 2220
FirstSort:
      
'      If FirstFlg = True Then       ' do a find for the first time through
      
      ' fix if first account has no detail 9/21/07
      If FirstFlg = True Or HFlg = False Then     ' do a find for the first time through
      
         If InStr(1, "NT", GLAccount.AcctType) = 0 Then
            ' see if history exists
            If RegBraCon <> "Con" Then
               X = "Account = " & GLAccount.Account
            Else
               X = "BaseNum = " & Int(GLAccount.Account / 10 ^ Digits)
            End If
            trs.Find X, 0, adSearchForward, 1
            
            If trs.EOF Then
               HFlg = False
            Else
               HFlg = True
            End If
         
         End If
      
      ElseIf trs.EOF = True Then
      
         HFlg = False
      
      Else                          ' find next if not first through
         
         If InStr(1, "NT", GLAccount.AcctType) = 0 Then
            If RegBraCon <> "Con" Then
               If trs!Account = GLAccount.Account Then
                  HFlg = True
               Else
                  HFlg = False
               End If
            Else
               If trs!BaseNum = Int(GLAccount.Account / 10 ^ Digits) Then
                  HFlg = True
               Else
                  HFlg = False
               End If
            End If
         End If
            
      End If
      
      ' 2410
      jFlg = True
      
      If RegBraCon <> "Con" Then jFlg = False
      If GLAccount.AcctType = "N" Then jFlg = False
      If PB <> 0 Then jFlg = False
      If GLAccount.GetCurrAmount(StartPD, StartPD) <> 0 Then jFlg = False
      If HFlg = True Then jFlg = False
      If GE >= GLAccount.Date1 Then jFlg = False
      
      If jFlg Then
         DtlGLClear
         If CAcct = 0 Then Exit Do
         GoTo Cycle
      End If
   
      ' skip zero bal and no activity accounts
      If PrtZero = False And HFlg = False Then
         If CY1 = 0 And PB = 0 And YB = 0 Then GoTo NextAcct
      End If
   
      FirstFlg = False
   
      ' 2430
      If Ln > MaxLines Then DtlGLHeader
      
      ' 2440 account header line
      Let X = "Acct # "
      If RegBraCon = "Con" Then
         X = X & Int(LastAcct / 10 ^ Digits)
      Else
         X = X & LastAcct
      End If
      
      PrintValue(1) = X
      FormatString(1) = "a20"
      
      PrintValue(2) = " "
      FormatString(2) = "a7"
      
      PrintValue(3) = GLAccount.FullDesc
      FormatString(3) = "a30"
      
      PrintValue(4) = " "
      FormatString(4) = "a18"
      
      PrintValue(5) = PB
      FormatString(5) = "d14"
      
      PrintValue(6) = " "
      FormatString(6) = "a2"
      
      PrintValue(7) = GB Mod (100) & "/" & Int(GB / 100)
      FormatString(7) = "r7"
      
      PrintValue(8) = " "
      FormatString(8) = "~"
      
      FormatPrint
      
      Ln = Ln + 1
            
      ' 2530 - Print N record
      If GLAccount.AcctType = "N" Then
         
         ' 2550
         If Ln > MaxLines - 3 - EndPd + StartPD Then DtlGLHeader
         
         ' 2560
         For I = StartPD To EndPd
         
             ' convert to mm/yyyy
             Mo = ((GLCompany.FirstPeriod + I - 1) Mod GLCompany.NumberPds)
             If I <= GLCompany.NumberPds - GLCompany.FirstPeriod + 1 Then
                Yr = FiscalYear
             Else
                Yr = FiscalYear - 1
             End If
             X = Format(Mo, "0#") & "/" & Format(Yr, "####")
    
             ' 2580
             c = GLAccount.GetCurrAmount(I, I)
             If c >= 0 Then
                Amount(4) = Amount(4) + c
             Else
                Amount(5) = Amount(5) + c
             End If
             Amount(6) = Amount(6) + c
             
             PrintValue(1) = X
             FormatString(1) = "a8"
             
             PrintValue(2) = "*"
             FormatString(2) = "a2"
             
             PrintValue(3) = "Net Acct."
             FormatString(3) = "a10"
             
             PrintValue(4) = "Current PR-LOSS"
             FormatString(4) = "a22"
             
             PrintValue(5) = Abs(c)
             FormatString(5) = "d14"
             
             If c >= 0 Then
                PrintValue(6) = " "
                FormatString(6) = "a16"
                w = 7
             Else
                w = 6
             End If
             
             PrintValue(w) = Amount(6)
             FormatString(w) = "d14"
             
             PrintValue(w + 1) = " "
             FormatString(w + 1) = "~"
             
             FormatPrint
             Ln = Ln + 1
         Next I
         
         DtlGLTotal
         
         GoTo NextAcct
      
      End If     ' if GLAccount.AcctType = "N"
      
      LastPd = 0
      
'      ' 2650 first detail record
'      If RegBraCon = "Reg" Then
'         x = "Account = " & GLAccount.Account
'      Else
'         x = "BaseNum = " & Int(GLAccount.Account / 10 ^ Digits)
'      End If
'
'      trs.Find x, 0, adSearchForward, 1
'
'      If trs.EOF Then GoTo GLGLTL
          
      If HFlg = False Then GoTo GLGLTL
          
      Do
   
         ' 2750
         If RegBraCon <> "Con" And trs!Account <> GLAccount.Account Then Exit Do
         
         ' 2760
         If RegBraCon = "Con" And trs!BaseNum <> Int(GLAccount.Account / 10 ^ Digits) Then Exit Do
         
         ' 2790 - type filter performed during temp record set creation
         
         ' >>>>>>>> assign fields to variables
         GLHAmt = trs!Amount
         GLHRef = trs!Reference
         GLHDes = trs!Description
         GLHPd = trs!Period
         GLHJS = trs!JS
          
         ' update totals 2830
         For I = 1 To 3
             If GLHAmt >= 0 Then
                Amount(I * 3 - 2) = Amount(I * 3 - 2) + GLHAmt
             Else
                Amount(I * 3 - 1) = Amount(I * 3 - 1) + GLHAmt
             End If
             Amount(I * 3) = Amount(I * 3) + GLHAmt
         Next I
         
         ' >>>>>>>> set up print line from variables
         
         ' convert to mm/yyyy
' *** need fix for 13 period companies ***
         Mo = ((GLCompany.FirstPeriod + GLHPd - 1) Mod GLCompany.NumberPds)
         If Mo = 0 Then Mo = GLCompany.NumberPds
         
         If GLCompany.FirstPeriod = 1 Then
            Yr = FiscalYear
         Else
            If GLHPd <= GLCompany.NumberPds - GLCompany.FirstPeriod + 1 Then
               Yr = FiscalYear - 1
            Else
               Yr = FiscalYear
            End If
         End If
         X = Format(Mo, "0#") & "/" & Format(Yr, "####")
         
'    MsgBox (GLHPd & " " & GLAccount.Account & " " & Mo & " " & i & " " & Yr & " " & X)
         
         PrintValue(1) = X
         FormatString(1) = "a8"
         
         X = Format(GLHJS, "#")
         PrintValue(2) = X
         FormatString(2) = "a2"
         
         PrintValue(3) = GLHRef
         FormatString(3) = "a10"
         
         PrintValue(4) = GLHDes
         FormatString(4) = "a20"
         
         PrintValue(5) = " "
         FormatString(5) = "a2"
         
         If GLHAmt >= 0 Then

            PrintValue(6) = Abs(GLHAmt)
            FormatString(6) = "d14"
            
            PrintValue(7) = " "
            FormatString(7) = "a16"
         
         Else

            PrintValue(6) = " "
            FormatString(6) = "a16"
            
            PrintValue(7) = Abs(GLHAmt)
            FormatString(7) = "d14"
            
         End If
         
         I = 7
         
         trs.MoveNext
         
         ' print the current line and exit if at end of temp rec set
         If trs.EOF Then
            I = I + 1
            PrintValue(I) = " "
            FormatString(I) = "~"
            FormatPrint
            Ln = Ln + 1
            Exit Do
         End If
         
         ' >>>>>>>> if break in period add mth subtl to format line
         If trs!Period <> GLHPd Then
            
            I = I + 1
            PrintValue(I) = " "
            FormatString(I) = "a3"
            
            I = I + 1
            PrintValue(I) = Amount(3)
            FormatString(I) = "d14"
         
         End If
         
         I = I + 1
         PrintValue(I) = " "
         FormatString(I) = "~"
         
         FormatPrint
         Ln = Ln + 1
         ' 2940
         Amount(1) = 0
         Amount(2) = 0
         If Ln > MaxLines Then
            ContFlg = True
            FormFeed
            DtlGLHeader
         End If
         
      Loop
   
GLGLTL:
      DtlGLTotal
   
NextAcct:
      If Not GLAccount.GetNext Then Exit Do
      
Cycle:
   Loop
   
   
   ' grand totals 3550
   
   ' next branch
   If RegBraCon = "Bra" And LoCons < HiBranch Then
      LoCons = LoCons + 1
      HiCons = LoCons
      GLAccount.FindFirst   ' go back to the first GLAccount record
      GoTo NextBranch
   End If

   Ln = Ln + 1
   
   ' ============================================================================================
   
   PrintValue(1) = " "
   FormatString(1) = "a21"
   
   PrintValue(2) = "Detail Total "
   FormatString(2) = "a21"
   
   For I = 1 To 3
       
       PrintValue(I * 2 + 1) = Abs(Amount(I + 6))
       FormatString(I * 2 + 1) = "d14"
   
       PrintValue(I * 2 + 2) = " "
       FormatString(I * 2 + 2) = "a2"
   
   Next I

   PrintValue(9) = " "
   FormatString(9) = "~"
   
   FormatPrint
   Ln = Ln + 1
   ' ============================================================================================
   
   If LoAcct = 0 And Amount(9) <> 0 Then
      PrintValue(1) = " ***** ERROR - Detail does not balance !!! *****"
      FormatString(1) = "a60"
      
      PrintValue(2) = " "
      FormatString(2) = "~"
      
      FormatPrint
      Ln = Ln + 1
   End If
   
   ' ============================================================================================
   
   PrintValue(1) = " "
   FormatString(1) = "a21"
   
   PrintValue(2) = "Total of GLM Balances"
   FormatString(2) = "a37"
   
   PrintValue(3) = Amount(11)
   FormatString(3) = "d11"
   
   PrintValue(4) = " "
   FormatString(4) = "~"
   
   FormatPrint
   Ln = Ln + 1
   ' ============================================================================================
      
   If LoAcct = 0 And Amount(11) <> 0 Then

      PrintValue(1) = " ***** ERROR - Total of GLM Balances does not balance !!! *****"
      FormatString(1) = "a60"
      
      PrintValue(2) = " "
      FormatString(2) = "~"
      
      FormatPrint
      Ln = Ln + 1
   End If

'   frmProgress.Visible = False
'   Unload frmProgress

End Sub


Private Sub DtlGLTotal()        ' 2970

Dim BalFlg As Boolean

    ' 2980
    If RBC <> "Con" And GLAccount.AcctType = "T" Then
       Amount(6) = YB
       Amount(10) = YB
    End If
    
    ' 3000 - print underline -------------
    PrintValue(1) = " "
    FormatString(1) = "a38"
    
    PrintValue(2) = String(60, "-")
    FormatString(2) = "a60"
    
    PrintValue(3) = " "
    FormatString(3) = "~"
    
    FormatPrint
    Ln = Ln + 1
        
    ' 3020 - balance check
    If RBC <> "Bra" And Amount(4) + Amount(5) + Amount(10) <> YB Then
       BalFlg = False
       Y = " ERROR"
    Else
       BalFlg = True
       Y = " "
    End If
    
    ' 3030 - 3040
    If RBC = "Bra" Then
       Amount(6) = YB
       BalFlg = True
    End If
       
    ' print the total line
    PrintValue(1) = " "
    FormatString(1) = "a22"
        
    PrintValue(2) = "Account Total"
    FormatString(2) = "a20"
    
    PrintValue(3) = Abs(Amount(4))
    FormatString(3) = "d14"
    
    PrintValue(4) = " "
    FormatString(4) = "a2"
    
    PrintValue(5) = Abs(Amount(5))
    FormatString(5) = "d14"
    
    PrintValue(6) = " "
    FormatString(6) = "a3"
    
    PrintValue(7) = Abs(Amount(6))
    PrintValue(7) = Amount(6)
    FormatString(7) = "d14"
    
    PrintValue(8) = Y
    FormatString(8) = "a6"
        
    PrintValue(9) = " "
    FormatString(9) = "~"
       
    FormatPrint
    Ln = Ln + 1
    
    If BalFlg = False Then
       
       PrintValue(1) = "DETAIL TOTAL DOES NOT EQUAL CONROL TOTAL.  CONTROL TOTAL ="
       FormatString(1) = "a75"
       
       PrintValue(2) = YB
       FormatString(2) = "d11"
       
       PrintValue(3) = " "
       FormatString(3) = "~"
       
       FormatPrint
       Ln = Ln + 1
       
    End If
    
    ' 3110
    Ln = Ln + 1
    
    ' 3130
    If RBC <> "Bra" And GLAccount.AcctType = "T" Then
       PrintValue(1) = String(79, "-")
       FormatString(1) = "a79"
       
       PrintValue(2) = " "
       FormatString(2) = "~"
       
       FormatPrint
       Ln = Ln + 1
    End If
    
    ' 3140 - 3150
    Ln = Ln + 2
    
    ' 3170
    If (PgBrk = True And Ln <= MaxLines) Or (Ln > MaxLines - 5) Then
       ContFlg = False
       FormFeed
       DtlGLHeader
    End If

    DtlGLClear

End Sub

Private Sub DtlGLClear()
    
    ' amount(11) - summ of PB
    ' 3200
    If RBC <> "Con" And GLAccount.AcctType = "0" Then
       Amount(11) = Amount(11) + Amount(6)
    End If
    
    ' 3210
    If RBC = "Con" And GLAccount.AcctType <> "N" Then
       Amount(11) = Amount(11) + Amount(6)
    End If
    
    Amount(4) = 0
    Amount(5) = 0
    Amount(6) = 0
    Amount(10) = 0
    
    ' cy[m1%] = 0 ???
    CY1 = 0
    
    GB = 0
    GE = 0
    PB = 0
    YB = 0
    
    ' get next consolidated GLAccount record
    If RBC <> "Con" Then Exit Sub
    
    CAcct = Int(LastAcct / 10 ^ Dgt) + 1
    
    Do Until Int(GLAccount.Account / 10 ^ Dgt) >= CAcct
       If GLAccount.GetNext = False Then
          CAcct = 0
          Exit Do
       End If
    Loop
    
End Sub


Private Sub DtlGLHeader()
      
      If Ln Then FormFeed
       
      X = "Detail General Ledger "
      If RBC = "Bra" Then X = X & "- Branch"
      If RBC = "Con" Then X = X & "- Consolidated"
      
      PageHeader X, glMsg(1), glMsg(2), glMsg(3)
   
      ' data header
      Ln = Ln + 1
      
      X = "        J                                       DEBIT           CREDIT            ACCOUNT     PRIOR"
      PrintValue(1) = X
      FormatString(1) = "a100"
      PrintValue(2) = " "
      FormatString(2) = "~"
      FormatPrint
      Ln = Ln + 1
      
      X = "DATE    S REF. #          DESCRIPTION          POSTINGS        POSTINGS           BALANCE      DATE"
      PrintValue(1) = X
      FormatString(1) = "a100"
      PrintValue(2) = " "
      FormatString(2) = "~"
      FormatPrint
      Ln = Ln + 1
      
      If ContFlg = False Then
         Ln = Ln + 1
         Exit Sub
      End If
      
      If RBC <> "Con" Then
         X = "Acct # " & GLAccount.Account & "   (Continued)"
      Else
         X = "Acct # " & Int(GLAccount.Account / 10 ^ Dgt) & "   (Continued)"
      End If

      PrintValue(1) = X
      FormatString(1) = "a60"
      
      PrintValue(2) = " "
      FormatString(2) = "~"
      
      FormatPrint
      Ln = Ln + 1

End Sub


Public Sub GLHistJnl(ByVal FiscalYear As Long, _
                     ByVal StartPD As Integer, _
                     ByVal EndPd As Integer, _
                     ByVal JS As Integer, _
                     ByVal Batch As Long, _
                     ByVal IncludeAcctDesc As Boolean)
                     
Dim TlDebits, TlCredits As Currency
Dim LastJS As Byte
Dim EMsg As String
Dim HashTotal As Long
Dim EntryCount As Long

   On Error GoTo ErrMsg

   SetEquates
   PrtInit ("Port")
   SetFont 8, Equate.Portrait
                              
   EMsg = "01"
    
   Prvw.Caption = GLCompany.Name & " - Data Entry Journal"
      
   EMsg = "02"
    
   frmProgress.lblMsg1 = "Printing Data Entry Journal for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering history information .... "
   frmProgress.Show
   ct = 0
    
   If Batch <> 0 Then
      
      If Not (GLBatch.GetBatch(Batch)) Then
         MsgBox "Batch info not found !!!", vbCritical + vbOKOnly, "GL Data Entry Journal"
         End
      End If
      
      If GLUser.GetBySQL("SELECT * FROM Users WHERE Users.ID = " & GLBatch.UpdateUser) Then
         glMsg(1) = glMsg(1) & " - Batch #: " & Batch & " User: " & GLUser.Name
      Else
         glMsg(1) = glMsg(1) & " - Batch #: " & Batch
      End If
         
      FiscalYear = GLBatch.FiscalYear
      StartPD = GLBatch.Period
      EndPd = GLBatch.Period
      JS = GLBatch.JournalSource
   
   End If
   
   glMsg(1) = "Fiscal Year: " & FiscalYear & " For Period #: " & StartPD
   If StartPD <> EndPd Then
      glMsg(1) = glMsg(1) & " To: " & EndPd
   End If
   
   EMsg = "03"
    
   ' set up the SQL string for History
   ' 2020-07-29 - order change
   If Batch <> 0 Then
      
      SString = "SELECT * FROM GLHistory WHERE BatchNumber = " & Batch & _
                " ORDER BY JournalSource, PostDate"
      
      SString = "SELECT * FROM GLHistory WHERE BatchNumber = " & Batch & _
                " ORDER BY JournalSource, ID"
   Else
      
      If JS = 0 Then    ' all journal sources
         
         If frmGLPrint.chkBudget = 0 Then
         
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND HisType <> 'B'" & _
                      " ORDER BY JournalSource, PostDate"
            
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND HisType <> 'B'" & _
                      " ORDER BY JournalSource, ID"
         Else
            
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND HisType = 'B'" & _
                      " ORDER BY JournalSource, PostDate"
      
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND HisType = 'B'" & _
                      " ORDER BY JournalSource, ID"
      
         End If
      
      Else              ' single journal
         
         If frmGLPrint.chkBudget = 0 Then
            
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND JournalSource = " & JS & _
                      " AND HisType <> 'B'" & _
                      " ORDER BY PostDate"
            
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND JournalSource = " & JS & _
                      " AND HisType <> 'B'" & _
                      " ORDER BY ID"
         Else
            
            SString = "SELECT * FROM GLHistory WHERE FiscalYear = " & FiscalYear & _
                      " AND Period >= " & StartPD & _
                      " AND Period <= " & EndPd & _
                      " AND JournalSource = " & JS & _
                      " AND HisType = 'B'" & _
                      " ORDER BY ID"
         
         End If
      
      End If
   
   End If
                     
   EMsg = "04"
    
   If Not GLHistory.GetByString(SString) Then
      Response = False
      MsgBox "No history found for the ranges given !", vbOKOnly + vbCritical, "GL History Journal"
      Exit Sub
   Else
      Response = True
   End If
   
   EMsg = "05"
    
   Do
   
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On History Record #: " & Format(ct, "##,###,##0")
         frmProgress.lblMsg2.Refresh
      End If
   
      If Ln = 0 Or Ln > MaxLines Then
         HistJnlNextPage
      End If
      
      ' print data line
      
      ' filter by glh type
' >>>>>>>>>>><<<<<<<<<<<<
'      If InStr(1, "NORTW", GLHistory.HisType, vbTextCompare) Then GoTo NextJnl
' >>>>>>>>>>><<<<<<<<<<<<
      
      EMsg = "07"
    
      X = CStr(GLHistory.JournalSource)
      I = 3 - Len(X)
      PrintValue(1) = Space(I) & X
      FormatString(1) = "a3"
        
      EMsg = "08"
    
      X = CStr(GLHistory.SourceCode)
      I = 3 - Len(X)
      PrintValue(2) = Space(I) & X
      FormatString(2) = "a3"
      
      EMsg = "09"
    
      X = GLHistory.HisType
      I = 3 - Len(X)
      PrintValue(3) = Space(I) & GLHistory.HisType
      FormatString(3) = "a3"
         
      EMsg = "10"
    
      ' convert to mm/yyyy
      Mo = ((GLCompany.FirstPeriod + GLHistory.Period - 1) Mod GLCompany.NumberPds)
      If Mo = 0 Then Mo = GLCompany.NumberPds
      If GLCompany.FirstPeriod <> 1 And GLHistory.Period <= GLCompany.NumberPds - GLCompany.FirstPeriod + 1 Then
         Yr = FiscalYear - 1
      Else
         Yr = FiscalYear
      End If
      X = Space(2) & Format(Mo, "0#") & "/" & Format(Yr, "####")
      PrintValue(4) = X
      FormatString(4) = "a10"
         
      EMsg = "11"
    
      X = Format(GLHistory.Account, "########0")
      I = 9 - Len(X) + 1
      PrintValue(5) = Space(I) & X
      FormatString(5) = "a11"
         
      EMsg = "12"
    
      PrintValue(6) = " " & GLHistory.Reference
      FormatString(6) = "a22"
         
      EMsg = "13"
    
      PrintValue(7) = "  " & GLHistory.Description & "  "
      FormatString(7) = "a24"
         
      EMsg = "14"
    
      If GLHistory.Amount >= 0 Then
         
         PrintValue(8) = GLHistory.Amount
         FormatString(8) = "d14"
         
         PrintValue(9) = " "
         FormatString(9) = "a16"
      
         TlDebits = TlDebits + GLHistory.Amount
      
      Else
      
         PrintValue(8) = " "
         FormatString(8) = "a16"
         
         PrintValue(9) = Abs(GLHistory.Amount)
         FormatString(9) = "d14"
      
         TlCredits = TlCredits + GLHistory.Amount
      
      End If
         
      ' update hash total
      HashTotal = (HashTotal + GLHistory.Account) Mod 10 ^ 9
         
      ' update entry count
      EntryCount = EntryCount + 1
         
      PrintValue(10) = " "
      FormatString(10) = "~"
         
      FormatPrint
      
      Ln = Ln + 1
      
      ' print acct description
      If IncludeAcctDesc Then
         If GLAccount.GetAcctRecSet(GLHistory.Account, GLHistory.Account) Then
         
            PrintValue(1) = ""
            FormatString(1) = "a24"
         
            PrintValue(2) = GLAccount.GetDesc
            FormatString(2) = "a30"
         
            PrintValue(3) = " "
            FormatString(3) = "~"
         
            FormatPrint
                     
            PrintValue(1) = " "
            FormatString(1) = "a1"
            
            PrintValue(2) = " "
            FormatString(2) = "~"
           
            FormatPrint
         
            Ln = Ln + 2
        
         End If
      End If
      
      LastJS = GLHistory.JournalSource
      
      If Not GLHistory.GetNext Then
         jFlg = False
      Else
         jFlg = True
      End If
      
      ' print total
      If Not jFlg Or GLHistory.JournalSource <> LastJS Then
         
         If Ln > MaxLines - 5 Then
            HistJnlNextPage
         End If
         
         PrintValue(1) = " "
         FormatString(1) = "a75"
         
         PrintValue(2) = String(14, "-") & "  " & String(14, "-")
         FormatString(2) = "a30"
         
         PrintValue(3) = " "
         FormatString(3) = "~"
         
         FormatPrint
         Ln = Ln + 1
         
         
         PrintValue(1) = " T O T A L   D E B I T S"
         FormatString(1) = "a76"
         
         PrintValue(2) = TlDebits
         FormatString(2) = "d14"
         
         PrintValue(3) = " "
         FormatString(3) = "~"
         
         FormatPrint
         Ln = Ln + 1
         
         
         PrintValue(1) = " T O T A L   C R E D I T S"
         FormatString(1) = "a92"
         
         PrintValue(2) = Abs(TlCredits)
         FormatString(2) = "d14"
         
         PrintValue(3) = " "
         FormatString(3) = "~"
         
         FormatPrint
         Ln = Ln + 1
         
         If TlDebits + TlCredits <> 0 Then
            
            PrintValue(1) = "OUT OF BALANCE BY"
            FormatString(1) = "a76"
            
            PrintValue(2) = TlDebits + TlCredits
            FormatString(2) = "d14"
            
            PrintValue(3) = " "
            FormatString(3) = "~"
            
            FormatPrint
            Ln = Ln + 1
         End If
         
         ' print the hash total
         PrintValue(1) = " HASH Total: "
         FormatString(1) = "a20"
         
         FormatString(2) = "a9"
         X = Format(HashTotal, "########0")
         PrintValue(2) = X
         
         FormatString(3) = "a5"
         PrintValue(3) = ""
         
         FormatString(4) = "a20"
         PrintValue(4) = "Entry Count:"
         
         FormatString(5) = "a11"
         X = Format(EntryCount, "###,###,##0")
         PrintValue(5) = X
         
         PrintValue(6) = " "
         FormatString(6) = "~"
         
         FormatPrint
         Ln = Ln + 1
         
         ' clear the totals
         TlDebits = 0
         TlCredits = 0
         HashTotal = 0
         EntryCount = 0
         
         If Not jFlg Then   ' end of report
            Prvw.vsp.EndDoc
'            FormFeed
            Exit Do
         Else               ' next journal
            FormFeed
            Ln = 0
         End If
         
      End If
   
   Loop

   frmProgress.Hide

   Exit Sub
   
ErrMsg:
   MsgBox "Error: " & Err.Description & " " & Err.Number & vbCrLf & _
          "Module err#: " & EMsg, vbExclamation + vbOKOnly, "Windows GL"
   On Error GoTo 0

End Sub
         
Private Sub HistJnlNextPage()
         
     If Ln <> 0 Then FormFeed
     
     X = "Journal # " & GLHistory.JournalSource & " "
     If GLJournal.GetData(GLHistory.JournalSource) Then
        X = X & GLJournal.JournalName
     End If
     
     PageHeader GLCompany.Name & " - Data Entry Journal", glMsg(1), X, glMsg(2)
           
     PrintValue(1) = "JS#"
     FormatString(1) = "a3"
     
     PrintValue(2) = " SC"
     FormatString(2) = "a3"
     
     PrintValue(3) = "  T"
     FormatString(3) = "a3"
     
     PrintValue(4) = "   Date"
     FormatString(4) = "a9"
     
     PrintValue(5) = "   Acct Num"
     FormatString(5) = "a11"
     
     PrintValue(6) = "  Reference"
     FormatString(6) = "a22"
     
     PrintValue(7) = "   Description"
     FormatString(7) = "a24"
     
     PrintValue(8) = "     Debit Amt"
     FormatString(8) = "a16"
     
     PrintValue(9) = "    Credit Amt"
     FormatString(9) = "a16"
     
     PrintValue(10) = " "
     FormatString(10) = "~"
     
     FormatPrint
     
     Ln = Ln + 1
                 

End Sub
         
Private Sub PageHeader(ByVal ReportName As String, _
                       ByVal Msg1 As String, _
                       ByVal Msg2 As String, _
                       ByVal Msg3 As String)
                       
   Ln = 0
   Pg = Pg + 1
   
   ' 29 characters for fixed left and right portion of first header line
   '    1             8       1   8                    10         1
   ' first line - system date & time / company name / page #
   X = GLCompany.Name
   Y = Format(Date, "mm/dd/yy ") & Format(Time, "hh:mm:ss")
   Z = "Page: " & Format(Pg, "####")
   
   If Len(X) > Columns - 29 Then
      X = Mid(GLCompany.Name, 1, Columns - 29)
   End If
   
   ' center the company name in the string
   I = ((Columns - 29 - Len(X)) / 2) - 1
   
   Ln = 1
   w = Y & Space(I) & X & Space(I) & Z
   PrtCenter Ln, w
   
   If ReportName <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, ReportName
   End If
   
   If Msg1 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, Msg1
   End If
   
   If Msg2 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, Msg2
   End If
   
   If Msg3 <> "" Then
      Ln = Ln + 1
      PrtCenter Ln, Msg3
   End If

   Ln = Ln + 1

End Sub
