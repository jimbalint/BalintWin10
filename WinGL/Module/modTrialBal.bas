Attribute VB_Name = "modTrialBal"
Option Explicit
Dim StartPD As Byte
Dim FiscalYear As Integer
Dim EndPd As Byte
Dim UnderLn As Boolean
Dim P1 As Long
Dim Amount As Currency
Dim TlDebits As Currency
Dim TlCredits As Currency
Dim UnderLines As Boolean
Dim ct As Long
Dim tm As Variant
Dim GLFName As String
Dim X As String
Public EndYM As Long
Dim CompanyID As Long
Dim flg As Boolean
Public DBName As String
Dim frmGLPring As New frmGLPrint
Dim BatchNum As Long


Public Sub GLTrialBal()

'   GLPrint.PrtZeroBal = False
   
   Dim SQLString As String
   Columns = 132
   UnderLn = False
   
   P1 = 2147483647
   PgNum = 1
   Columns = 132
   Condensed = True
   UnderLines = True
   
   ' initialize the print screen
   PrtInit ("Port")
   SetFont 7, Equate.Portrait             ' compressed - portrait
   
   ' opens company record and assigns date variables
   
   ' !!!!! glprint.begindate and glprint.enddate
   ' !!!!! STUFFED in the frmGLPrint code
   
   GetDates CompanyID
   
   StartPD = GetPeriod(GLCompany.NumberPds, GLPrint.BeginDate Mod 100, GLCompany.FirstPeriod)
   EndPd = GetPeriod(GLCompany.NumberPds, GLPrint.EndDate Mod 100, GLCompany.FirstPeriod)
   
   PgHeader

'   SQLString = "Select [Account],[Description],[AcctType] from GLAccount " & _
'               "Where Account >= " & GLPrint.LowAccount & " " & _
'               "and Account <= " & GLPrint.HiAccount & " " & _
'               "and (AcctType = ""0"" or AcctType = ""P"") " & _
'               "Order By Account"
'
'   Dim rs As ADODB.Recordset
'   Set rs = New ADODB.Recordset
'   rs.Source = SQLString
'   rs.ActiveConnection = cn
'   rs.CursorType = adOpenDynamic
'   rs.Open
'
'   ' Check to make sure a record
'   ' actually came back
'   If rs.EOF = True And rs.BOF = True Then
'      MsgBox "No accounts found !!!", vbExclamation
'      rs.Close
'      End
'   End If
'
''   rs.MoveLast
''   frmProgress.prgBar1.Max = CLng(rs.RecordCount)
'
'   rs.MoveFirst
   
   
   frmProgress.lblMsg1 = "Printing Trial Balance for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering Account Information ..."
   frmProgress.Show
   
   GLAccount.GetAllAccounts
   
   ct = 0
   
   Do
      
      ct = ct + 1
      If ct = 1 Or ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
         frmProgress.lblMsg2.Refresh
      End If
      
      If GLAccount.Account < GLPrint.LowAccount Then GoTo Cycle
      If GLAccount.Account > GLPrint.HiAccount Then Exit Do
      If GLAccount.AcctType <> "0" And GLAccount.AcctType <> "P" Then GoTo Cycle
      
      Amount = GLAmount.GetAmount(GLAccount.Account, _
                                    GLPrint.FiscalYear, _
                                    1, _
                                    EndPd)
           
      ' 0940
      If GLAccount.AcctType = "P" And P1 = 2147483647 Then
         PrintTotals
         P1 = GLAccount.Account
         Prvw.vsp.NewPage
         PgHeader
         GoTo Cycle
      End If
           
      If Amount = 0# And GLPrint.PrtZeroBal = False Then
         GoTo Cycle
      End If
      
      If Amount >= 0 Then     ' Debit Balance
         
         FormatString(1) = "n9"
         FormatString(2) = "x2"
         FormatString(3) = "a27"
         FormatString(4) = "x2"
         FormatString(5) = "d14"
         FormatString(6) = "~"
      
         TlDebits = TlDebits + Amount
      
      Else
         
         FormatString(1) = "n9"
         FormatString(2) = "x2"
         FormatString(3) = "a27"
         FormatString(4) = "t57"
         FormatString(5) = "d14"
         FormatString(6) = "~"
      
         TlCredits = TlCredits + Amount
      
      End If
      
      PrintValue(1) = GLAccount.Account
      PrintValue(3) = GLAccount.FullDesc
      
      PrintValue(5) = Amount
      
      FormatPrint
      Ln = Ln + 1
      
      If UnderLines = True Then
'         Prt Ln, 1, String(132, "-")
         Prt Ln, 1, String(129, "-")
         Ln = Ln + 1
      End If
      
      If Ln > 55 Then
         Prvw.vsp.NewPage
         PgHeader
      End If
   
Cycle:
      If Not GLAccount.GetNext Then Exit Do
   Loop
   
   ' close the progress window
   frmProgress.MousePointer = vbArrow
   frmProgress.Hide
      
   PrintTotals
   
   ' open the preview window
   Prvw.Caption = GLCompany.Name & " Trial Balance"
   Prvw.vsp.EndDoc
         
End Sub
Private Sub PrintTotals()
   
Dim AbsAmt As Currency
   
   If Ln > 50 Then
      Prvw.vsp.NewPage
      PgHeader
   End If
   
   ' 1490 ----------------------------------------------------------------
   FormatString(1) = "x9"
   FormatString(2) = "x2"
   FormatString(3) = "x27"
   FormatString(4) = "x2"
   FormatString(5) = "d14"
   FormatString(6) = "t57"
   FormatString(7) = "d14"
   FormatString(8) = "~"
   
   PrintValue(5) = TlDebits
   PrintValue(7) = TlCredits
   
   FormatPrint
   Ln = Ln + 1
    
   ' 1500 / 1510 ---------------------------------------------------------
   If Abs(TlDebits) > Abs(TlCredits) Then
      
      FormatString(1) = "x9"
      FormatString(2) = "x2"
      FormatString(3) = "a27"
      FormatString(4) = "x2"
      FormatString(5) = "t57"
      FormatString(6) = "d14"
      FormatString(7) = "~"
      
      PrintValue(3) = "Net Change"
      PrintValue(6) = Abs(TlDebits + TlCredits) * (-1)
   
   Else
   
      FormatString(1) = "x9"
      FormatString(2) = "x2"
      FormatString(3) = "a27"
      FormatString(4) = "x2"
      FormatString(5) = "d14"
      FormatString(6) = "~"
      
      PrintValue(3) = "Net Change"
      PrintValue(5) = Abs(TlDebits + TlCredits)
   
   End If
      
   FormatPrint
   Ln = Ln + 1
    
   ' 1520 ---------------------------------------------------
   FormatString(1) = "x9"
   FormatString(2) = "x2"
   FormatString(3) = "x27"
   FormatString(4) = "x2"
   FormatString(5) = "a14"
   FormatString(6) = "t57"
   FormatString(7) = "a14"
   FormatString(8) = "~"
   
   PrintValue(5) = String(28, "-")
   PrintValue(7) = String(28, "-")
   
   FormatPrint
   Ln = Ln + 1
   
   ' 1530 ----------------------------------------------------
   FormatString(1) = "x9"
   FormatString(2) = "x2"
   FormatString(3) = "a27"
   FormatString(4) = "x2"
   FormatString(5) = "d14"
   FormatString(6) = "t57"
   FormatString(7) = "d14"
   FormatString(8) = "~"
   
   If Abs(TlDebits) > Abs(TlCredits) Then
      AbsAmt = Abs(TlDebits)
   Else
      AbsAmt = Abs(TlCredits)
   End If
   
   PrintValue(3) = "T O T A L"
   PrintValue(5) = AbsAmt
   PrintValue(7) = AbsAmt * (-1)
   
   FormatPrint
   Ln = Ln + 1
   
   ' --------------------------------------
   FormatString(1) = "x9"
   FormatString(2) = "x2"
   FormatString(3) = "x27"
   FormatString(4) = "x2"
   FormatString(5) = "a14"
   FormatString(6) = "t57"
   FormatString(7) = "a14"
   FormatString(8) = "~"
   
   PrintValue(5) = String(14, "=")
   PrintValue(7) = String(14, "=")
   
   FormatPrint
   Ln = Ln + 1
   
   TlDebits = 0
   TlCredits = 0
   
End Sub


Private Sub PgHeader()

   Ln = 2
   PrtCenter Ln, GLCompany.Name
   Prt Ln, Columns - 10, "Page:" & CStr(PgNum)
   Ln = Ln + 1
   PgNum = PgNum + 1
   
   ' ----------- trial balance ending date
   PrtCenter Ln, "Trial Balance Ending: " & Format(CurrYrPdEnd, "Long Date")
   Ln = Ln + 1
   
   ' ------------ system date
   PrtCenter Ln, "System Date: " & Format(Now(), "Long Date")
   Ln = Ln + 2
   
   ' -----------------------
   FormatString(1) = "t49"
   FormatString(2) = "a22"
   FormatString(3) = "t82"
   FormatString(4) = "a22"
   FormatString(5) = "t118"
   FormatString(6) = "a10"
   FormatString(7) = "~"
   
   PrintValue(2) = "P r e l i m i n a r y"
   PrintValue(4) = "A d j u s t m e n t s"
   PrintValue(6) = "F i n a l"
   
   FormatPrint
   Ln = Ln + 1
   
   ' -------------------------
   FormatString(1) = "r9"
   FormatString(2) = "x2"
   FormatString(3) = "a27"
   FormatString(4) = "x2"
   
   FormatString(5) = "r14"
   FormatString(6) = "x2"
   FormatString(7) = "r14"
   FormatString(8) = "x2"
   
   FormatString(9) = "r14"
   FormatString(10) = "x2"
   FormatString(11) = "r14"
   FormatString(12) = "x2"
   
   FormatString(13) = "r12"
   FormatString(14) = "x2"
   FormatString(15) = "r12"
   FormatString(16) = "x2"
   
   FormatString(17) = "~"
   
   PrintValue(1) = "ACCT #"
   PrintValue(2) = "Account Description"
   
   PrintValue(5) = "DEBIT"
   PrintValue(7) = "CREDIT"
   
   PrintValue(9) = "DEBIT"
   PrintValue(11) = "CREDIT"
   
   PrintValue(13) = "DEBIT"
   PrintValue(15) = "CREDIT"
   
   FormatPrint
   Ln = Ln + 2
   
End Sub


