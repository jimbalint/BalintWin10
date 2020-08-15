Option Explicit

Dim cn As ADODB.Connection
Dim rsAmt As ADODB.Recordset
Dim rsAcct As ADODB.Recordset

Dim cnSys As ADODB.Connection
Dim rsDesc As ADODB.Recordset
Dim rsCompany As ADODB.Recordset

Dim FName As String
Dim flg As Boolean

Dim SQLString As String

Dim i As Integer
Dim j As Integer

Dim bytConnected As Byte

Dim CmdBar As CommandBar
Dim CmdBarBut As CommandBarButton


Sub Connect()
   
Dim ErrMsg As String
   
   ' data disconnect if already connected
   If bytConnected Then Disconnect
   
   bytConnected = 1
   
   Application.Calculation = xlCalculationManual
   
   FName = Range("GLFile")
   
   If FName = "" Then
      MsgBox "Please enter a dBase file name", vbCritical + vbOKOnly, "File Open"
      Range("GLFile").Select
      Exit Sub
   End If
   
'   On Error GoTo ConnectErr
   
   ErrMsg = "Data Base Connect"
   
   Set cn = New ADODB.Connection
   cn.Provider = "Microsoft.Jet.OLEDB.4.0"
   cn.ConnectionString = FName
   cn.Open
   
   ErrMsg = "GLAmount Get"
   
   Set rsAmt = New ADODB.Recordset
   rsAmt.Source = "SELECT * FROM GLAmount"
   rsAmt.LockType = adLockOptimistic
   rsAmt.CursorType = adOpenKeyset
   Set rsAmt.ActiveConnection = cn
   rsAmt.Open

   If rsAmt.BOF And rsAmt.EOF Then
      ErrMsg = "No GLAmount records found !"
      GoTo ConnectErr
   End If
   
   ErrMsg = "GLAccount Get"
   
   Set rsAcct = New ADODB.Recordset
   
   rsAcct.Source = "SELECT * FROM GLAccount ORDER BY Account"
   
   rsAcct.LockType = adLockOptimistic
   rsAcct.CursorType = adOpenKeyset
   Set rsAcct.ActiveConnection = cn
   rsAcct.Open
   
   If rsAcct.BOF And rsAcct.EOF Then
      ErrMsg = "No GLAccount records found !"
      GoTo ConnectErr
   End If
   
   ErrMsg = "GLSystem Connect"
   
   ' open the GLDesc file
   Set cnSys = New ADODB.Connection
   cnSys.Provider = "Microsoft.Jet.OLEDB.4.0"
   cnSys.ConnectionString = "v:\balint\data\GLSystem.mdb"
   cnSys.Open
   
   ErrMsg = "GLDescription Connect"
   
   Set rsDesc = New ADODB.Recordset
   rsDesc.Source = "SELECT * FROM GLDescriptions"
   rsDesc.LockType = adLockOptimistic
   rsDesc.CursorType = adOpenKeyset
   Set rsDesc.ActiveConnection = cnSys
   rsDesc.Open
   
   ErrMsg = "GLCompany Get"
   
   ' open the company file and get the first record
   Set rsCompany = New ADODB.Recordset
   rsCompany.Source = "SELECT * FROM GLCompany"
   rsCompany.LockType = adLockOptimistic
   rsCompany.CursorType = adOpenKeyset
   Set rsCompany.ActiveConnection = cnSys
   rsCompany.Open
   
   If rsCompany.BOF And rsCompany.EOF Then
      ErrMsg = "No Company records found!"
      GoTo ConnectErr
   End If
   
   ' loop through the company file - stop on the first one that has the same file name
   flg = False
   rsCompany.MoveFirst
   Do Until rsCompany.EOF
      If StrConv(rsCompany!FileName, vbLowerCase) = StrConv(FName, vbLowerCase) Then
         Range("LowBranch").Value = rsCompany!LowBranch
         Range("HiBranch").Value = rsCompany!HiBranch
         Range("LowConsolidated").Value = rsCompany!LowConsolidated
         Range("HiConsolidated").Value = rsCompany!HiConsolidated
         flg = True
         Exit Do
      End If
      rsCompany.MoveNext
   Loop
   
   If Not flg Then
      MsgBox "Warning - Company record not found for: " & vbCrLf & vbCrLf & _
             FName, vbInformation + vbOKOnly
   End If
   
   ' force recalc of all cells
   For i = 1 To Sheets.Count
       Sheets(i).Select
       Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell)).Dirty
       Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell)).Calculate
   Next i
   
   Application.CalculateFull

'   Application.Calculate
   
   ' home the cursor
   Sheets(1).Select
   Range("A2").Select
   
   MsgBox "Connect to " & FName & " complete" & vbCrLf & vbCrLf & _
          "Press Alt-G to refresh or if the data base changes", vbInformation
   Exit Sub
   
ConnectErr:
   
   MsgBox ErrMsg & " - " & Err.Description & " " & Str$(Err), vbCritical + vbOKOnly, _
          "Init Error!"
   Range("GLFile").Select
          
End Sub

Sub Disconnect()
    
'    rsAmt.Close
'    Set rsAmt = Nothing

    rsAcct.Close
    Set rsAcct = Nothing

    cn.Close
    Set cn = Nothing

    rsDesc.Close
    Set rsDesc = Nothing

    rsCompany.Close
    Set rsCompany = Nothing

    cnSys.Close
    Set cnSys = Nothing

End Sub

Function GetAmount(Acct As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency

    ' check to see if a consolidated account
    If Range("HiConsolidated").Value <> 0 And Acct Mod 10 ^ rsCompany!subdigits = Range("HiConsolidated").Value Then
       
       rsAcct.MoveFirst
       Do Until rsAcct.EOF
          If rsAcct!account = Acct Then
             GetAmount = GetCons(Int(rsAcct!account / 10 ^ rsCompany!subdigits), FY1, FY2, Pd1, Pd2)
          End If
          rsAcct.MoveNext
       Loop
       Exit Function
    End If

    rsAmt.MoveFirst
    
    Do Until rsAmt.EOF
       
       If rsAmt!account = Acct And rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
          If 1 >= Pd1 And 1 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount01
          If 2 >= Pd1 And 2 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount02
          If 3 >= Pd1 And 3 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount03
          If 4 >= Pd1 And 4 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount04
          If 5 >= Pd1 And 5 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount05
          If 6 >= Pd1 And 6 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount06
          If 7 >= Pd1 And 7 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount07
          If 8 >= Pd1 And 8 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount08
          If 9 >= Pd1 And 9 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount09
          If 10 >= Pd1 And 10 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount10
          If 11 >= Pd1 And 11 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount11
          If 12 >= Pd1 And 12 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount12
          If 13 >= Pd1 And 13 <= Pd2 Then GetAmount = GetAmount + rsAmt!Amount13
       
       End If
       
       rsAmt.MoveNext

    Loop

End Function

Function GetCons(Base As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency

Dim LC As Long
Dim HC As Long

    LC = 0
    HC = 3

    rsAmt.MoveFirst
  
    Do Until rsAmt.EOF
       
'       If CLng(Int(rsAmt!account / 10 ^ rsCompany!subdigits)) = Base And _
'          CLng(rsAmt!account Mod 10 ^ rsCompany!subdigits) >= LC And _
'          CLng(rsAmt!account Mod 10 ^ rsCompany!subdigits) <= HC And _
'          rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
       If CLng(Int(rsAmt!account / 10 ^ rsCompany!subdigits)) = Base And _
          rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
       
          If 1 >= Pd1 And 1 <= Pd2 Then GetCons = GetCons + rsAmt!Amount01
          If 2 >= Pd1 And 2 <= Pd2 Then GetCons = GetCons + rsAmt!Amount02
          If 3 >= Pd1 And 3 <= Pd2 Then GetCons = GetCons + rsAmt!Amount03
          If 4 >= Pd1 And 4 <= Pd2 Then GetCons = GetCons + rsAmt!Amount04
          If 5 >= Pd1 And 5 <= Pd2 Then GetCons = GetCons + rsAmt!Amount05
          If 6 >= Pd1 And 6 <= Pd2 Then GetCons = GetCons + rsAmt!Amount06
          If 7 >= Pd1 And 7 <= Pd2 Then GetCons = GetCons + rsAmt!Amount07
          If 8 >= Pd1 And 8 <= Pd2 Then GetCons = GetCons + rsAmt!Amount08
          If 9 >= Pd1 And 9 <= Pd2 Then GetCons = GetCons + rsAmt!Amount09
          If 10 >= Pd1 And 10 <= Pd2 Then GetCons = GetCons + rsAmt!Amount10
          If 11 >= Pd1 And 11 <= Pd2 Then GetCons = GetCons + rsAmt!Amount11
          If 12 >= Pd1 And 12 <= Pd2 Then GetCons = GetCons + rsAmt!Amount12
          If 13 >= Pd1 And 13 <= Pd2 Then GetCons = GetCons + rsAmt!Amount13
       
       End If
       
       rsAmt.MoveNext

    Loop

End Function

Function GetConsBud(Base As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency

Dim LC As Long
Dim HC As Long

    LC = 0
    HC = 3

    rsAmt.MoveFirst
  
    Do Until rsAmt.EOF
       
'       If CLng(Int(rsAmt!account / 10 ^ rsCompany!subdigits)) = Base And _
'          CLng(rsAmt!account Mod 10 ^ rsCompany!subdigits) >= LC And _
'          CLng(rsAmt!account Mod 10 ^ rsCompany!subdigits) <= HC And _
'          rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
       If CLng(Int(rsAmt!account / 10 ^ rsCompany!subdigits)) = Base And _
          rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
       
          If 1 >= Pd1 And 1 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget01
          If 2 >= Pd1 And 2 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget02
          If 3 >= Pd1 And 3 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget03
          If 4 >= Pd1 And 4 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget04
          If 5 >= Pd1 And 5 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget05
          If 6 >= Pd1 And 6 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget06
          If 7 >= Pd1 And 7 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget07
          If 8 >= Pd1 And 8 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget08
          If 9 >= Pd1 And 9 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget09
          If 10 >= Pd1 And 10 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget10
          If 11 >= Pd1 And 11 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget11
          If 12 >= Pd1 And 12 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget12
          If 13 >= Pd1 And 13 <= Pd2 Then GetConsBud = GetConsBud + rsAmt!Budget13
       
       End If
       
       rsAmt.MoveNext

    Loop

End Function



Function GetBudget(Acct As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency

    ' check to see if a consolidated account
    If Range("HiConsolidated").Value <> 0 And Acct Mod 10 ^ rsCompany!subdigits = Range("HiConsolidated").Value Then
       
       rsAcct.MoveFirst
       Do Until rsAcct.EOF
          If rsAcct!account = Acct Then
             GetBudget = GetConsBud(Int(rsAcct!account / 10 ^ rsCompany!subdigits), FY1, FY2, Pd1, Pd2)
          End If
          rsAcct.MoveNext
       Loop
       Exit Function
    End If

    rsAmt.MoveFirst
    
    Do Until rsAmt.EOF
       
       If rsAmt!account = Acct And rsAmt!fiscalyear >= FY1 And rsAmt!fiscalyear <= FY2 Then
       
          If 1 >= Pd1 And 1 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget01
          If 2 >= Pd1 And 2 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget02
          If 3 >= Pd1 And 3 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget03
          If 4 >= Pd1 And 4 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget04
          If 5 >= Pd1 And 5 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget05
          If 6 >= Pd1 And 6 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget06
          If 7 >= Pd1 And 7 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget07
          If 8 >= Pd1 And 8 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget08
          If 9 >= Pd1 And 9 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget09
          If 10 >= Pd1 And 10 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget10
          If 11 >= Pd1 And 11 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget11
          If 12 >= Pd1 And 12 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget12
          If 13 >= Pd1 And 13 <= Pd2 Then GetBudget = GetBudget + rsAmt!Budget13
       
       End If
       
       rsAmt.MoveNext

    Loop
    
End Function


Function GetDesc(ByVal Acct As Long) As String
   
    GetDesc = "Not Found!"
    rsAcct.MoveFirst
    Do Until rsAcct.EOF
       If rsAcct!account = Acct Then
          If rsAcct!DescNumber = 0 Then
             GetDesc = rsAcct!Description
          ElseIf rsAcct!DescNumber = 1 Then
             GetDesc = rsCompany!Name
          Else
             GetDesc = "Description #" & rsAcct!DescNumber & " not found"
             rsDesc.MoveFirst
             Do Until rsDesc.EOF
                If rsDesc!Number = rsAcct!DescNumber Then
                   GetDesc = rsAcct!Description & rsDesc!Description
                   Exit Do
                End If
                rsDesc.MoveNext
             Loop
          End If
          Exit Do
       End If
       rsAcct.MoveNext
    Loop
    
End Function

Function GetDate(ByVal FormatNum As Byte, _
                 ByVal FY As Long, _
                 ByVal Pd1 As Byte, _
                 ByVal Pd2 As Byte) As String

Dim CurrYrPdEnd As String
Dim CurrYrCurrPdBeg As String
Dim CurrYrFYBeg As String
Dim CurrYrFYEnd As String
Dim PrevYrPdEnd As String
Dim PrevYrPDBeg As String
Dim PrevYrFYBeg As String
Dim ReportDate As String
Dim x As String
Dim y As String

Dim Year1 As Long
Dim Month1 As Byte
Dim Year2 As Long
Dim Month2 As Byte
      
Dim DString As String
Dim Months As String
Dim Digit(13) As String
Dim fmt As String
Dim ii As Double
      
      
    ' get the calendar year and month for the start/end periods
    If rsCompany!firstperiod = 1 Then
       
       Month1 = Pd1
       Year1 = FY
       
       Month2 = Pd2
       Year2 = FY
    
    Else
       
       x = DateSerial(FY - 1, rsCompany!firstperiod, 1)
       
       y = DateAdd("m", Pd1 - 1, x)
       Month1 = Month(y)
       Year1 = Year(y)
       
       y = DateAdd("m", Pd2 - 1, x)
       Month2 = Month(y)
       Year2 = Year(y)
       
    End If
    
    ' curr yr pd ending      1
    CurrYrPdEnd = DateSerial(Year2, Month2, 1)
    CurrYrPdEnd = DateAdd("m", 1, CurrYrPdEnd)
    CurrYrPdEnd = DateAdd("d", -1, CurrYrPdEnd)
   
    ' curr yr pd beg        2
    CurrYrCurrPdBeg = DateSerial(Year1, Month1, 1)
                               
    ' curr yr fy beg        3
    CurrYrFYBeg = DateSerial(Year1, rsCompany!firstperiod, 1)
    If rsCompany!firstperiod > Month1 Then
       CurrYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)
    End If
      
    ' FY end date
    CurrYrFYEnd = DateAdd("yyyy", 1, CurrYrFYBeg)
    CurrYrFYEnd = DateAdd("d", -1, CurrYrFYEnd)
      
    ' prev yr pd end        4
    PrevYrPdEnd = DateAdd("yyyy", -1, CurrYrPdEnd)
      
    ' prev yr pd beg        5
    PrevYrPDBeg = DateAdd("yyyy", -1, CurrYrCurrPdBeg)
      
    ' prev yr fy beg        6
    PrevYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)

    fmt = "mmmm dd, yyyy"
   
    Digit(0) = " Zero "
    Digit(1) = " One "
    Digit(2) = " Two "
    Digit(3) = " Three "
    Digit(4) = " Four "
    Digit(5) = " Five "
    Digit(6) = " Six "
    Digit(7) = " Seven "
    Digit(8) = " Eight "
    Digit(9) = " Nine "
    Digit(10) = " Ten "
    Digit(11) = " Eleven "
    Digit(12) = " Twelve "
    Digit(13) = " Thirteen "
   
   
    Select Case FormatNum
    
       Case 0
          GetDate = Format(CurrYrPdEnd, fmt)
   
       Case 1
          If Pd2 = 1 Then
             GetDate = CStr(Digit(Pd2)) & " Month Ended " & Format(CurrYrPdEnd, fmt)
          Else
             GetDate = CStr(Digit(Pd2)) & " Months Ended " & Format(CurrYrPdEnd, fmt)
          End If
   
       Case 2
          GetDate = Format(CurrYrCurrPdBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
       Case 3
          GetDate = Format(CurrYrFYBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
       Case 4
          If Pd2 = 1 Then
             GetDate = CStr(Digit(Pd2)) & " Month Ended " & Format(PrevYrPdEnd, fmt)
          Else
             GetDate = CStr(Digit(Pd2)) & " Months Ended " & Format(PrevYrPdEnd, fmt)
          End If
   
       Case 5
          GetDate = Format(PrevYrPDBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
       Case 6
          GetDate = Format(PrevYrFYBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
       Case 7
          GetDate = Format(Now, fmt)
   
       Case 8
          GetDate = CStr(Pd2 * 4) & " Weeks Ended " & Format(CurrYrPdEnd, fmt)
   
       Case 9
          GetDate = CStr(Pd2 * 4) & " Weeks Ended " & Format(PrevYrPdEnd, fmt)
   
       Case 10
          GetDate = Space(18 - Len(Format(CurrYrPdEnd, fmt))) & Format(CurrYrPdEnd, fmt)
   
       Case 11
          GetDate = Space(18 - Len(Format(CurrYrCurrPdBeg, fmt))) & Format(CurrYrCurrPdBeg, fmt)
   
       Case 12
          GetDate = Space(18 - Len(Format(CurrYrFYBeg, fmt))) & Format(CurrYrFYBeg, fmt)
   
       Case 13
          GetDate = Space(18 - Len(Format(PrevYrPdEnd, fmt))) & Format(PrevYrPdEnd, fmt)
   
       Case 14
          GetDate = Space(18 - Len(Format(PrevYrPDBeg, fmt))) & Format(PrevYrPDBeg, fmt)
   
       Case 15
          GetDate = Space(18 - Len(Format(PrevYrFYBeg, fmt))) & Format(PrevYrFYBeg, fmt)
   
       Case 16
          If Pd2 - Pd1 + 1 = 1 Then
             GetDate = CStr(Digit(Pd2 - Pd1 + 1)) & " Month"
          Else
             GetDate = CStr(Digit(Pd2 - Pd1 + 1)) & " Months"
          End If

       Case 17
          If Pd2 = 1 Then
             GetDate = CStr(Pd2) & " Month"
          Else
             GetDate = CStr(Pd2) & " Months"
          End If
   
       Case 18
          GetDate = " And " & Year(CurrYrFYEnd)
   
       Case 19
          GetDate = " And " & Year(CurrYrFYEnd) - 1
   
    End Select

End Function


