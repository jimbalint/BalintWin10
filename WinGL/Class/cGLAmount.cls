VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public ID As Long
Public Account As Long
Public FiscalYear As Long
Public Amount01 As Currency
Public Amount02 As Currency
Public Amount03 As Currency
Public Amount04 As Currency
Public Amount05 As Currency
Public Amount06 As Currency
Public Amount07 As Currency
Public Amount08 As Currency
Public Amount09 As Currency
Public Amount10 As Currency
Public Amount11 As Currency
Public Amount12 As Currency
Public Amount13 As Currency
Public TotalAmount As Currency
Public Budget01 As Currency
Public Budget02 As Currency
Public Budget03 As Currency
Public Budget04 As Currency
Public Budget05 As Currency
Public Budget06 As Currency
Public Budget07 As Currency
Public Budget08 As Currency
Public Budget09 As Currency
Public Budget10 As Currency
Public Budget11 As Currency
Public Budget12 As Currency
Public Budget13 As Currency
Public TotalBudget As Currency
Public mintFiscalYear As Integer

Dim rs As ADODB.Recordset
Dim OK As Boolean
Dim SQLString As String
Dim Pointer As Long

Dim cmd As ADODB.Command
Dim param As ADODB.Parameter
Dim xDB As New XArrayDB

Public Sub OpenRS()
   
   SQLString = "GLAmount"
   rsInit SQLString, cn, rs

End Sub

Public Sub CloseRS()

   rs.Close
   Set rs = Nothing

End Sub

Public Function GetRecordSet(ByVal LoAcct As Long, _
                             ByVal HiAcct As Long, _
                             ByVal StartFY As Long, _
                             ByVal EndFY As Long) As Boolean
   
   If LoAcct = 0 And HiAcct = 0 Then
      SQLString = "SELECT * FROM GLAmount WHERE " & _
                  "FiscalYear >= " & StartFY & " AND " & _
                  "FiscalYear <= " & EndFY & _
                  " ORDER BY Account, FiscalYear"
   Else
      SQLString = "SELECT * FROM GLAmount WHERE " & _
                  "Account >= " & LoAcct & " AND " & _
                  "Account <= " & HiAcct & " AND " & _
                  "FiscalYear >= " & StartFY & " AND " & _
                  "FiscalYear <= " & EndFY & _
                  " ORDER BY Account, FiscalYear"
   End If
   
   rsInit SQLString, cn, rs
   
   If rs.BOF And rs.EOF Then
      GetRecordSet = False
      Exit Function
   Else
      GetRecordSet = True
   End If
   
   rs.MoveFirst
   
   AssignVars
   
End Function


Public Function Find(ByVal Acct As Long, ByVal FY As Long) As Boolean
   
Dim R1 As Long
   
   SQLString = "Account = " & Acct
   
   R1 = 0
   
   Do
      
      rs.Find SQLString, R1, adSearchForward, 1
      
      If rs.EOF Then
         Find = False
         Exit Function
      End If
   
      If rs!FiscalYear = FY Then Exit Do
      R1 = rs.AbsolutePosition
   
   Loop
   
   AssignVars
   Find = True

End Function
Public Function sFind(ByVal Acct As Long, _
                      ByVal FY As Long, _
                      ByVal SkipRecs As Long) As Long
   
Dim a As Long
   
   SQLString = "Account = " & Acct
   
   Do
   
      a = SkipRecs + 1
      rs.Find SQLString, a, adSearchForward, 1
   
      If rs.EOF Then       ' try from top of record set
         
         rs.Find SQLString, 0, adSearchForward, 1
         If rs.EOF Then
            sFind = 0
            Exit Function
         Else
            sFind = rs.AbsolutePosition
         End If
      End If
   
      If rs!FiscalYear = FY Then
         sFind = rs.AbsolutePosition
         AssignVars
         Exit Function
      Else
         SkipRecs = rs.AbsolutePosition
      End If
   
   Loop
   

End Function


Public Sub AssignVars()
         
         Account = rs!Account
         Amount01 = rs!Amount01
         Amount02 = rs!Amount02
         Amount03 = rs!Amount03
         Amount04 = rs!Amount04
         Amount05 = rs!Amount05
         Amount06 = rs!Amount06
         Amount07 = rs!Amount07
         Amount08 = rs!Amount08
         Amount09 = rs!Amount09
         Amount10 = rs!Amount10
         Amount11 = rs!Amount11
         Amount12 = rs!Amount12
         Amount13 = rs!Amount13
         
         TotalAmount = rs!Amount01 + rs!Amount02 + rs!Amount03 + rs!Amount04 + rs!Amount05
         TotalAmount = TotalAmount + rs!Amount06 + rs!Amount07 + rs!Amount08 + rs!Amount09
         TotalAmount = TotalAmount + rs!Amount10 + rs!Amount11 + rs!Amount12 + rs!Amount13
         
         Budget01 = rs!Budget01
         Budget02 = rs!Budget02
         Budget03 = rs!Budget03
         Budget04 = rs!Budget04
         Budget05 = rs!Budget05
         Budget06 = rs!Budget06
         Budget07 = rs!Budget07
         Budget08 = rs!Budget08
         Budget09 = rs!Budget09
         Budget10 = rs!Budget10
         Budget11 = rs!Budget11
         Budget12 = rs!Budget12
         Budget13 = rs!Budget13
         
         TotalBudget = rs!Budget01 + rs!Budget02 + rs!Budget03 + rs!Budget04 + rs!Budget05
         TotalBudget = TotalBudget + rs!Budget06 + rs!Budget07 + rs!Budget08 + rs!Budget09
         TotalBudget = TotalBudget + rs!Budget10 + rs!Budget11 + rs!Budget12 + rs!Budget13
         
         FiscalYear = rs!FiscalYear

End Sub


Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
         
    rs.Fields("Account") = Account
    rs.Fields("Amount01") = Amount01
    rs.Fields("Amount02") = Amount02
    rs.Fields("Amount03") = Amount03
    rs.Fields("Amount04") = Amount04
    rs.Fields("Amount05") = Amount05
    rs.Fields("Amount06") = Amount06
    rs.Fields("Amount07") = Amount07
    rs.Fields("Amount08") = Amount08
    rs.Fields("Amount09") = Amount09
    rs.Fields("Amount10") = Amount10
    rs.Fields("Amount11") = Amount11
    rs.Fields("Amount12") = Amount12
    rs.Fields("Amount13") = Amount13
    rs.Fields("Budget01") = Budget01
    rs.Fields("Budget02") = Budget02
    rs.Fields("Budget03") = Budget03
    rs.Fields("Budget04") = Budget04
    rs.Fields("Budget05") = Budget05
    rs.Fields("Budget06") = Budget06
    rs.Fields("Budget07") = Budget07
    rs.Fields("Budget08") = Budget08
    rs.Fields("Budget09") = Budget09
    rs.Fields("Budget10") = Budget10
    rs.Fields("Budget11") = Budget11
    rs.Fields("Budget12") = Budget12
    rs.Fields("Budget13") = Budget13
    rs.Fields("FiscalYear") = FiscalYear
   
    rs.Update

'  need out for GLImport
'    rs.Close
'    Set rs = Nothing
    
End Sub

Public Sub Clear()
      ID = 0
      Account = 0
      Amount01 = 0
      Amount02 = 0
      Amount03 = 0
      Amount04 = 0
      Amount05 = 0
      Amount06 = 0
      Amount07 = 0
      Amount08 = 0
      Amount09 = 0
      Amount10 = 0
      Amount11 = 0
      Amount12 = 0
      Amount13 = 0
      Budget01 = 0
      Budget02 = 0
      Budget03 = 0
      Budget04 = 0
      Budget05 = 0
      Budget06 = 0
      Budget07 = 0
      Budget08 = 0
      Budget09 = 0
      Budget10 = 0
      Budget11 = 0
      Budget12 = 0
      Budget13 = 0
      FiscalYear = 0
End Sub


Public Function GetAmount(ByVal lngAccount As Long, _
                          ByVal intFiscalYear As Integer, _
                          ByVal bytPeriod1 As Byte, _
                          ByVal bytPeriod2 As Byte) As Currency
                     
    ' find the GLAmount record first
    If Not QueryByAccountByFiscalYear(lngAccount, intFiscalYear) Then
       GetAmount = 0
       Exit Function
    End If

    If 1 >= bytPeriod1 And 1 <= bytPeriod2 Then GetAmount = GetAmount + Amount01
    If 2 >= bytPeriod1 And 2 <= bytPeriod2 Then GetAmount = GetAmount + Amount02
    If 3 >= bytPeriod1 And 3 <= bytPeriod2 Then GetAmount = GetAmount + Amount03
    If 4 >= bytPeriod1 And 4 <= bytPeriod2 Then GetAmount = GetAmount + Amount04
    If 5 >= bytPeriod1 And 5 <= bytPeriod2 Then GetAmount = GetAmount + Amount05
    If 6 >= bytPeriod1 And 6 <= bytPeriod2 Then GetAmount = GetAmount + Amount06
    If 7 >= bytPeriod1 And 7 <= bytPeriod2 Then GetAmount = GetAmount + Amount07
    If 8 >= bytPeriod1 And 8 <= bytPeriod2 Then GetAmount = GetAmount + Amount08
    If 9 >= bytPeriod1 And 9 <= bytPeriod2 Then GetAmount = GetAmount + Amount09
    If 10 >= bytPeriod1 And 10 <= bytPeriod2 Then GetAmount = GetAmount + Amount10
    If 11 >= bytPeriod1 And 11 <= bytPeriod2 Then GetAmount = GetAmount + Amount11
    If 12 >= bytPeriod1 And 12 <= bytPeriod2 Then GetAmount = GetAmount + Amount12
    If 13 >= bytPeriod1 And 13 <= bytPeriod2 Then GetAmount = GetAmount + Amount13
    
End Function

Public Function GetBudget(ByVal lngAccount As Long, _
                          ByVal intFiscalYear As Integer, _
                          ByVal bytPeriod1 As Byte, _
                          ByVal bytPeriod2 As Byte) As Currency
                     
    ' find the GLAmount record first
    If Not QueryByAccountByFiscalYear(lngAccount, intFiscalYear) Then
       GetBudget = 0
       Exit Function
    End If
                     
    If 1 >= bytPeriod1 And 1 <= bytPeriod2 Then GetBudget = GetBudget + Budget01
    If 2 >= bytPeriod1 And 2 <= bytPeriod2 Then GetBudget = GetBudget + Budget02
    If 3 >= bytPeriod1 And 3 <= bytPeriod2 Then GetBudget = GetBudget + Budget03
    If 4 >= bytPeriod1 And 4 <= bytPeriod2 Then GetBudget = GetBudget + Budget04
    If 5 >= bytPeriod1 And 5 <= bytPeriod2 Then GetBudget = GetBudget + Budget05
    If 6 >= bytPeriod1 And 6 <= bytPeriod2 Then GetBudget = GetBudget + Budget06
    If 7 >= bytPeriod1 And 7 <= bytPeriod2 Then GetBudget = GetBudget + Budget07
    If 8 >= bytPeriod1 And 8 <= bytPeriod2 Then GetBudget = GetBudget + Budget08
    If 9 >= bytPeriod1 And 9 <= bytPeriod2 Then GetBudget = GetBudget + Budget09
    If 10 >= bytPeriod1 And 10 <= bytPeriod2 Then GetBudget = GetBudget + Budget10
    If 11 >= bytPeriod1 And 11 <= bytPeriod2 Then GetBudget = GetBudget + Budget11
    If 12 >= bytPeriod1 And 12 <= bytPeriod2 Then GetBudget = GetBudget + Budget12
    If 13 >= bytPeriod1 And 13 <= bytPeriod2 Then GetBudget = GetBudget + Budget13
    
End Function

Public Sub DeleteAll()

    SQLString = "Delete * from GLAmount"
    rsInit SQLString, cn, rs

'    CloseRS
    
End Sub


Public Function QueryByAccountByFiscalYear(ByVal Acct As Long, ByVal FY As Long) As Boolean

   SQLString = "SELECT * FROM GLAmount WHERE Account = " & Acct & _
             " AND FiscalYear = " & FY

   rsInit SQLString, cn, rs

' ---------------------------------------
'  using the Command object is read only record set ???
'
'   Set cmd = New ADODB.Command
'   Set cmd.ActiveConnection = cn
'   cmd.CommandType = adCmdStoredProc
'   cmd.CommandText = "qryGLAmountByAccountByFiscalYear"
'
'   Set param = cmd.CreateParameter("qAccount", adNumeric, adParamInput, , Acct)
'   cmd.Parameters.Append param
'
'   Set param = cmd.CreateParameter("qFiscalYear", adNumeric, adParamInput, , FY)
'   cmd.Parameters.Append param
'
'   Set rs = New ADODB.Recordset
'   rs.CursorType = adOpenKeyset
'   rs.LockType = adLockOptimistic
'   Set rs = cmd.Execute
'
' ----------------------------------------
   
   If rs.BOF And rs.EOF Then
      QueryByAccountByFiscalYear = False
      Exit Function
   End If

   rs.MoveFirst
   
   AssignVars
   
   QueryByAccountByFiscalYear = True

End Function

Public Sub AddAmount(ByVal Pd As Byte, ByVal Amount As Currency)
              
      If Pd = 1 Then Amount01 = Amount01 + Amount
      If Pd = 2 Then Amount02 = Amount02 + Amount
      If Pd = 3 Then Amount03 = Amount03 + Amount
      If Pd = 4 Then Amount04 = Amount04 + Amount
      If Pd = 5 Then Amount05 = Amount05 + Amount
      If Pd = 6 Then Amount06 = Amount06 + Amount
      If Pd = 7 Then Amount07 = Amount07 + Amount
      If Pd = 8 Then Amount08 = Amount08 + Amount
      If Pd = 9 Then Amount09 = Amount09 + Amount
      If Pd = 10 Then Amount10 = Amount10 + Amount
      If Pd = 11 Then Amount11 = Amount11 + Amount
      If Pd = 12 Then Amount12 = Amount12 + Amount
      If Pd = 13 Then Amount13 = Amount13 + Amount

End Sub
 

Public Sub AddBudgAmount(ByVal Pd As Byte, ByVal Amount As Currency)
              
      If Pd = 1 Then Budget01 = Budget01 + Amount
      If Pd = 2 Then Budget02 = Budget02 + Amount
      If Pd = 3 Then Budget03 = Budget03 + Amount
      If Pd = 4 Then Budget04 = Budget04 + Amount
      If Pd = 5 Then Budget05 = Budget05 + Amount
      If Pd = 6 Then Budget06 = Budget06 + Amount
      If Pd = 7 Then Budget07 = Budget07 + Amount
      If Pd = 8 Then Budget08 = Budget08 + Amount
      If Pd = 9 Then Budget09 = Budget09 + Amount
      If Pd = 10 Then Budget10 = Budget10 + Amount
      If Pd = 11 Then Budget11 = Budget11 + Amount
      If Pd = 12 Then Budget12 = Budget12 + Amount
      If Pd = 13 Then Budget13 = Budget13 + Amount

End Sub
 

Public Function GetYears(ByVal LoYear As Integer, _
                         ByVal HiYear As Integer) _
                         As XArrayDB

Dim I As Integer

   SQLString = "Select DISTINCT FiscalYear from GLAmount order by FiscalYear Desc"
    
   rsInit SQLString, cn, rs
   
   xDB.ReDim 0, 0, 0, 0

   xDB(0, 0) = "-1"

   If rs.BOF And rs.EOF Then
      Set GetYears = xDB
      Exit Function
   End If

   rs.MoveFirst
   
   I = 0
   
   Do Until rs.EOF
      
      If LoYear <> 0 And rs!FiscalYear < LoYear Then GoTo NextYr
      If HiYear <> 0 And rs!FiscalYear > HiYear Then GoTo NextYr
      
      xDB.AppendRows 1
      I = I + 1
      
      xDB(0, 0) = ""
      xDB(I, 0) = CStr(rs!FiscalYear)
   
NextYr:
      rs.MoveNext
   
   Loop

   Set GetYears = xDB

   rs.Close
   Set rs = Nothing

End Function
