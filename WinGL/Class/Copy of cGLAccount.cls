VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Private SQLString As String

' private member variables for object properties
Public ID As Long
Public DescNumber As Long
Public Account As Long
Public AllSchedules As Boolean
Public AllStatements As Boolean
Public BranchAcct As Boolean
Public BSColumn As Byte
Public ConsAcct As Boolean
Public Date1 As Long
Public Date2 As Long
Public Description As String
Public DollarSign As Boolean
Public LineFeeds As Byte
Public PrintTab As Byte
Public SignRevSched As Boolean
Public SignRevStmt As Boolean
Public TotalLevel As Byte
Public TotalOnLedger As Boolean
Public AcctType As String

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim xdb As New XArrayDB


Public Function Save(ByVal booAddPut As Boolean) As String

    Set rs = New ADODB.Recordset
    
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    If Account = 0 Then
       Save = "Account number not assigned!"
       Exit Function
    End If
    
    Set rs.ActiveConnection = cn
 
    If booAddPut = True Then      ' Add
    
         ' make sure the account does not already exist !!!
         rs.Source = "select * " & _
                     "from GLAccount " & _
                     "where Account = " & Account
         rs.Open
         
         If rs.BOF = False And rs.EOF = False Then
            Save = "Account already exists: " & Account
            rs.Close
            Set rs = Nothing
            Exit Function
         End If
         
         rs.Close

         rs.Source = "GLAccount"
         rs.Open
         rs.AddNew
    
    Else                           ' Put
    
         rs.Source = "select * " & _
                     "from GLAccount " & _
                     "where Account = " & Account
         rs.Open
         
         If rs.BOF = True And rs.EOF = True Then
            Save = "Account not found: " & Account
            Exit Function
         End If
         
    
    End If
    
    rs.Fields("Account") = Account
    rs.Fields("AllSchedules") = AllSchedules
    rs.Fields("AllStatements") = AllStatements
    rs.Fields("BranchAcct") = BranchAcct
    rs.Fields("BSColumn") = BSColumn
    rs.Fields("ConsAcct") = ConsAcct
    rs.Fields("Date1") = Date1
    rs.Fields("Date2") = Date2
    rs.Fields("Description") = Description
    rs.Fields("DescNumber") = DescNumber
    rs.Fields("DollarSign") = DollarSign
    rs.Fields("LineFeeds") = LineFeeds
    rs.Fields("PrintTab") = PrintTab
    rs.Fields("SignRevSched") = SignRevSched
    rs.Fields("SignRevStmt") = SignRevStmt
    rs.Fields("TotalLevel") = TotalLevel
    rs.Fields("TotalOnLedger") = TotalOnLedger
    rs.Fields("AcctType") = AcctType
    
    rs.Update
    
    rs.Close
    
    Set rs = Nothing
    
End Function

Public Sub Clear()
      
      ID = 0
      Account = 0
      Description = ""
      DescNumber = 0
      AllSchedules = False
      AllStatements = False
      BranchAcct = False
      BSColumn = 0
      ConsAcct = False
      Date1 = 0
      Date2 = 0
      DollarSign = False
      LineFeeds = 0
      PrintTab = 0
      SignRevSched = False
      SignRevStmt = False
      TotalLevel = 0
      TotalOnLedger = False
      AcctType = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    
    rs.Source = "Delete * from GLAccount"
    Set rs.ActiveConnection = cn
    rs.Open
    Set rs = Nothing
    
End Sub


Public Function DeleteRecord(ByVal lngAcct As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngAcct = 0 Then
       ErrMessage = "Account number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    rs.Source = "Delete * from GLAccount where Account = " & lngAcct
    Set rs.ActiveConnection = cn
   
    If rs.BOF = True And rs.EOF = True Then
       ErrMessage = "Account not found!"
       DeleteRecord = False
    Else
       ErrMessage = ""
       DeleteRecord = True
    End If
    
    Set rs = Nothing
    
End Function

Public Function Records() As Long

    SQLString = "select Account from GLAccount Order By Account"
    
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = cn
    rs.CursorType = adOpenStatic
    rs.Source = SQLString
    rs.Open
    
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
       Records = 0
    Else
       rs.MoveLast
       Records = rs.RecordCount
    End If

    Set rs = Nothing

End Function

Public Function GetRecordSets(ByVal CurrFY As Integer, ByVal PrevFY As Integer) As Boolean

   If GLPrint.RegBraCon <> Equate.Budget Then   ' don't include budget fields
      
      SQLString = "select GLAccount.*, GLAmount.FiscalYear, " & _
                  "GLAmount.Amount01, GLAmount.Amount02, GLAmount.Amount03, " & _
                  "GLAmount.Amount04, GLAmount.Amount05, GLAmount.Amount06, " & _
                  "GLAmount.Amount07, GLAmount.Amount08, GLAmount.Amount09, " & _
                  "GLAmount.Amount10, GLAmount.Amount11, GLAmount.Amount12, GLAmount.Amount13 " & _
                  "from GLAccount Left Join GLAmount on " & _
                  "(GLAccount.Account = GLAmount.Account and " & _
                  CurrFY & " = GLAmount.FiscalYear) order by GLAccount.Account"
                  
   Else
   
      SQLString = "select GLAccount.*, GLAmount.FiscalYear, " & _
                  "GLAmount.Amount01, GLAmount.Amount02, GLAmount.Amount03, " & _
                  "GLAmount.Amount04, GLAmount.Amount05, GLAmount.Amount06, " & _
                  "GLAmount.Amount07, GLAmount.Amount08, GLAmount.Amount09, " & _
                  "GLAmount.Amount10, GLAmount.Amount11, GLAmount.Amount12, GLAmount.Amount13, " & _
                  "GLAmount.Budget01, GLAmount.Budget02, GLAmount.Budget03, " & _
                  "GLAmount.Budget04, GLAmount.Budget05, GLAmount.Budget06, " & _
                  "GLAmount.Budget07, GLAmount.Budget09, GLAmount.Budget10, " & _
                  "GLAmount.Budget11, GLAmount.Budget12, GLAmount.Budget13, " & _
                  "from GLAccount Left Join GLAmount on " & _
                  "(GLAccount.Account = GLAmount.Account and " & _
                  CurrFY & " = GLAmount.FiscalYear) order by GLAccount.Account"
   
   End If
   
   Set rs1 = New ADODB.Recordset
   rs1.Source = SQLString
   rs1.ActiveConnection = cn
   rs1.CursorType = adOpenDynamic
   rs1.LockType = adLockBatchOptimistic
   
   rs1.Open

   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !", vbExclamation
      rs1.Close
      GetRecordSets = False
      Exit Function
   End If

   rs1.MoveFirst
   
   AssignFields

   ' get PrevFY amounts
   SQLString = "Select * from GLAmount where GLAmount.FiscalYear = " & PrevFY
      
   Set rs2 = New ADODB.Recordset
   rs2.Source = SQLString
   rs2.ActiveConnection = cn
   rs2.CursorType = adOpenDynamic
   rs2.Open

   ' Check to make sure a record
   ' actually came back
   If rs2.EOF = True And rs2.BOF = True Then
      MsgBox "No AMOUNTS found !!!", vbExclamation
      rs2.Close
      GetRecordSets = False
      Exit Function
   End If
   
   SQLString = "Account = " & CStr(rs1!Account)
   
   rs2.Find SQLString, 0, adSearchForward, 1

   GetRecordSets = True

End Function

Public Function GetAllAccounts() As Boolean
   
   SQLString = "Select * from GLAccount"
   
   Set rs1 = New ADODB.Recordset
   rs1.Source = SQLString
   rs1.ActiveConnection = cn
   rs1.CursorType = adOpenDynamic
   rs1.LockType = adLockBatchOptimistic
   
   rs1.Open

   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !!!", vbExclamation
      rs1.Close
      GetAllAccounts = False
      Exit Function
   End If

End Function



Public Function GetNext() As Boolean

   rs1.MoveNext
   
   If rs1.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignFields
      GetNext = True
   End If

   SQLString = "Account = " & CStr(rs1!Account)
   
   rs2.Find SQLString, 0, adSearchForward, 1
   
End Function

Public Sub AssignFields()
    
    ' Populate properties with
    ' results of the query.
    Account = rs1!Account
    AllSchedules = rs1!AllSchedules
    AllStatements = rs1!AllStatements
    BranchAcct = rs1!BranchAcct
    BSColumn = rs1!BSColumn
    ConsAcct = rs1!ConsAcct
    Date1 = rs1!Date1
    Date2 = rs1!Date2
    Description = rs1!Description & ""
    DescNumber = rs1!DescNumber
    DollarSign = rs1!DollarSign
    LineFeeds = rs1!LineFeeds
    PrintTab = rs1!PrintTab
    SignRevSched = rs1!SignRevSched
    SignRevStmt = rs1!SignRevStmt
    TotalLevel = rs1!TotalLevel
    TotalOnLedger = rs1!TotalOnLedger
    AcctType = rs1!AcctType

End Sub

Public Function GetCurrAmount(ByVal bytPeriod1 As Byte, _
                              ByVal bytPeriod2 As Byte) As Currency
                     
    If 1 >= bytPeriod1 And 1 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount01
    If 2 >= bytPeriod1 And 2 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount02
    If 3 >= bytPeriod1 And 3 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount03
    If 4 >= bytPeriod1 And 4 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount04
    If 5 >= bytPeriod1 And 5 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount05
    If 6 >= bytPeriod1 And 6 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount06
    If 7 >= bytPeriod1 And 7 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount07
    If 8 >= bytPeriod1 And 8 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount08
    If 9 >= bytPeriod1 And 9 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount09
    If 10 >= bytPeriod1 And 10 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount10
    If 11 >= bytPeriod1 And 11 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount11
    If 12 >= bytPeriod1 And 12 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount12
    If 13 >= bytPeriod1 And 13 <= bytPeriod2 Then GetCurrAmount = GetCurrAmount + rs1!Amount13
    
End Function

Public Function GetPrevAmount(ByVal bytPeriod1 As Byte, _
                              ByVal bytPeriod2 As Byte) As Currency
                     
    ' find for prev amount was not successful
    If rs2.EOF Then
       GetPrevAmount = 0
       Exit Function
    End If
                     
    If 1 >= bytPeriod1 And 1 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount01
    If 2 >= bytPeriod1 And 2 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount02
    If 3 >= bytPeriod1 And 3 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount03
    If 4 >= bytPeriod1 And 4 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount04
    If 5 >= bytPeriod1 And 5 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount05
    If 6 >= bytPeriod1 And 6 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount06
    If 7 >= bytPeriod1 And 7 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount07
    If 8 >= bytPeriod1 And 8 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount08
    If 9 >= bytPeriod1 And 9 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount09
    If 10 >= bytPeriod1 And 10 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount10
    If 11 >= bytPeriod1 And 11 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount11
    If 12 >= bytPeriod1 And 12 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount12
    If 13 >= bytPeriod1 And 13 <= bytPeriod2 Then GetPrevAmount = GetPrevAmount + rs2!Amount13
    
End Function


Public Function GetAccount(ByVal Acct As Long) As Boolean

   rs1.Find "Account = " & Acct, 0, adSearchForward, 1
   
   If rs1.EOF Then
      GetAccount = False
   Else
      AssignFields
      GetAccount = True
   End If

End Function
Public Function GetBudget(ByVal bytPeriod1 As Byte, _
                              ByVal bytPeriod2 As Byte) As Currency
                     
    If 1 >= bytPeriod1 And 1 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget01
    If 2 >= bytPeriod1 And 2 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget02
    If 3 >= bytPeriod1 And 3 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget03
    If 4 >= bytPeriod1 And 4 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget04
    If 5 >= bytPeriod1 And 5 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget05
    If 6 >= bytPeriod1 And 6 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget06
    If 7 >= bytPeriod1 And 7 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget07
    If 8 >= bytPeriod1 And 8 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget08
    If 9 >= bytPeriod1 And 9 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget09
    If 10 >= bytPeriod1 And 10 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget10
    If 11 >= bytPeriod1 And 11 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget11
    If 12 >= bytPeriod1 And 12 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget12
    If 13 >= bytPeriod1 And 13 <= bytPeriod2 Then GetBudget = GetBudget + rs1!Budget13
    
End Function

Public Function GetAcctRecSet(ByVal LoAcct As Long, ByVal HiAcct As Long) As Boolean

   If LoAcct = 0 And HiAcct = 0 Then
      SQLString = "SELECT * FROM GLAccount"
   
   Else
      SQLString = "SELECT * FROM GLAccount WHERE GLAccount.Account >= LoAcct AND " & _
                  "GLAccount.Account <= HiAcct)"
   End If
               
   Set rs1 = New ADODB.Recordset
   rs1.Source = SQLString
   rs1.ActiveConnection = cn
   rs1.CursorType = adOpenDynamic
   rs1.LockType = adLockBatchOptimistic
   
   rs1.Open

   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !", vbExclamation
      rs1.Close
      GetAcctRecSet = False
      Exit Function
   End If

   rs1.MoveFirst
   
   AssignFields

   GetAcctRecSet = True

End Function


Public Function GetNextAcct() As Boolean

   rs1.MoveNext
   
   If rs1.EOF Then
      GetNextAcct = False
      Exit Function
   Else
      AssignFields
      GetNextAcct = True
   End If

End Function


Public Function xGetRecordSet() As Boolean

Dim r1 As Long

   ' load the accounts
   SQLString = "SELECT * FROM GLAccount"
   
   rsInit SQLString, cn, rs
   
   If rs.BOF And rs.EOF Then
      xGetRecordSet = False
      Exit Function
   End If
   
   ' load into the arrays
   rs.MoveLast
   
   xdb.ReDim 1, rs.RecordCount, 0, 19
   
   rs.MoveFirst
   
   Do Until rs.EOF
      
      r1 = r1 + 1
      
      xdb(r1, 0) = CLng(rs!ID)
      xdb(r1, 1) = CLng(rs!Account)
      xdb(r1, 2) = CBool(rs!AllSchedules)
      xdb(r1, 3) = CBool(rs!AllStatements)
      xdb(r1, 4) = CBool(rs!BranchAcct)
      xdb(r1, 5) = CByte(rs!BSColumn)
      xdb(r1, 6) = CBool(rs!ConsAcct)
      xdb(r1, 7) = CLng(rs!Date1)
      xdb(r1, 8) = CLng(rs!Date2)
      xdb(r1, 9) = CStr(rs!Description) & ""
      xdb(r1, 10) = CLng(rs!DescNumber)
      xdb(r1, 11) = CBool(rs!DollarSign)
      xdb(r1, 12) = CByte(rs!LineFeeds)
      xdb(r1, 13) = CByte(rs!PrintTab)
      xdb(r1, 14) = CBool(rs!SignRevSched)
      xdb(r1, 15) = CBool(rs!SignRevStmt)
      xdb(r1, 16) = CByte(rs!TotalLevel)
      xdb(r1, 17) = CBool(rs!TotalOnLedger)
      xdb(r1, 18) = CStr(rs!AcctType) & ""
      xdb(r1, 19) = False
      
      rs.MoveNext
   
   Loop

   xGetRecordSet = True
   Set rs = Nothing
   
End Function
                   
Public Function xGetAccount(ByVal Acct As Long) As Long

   Clear      ' clear values

   xGetAccount = xdb.Find(1, 1, Acct)
   
   ' not found !!!
   If xGetAccount = 0 Then Exit Function

   xAssignVars (xGetAccount)

End Function

Sub xAssignVars(ByVal RowNum As Long)

    ID = xdb(RowNum, 0)
    Account = xdb(RowNum, 1)
    AllSchedules = xdb(RowNum, 2)
    AllStatements = xdb(RowNum, 3)
    BranchAcct = xdb(RowNum, 4)
    BSColumn = xdb(RowNum, 5)
    ConsAcct = xdb(RowNum, 6)
    Date1 = xdb(RowNum, 7)
    Date2 = xdb(RowNum, 8)
    Description = xdb(RowNum, 9)
    DescNumber = xdb(RowNum, 10)
    DollarSign = xdb(RowNum, 11)
    LineFeeds = xdb(RowNum, 12)
    PrintTab = xdb(RowNum, 13)
    SignRevSched = xdb(RowNum, 14)
    SignRevStmt = xdb(RowNum, 15)
    TotalLevel = xdb(RowNum, 16)
    TotalOnLedger = xdb(RowNum, 17)
    AcctType = xdb(RowNum, 18)

End Sub

Public Sub xSaveRow(ByVal RowNum As Long)

   ' add a record to the xDBArray
   If RowNum = 0 Then
      xdb.AppendRows 1
      RowNum = xdb.UpperBound(1)
      xdb(RowNum, 0) = 0          ' set the ID to zero
   End If
   
   ' assign the rest of the fields
   xdb(RowNum, 1) = Account
   xdb(RowNum, 2) = AllSchedules
   xdb(RowNum, 3) = AllStatements
   xdb(RowNum, 4) = BranchAcct
   xdb(RowNum, 5) = BSColumn
   xdb(RowNum, 6) = ConsAcct
   xdb(RowNum, 7) = Date1
   xdb(RowNum, 8) = Date2
   xdb(RowNum, 9) = Description
   xdb(RowNum, 10) = DescNumber
   xdb(RowNum, 11) = DollarSign
   xdb(RowNum, 12) = LineFeeds
   xdb(RowNum, 13) = PrintTab
   xdb(RowNum, 14) = SignRevSched
   xdb(RowNum, 15) = SignRevStmt
   xdb(RowNum, 16) = TotalLevel
   xdb(RowNum, 17) = TotalOnLedger
   xdb(RowNum, 18) = AcctType
   xdb(RowNum, 19) = True          ' set the update flag
      
End Sub

Public Function xSaveFile() As Boolean

Dim r1 As Long
Dim r2 As Long
Dim UCount As Long
Dim NCount As Long

   ' load the accounts
   SQLString = "SELECT * FROM GLAccount"
   rsInit SQLString, cn, rs
   
   If rs.BOF And rs.EOF Then
      xSaveFile = False
      Exit Function
   End If

   rs.MoveFirst

   MsgBox "A"
   xdb.QuickSort 1, xdb.UpperBound(1), 19, XORDER_ASCEND, XTYPE_BOOLEAN
   MsgBox xdb(1, 19)
   
   Do Until rs.EOF

      ' find the account in xdb
'      r1 = xdb.Find(1, 1, CLng(rs!Account))
      If r1 = 0 Then GoTo sCycle             ' ????
      If xdb(r1, 0) = 0 Then GoTo sCycle     ' ????  ID not assigned - new rec
      If Not xdb(r1, 19) Then GoTo sCycle    ' update not needed

      ' assign to the record set
      xAssignFields (r1)

      UCount = UCount + 1

sCycle:
      rs.MoveNext

   Loop
 
 MsgBox "Records Updated: " & UCount

   ' look for new accounts to add - ID = 0
   r1 = 1
   Do
      
      r2 = xdb.Find(r1, 0, 0)
      
      If r2 = 0 Then
         Exit Do
      End If
      
      ' add it
      rs.AddNew
      xAssignFields (r2)
      rs.Update
      
      NCount = NCount + 1
      
      r1 = r2 + 1
      If r1 > xdb.UpperBound(1) Then Exit Do
   
   Loop

   MsgBox "Adds: " & NCount

End Function

Sub xAssignFields(ByVal RowNum As Long)
      
      rs.Fields("Account") = xdb(RowNum, 1)
      rs.Fields("AllSchedules") = xdb(RowNum, 2)
      rs.Fields("AllStatements") = xdb(RowNum, 3)
      rs.Fields("BranchAcct") = xdb(RowNum, 4)
      rs.Fields("BSColumn") = xdb(RowNum, 5)
      rs.Fields("ConsAcct") = xdb(RowNum, 6)
      rs.Fields("Date1") = xdb(RowNum, 7)
      rs.Fields("Date2") = xdb(RowNum, 8)
      rs.Fields("Description") = xdb(RowNum, 9)
      rs.Fields("DescNumber") = xdb(RowNum, 10)
      rs.Fields("DollarSign") = xdb(RowNum, 11)
      rs.Fields("LineFeeds") = xdb(RowNum, 12)
      rs.Fields("PrintTab") = xdb(RowNum, 13)
      rs.Fields("SignRevSched") = xdb(RowNum, 14)
      rs.Fields("SignRevStmt") = xdb(RowNum, 15)
      rs.Fields("TotalLevel") = xdb(RowNum, 16)
      rs.Fields("TotalOnLedger") = xdb(RowNum, 17)
      rs.Fields("AcctType") = xdb(RowNum, 18)

      rs.Update

End Sub
