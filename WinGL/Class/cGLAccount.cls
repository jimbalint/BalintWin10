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
' 2020-01-25
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
Public FullDesc As String
Public TypeLevel As String

Public RecCt As Long

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsd As New ADODB.Recordset

Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Public Sub OpenRS()

   SQLString = "SELECT * FROM GLAccount ORDER BY Account"
   rsInit SQLString, cn, rs1

   If rs1.BOF And rs1.EOF Then
      Clear
      Exit Sub
   End If

   rs1.MoveFirst
   
   SQLString = "SELECT * FROM GLDescriptions"
   rsInit SQLString, cnDes, rsd
   
   AssignFields

End Sub

Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
       rs1.AddNew
    End If
    
    rs1.Fields("Account") = Account
    rs1.Fields("AllSchedules") = AllSchedules
    rs1.Fields("AllStatements") = AllStatements
    rs1.Fields("BranchAcct") = BranchAcct
    rs1.Fields("BSColumn") = BSColumn
    rs1.Fields("ConsAcct") = ConsAcct
    rs1.Fields("Date1") = Date1
    rs1.Fields("Date2") = Date2
    rs1.Fields("Description") = Description
    rs1.Fields("DescNumber") = DescNumber
    rs1.Fields("DollarSign") = DollarSign
    rs1.Fields("LineFeeds") = LineFeeds
    rs1.Fields("PrintTab") = PrintTab
    rs1.Fields("SignRevSched") = SignRevSched
    rs1.Fields("SignRevStmt") = SignRevStmt
    rs1.Fields("TotalLevel") = TotalLevel
    rs1.Fields("TotalOnLedger") = TotalOnLedger
    rs1.Fields("AcctType") = AcctType
    
    rs1.Update
    
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
      FullDesc = ""
      TypeLevel = ""
      
End Sub

Public Sub DeleteAll()

    Set rs1 = New ADODB.Recordset
    
'    rs1.Source = "Delete * from GLAccount"
'    Set rs1.ActiveConnection = cn
'    rs1.Open
    
    rsInit "DELETE * FROM GLAccount", cn, rs1
    
'    CloseRS
    
End Sub


Public Function DeleteRecord(ByVal lngAcct As Long) As Boolean

    Set rs1 = New ADODB.Recordset
    
    If lngAcct = 0 Then
       ErrMessage = "Account number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from GLAccount where Account = " & lngAcct
    
    rsInit SQLString, cn, rs1
    
    DeleteRecord = True
    
'    If rs1.BOF = True And rs1.EOF = True Then
'       ErrMessage = "Account not found!"
'       DeleteRecord = False
'    Else
'       ErrMessage = ""
'       DeleteRecord = True
'    End If
'
'    CloseRS
    
End Function

Public Function DeleteCurrentRecord() As Boolean
   DeleteCurrentRecord = False
'   On Error Resume Next
   rs1.Delete
   rs1.Update
   DeleteCurrentRecord = True
End Function

Public Function Records() As Long

    SQLString = "select Account from GLAccount Order By Account"
    
    rsInit SQLString, cn, rs1
    
    ' Check to make sure a record
    ' actually came back
    If rs1.EOF = True And rs1.BOF = True Then
       Records = 0
    Else
       rs1.MoveLast
       Records = rs1.RecordCount
    End If

    CloseRS

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
                  "GLAmount.Budget07, GLAmount.Budget08, GLAmount.Budget09, " & _
                  "GLAmount.Budget10, GLAmount.Budget11, GLAmount.Budget12, GLAmount.Budget13 " & _
                  "from GLAccount Left Join GLAmount on " & _
                  "(GLAccount.Account = GLAmount.Account and " & _
                  CurrFY & " = GLAmount.FiscalYear) order by GLAccount.Account"
   
   End If
   
   rsInit SQLString, cn, rs1
   
   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !", vbExclamation
      rs1.Close
      GetRecordSets = False
      CloseRS
      Exit Function
   End If

   rs1.MoveFirst
   
   SQLString = "SELECT * FROM GLDescriptions"
   rsInit SQLString, cnDes, rsd
   
   AssignFields

   GetRecordSets = True
   
   ' only open the second record set if comparative stmts
   If CompFlag Then
   
       ' get PrevFY amounts
       SQLString = "SELECT * FROM GLAmount WHERE GLAmount.FiscalYear = " & PrevFY
    
       rsInit SQLString, cn, rs2
    
       ' Check to make sure a record
       ' actually came back
       If rs2.EOF = True And rs2.BOF = True Then
    
    '      MsgBox "No AMOUNTS found !!!", vbExclamation
    '      rs2.Close
    '      Set rs2 = Nothing
    '      GetRecordSets = False
    
          Exit Function
       End If
    
       SQLString = "Account = " & CStr(rs1!Account)
    
       rs2.Find SQLString, 0, adSearchForward, 1
    
       GetRecordSets = True

   End If

End Function

Public Function GetRecordSetsNoBudget(ByVal CurrFY As Integer, ByVal PrevFY As Integer) As Boolean

   SQLString = "select GLAccount.*, GLAmount.FiscalYear, " & _
                  "GLAmount.Amount01, GLAmount.Amount02, GLAmount.Amount03, " & _
                  "GLAmount.Amount04, GLAmount.Amount05, GLAmount.Amount06, " & _
                  "GLAmount.Amount07, GLAmount.Amount08, GLAmount.Amount09, " & _
                  "GLAmount.Amount10, GLAmount.Amount11, GLAmount.Amount12, GLAmount.Amount13 " & _
                  "from GLAccount Left Join GLAmount on " & _
                  "(GLAccount.Account = GLAmount.Account and " & _
                  CurrFY & " = GLAmount.FiscalYear) order by GLAccount.Account"
                  
   rsInit SQLString, cn, rs1
   
   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !", vbExclamation
      rs1.Close
      GetRecordSetsNoBudget = False
      CloseRS
      Exit Function
   End If

   rs1.MoveFirst
   
   SQLString = "SELECT * FROM GLDescriptions"
   rsInit SQLString, cnDes, rsd
   
   AssignFields

   ' get PrevFY amounts
   SQLString = "Select * from GLAmount where GLAmount.FiscalYear = " & PrevFY
      
   rsInit SQLString, cn, rs2
      
   ' Check to make sure a record
   ' actually came back
   If rs2.EOF = True And rs2.BOF = True Then
      MsgBox "No AMOUNTS found !!!", vbExclamation
      rs2.Close
      Set rs2 = Nothing
      GetRecordSetsNoBudget = False
      Exit Function
   End If
   
   SQLString = "Account = " & CStr(rs1!Account)
   
   rs2.Find SQLString, 0, adSearchForward, 1

   GetRecordSetsNoBudget = True

End Function


Public Function GetAllAccounts() As Boolean
   
   SQLString = "SELECT * FROM GLAccount ORDER BY Account"
   
   rsInit SQLString, cn, rs1
   
   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !!!", vbExclamation
'      CloseRS
      GetAllAccounts = False
      Exit Function
   End If

   GetAllAccounts = True

   rs1.MoveFirst
   
   SQLString = "SELECT * FROM GLDescriptions"
   rsInit SQLString, cnDes, rsd
   
   AssignFields

   ' open the gldesc record set
'   Set GLDescription = New cGLDescription
'   GLDescription.OpenRS

End Function


Public Function Find(ByVal Acct As Long) As Boolean
   
   SQLString = "Account = " & Acct
   
   rs1.Find SQLString, 0, adSearchForward, 1
   
   If rs1.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

Public Function FindFirst() As Boolean
   
   rs1.MoveFirst
   
   If rs1.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
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

   If CompFlag Then
      SQLString = "Account = " & CStr(rs1!Account)
      rs2.Find SQLString, 0, adSearchForward, 1
   End If
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
    DescNumber = rs1!DescNumber
    
    DollarSign = rs1!DollarSign
    LineFeeds = rs1!LineFeeds
    PrintTab = rs1!PrintTab
    SignRevSched = rs1!SignRevSched
    SignRevStmt = rs1!SignRevStmt
    TotalLevel = rs1!TotalLevel
    TotalOnLedger = rs1!TotalOnLedger
    AcctType = rs1!AcctType
    Description = Trim(rs1!Description & "")
    
''' wtF
    FullDesc = ""
    TypeLevel = ""

    FullDesc = GetDesc()
    TypeLevel = AcctType
    If TotalLevel <> 0 Then
        TypeLevel = TypeLevel & TotalLevel
    End If
    
    RecCt = rs1.RecordCount
    
'  If rs1!Account = 2151 Then MsgBox Description

End Sub

Public Function GetDesc() As String
    
    If rs1!DescNumber = 0 Then
       GetDesc = RTrim(rs1!Description) & ""
    Else
       rsd.Find "Number = " & rs1!DescNumber, 0, adSearchForward, 1
       If rsd.EOF Then
          GetDesc = RTrim(rs1!Description & "")
       Else
          GetDesc = RTrim(rs1!Description & "") & RTrim(rsd!Description & "")
       End If
    End If
       
End Function

Public Function GetCurrAmount(ByVal bytPeriod1 As Byte, _
                              ByVal bytPeriod2 As Byte) As Currency
                     
    ' did a left join for rs1 record set
    ' if any value is null then there was no GLAmount record for the
    ' GLAccount record --> return zero
    If IsNull(rs1!Amount01) Then
       GetCurrAmount = 0
       Exit Function
    End If
                     
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
    If Not CompFlag Then
       GetPrevAmount = 0
       Exit Function
    End If
    
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
      Exit Function
   Else
      AssignFields
      GetAccount = True
   End If
   
   If CompFlag Then
      SQLString = "Account = " & CStr(rs1!Account)
      rs2.Find SQLString, 0, adSearchForward, 1
   End If

End Function
Public Function GetBudget(ByVal bytPeriod1 As Byte, _
                              ByVal bytPeriod2 As Byte) As Currency
                     
                         
    ' did a left join for rs1 record set
    ' if any value is null then there was no GLAmount record for the
    ' GLAccount record --> return zero
    If IsNull(rs1!Budget01) Then
       GetBudget = 0
       Exit Function
    End If
                     
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
      SQLString = "SELECT * FROM GLAccount ORDER BY Account"
   
   Else
      SQLString = "SELECT * FROM GLAccount WHERE (GLAccount.Account >= " & LoAcct & " AND " & _
                  "GLAccount.Account <= " & HiAcct & ") ORDER BY Account"
   End If
               
   rsInit SQLString, cn, rs1
               
   ' Check to make sure a record
   ' actually came back
   If rs1.EOF = True And rs1.BOF = True Then
      MsgBox "No accounts found !", vbExclamation
      CloseRS
      GetAcctRecSet = False
      Exit Function
   End If

   rs1.MoveFirst
   
   SQLString = "SELECT * FROM GLDescriptions"
   rsInit SQLString, cnDes, rsd
   
   AssignFields

   GetAcctRecSet = True

End Function

Public Function GetAcctsByType(ByVal aType As String) As Boolean

    SQLString = " SELECT * FROM GLAccount WHERE AcctType = '" & aType & "' ORDER BY Account"
    rsInit SQLString, cn, rs1
    
    If rs1.EOF = True And rs1.BOF = True Then
        MsgBox "No accounts found !", vbExclamation
        CloseRS
        GetAcctsByType = False
        Exit Function
    End If
    
    rs1.MoveFirst
   
    SQLString = "SELECT * FROM GLDescriptions"
    rsInit SQLString, cnDes, rsd
   
    AssignFields

    GetAcctsByType = True

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


'Public Function QueryByAccount(ByVal Acct As Long) As Boolean
'
'   Set cmd = New ADODB.Command
'   Set cmd.ActiveConnection = cn
'   cmd.CommandType = adCmdStoredProc
'   cmd.CommandText = "qryGLAccountByAccount"
'
'   Set param = cmd.CreateParameter("qAccount", adNumeric, adParamInput, , Acct)
'   cmd.Parameters.Append param
'
'   Set rs1 = cmd.Execute
'
'   If rs1.BOF And rs1.EOF Then
'      QueryByAccount = False
'      CloseRS
'      Exit Function
'   End If
'
'   rs1.MoveFirst
'
'   SQLString = "SELECT * FROM GLDescriptions"
'   rsInit SQLString, cnDes, rsd
'
'   AssignFields
'
'   QueryByAccount = True
'
'   CloseRS
'
'End Function


Public Sub CloseRS()
   rs1.Close
   Set rs1 = Nothing
End Sub


