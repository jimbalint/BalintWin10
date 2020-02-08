VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim adoRS As ADODB.Recordset
Dim adoCN As ADODB.Connection

Public Function GetBatch(ByVal DataFile As String, ByVal SQL As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 1, 0, 0, 7
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset(SQL)
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 7
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!BatchNumber)
        xdb.Value(ndx, 1) = ShowDate(rs!Updated)
        xdb.Value(ndx, 2) = CStr(rs!fiscalYear) & "-" & Format(rs!period, "00")
        xdb.Value(ndx, 3) = ShowValue(rs!debits)
        xdb.Value(ndx, 4) = ShowValue(rs!credits)
        xdb.Value(ndx, 5) = CStr(rs!Records)
        xdb.Value(ndx, 6) = CStr(rs!JournalSource)
        xdb.Value(ndx, 7) = CStr(rs!updateUser)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetBatch = xdb
End Function
Public Function GetBatchADO(ByVal DataFile As String, ByVal SQL As String) As XArrayDB
    
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
'    On Error GoTo glErr
    xdb.ReDim 1, 0, 0, 7
    
 Set adoCN = New ADODB.Connection
 adoCN.Provider = "Microsoft.Jet.OLEDB.4.0"
 adoCN.ConnectionString = DataFile
 adoCN.Mode = adModeReadWrite
 adoCN.Open
    
    rsInit SQL, adoCN, adoRS
    
    adoRS.MoveLast
    N = adoRS.RecordCount
    xdb.ReDim 1, N, 0, 7
    adoRS.MoveFirst
    
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(adoRS!BatchNumber)
        xdb.Value(ndx, 1) = ShowDate(adoRS!Updated)
        xdb.Value(ndx, 2) = CStr(adoRS!fiscalYear) & "-" & Format(adoRS!period, "00")
        xdb.Value(ndx, 3) = ShowValue(adoRS!debits)
        xdb.Value(ndx, 4) = ShowValue(adoRS!credits)
        xdb.Value(ndx, 5) = CStr(adoRS!Records)
        xdb.Value(ndx, 6) = CStr(adoRS!JournalSource)
        xdb.Value(ndx, 7) = CStr(adoRS!updateUser)
        adoRS.MoveNext
    Next ndx
    adoRS.Close
glErr:
    Set GetBatchADO = xdb
End Function

Public Function GetCompany(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM GLCompany ORDER BY Name")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, N - 1, 0, 2
    rs.MoveFirst
    For ndx = 0 To N - 1
        xdb.Value(ndx, 0) = CStr(rs!Name)
        xdb.Value(ndx, 1) = CStr(rs!FileName)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetCompany = xdb
End Function

Public Function GetAccounts(ByVal DataFile As String, ByVal strType As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx, dbx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM GLAccount WHERE AcctType= '" & strType & "' ORDER BY Account")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, N - 1, 0, 2
    rs.MoveFirst
    ndx = -1
    For dbx = 0 To N - 1
        If strType = "ALL" Or strType = rs!AcctType Then
            ndx = ndx + 1
            xdb.Value(ndx, 0) = CStr(rs!Account)
            xdb.Value(ndx, 1) = CStr(rs!Description)
            xdb.Value(ndx, 2) = CStr(rs!ID)
        End If
        rs.MoveNext
    Next dbx
    rs.Close
glErr:
    Set GetAccounts = xdb
End Function

Public Function Get0Accounts(ByVal DataFile As String, ByVal strType As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx, dbx As Long
    Dim LastType As String
    
    Dim SysDB As DAO.Database
    Dim rsd As DAO.Recordset
    
    LastType = "B"   ' init to balance sheet accts
    
'    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    If BalintFolder = "" Then
        Set SysDB = OpenDatabase("\Balint\Data\GLSystem.mdb")
    Else
        Set SysDB = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    
    Set rsd = SysDB.OpenRecordset("SELECT * FROM GLDescriptions ORDER BY Number")
    
    Set rs = db.OpenRecordset("SELECT * FROM GLAccount ORDER BY Account")
    
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, 1, 0, 2
    rs.MoveFirst
    ndx = -1
    
    For dbx = 0 To N - 1
        
        If InStr(1, "NIE", rs!AcctType, vbTextCompare) <> 0 Then
           If rs!AcctType = "N" Then
              LastType = "I"
           Else
              LastType = rs!AcctType
           End If
        End If
        
        ' "A" for ALL else    B=Bal Sht  I=Inc  E=Expense
        If (strType = "A" Or strType = LastType) And rs!AcctType = "0" Then
            xdb.AppendRows (1)
            ndx = ndx + 1
            xdb.Value(ndx, 0) = CStr(rs!Account)
            
            If IsNumeric(rs!DescNumber) And rs!DescNumber <> 0 Then
               rsd.MoveLast
               rsd.FindFirst ("Number = " & rs!DescNumber)
               If rsd.NoMatch Then
                  xdb.Value(ndx, 1) = CStr("Desc#: " & rs!DescNumber)
               Else
                  xdb.Value(ndx, 1) = CStr(rsd!Description)
               End If
            ElseIf IsNull(rs!Description) Then
               xdb.Value(ndx, 1) = " "
            Else
               xdb.Value(ndx, 1) = CStr(rs!Description)
            End If
            
            xdb.Value(ndx, 2) = CStr(rs!ID)
        End If
        rs.MoveNext
    
    Next dbx
    rs.Close
glErr:
    Set Get0Accounts = xdb
End Function



Public Function GetDescriptions(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glDescriptions ORDER BY Number")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 2
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!Number)
        xdb.Value(ndx, 1) = CStr(rs!Description)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetDescriptions = xdb
End Function

Public Sub PutDescriptions(ByVal DataFile As String, ByRef xdb As XArrayDB)
    On Error GoTo glErr
    Dim N, ndx As Long
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        If "ADD" = xdb.Value(ndx, 2) Then
            Set rs = db.OpenRecordset("glDescriptions")
            rs.AddNew
            xdb.Value(ndx, 2) = rs!ID
            rs!Number = CLng(xdb.Value(ndx, 0))
            rs!Description = Mid(xdb.Value(ndx, 1), 1, 60)
            rs.Update
            rs.Close
        Else
            Set rs = db.OpenRecordset("SELECT * FROM glDescriptions WHERE ID = " & xdb.Value(ndx, 2))
            If rs.RecordCount = 1 Then
                rs.Edit
                rs!Number = CLng(xdb.Value(ndx, 0))
                rs!Description = xdb.Value(ndx, 1)
                rs.Update
            End If
            rs.Close
        End If
    Next ndx
glErr:
End Sub

Public Function GetJournals(ByVal DataFile As String) As XArrayDB
    
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glJournal ORDER BY JournalSource")
    
    ' init if not assigned
    If rs.BOF And rs.EOF Then
       xdb.ReDim 1, 9, 0, 2
       For ndx = 1 To 9
           xdb.Value(ndx, 0) = CStr(ndx)
           xdb.Value(ndx, 1) = "Jnl " & CStr(ndx)
           xdb.Value(ndx, 2) = 0
       Next ndx
       Set GetJournals = xdb
       Exit Function
    End If
    
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 2
    rs.MoveFirst
    
    For ndx = 1 To N
        If Not IsNull(rs!JournalSource) Then
           xdb.Value(ndx, 0) = CStr(rs!JournalSource)
           If Not IsNull(rs!JournalName) Then
              xdb.Value(ndx, 1) = CStr(rs!JournalName)
           Else
              xdb.Value(ndx, 1) = "Jnl " & rs!JournalSource
           End If
           xdb.Value(ndx, 2) = CStr(rs!ID)
        Else
           xdb.Value(ndx, 0) = ""
           xdb.Value(ndx, 1) = ""
           xdb.Value(ndx, 2) = ""
        End If
    
        rs.MoveNext
    
    Next ndx
    rs.Close
glErr:
    Set GetJournals = xdb

End Function

Public Sub PutJournals(ByVal DataFile As String, ByRef xdb As XArrayDB)
    On Error GoTo glErr
    Dim N, ndx As Long
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        If "ADD" = xdb.Value(ndx, 2) Then
            Set rs = db.OpenRecordset("glJournal")
            rs.AddNew
            xdb.Value(ndx, 2) = rs!ID
            rs!JournalSource = CLng(xdb.Value(ndx, 0))
            rs!JournalName = Mid(xdb.Value(ndx, 1), 1, 60)
            rs.Update
            rs.Close
        Else
            Set rs = db.OpenRecordset("SELECT * FROM glJournal WHERE ID = " & xdb.Value(ndx, 2))
            If rs.RecordCount = 1 Then
                rs.Edit
                rs!JournalSource = CLng(xdb.Value(ndx, 0))
                rs!JournalName = xdb.Value(ndx, 1)
                rs.Update
                rs.Close
            End If
        End If
    Next ndx
glErr:
End Sub

Public Function GetBranch(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glBranch ORDER BY BranchNumber")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 2
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!BranchNumber)
        xdb.Value(ndx, 1) = CStr(rs!Name)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetBranch = xdb
    If Err.Number = 0 Then Exit Function
    MsgBox Error(Err.Number)
End Function

Public Sub PutBranch(ByVal DataFile As String, ByRef xdb As XArrayDB)
    On Error GoTo glErr
    Dim N, ndx As Long
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        If "ADD" = xdb.Value(ndx, 2) Then
            Set rs = db.OpenRecordset("glBranch")
            rs.AddNew
            xdb.Value(ndx, 2) = rs!ID
            rs!BranchNumber = CInt(xdb.Value(ndx, 0))
            rs!Name = Mid(xdb.Value(ndx, 1), 1, 60)
            rs.Update
            rs.Close
        Else
            Set rs = db.OpenRecordset("SELECT * FROM glBranch WHERE ID = " & xdb.Value(ndx, 2))
            If rs.RecordCount = 1 Then
                rs.Edit
                rs!JournalSource = CLng(xdb.Value(ndx, 0))
                rs!JournalName = xdb.Value(ndx, 1)
                rs.Update
                rs.Close
            End If
        End If
    Next ndx
glErr:
End Sub

Public Function Locator(ByVal Value As String, ByVal x As XArrayDB, ByVal col As Long, ByVal start As Long) As Long
    Locator = start
    If Value = "" Then Exit Function
    Dim sublen, ndx As Long
    sublen = Len(Value)
    If sublen = 1 Then start = x.LowerBound(1)
    For ndx = start To x.UpperBound(1)
        If Mid(x.Value(ndx, col), 1, sublen) = CStr(Value) Then
            Locator = ndx
            Exit Function
        End If
    Next ndx
End Function

Public Function GetHistory(ByVal DataFile As String, ByVal BatchNumber As Long) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx, Acct As Long
    Dim x As String
'    On Error GoTo glErr
    xdb.ReDim 1, 0, 0, 6
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glHistory WHERE BatchNumber=" & CStr(BatchNumber))
    If rs.EOF = True Then GoTo glErr
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 6
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!Account + 0)
        xdb.Value(ndx, 1) = glAccountName(rs!Account, x)
        xdb.Value(ndx, 2) = CStr(rs!Reference & "")
        xdb.Value(ndx, 3) = CStr(rs!Description & "")
        xdb.Value(ndx, 4) = CStr(rs!Amount + 0)
        xdb.Value(ndx, 5) = CStr(rs!ID + 0)
        xdb.Value(ndx, 6) = CStr(rs!PostDate + 0)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetHistory = xdb
    If Err.Number = 0 Then Exit Function
    MsgBox Error(Err.Number)
End Function

Public Function PutHistory(ByVal DataFile As String, _
                           ByVal BatchNumber As Long, _
                           ByVal SourceCode As Byte, _
                           ByVal JournalSource As Byte, _
                           ByVal fiscalYear As Integer, _
                           ByVal period As Byte, _
                           ByRef xdb As XArrayDB) As Long
    
    PutHistory = 0
    Dim N, ndx, numAcct As Long
'    On Error GoTo glErr
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    
    ' delete all records for batch first
    db.Execute ("DELETE * FROM GLHistory WHERE BatchNumber = " & CStr(BatchNumber))
    
    ' add all as new
    Set rs = db.OpenRecordset("SELECT * FROM glHistory")
    
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        numAcct = CLng(xdb.Value(ndx, 0))
        If numAcct > 0 Or numAcct = -1 Then
            
'            If xdb.Value(ndx, 5) = 0 Then
'                Set rs = db.OpenRecordset("SELECT * FROM glHistory")
'                rs.AddNew
'                xdb.Value(ndx, 5) = CStr(rs!ID)
'                rs!BatchNumber = BatchNumber
'                rs!SourceCode = SourceCode
'                rs!HisType = "0"
'                rs!UpdateFlag = False
'            Else
'                Set rs = db.OpenRecordset("SELECT * FROM glHistory WHERE ID = " & xdb.Value(ndx, 5))
'                rs.Edit
'            End If
            
            rs.AddNew
            xdb.Value(ndx, 5) = CStr(rs!ID)
            rs!BatchNumber = BatchNumber
            rs!SourceCode = SourceCode
            
            If JournalSource >= 100 Then
               rs!HisType = "B"
            Else
               rs!HisType = "A"
            End If
            
            rs!UpdateFlag = False
            
            If IsEmpty(xdb.Value(ndx, 2)) Then xdb.Value(ndx, 2) = " "
            If IsEmpty(xdb.Value(ndx, 3)) Then xdb.Value(ndx, 3) = " "
            
' skip null amount records
'            If IsEmpty(xdb.Value(ndx, 4)) Then xdb.Value(ndx, 4) = "0"
            
            If IsEmpty(xdb.Value(ndx, 4)) Then GoTo NextHist
            
            rs!Account = CLng(xdb.Value(ndx, 0))
            rs!Reference = Left(xdb.Value(ndx, 2), 20)
            rs!Description = Left(xdb.Value(ndx, 3), 20)
            rs!Amount = CCur(xdb.Value(ndx, 4))
            rs!PostDate = CDate(xdb.Value(ndx, 6))
            rs!fiscalYear = fiscalYear
            rs!period = period
            rs!JournalSource = JournalSource
            rs!SourceCode = SourceCode
            rs.Update
            PutHistory = PutHistory + 1
        
        End If
    
NextHist:
    Next ndx
    
    rs.Close
    
    Exit Function

glErr:
    
    MsgBox Error(Err.Number)
End Function

