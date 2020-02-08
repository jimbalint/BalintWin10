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

Dim db As dao.Database
Dim rs As dao.Recordset

Public Function GetBatch(ByVal DataFile As String, ByVal SQL As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 1
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset(SQL)
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, N - 1, 0, 5
    rs.MoveFirst
    For ndx = 0 To N - 1
        xdb.Value(ndx, 0) = CStr(rs!BatchNumber)
        xdb.Value(ndx, 1) = gl.ShowDate(rs!Updated)
        xdb.Value(ndx, 2) = CStr(rs!FiscalYear) & "-" & CStr(rs!Period)
        xdb.Value(ndx, 3) = gl.ShowValue(rs!debits)
        xdb.Value(ndx, 4) = gl.ShowValue(rs!credits)
        xdb.Value(ndx, 5) = CStr(rs!Records)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetBatch = xdb
End Function

Public Function GetCompany(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM GLCompany ORDER BY Name")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, N - 1, 0, 2
    rs.MoveFirst
    For ndx = 0 To N - 1
        xdb.Value(ndx, 0) = CStr(rs!name)
        xdb.Value(ndx, 1) = CStr(rs!FileName)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetCompany = xdb
End Function

Public Function GetAccounts(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM GLAccount WHERE AcctType='0' ORDER BY Account")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 0, N - 1, 0, 2
    rs.MoveFirst
    For ndx = 0 To N - 1
        xdb.Value(ndx, 0) = CStr(rs!Account)
        xdb.Value(ndx, 1) = CStr(rs!Description)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetAccounts = xdb
End Function

Public Function GetDescriptions(ByVal DataFile As String) As XArrayDB
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 2
    Set db = OpenDatabase(DataFile)
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
    Set db = OpenDatabase(DataFile)
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
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM glJournal ORDER BY JournalSource")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 2
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!JournalSource)
        xdb.Value(ndx, 1) = CStr(rs!JournalName)
        xdb.Value(ndx, 2) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetJournals = xdb
End Function

Public Sub PutJournals(ByVal DataFile As String, ByRef xdb As XArrayDB)
    On Error GoTo glErr
    Dim N, ndx As Long
    Set db = OpenDatabase(DataFile)
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
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM glBranch ORDER BY BranchNumber")
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 2
    rs.MoveFirst
    For ndx = 1 To N
        
        xdb.Value(ndx, 0) = CStr(rs!BranchNumber)
        xdb.Value(ndx, 1) = CStr(rs!name)
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
    Set db = OpenDatabase(DataFile)
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        If "ADD" = xdb.Value(ndx, 2) Then
            Set rs = db.OpenRecordset("glBranch")
            rs.AddNew
            xdb.Value(ndx, 2) = rs!ID
            rs!BranchNumber = CInt(xdb.Value(ndx, 0))
            rs!name = Mid(xdb.Value(ndx, 1), 1, 60)
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
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 4
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM glHistory WHERE BatchNumber=" & CStr(BatchNumber))
    If rs.EOF = True Then GoTo glErr
    rs.MoveLast
    N = rs.RecordCount
    xdb.ReDim 1, N, 0, 4
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!Account)
        xdb.Value(ndx, 1) = CStr(rs!Amount)
        xdb.Value(ndx, 2) = CStr(rs!Reference)
        xdb.Value(ndx, 3) = CStr(rs!Description)
        xdb.Value(ndx, 4) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetHistory = xdb
    If Err.Number = 0 Then Exit Function
    MsgBox Error(Err.Number)
End Function


Public Function GetHist(ByVal DataFile As String, _
                        ByVal FiscalYear As Long, _
                        ByVal Period As Integer) As XArrayDB
    
    Dim xdb As New XArrayDB
    Dim N, ndx As Long
    On Error GoTo glErr
    xdb.ReDim 0, 0, 0, 5
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("SELECT * FROM glHistory WHERE FiscalYear = " & FiscalYear & _
                              " AND Period = " & Period)
    If rs.EOF = True Then GoTo glErr
    rs.MoveLast
    N = rs.RecordCount
    
    xdb.ReDim 1, N, 0, 5
    rs.MoveFirst
    For ndx = 1 To N
        xdb.Value(ndx, 0) = CStr(rs!JournalSource)
        xdb.Value(ndx, 1) = CStr(rs!Account)
        xdb.Value(ndx, 2) = CStr(rs!Reference)
        xdb.Value(ndx, 3) = CStr(rs!Description)
        xdb.Value(ndx, 4) = CStr(rs!Amount)
        xdb.Value(ndx, 5) = CStr(rs!ID)
        rs.MoveNext
    Next ndx
    rs.Close
glErr:
    Set GetHist = xdb
    If Err.Number = 0 Then Exit Function
    MsgBox Error(Err.Number)
End Function

Public Sub PutHist(ByVal DataFile As String, _
                   ByVal FiscalYear As Long, _
                   ByVal Period As Integer, _
                   ByRef xdb As XArrayDB)

    On Error GoTo glErr
    Dim N, ndx As Long
    Set db = OpenDatabase(DataFile)
    N = xdb.UpperBound(1)
    For ndx = 1 To N
        If IsNull(xdb.Value(ndx, 1)) Or xdb.Value(ndx, 1) = 0 Then GoTo NextN
        If "ADD" = xdb.Value(ndx, 5) Then
            Set rs = db.OpenRecordset("GLHistory")
            rs.AddNew
            xdb.Value(ndx, 5) = rs!ID
            rs!JournalSource = CByte(xdb.Value(ndx, 0))
            rs!Account = CLng(xdb.Value(ndx, 1))
            rs!Reference = xdb.Value(ndx, 2)
            rs!Description = xdb.Value(ndx, 3)
            rs!Amount = CCur(xdb.Value(ndx, 4))
            rs!FiscalYear = FiscalYear
            rs!Period = Period
            rs!BatchNumber = 0
            rs!SourceCode = 0
            rs!UpdateFlag = False
            rs!HisType = ""
            rs.Update
            rs.Close
        Else
            Set rs = db.OpenRecordset("SELECT * FROM GLHistory WHERE ID = " & xdb.Value(ndx, 5))
            If rs.RecordCount = 1 Then
                rs.Edit
                rs!JournalSource = CByte(xdb.Value(ndx, 0))
                rs!Account = CLng(xdb.Value(ndx, 1))
                rs!Reference = xdb.Value(ndx, 2)
                rs!Description = xdb.Value(ndx, 3)
                rs!Amount = CCur(xdb.Value(ndx, 4))
                rs.Update
            End If
            rs.Close
        End If
NextN:
    Next ndx
glErr:
End Sub

