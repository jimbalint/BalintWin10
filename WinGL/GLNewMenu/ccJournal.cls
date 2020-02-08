VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public JournalName As String    ' Field Names
Public ID As Long
Public JournalSource As Long

Private mcol As New Collection

Public Function Records() As Long
    Records = mcol.Count
End Function

Public Function Item(ByVal index As Variant) As ccJournal
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)
End Function

Public Function PutRecord(ByVal DataFile As String, ByVal curID As Long) As Boolean
    PutRecord = True
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("glJournal")  ' Like share in Clarion
    With rs
        If curID = 0 Then           ' New Record (ADD)
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else                        ' Existing (EDIT)
            .index = "PrimaryKey"
            .Seek "=", curID
            .Edit
        End If
        !JournalSource = JournalSource
        !JournalName = JournalName
        .Update
    End With
    Exit Function
glErr:
    PutRecord = False           ' Inform calling procedure of failure
    MsgBox Error(Err.Number)
End Function

Public Function AddRecord(ByVal DataFile As String) As Long
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("glJournal")  ' Like share in Clarion
    With rs
        If curID = 0 Then           ' New Record (ADD)
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else                        ' Existing (EDIT)
            .index = "PrimaryKey"
            .Seek "=", curID
            .Edit
        End If
        !JournalSource = JournalSource
        !JournalName = JournalName
        .Update
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function CurrentID(ByVal DataFile As String, ByVal JSource As Long) As Long
    CurrentID = 0           ' used to find the primary field value
    On Error GoTo glErr     ' from OnMove function in List boxes
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("glJournal")
    With rs
        .index = "JournalKey"
        .Seek "=", JSource
        CurrentID = .Fields("ID").Value
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Sub DeleteRecord(ByVal DataFile As String, ByVal ID As Long)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("glJournal")
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        .Delete
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Function GetSQL(ByVal DataFile As String, ByVal SQL As String) As Integer
    GetSQL = 0
    On Error GoTo glErr
    Free
    Dim db As Database
    Set db = OpenDatabase(DataFile)
    Dim rs As Recordset
    Set rs = db.OpenRecordset(SQL)
    If rs.RecordCount = 0 Then Exit Function
    rs.MoveFirst
    Do Until rs.EOF = True
        Add rs
        GetSQL = GetSQL + 1
        rs.MoveNext
    Loop
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Private Function Add(rs As Recordset)
    Dim cc As New ccJournal
    With rs
        cc.JournalName = !JournalName & ""
        cc.JournalSource = CInt(!JournalSource) + 0
        cc.ID = !ID + 0
        mcol.Add cc, CStr(cc.JournalSource)
    End With
End Function

Private Function Free()
    Dim index As Variant
    For index = mcol.Count To 1 Step -1
        mcol.Remove index
    Next index
End Function

Private Sub Delete(ByVal index As Variant)
    mcol.Remove index
End Sub


