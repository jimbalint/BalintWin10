VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, batchNumber, createUser, updateUser, nRecords As Long
Public fiscalYear, JournalSource As Integer
Public debits, credits As Currency
Public period As Byte
Public Created, Updated As Date

Private mcol As New Collection

Public Function Records() As Long
    Records = mcol.Count
End Function

Public Function Item(ByVal index As Variant) As ccBatch
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)
End Function

Public Sub PutRecord(ByVal ID As Long, ByVal FileName As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("glBatch")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !debits = debits
        !credits = credits
        !batchNumber = batchNumber
        !Created = Created
        !createUser = createUser
        !fiscalYear = fiscalYear
        !Records = nRecords
        !period = period
        !Updated = Now
        !updateUser = updateUser
        !JournalSource = JournalSource
        .Update
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub DeleteRecord(ByVal ID As Long, ByVal FileName As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("glBatch")
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        .Delete
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Function GetSQL(ByVal SQL As String, ByVal DataFile As String) As Integer
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
    Dim cc As New ccBatch
    With rs
        cc.ID = !ID + 0
        cc.debits = !debits + 0
        cc.credits = !credits + 0
        cc.batchNumber = !batchNumber + 0
        cc.Created = !Created
        cc.createUser = !createUser + 0
        cc.fiscalYear = !fiscalYear + 0
        cc.nRecords = !Records + 0
        cc.period = !period + 0
        cc.Updated = !Updated + 0
        cc.updateUser = !updateUser + 0
        cc.JournalSource = !JournalSource + 0
        mcol.Add cc, CStr(cc.batchNumber)
    End With
    Set Add = cc
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

