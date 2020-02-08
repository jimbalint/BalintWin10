VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, BatchNumber, AccountNumber As Long
Public Amount As Currency
Public SourceCode As Byte
Public Reference, Description, Histype As String
Public Updated As Boolean

Private mcol As New Collection

Public Function Records() As Long
    Records = mcol.Count
End Function

Public Function Item(ByVal index As Variant) As ccHistory
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)
End Function

Public Function PutRecord(ByVal ID As Long, ByVal FileName As String) As Long
    PutRecord = 0
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("glHistory")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !Amount = Amount
        !BatchNumber = BatchNumber
        !Account = AccountNumber
        !SourceCode = SourceCode
        !Reference = Reference
        !Description = Description
        !Histype = Histype
        !UpdateFlag = Updated
        .Update
        PutRecord = !ID
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Sub DeleteRecord(ByVal ID As Long, ByVal FileName As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("glHistory")
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
    Dim cc As New ccHistory
    With rs
        cc.ID = !ID + 0
        cc.Amount = !Amount + 0
        cc.AccountNumber = !Account + 0
        cc.BatchNumber = !BatchNumber + 0
        cc.SourceCode = !SourceCode
        cc.Reference = !Reference
        cc.Description = !Description
        cc.Histype = !Histype
        cc.Updated = !UpdateFlag
        mcol.Add cc, CStr(cc.ID)
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


