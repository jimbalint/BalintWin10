VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long       ' Field Names
Public number As Long
Public description As String

Private mcol As New Collection      ' Used for multiple records

Public Function Records() As Long
    Records = mcol.Count            ' Number of Records
End Function

Public Function Item(ByVal index As Variant) As ccDescriptions
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)     ' Default Behavior Num=Record Str=Find
End Function

Public Sub PutRecord(ByVal ID As Long)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glDescriptions")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !number = number
        !description = description
        .Update
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Public Sub DeleteRecord(ByVal ID As Long)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glDescriptions")
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        .Delete
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Public Function GetSQL(ByVal SQL As String) As Integer
    GetSQL = 0
    On Error GoTo glErr
    Free
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
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
    MsgBox Error(Err.number)
End Function

Private Function Add(rs As Recordset)
    Dim cc As New ccDescriptions
    With rs
        cc.description = !description & ""
        cc.number = !number + 0
        cc.ID = !ID + 0
        mcol.Add cc, CStr(cc.number)
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


