VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public name, address1, address2, address3, city, state, FileName, zipcode As String
Public lastUpdate, lastClose As Date
Public SuspAcct, NetProfitAcct, FirstPAcct, PctBaseAcct, LastBatch As Long
Public nPeriods, FirstPeriod, curPeriod As Byte
Public curFiscalYear, FirstFiscalYear As Integer
Private mcol As New Collection

Public Function Records() As Long
    Records = mcol.Count
End Function

Public Function Item(ByVal index As Variant) As ccCompany
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)
End Function

Public Sub PutRecord(ByVal ID As Long)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glcompany")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !name = name
        !FileName = FileName & ""
        !address1 = address1 & ""
        !address2 = address2 & ""
        !address3 = address3 & ""
        !city = city & ""
        !state = state & ""
        !zipcode = zipcode & ""
        !LastBatch = LastBatch + 0
        !curFiscalYear = curFiscalYear
        !curPeriod = curPeriod
        !FirstFiscalYear = FirstFiscalYear
        !NumberPds = nPeriods
        !lastUpdate = lastUpdate
        !lastClose = lastClose
        .Update
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub DeleteRecord(ByVal ID As Long)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glcompany")
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        .Delete
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Function GetSQL(ByVal SQL As String) As Integer
    GetSQL = 0
    On Error GoTo glErr
    Free
    Dim db As Database
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
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
    Dim cc As New ccCompany
    With rs
        cc.ID = !ID + 0
        cc.name = !name & ""
        cc.FileName = !FileName & ""
        cc.address1 = !address1 & ""
        cc.address2 = !address2 & ""
        cc.address3 = !address3 & ""
        cc.city = !city & ""
        cc.state = !state & ""
        cc.zipcode = !zipcode & ""
        cc.LastBatch = !LastBatch + 0
        cc.curFiscalYear = !curFiscalYear + 0
        cc.curPeriod = !curPeriod + 0
        cc.FirstFiscalYear = !FirstFiscalYear + 0
        cc.nPeriods = !NumberPds + 0
        cc.lastUpdate = !lastUpdate + 0
        cc.lastClose = !lastClose
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


