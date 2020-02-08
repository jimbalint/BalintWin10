VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ccUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Login, name, Password As String    ' Field Names
Public ID, LastCompany As Long
Public LoadLastCompany As Boolean

Private mcol As New Collection

Public Function Records() As Long
    Records = mcol.Count
End Function

Public Function Item(ByVal index As Variant) As ccUsers
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item(index)
End Function

Public Sub PutRecord(ByVal ID As Long)
    On Error GoTo glErr
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase("\balint\data\glSystem.mdb")

    Set rs = db.OpenRecordset("users")
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
        !LogOn = Login
        
        If IsNull(Password) Or Password = "" Then
           !Password = " "
        Else
           !Password = Password
        End If
        !LastCompany = LastCompany
        !LoadLastCompany = LoadLastCompany
        .Update
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub DeleteRecord(ByVal ID As Long)
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("users")
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
    Dim db As DAO.Database
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Dim rs As DAO.Recordset
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
    Dim cc As New ccUsers
    With rs
        cc.name = !name & ""
        cc.Login = !LogOn & ""
        cc.Password = !Password & ""
        cc.ID = !ID + 0
        cc.LastCompany = !LastCompany + 0
        cc.LoadLastCompany = !LoadLastCompany + 0
        mcol.Add cc, cc.Login
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

