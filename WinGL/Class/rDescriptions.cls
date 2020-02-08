VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long       ' Field Names
Public Number As Long
Public Description As String

Public Function GetDescription(ByVal nDesc As Long) As String
    GetDescription = ""
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    Set rs = db.OpenRecordset("SELECT * FROM glDescriptions WHERE number=" & CStr(nDesc))
    rs.MoveFirst
    GetDescription = rs!Description
    Exit Function
glErr:
End Function
