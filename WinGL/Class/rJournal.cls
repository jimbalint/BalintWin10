VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public Name As String
Public Source As Integer

Public Function GetJournal(ByVal nSource As Integer, ByVal DataFile As String)
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(Name:=DataFile, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glJournal WHERE JournalSource=" & CStr(nSource))
    rs.MoveFirst
    ID = rs!ID
    Name = rs!JournalName
    Source = rs!JournalSource
    Exit Function
glErr:
    ID = 0
    Name = ""
    Source = 0
End Function

