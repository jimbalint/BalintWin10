VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Field List for glHistory:
Public ID, Account, batchNumber As Long
Public Amount As Currency
Public Reference, Description, Histype As String
Public SourceCode As Byte
Public UpdateFlag As Boolean

Dim db As Database
Dim rs As Recordset

Public Function AddHistory(ByVal DataFile As String, ByVal batchNumber As Long, ByVal SourceCode As Byte, ByVal Histype As String) As Long
    On Error GoTo glErr
    Set db = OpenDatabase(DataFile)
    Set rs = db.OpenRecordset("glHistory")
    rs.AddNew
    rs!batchNumber = batchNumber
    rs!SourceCode = SourceCode
    rs!Histype = Histype
    rs!UpdateFlag = False
    AddHistory = rs!ID
    rs.Update
    rs.Close
glErr:
    If Err.Number = 0 Then Exit Function
    AddHistory = 0
    MsgBox Error(Error.Number)
End Function

