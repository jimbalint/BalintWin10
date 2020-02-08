VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, BatchNumber, createUser, updateUser, nRecords As Long
Public fiscalYear, JournalSource As Integer
Public debits, credits As Currency
Public period As Byte
Public Created, Updated As Date

Public Sub AddBatch(ByRef LastBatch As Long, ByVal fiscalYear As Integer, ByVal curPeriod As Byte, ByVal FileName As String)
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(Name:=FileName, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    ' get the last batch number from the data
    Set rs = db.OpenRecordset("SELECT * FROM glBatch ORDER BY BatchNumber DESC")
    If rs.RecordCount = 0 Then
        LastBatch = 0
    Else
        rs.MoveFirst
        LastBatch = rs!BatchNumber
    End If
    
    With rs
        .AddNew
        !BatchNumber = LastBatch + 1
        !debits = 0
        !credits = 0
        !Created = Now
        !createUser = curUser
        !fiscalYear = fiscalYear
        !Records = 0
        !JournalSource = 0
        !period = curPeriod
        !Updated = Now
        !updateUser = curUser
        .Update
    End With
    LastBatch = LastBatch + 1
    
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub PutRecord(ByVal BatchNum As Long, ByVal FileName As String)
'    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(Name:=FileName, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glBatch WHERE BatchNumber=" & BatchNum)
    With rs
        .Edit
        !debits = debits
        !credits = credits
        !BatchNumber = BatchNumber
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

Public Function GetBatch(ByVal BatchNum As Long, ByVal FileName As String) As Boolean
    GetBatch = False
'    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(Name:=FileName, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM glBatch WHERE BatchNumber=" & BatchNum)
    With rs
        ID = !ID + 0
        debits = !debits + 0
        credits = !credits + 0
        BatchNumber = !BatchNumber + 0
        Created = !Created
        createUser = !createUser + 0
        fiscalYear = !fiscalYear + 0
        nRecords = !Records + 0
        period = !period + 0
        Updated = !Updated + 0
        updateUser = !updateUser + 0
        JournalSource = !JournalSource + 0
    End With
    GetBatch = True
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Sub DeleteRecord(ByVal BatchNum As Long, ByVal FileName As String)
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = OpenDatabase(Name:=FileName, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    db.Execute ("DELETE * FROM GLBatch WHERE BatchNumber = " & BatchNum)
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

