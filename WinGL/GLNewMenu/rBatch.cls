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
Public fiscalyear, JournalSource As Integer
Public debits, credits As Currency
Public period As Byte
Public Created, Updated As Date

Public Sub AddRecord(ByVal lastbatch As Long, ByVal fiscalyear As Integer, ByVal curPeriod As Byte, ByVal filename As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(filename)
    Set rs = db.OpenRecordset("SELECT * FROM glBatch")
    With rs
        .AddNew
        !BatchNumber = lastbatch
        !debits = 0
        !credits = 0
        !Created = Now
        !createUser = curUser
        !fiscalyear = fiscalyear
        !Records = 0
        !period = curPeriod
        !Updated = Now
        !updateUser = curUser
        .Update
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub PutRecord(ByVal BatchNum As Long, ByVal filename As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(filename)
    Set rs = db.OpenRecordset("SELECT * FROM glBatch WHERE BatchNumber=" & BatchNum)
    With rs
        .Edit
        !debits = debits
        !credits = credits
        !BatchNumber = BatchNumber
        !Created = Created
        !createUser = createUser
        !fiscalyear = fiscalyear
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

Public Sub GetRecord(ByVal BatchNum As Long, ByVal filename As String)
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(filename)
    Set rs = db.OpenRecordset("SELECT * FROM glBatch WHERE BatchNumber=" & BatchNum)
    With rs
        ID = !ID + 0
        debits = !debits + 0
        credits = !credits + 0
        BatchNumber = !BatchNumber + 0
        Created = !Created
        createUser = !createUser + 0
        fiscalyear = !fiscalyear + 0
        nRecords = !Records + 0
        period = !period + 0
        Updated = !Updated + 0
        updateUser = !updateUser + 0
        JournalSource = !JournalSource + 0
    End With
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

