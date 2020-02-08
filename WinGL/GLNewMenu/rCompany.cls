VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public name, address1, address2, address3, city, state, FileName, zipcode As String
Public lastUpdate, lastClose As Long
Public SuspAcct, NetProfitAcct, FirstPAcct, PctBaseAcct, LastBatch, RetEarnAcct As Long
Public NumberPds, FirstPeriod, curPeriod, SubDigits As Byte
Public curFiscalYear, FirstFiscalYear As Integer
Public LowBranch, HiBranch, LowConsolidated, HiConsolidated As Long

Public Function PutRecord(ByVal ID As Long) As Long
    On Error GoTo glErr
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
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
        !FileName = FileName
        !address1 = address1
        !address2 = address2
        !address3 = address3
        !city = city
        !state = state
        !zipcode = zipcode
        !LastBatch = LastBatch
        !curFiscalYear = curFiscalYear
        !curPeriod = curPeriod
        !FirstFiscalYear = FirstFiscalYear
        !NumberPds = NumberPds
        !lastUpdate = lastUpdate
        !lastClose = lastClose
        !SuspAcct = SuspAcct
        !NetProfitAcct = NetProfitAcct
        !FirstPAcct = FirstPAcct
        !PctBaseAcct = PctBaseAcct
        !FirstPeriod = FirstPeriod
        !RetEarnAcct = RetEarnAcct
        !SubDigits = SubDigits
        !LowBranch = LowBranch
        !HiBranch = HiBranch
        !LowConsolidated = LowConsolidated
        !HiConsolidated = HiConsolidated
        PutRecord = !ID
        .Update
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function GetRecord(ByVal ID As Long) As Boolean
    
    On Error GoTo glErr
    
    GetRecord = False
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Set rs = db.OpenRecordset("glcompany")
    
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        name = !name & ""
        FileName = !FileName & ""
        address1 = !address1 & ""
        address2 = !address2 & ""
        address3 = !address3 & ""
        city = !city & ""
        state = !state & ""
        zipcode = !zipcode & ""
        LastBatch = !LastBatch + 0
        curFiscalYear = !curFiscalYear + 0
        curPeriod = !curPeriod + 0
        FirstFiscalYear = !FirstFiscalYear + 0
        NumberPds = !NumberPds + 0
        lastUpdate = !lastUpdate + 0
        lastClose = !lastClose + 0
        SuspAcct = !SuspAcct + 0
        NetProfitAcct = !NetProfitAcct + 0
        FirstPAcct = !FirstPAcct + 0
        PctBaseAcct = !PctBaseAcct + 0
        FirstPeriod = !FirstPeriod + 0
        RetEarnAcct = !RetEarnAcct
        SubDigits = !SubDigits
        GetRecord = True
        LowBranch = !LowBranch
        HiBranch = !HiBranch
        LowConsolidated = !LowConsolidated
        HiConsolidated = !HiConsolidated
    End With
    
    GetRecord = True
    
    Exit Function
glErr:
    
'    MsgBox Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "Windows GL"

End Function

