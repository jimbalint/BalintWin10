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
Public Name, Address1, Address2, Address3, City, State, FileName, ZipCode As String
Public LastUpdate, LastClose As Long
Public SuspAcct, NetProfitAcct, FirstPAcct, PctBaseAcct, LastBatch, RetEarnAcct As Long
Public NumberPds, FirstPeriod, CurPeriod, SubDigits As Byte
Public CurFiscalYear, FirstFiscalYear As Integer
Public cID As Long

Dim rs As New ADODB.Recordset
Public Sub OpenRS()

   SQLString = "SELECT * FROM GLCompany"
   rsInit SQLString, cnDES, rs

   If rs.BOF And rs.EOF Then
      ' Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   ' AssignFields

End Sub

Public Function MonthName(ByVal mon As Byte, ByVal yr As Integer) As String
    If com.FirstPeriod <> 1 Then
       If mon <= 12 - com.FirstPeriod + 1 Then yr = yr - 1
    End If
    mon = mon + com.FirstPeriod - 1
    If mon > 12 Then mon = mon - 12
    MonthName = Format(CStr(mon) & "/1/" & yr, "mmmm") & ", " & CStr(yr)
End Function

Public Function PutRecord(ByVal ID As Long) As Long
    
    On Error GoTo glErr
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    Set rs = db.OpenRecordset("glcompany")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
            Name = " "
            FileName = " "
            Address1 = " "
            Address2 = " "
            Address3 = " "
            City = " "
            State = " "
            ZipCode = " "
            LastBatch = 0
            CurFiscalYear = 0
            CurPeriod = 0
            FirstFiscalYear = 0
            NumberPds = 0
            LastUpdate = 0
            LastClose = 0
            SuspAcct = 0
            NetProfitAcct = 0
            FirstPAcct = 0
            PctBaseAcct = 0
            FirstPeriod = 0
            RetEarnAcct = 0
            SubDigits = 0
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !Name = Name
        !FileName = FileName
        !Address1 = Address1
        !Address2 = Address2
        !Address3 = Address3
        !City = City
        !State = State
        !ZipCode = ZipCode
        !LastBatch = LastBatch
        !CurFiscalYear = CurFiscalYear
        !CurPeriod = CurPeriod
        !FirstFiscalYear = FirstFiscalYear
        !NumberPds = NumberPds
        !LastUpdate = LastUpdate
        !LastClose = LastClose
        !SuspAcct = SuspAcct
        !NetProfitAcct = NetProfitAcct
        !FirstPAcct = FirstPAcct
        !PctBaseAcct = PctBaseAcct
        !FirstPeriod = FirstPeriod
        !RetEarnAcct = RetEarnAcct
        !SubDigits = SubDigits
        PutRecord = !ID
        .Update
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function


Public Function PutRecord_DAO(ByVal ID As Long) As Long
    
'    On Error GoTo glErr
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    If BalintFolder = "" Then
'        Set db = OpenDatabase("\balint\data\glSystem.mdb")
'    Else
'        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
'    End If
'    Set rs = db.OpenRecordset("glcompany")
'    With rs
'        If ID = 0 Then
'            If rs.RecordCount > 0 Then .MoveLast
'            .AddNew
'            Name = " "
'            FileName = " "
'            Address1 = " "
'            Address2 = " "
'            Address3 = " "
'            City = " "
'            State = " "
'            ZipCode = " "
'            LastBatch = 0
'            CurFiscalYear = 0
'            CurPeriod = 0
'            FirstFiscalYear = 0
'            NumberPds = 0
'            LastUpdate = 0
'            LastClose = 0
'            SuspAcct = 0
'            NetProfitAcct = 0
'            FirstPAcct = 0
'            PctBaseAcct = 0
'            FirstPeriod = 0
'            RetEarnAcct = 0
'            SubDigits = 0
'        Else
'            .index = "PrimaryKey"
'            .Seek "=", ID
'            .Edit
'        End If
'        !Name = Name
'        !FileName = FileName
'        !Address1 = Address1
'        !Address2 = Address2
'        !Address3 = Address3
'        !City = City
'        !State = State
'        !ZipCode = ZipCode
'        !LastBatch = LastBatch
'        !CurFiscalYear = CurFiscalYear
'        !CurPeriod = CurPeriod
'        !FirstFiscalYear = FirstFiscalYear
'        !NumberPds = NumberPds
'        !LastUpdate = LastUpdate
'        !LastClose = LastClose
'        !SuspAcct = SuspAcct
'        !NetProfitAcct = NetProfitAcct
'        !FirstPAcct = FirstPAcct
'        !PctBaseAcct = PctBaseAcct
'        !FirstPeriod = FirstPeriod
'        !RetEarnAcct = RetEarnAcct
'        !SubDigits = SubDigits
'        PutRecord = !ID
'        .Update
'    End With
'    Exit Function
'glErr:
'    MsgBox Error(Err.Number)
End Function

Public Function GetRecord(ByVal ID As Long) As Boolean

'    On Error GoTo glErr
    
    GetRecord = False
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\glSystem.mdb")
    End If
    Set rs = db.OpenRecordset("glcompany")
    
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        cID = !ID + 0
        Name = !Name & ""
        FileName = !FileName & ""
        Address1 = !Address1 & ""
        Address2 = !Address2 & ""
        Address3 = !Address3 & ""
        City = !City & ""
        State = !State & ""
        ZipCode = !ZipCode & ""
        LastBatch = !LastBatch + 0
        CurFiscalYear = !CurFiscalYear + 0
        CurPeriod = !CurPeriod + 0
        FirstFiscalYear = !FirstFiscalYear + 0
        NumberPds = !NumberPds + 0
        LastUpdate = !LastUpdate + 0
        LastClose = !LastClose + 0
        SuspAcct = !SuspAcct + 0
        NetProfitAcct = !NetProfitAcct + 0
        FirstPAcct = !FirstPAcct + 0
        PctBaseAcct = !PctBaseAcct + 0
        FirstPeriod = !FirstPeriod + 0
        RetEarnAcct = !RetEarnAcct
        SubDigits = !SubDigits
        GetRecord = True
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function


