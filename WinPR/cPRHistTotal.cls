VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim trs As New ADODB.Recordset

Public RecType As Integer
Public RecID As Double
Public RegHours As Single
Public RegAmount As Currency

'Public RegAmount As Currency
'
'Public OTHours As Single
'Public OTRate As Currency
'Public OTAmount As Currency
'
'Public OEHours As Single
'Public OERate As Currency
'Public OEAmount As Currency
'
'Public SSWage As Currency
'Public SSTax As Currency
'
'Public MedWage As Currency
'Public MedTax As Currency
'
'Public FWTWage As Currency
'Public FWTTax As Currency
'
'Public SWTWage As Currency
'Public SWTTax As Currency
'
'Public CWTWage As Currency
'Public CWTTax As Currency
'
'Public Deductions As Currency
'Public DirectDeposit As Currency
'
'Public Gross As Currency
'Public Net As Currency
'
'Public FUNWage As Currency
'Public SUNWage As Currency
'
'Dim rs As New ADODB.Recordset
    
Public Sub CreateRS()
    
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "RecType", adInteger
    trs.Fields.Append "RecID", adDouble
    trs.Fields.Append "RegHours", adSingle
    trs.Fields.Append "RegAmount", adCurrency

    trs.Open , , adOpenDynamic, adLockOptimistic

End Sub

Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
        trs.AddNew
    End If

    trs.Fields("RecType") = RecType
    trs.Fields("RecID") = RecID
    trs.Fields("RegHours") = RegHours
    trs.Fields("RegAmount") = RegAmount
    
    trs.Update

End Function

Public Sub AssignFields()

    RecID = trs!RecID
    RecType = trs!RecType
    RegHours = trs!RegHours
    RegAmount = trs!RegAmount

End Sub

Public Function tFind(ByVal RecType As Integer, RecID As Double) As Boolean
    
    If trs.EOF And trs.BOF Then
        tFind = False
        Exit Function
    End If
    
    trs.MoveFirst
    
    If trs.EOF Then
        tFind = False
        Exit Function
    End If
    
    Do
    
        If trs!RecType = RecType And trs!RecID = RecID Then
            AssignFields
            tFind = True
            Exit Function
        End If
        
        trs.MoveNext
        If trs.EOF Then
            tFind = False
            Exit Function
        End If
        
    Loop
    
End Function


'Public Function DeleteRecord(ByVal lngHistID As Long) As Boolean
'
'    Set rs = New ADODB.Recordset
'
'    If lngHistID = 0 Then
'       ErrMessage = "HistID not specified!"
'       DeleteRecord = False
'       Exit Function
'    End If
'
'    SQLString = "Delete * from PRHist where RecID = " & lngHistID
'
'    rsInit SQLString, cn, rs
'
'    DeleteRecord = True
'
'End Function
'
'Public Function Records() As Long
'
'    SQLString = "SELECT EmployeeID from PRHist"
'
'    rsInit SQLString, cn, rs
'
'    ' Check to make sure a record
'    ' actually came back
'    If rs.EOF = True And rs.BOF = True Then
'       Records = 0
'    Else
'       rs.MoveLast
'       Records = rs.RecordCount
'    End If
'
'    rs.Close
'
'End Function
'
'Public Function Find(ByVal lngHistID As Long) As Boolean
'
'   SQLString = "RecID = " & lngHistID
'
'   rs.Find SQLString, 0, adSearchForward, 1
'
'   If rs.EOF Then
'      Find = False
'      Exit Function
'   End If
'
'   Find = True
'
'   AssignFields
'
'End Function
'
'Public Function FindFirst() As Boolean
'
'   rs.MoveFirst
'
'   If rs.EOF Then
'      FindFirst = False
'      Exit Function
'   End If
'
'   FindFirst = True
'
'   AssignFields
'
'End Function
'
'
'Public Function GetNext() As Boolean
'
'   rs.MoveNext
'
'   If rs.EOF Then
'      GetNext = False
'      Exit Function
'   Else
'      AssignFields
'      GetNext = True
'   End If
'
'End Function
'
'
'
'Public Sub Clear()
'
'    HistID = 0
'    EmployeeID = 0
'    CheckNumber = 0
'    BatchID = 0
'
'    PEDate = 0
'    DepartmentID = 0
'    YearMonth = 0
'
'    RegHours = 0
'    RegRate = 0
'    RegAmount = 0
'
'    OTHours = 0
'    OTRate = 0
'    OTAmount = 0
'
'    OEHours = 0
'    OERate = 0
'    OEAmount = 0
'
'    SSWage = 0
'    SSTax = 0
'
'    MedWage = 0
'    MedTax = 0
'
'    FWTWage = 0
'    FWTTax = 0
'
'    SWTWage = 0
'    SWTTax = 0
'
'    CWTWage = 0
'    CWTTax = 0
'
'    Deductions = 0
'    DirectDeposit = 0
'
'    FUNWage = 0
'    SUNWage = 0
'
'    Gross = 0
'    Net = 0
'
'End Sub
'
'Public Sub DeleteAll()
'
'    Set rs = New ADODB.Recordset
'    rsInit "DELETE * FROM PRHist", cn, rs
'
'End Sub
'
'Public Sub AssignFields()
'
'    HistID = rs!HistID
'    EmployeeID = rs!EmployeeID
'    CheckNumber = rs!CheckNumber
'    BatchID = rs!BatchID
'
'    YearMonth = rs!YearMonth
'
'    PEDate = rs!PEDate
'    DepartmentID = rs!DepartmentID
'
'    RegHours = rs!RegHours
'    RegRate = rs!RegRate
'    RegAmount = rs!RegAmount
'
'    OTHours = rs!OTHours
'    OTRate = rs!OTRate
'    OTAmount = rs!OTAmount
'
'    OEHours = rs!OEHours
'    OERate = rs!OERate
'    OEAmount = rs!OEAmount
'
'    SSWage = rs!SSWage
'    SSTax = rs!SSTax
'
'    MedWage = rs!MedWage
'    MedTax = rs!MedTax
'
'    FWTWage = rs!FWTWage
'    FWTTax = rs!FWTTax
'
'    SWTWage = rs!SWTWage
'    SWTTax = rs!SWTTax
'
'    CWTWage = rs!CWTWage
'    CWTTax = rs!CWTTax
'
'    Deductions = rs!Deductions
'    DirectDeposit = rs!DirectDeposit
'
'    FUNWage = rs!FUNWage
'    SUNWage = rs!SUNWage
'
'    Gross = rs!Gross
'    Net = rs!Net
'
'End Sub
'
'Private Function MakeByte(ByVal BooVal As Boolean) As Integer
'
'    If BooVal = True Then
'        MakeByte = 1
'    Else
'        MakeByte = 0
'    End If
'
'End Function
'
'Public Function GetBySQL(ByVal SString As String) As Boolean
'
'   rsInit SString, cn, rs
'
'   If rs.BOF And rs.EOF Then
'      GetBySQL = False
'      Exit Function
'   End If
'
'   rs.MoveFirst
'
'   AssignFields
'
'   GetBySQL = True
'
'End Function
'
'
