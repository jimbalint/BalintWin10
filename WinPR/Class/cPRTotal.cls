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

Public PRTotalID As Long
Public RecType As Byte
Public RecID As Long
Public IDNumber As Long
Public SSN As Long
Public DeptID As Long
Public EmployeeID As Long
Public Name As String
Public DepartmentID As Long
Public PeriodType As Byte
Public RegHours As Single
Public RegRate As Currency
Public RegAmount As Currency
Public OTHours As Single
Public OTRate As Currency
Public OTAmount As Currency
Public OEHours As Single
Public OERate As Currency
Public OEAmount As Currency
Public SSWageBase As Currency
Public SSWage As Currency
Public SSTax As Currency
Public MEDWage As Currency
Public MedTax As Currency
Public FWTWage As Currency
Public FWTTax As Currency
Public Deductions As Currency
Public StateWage As Currency
Public StateTax As Currency
Public CityWage As Currency
Public CityTax As Currency
Public FUNWageBase As Currency
Public FUNWage As Currency
Public FUNTax As Currency
Public SUNWageBase As Currency
Public SUNWage As Currency
Public SUNTax As Currency
Public YTDGross As Currency
Public YTDFUNWageBase As Currency
Public YTDSUNWageBase As Currency
Public Gross As Currency
Public Net As Currency
Public DirectDeposit As Currency
Public Count As Long


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
    
    trs.Fields.Append "PRTotalID", adDouble
    trs.Fields.Append "RecType", adInteger
    trs.Fields.Append "RecID", adDouble
    trs.Fields.Append "IDNumber", adDouble
    trs.Fields.Append "SSN", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    trs.Fields.Append "Name", adVarChar, 40, adFldIsNullable
    trs.Fields.Append "DepartmentID", adDouble
    trs.Fields.Append "PeriodType", adInteger
    trs.Fields.Append "RegHours", adSingle
    trs.Fields.Append "RegRate", adCurrency
    trs.Fields.Append "RegAmount", adCurrency
    trs.Fields.Append "OTHours", adCurrency
    trs.Fields.Append "OTRate", adCurrency
    trs.Fields.Append "OTAmount", adCurrency
    trs.Fields.Append "OEHours", adCurrency
    trs.Fields.Append "OERate", adCurrency
    trs.Fields.Append "OEAmount", adCurrency
    trs.Fields.Append "SSWageBase", adCurrency
    trs.Fields.Append "SSWage", adCurrency
    trs.Fields.Append "SSTax", adCurrency
    trs.Fields.Append "MedWage", adCurrency
    trs.Fields.Append "MedTax", adCurrency
    trs.Fields.Append "FWTWage", adCurrency
    trs.Fields.Append "FWTTax", adCurrency
    trs.Fields.Append "Deductions", adCurrency
    trs.Fields.Append "StateWage", adCurrency
    trs.Fields.Append "StateTax", adCurrency
    trs.Fields.Append "CityWage", adCurrency
    trs.Fields.Append "CityTax", adCurrency
    trs.Fields.Append "FUNWageBase", adCurrency
    trs.Fields.Append "FUNWage", adCurrency
    trs.Fields.Append "FUNTax", adCurrency
    trs.Fields.Append "SUNWageBase", adCurrency
    trs.Fields.Append "SUNWage", adCurrency
    trs.Fields.Append "SUNTax", adCurrency
    trs.Fields.Append "Gross", adCurrency
    trs.Fields.Append "YTDGross", adCurrency
    trs.Fields.Append "Net", adCurrency
    trs.Fields.Append "DirectDeposit", adCurrency
    trs.Fields.Append "Count", adDouble

    trs.Fields.Append "YTDFUNWageBase", adCurrency
    trs.Fields.Append "YTDSUNWageBase", adCurrency

    trs.Open , , adOpenDynamic, adLockOptimistic

End Sub

Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
        trs.AddNew
    End If
    If RecType = 0 Then
        Exit Function
    End If
    trs.Fields("RecType") = RecType
    trs.Fields("RecID") = RecID
    trs.Fields("IDNumber") = IDNumber
    trs.Fields("DepartmentID") = DepartmentID
    trs.Fields("EmployeeID") = EmployeeID
    trs.Fields("SSN") = SSN
    trs.Fields("Name") = Mid(Name, 1, 40)
    trs.Fields("PeriodType") = PeriodType
    trs.Fields("RegHours") = RegHours
    trs.Fields("RegRate") = RegRate
    trs.Fields("RegAmount") = RegAmount
    trs.Fields("OTHours") = OTHours
    trs.Fields("OTRate") = OTRate
    trs.Fields("OTAmount") = OTAmount
    trs.Fields("OEHours") = OEHours
    trs.Fields("OERate") = OERate
    trs.Fields("OEAmount") = OEAmount
    trs.Fields("SSWageBase") = SSWageBase
    trs.Fields("SSWage") = SSWage
    trs.Fields("SSTax") = SSTax
    trs.Fields("MedWage") = MEDWage
    trs.Fields("MedTax") = MedTax
    trs.Fields("FWTWage") = FWTWage
    trs.Fields("FWTTax") = FWTTax
    trs.Fields("Deductions") = Deductions
    trs.Fields("StateWage") = StateWage
    trs.Fields("StateTax") = StateTax
    trs.Fields("CityWage") = CityWage
    trs.Fields("CityTax") = CityTax
    trs.Fields("FUNWageBase") = FUNWageBase
    trs.Fields("FUNWage") = FUNWage
    trs.Fields("FUNTax") = FUNTax
    trs.Fields("SUNWageBase") = SUNWageBase
    trs.Fields("SUNWage") = SUNWage
    trs.Fields("SUNTax") = SUNTax
    trs.Fields("Gross") = Gross
    trs.Fields("YTDGross") = YTDGross
    trs.Fields("Net") = Net
    trs.Fields("DirectDeposit") = DirectDeposit
    trs.Fields("YTDFUNWageBase") = YTDFUNWageBase
    trs.Fields("YTDSUNWageBase") = YTDSUNWageBase
    trs!Count = Count
    trs.Update
    
    If booAddPut = True Then      ' Add
        trs!PRTotalID = trs.RecordCount
        PRTotalID = trs.RecordCount
        trs.Update
    End If

End Function

Public Sub AssignFields()

    PRTotalID = nNull(trs!PRTotalID)
    RecID = nNull(trs!RecID)
    DepartmentID = nNull(trs!DepartmentID)
    EmployeeID = nNull(trs!EmployeeID)
    Name = trs!Name & ""
    IDNumber = nNull(trs!IDNumber)
    SSN = nNull(trs!SSN)
    PeriodType = nNull(trs!PeriodType)
    RecType = nNull(trs!RecType)
    RegHours = nNull(trs!RegHours)
    RegRate = nNull(trs!RegRate)
    RegAmount = nNull(trs!RegAmount)
    OTHours = nNull(trs!OTHours)
    OTRate = nNull(trs!OTRate)
    OTAmount = nNull(trs!OTAmount)
    OEHours = nNull(trs!OEHours)
    OERate = nNull(trs!OERate)
    OEAmount = nNull(trs!OEAmount)
    SSWageBase = nNull(trs!SSWageBase)
    SSWage = nNull(trs!SSWage)
    SSTax = nNull(trs!SSTax)
    MEDWage = nNull(trs!MEDWage)
    MedTax = nNull(trs!MedTax)
    FWTWage = nNull(trs!FWTWage)
    FWTTax = nNull(trs!FWTTax)
    StateWage = nNull(trs!StateWage)
    StateTax = nNull(trs!StateTax)
    Deductions = nNull(trs!Deductions)
    CityWage = nNull(trs!CityWage)
    CityTax = nNull(trs!CityTax)
    FUNWageBase = nNull(trs!FUNWageBase)
    FUNWage = nNull(trs!FUNWage)
    FUNTax = nNull(trs!FUNTax)
    SUNWageBase = nNull(trs!SUNWageBase)
    SUNWage = nNull(trs!SUNWage)
    SUNTax = nNull(trs!SUNTax)
    Gross = nNull(trs!Gross)
    YTDGross = nNull(trs!YTDGross)
    Net = nNull(trs!Net)
    DirectDeposit = nNull(trs!DirectDeposit)
    Count = nNull(trs!Count)

    YTDFUNWageBase = nNull(trs!YTDFUNWageBase)
    YTDSUNWageBase = nNull(trs!YTDSUNWageBase)

End Sub

Public Function tFind(ByVal RecType As Integer, RecID As Long) As Boolean
    
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

Public Function pFind(ByVal RecType As Integer, _
                      ByVal PeriodType As Byte, _
                      ByVal RecID As Long) As Boolean
                      

    If trs.EOF And trs.BOF Then
        pFind = False
        Exit Function
    End If
    
    trs.MoveFirst
    
    If trs.EOF Then
        pFind = False
        Exit Function
    End If
    
    Do
    
        If trs!RecType = RecType And trs!RecID = RecID And trs!PeriodType = PeriodType Then
            AssignFields
            pFind = True
            Exit Function
        End If
        
        trs.MoveNext
        If trs.EOF Then
            pFind = False
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
Public Function FindFirst() As Boolean
   
   If trs.BOF And trs.EOF Then
      FindFirst = False
      Exit Function
   End If
   
   trs.MoveFirst
   FindFirst = True

   AssignFields

End Function

Public Sub Clear()
    
    PRTotalID = 0
    PeriodType = 0
    Name = ""
    SSN = 0
    PRTotal.RegHours = 0
    PRTotal.RegAmount = 0
    PRTotal.OTHours = 0
    PRTotal.OTAmount = 0
    PRTotal.OEHours = 0
    PRTotal.OEAmount = 0
    PRTotal.SSWageBase = 0
    PRTotal.SSWage = 0
    PRTotal.SSTax = 0
    PRTotal.MEDWage = 0
    PRTotal.MedTax = 0
    PRTotal.FWTWage = 0
    PRTotal.FWTTax = 0
    PRTotal.StateWage = 0
    PRTotal.StateTax = 0
    PRTotal.Deductions = 0
    PRTotal.CityWage = 0
    PRTotal.CityTax = 0
    PRTotal.FUNWageBase = 0
    PRTotal.FUNWage = 0
    PRTotal.FUNTax = 0
    PRTotal.SUNWageBase = 0
    PRTotal.SUNWage = 0
    PRTotal.SUNTax = 0
    PRTotal.Gross = 0
    PRTotal.YTDGross = 0
    PRTotal.Net = 0
    
    PRTotal.YTDFUNWageBase = 0
    PRTotal.YTDSUNWageBase = 0
    
    RegHours = 0
    RegRate = 0
    RegAmount = 0

    OTHours = 0
    OTRate = 0
    OTAmount = 0

    OEHours = 0
    OERate = 0
    OEAmount = 0

    SSWage = 0
    SSTax = 0

    MEDWage = 0
    MedTax = 0

    FWTWage = 0
    FWTTax = 0

    StateWage = 0
    StateTax = 0

    CityWage = 0
    CityTax = 0

    PRTotal.RegAmount = 0
    
    TGrossPay = 0
    TTotTaxes = 0
    TNetPay = 0
    TTotHours = 0

    DirectDeposit = 0

'    Deductions = 0
'    DirectDeposit = 0
'
'    FUNWage = 0
'    SUNWage = 0

'    GrossPay = 0
'    NetPay = 0
    Count = 0

End Sub

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
'    DepartmentID = rs!PREmployee.DepartmentID]
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


Public Function GetNext() As Boolean
    
    On Error Resume Next
    trs.MoveNext
    If Err.Number <> 0 Then
        On Error GoTo 0
       GetNext = False
       Exit Function
    End If
    
    On Error GoTo 0
    If trs.EOF Then
       GetNext = False
       Exit Function
    Else
       AssignFields
       GetNext = True
    End If
    
End Function
Public Function GetPrev() As Boolean

    On Error Resume Next
    trs.MovePrevious
    If Err.Number <> 0 Then
        On Error GoTo 0
        GetPrev = False
        Clear
        Exit Function
    End If
        
    On Error GoTo 0
    AssignFields
    GetPrev = True

End Function

Public Sub TSort()
   trs.Sort = "RecType"
End Sub
Public Sub TSortByString(ByVal SortString As String)
    trs.Sort = SortString
End Sub

Public Sub TClose()
    trs.Close
    Set trs = Nothing
End Sub

Public Function GetByID(ByVal PRTID As Long) As Boolean

    trs.Find "PRTotalID = " & PRTID, 0, adSearchForward, 1
    If trs.EOF Then
        GetByID = False
    Else
        GetByID = True
        AssignFields
    End If

End Function

Public Sub TFilter(ByVal fString As String)
    trs.Filter = adFilterNone
    trs.Filter = fString
End Sub
