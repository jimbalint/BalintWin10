VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HistID As Long
Public EmployeeID As Long
Public BatchID As Long
Public StateID As Long
Public CheckNumber As Long

Public PEDate As Date
Public DepartmentID As Long
Public YearMonth As Long
Public CheckDate As Date

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
Public ManualSSTax As Byte

Public MEDWage As Currency
Public MedTax As Currency
Public MedAddAmt As Currency
Public ManualMedTax As Byte

Public FWTWage As Currency
Public FWTTax As Currency
Public ManualFWTTax As Byte

Public SWTWage As Currency
Public SWTTax As Currency

Public CWTWage As Currency
Public CWTTax As Currency

Public Deductions As Currency
Public DirectDeposit As Currency

Public Gross As Currency
Public Net As Currency

Public FUNWageBase As Currency
Public FUNWage As Currency
Public SUNWageBase As Currency
Public SUNWage As Currency
Public GLUpdate As Byte

Public WkcAmount As Currency
Public NotInNetAmount As Currency
Public SDTax As Currency

Public Records As Long
Public QBUpdateFlag As Byte


Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRHist"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngHistID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngHistID = 0 Then
       ErrMessage = "HistID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRHist where RecID = " & lngHistID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Find(ByVal lngHistID As Long) As Boolean
   
   SQLString = "RecID = " & lngHistID
   
   rs.Find SQLString, 0, adSearchForward, 1
   
   If rs.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

Public Function FindFirst() As Boolean
   
   rs.MoveFirst
   
   If rs.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
End Function


Public Function GetNext() As Boolean

   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignFields
      GetNext = True
   End If

End Function


Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("EmployeeID") = EmployeeID
    rs.Fields("BatchID") = BatchID
    rs.Fields("CheckNumber") = CheckNumber
    rs.Fields("YearMonth") = YearMonth
    rs.Fields("CheckDate") = CheckDate
    rs.Fields("StateID") = StateID
    
    rs.Fields("PEDate") = PEDate
    rs.Fields("DepartmentID") = DepartmentID
    
    rs.Fields("RegHours") = RegHours
    rs.Fields("RegRate") = RegRate
    rs.Fields("RegAmount") = RegAmount

    rs.Fields("OTHours") = OTHours
    rs.Fields("OTRate") = OTRate
    rs.Fields("OTAmount") = OTAmount

    rs.Fields("OEHours") = OEHours
    rs.Fields("OERate") = OERate
    rs.Fields("OEAmount") = OEAmount

    rs.Fields("SSWageBase") = SSWageBase
    rs.Fields("SSWage") = SSWage
    rs.Fields("SSTax") = SSTax
    rs.Fields("ManualSSTax") = ManualSSTax
    
    rs.Fields("MedWage") = MEDWage
    rs.Fields("MedTax") = MedTax
    rs.Fields("ManualMedTax") = ManualMedTax
    rs.Fields("MedAddAmt") = MedAddAmt

    rs.Fields("FWTWage") = FWTWage
    rs.Fields("FWTTax") = FWTTax
    rs.Fields("ManualFWTTax") = ManualFWTTax

    rs.Fields("SWTWage") = SWTWage
    rs.Fields("SWTTax") = SWTTax

    rs.Fields("CWTWage") = CWTWage
    rs.Fields("CWTTax") = CWTTax

    rs.Fields("Deductions") = Deductions
    rs.Fields("DirectDeposit") = DirectDeposit

    rs.Fields("Gross") = Gross
    rs.Fields("Net") = Net

    rs.Fields("FUNWageBase") = FUNWageBase
    rs.Fields("FUNWage") = FUNWage
    rs.Fields("SUNWageBase") = SUNWageBase
    rs.Fields("SUNWage") = SUNWage
    rs.Fields("GLUpdate") = GLUpdate
    rs.Fields("WkcAmount") = WkcAmount
    rs.Fields("NotInNetAmount") = NotInNetAmount
    rs.Fields("SDTax") = SDTax
    rs.Fields("QBUpdateFlag") = QBUpdateFlag
    rs.Update
    
    HistID = rs!HistID
    
End Function

Public Sub Clear()
      
    HistID = 0
    EmployeeID = 0
    CheckNumber = 0
    BatchID = 0
    StateID = 0

    PEDate = 0
    DepartmentID = 0
    YearMonth = 0
    CheckDate = 0

    RegHours = 0
    RegRate = 0
    RegAmount = 0

    OTHours = 0
    OTRate = 0
    OTAmount = 0

    OEHours = 0
    OERate = 0
    OEAmount = 0

    SSWageBase = 0
    SSWage = 0
    SSTax = 0
    ManualSSTax = 0

    MEDWage = 0
    MedTax = 0
    MedAddAmt = 0
    ManualMedTax = 0

    FWTWage = 0
    FWTTax = 0
    ManualFWTTax = 0

    SWTWage = 0
    SWTTax = 0

    CWTWage = 0
    CWTTax = 0

    Deductions = 0
    DirectDeposit = 0

    FUNWageBase = 0
    FUNWage = 0
    SUNWageBase = 0
    SUNWage = 0
    GLUpdate = 0
    WkcAmount = 0
    NotInNetAmount = 0
    SDTax = 0
    
    Gross = 0
    Net = 0
    Records = 0
    QBUpdateFlag = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRHist", cn, rs
    
End Sub

Public Sub AssignFields()
        
    HistID = nNull(rs!HistID)
    EmployeeID = nNull(rs!EmployeeID)
    CheckNumber = nNull(rs!CheckNumber)
    BatchID = nNull(rs!BatchID)
    StateID = nNull(rs!StateID)
    YearMonth = nNull(rs!YearMonth)
    CheckDate = nNull(rs!CheckDate)
    
    PEDate = nNull(rs!PEDate)
    DepartmentID = nNull(rs!DepartmentID)
    
    RegHours = nNull(rs!RegHours)
    RegRate = nNull(rs!RegRate)
    RegAmount = nNull(rs!RegAmount)

    OTHours = nNull(rs!OTHours)
    OTRate = nNull(rs!OTRate)
    OTAmount = nNull(rs!OTAmount)
    
    OEHours = nNull(rs!OEHours)
    OERate = nNull(rs!OERate)
    OEAmount = nNull(rs!OEAmount)

    SSWageBase = nNull(rs!SSWageBase)
    SSWage = nNull(rs!SSWage)
    SSTax = nNull(rs!SSTax)
    ManualSSTax = nNull(rs!ManualSSTax)

    MEDWage = nNull(rs!MEDWage)
    MedTax = nNull(rs!MedTax)
    MedAddAmt = nNull(rs!MedAddAmt)
    ManualMedTax = nNull(rs!ManualMedTax)

    FWTWage = nNull(rs!FWTWage)
    FWTTax = nNull(rs!FWTTax)
    ManualFWTTax = nNull(rs!ManualFWTTax)

    SWTWage = nNull(rs!SWTWage)
    SWTTax = nNull(rs!SWTTax)

    CWTWage = nNull(rs!CWTWage)
    CWTTax = nNull(rs!CWTTax)

    Deductions = nNull(rs!Deductions)
    DirectDeposit = nNull(rs!DirectDeposit)

    FUNWageBase = nNull(rs!FUNWageBase)
    FUNWage = nNull(rs!FUNWage)
    SUNWageBase = nNull(rs!SUNWageBase)
    SUNWage = nNull(rs!SUNWage)
    GLUpdate = nNull(rs!GLUpdate)
    WkcAmount = nNull(rs!WkcAmount)
    NotInNetAmount = nNull(rs!NotInNetAmount)
    SDTax = nNull(rs!SDTax)

    Gross = nNull(rs!Gross)
    Net = nNull(rs!Net)
    QBUpdateFlag = nNull(rs!QBUpdateFlag)
        
End Sub

Private Function MakeByte(ByVal BooVal As Boolean) As Integer
    
    If BooVal = True Then
        MakeByte = 1
    Else
        MakeByte = 0
    End If
    
End Function

Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cn, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
   Records = rs.RecordCount
                             
End Function

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRHist WHERE HistID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

    Records = 1

End Function
Public Function Count() As Long
    Count = rs.RecordCount
End Function

