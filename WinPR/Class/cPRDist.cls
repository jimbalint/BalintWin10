VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DistID As Long
Public EmployeeID As Long
Public BatchID As Long
Public HistID As Long
Public StateID As Long
Public CityID As Long
Public CourtesyCityID As Long
Public JobID As Long
Public CustomerID As Long
Public DepartmentID As Long

Public YearMonth As Long
Public CheckDate As Date
Public PEDate As Date

Public DistType As Byte
Public ItemType As Byte
Public ItemID As Long
Public EmployerItemID As Long
Public Hours As Single
Public Rate As Currency
Public Amount As Currency
Public ManualAmount As Byte

Public BillingRate As Currency

Public GrossWage As Currency
Public StateWage As Currency
Public StateTax As Currency
Public ManualStateTax As Byte

Public CityWage As Currency
Public CityTax As Currency
Public CourtesyCityTax As Currency
Public ManualCityTax As Byte
Public ManualCourtesyCityTax As Byte

Public SUNWage As Currency
Public HistFlag As Byte

Public NotInNet As Byte
Public SDTax As Currency
Public Records As Long

Public QBInvoiceID As String

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRDist"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

   Records = rs.RecordCount

End Sub

Public Function DeleteRecord(ByVal lngHistID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngHistID = 0 Then
       ErrMessage = "HistID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRDist where RecID = " & lngHistID
    
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
   
   Records = 1
   
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
    rs.Fields("HistID") = HistID
    rs.Fields("StateID") = StateID
    rs.Fields("CityID") = CityID
    rs.Fields("CourtesyCityID") = CourtesyCityID
    rs.Fields("JobID") = JobID
    rs.Fields("CustomerID") = CustomerID
    rs.Fields("DepartmentID") = DepartmentID
    
    rs.Fields("YearMonth") = YearMonth
    rs.Fields("CheckDate") = CheckDate
    rs.Fields("PEDate") = PEDate
    
    rs.Fields("DistType") = DistType
    rs.Fields("ItemID") = ItemID
    rs.Fields("EmployerItemID") = EmployerItemID
    rs.Fields("ItemType") = ItemType
    rs.Fields("Hours") = Hours
    rs.Fields("Rate") = Rate
    rs.Fields("Amount") = Amount
    rs.Fields("ManualAmount") = ManualAmount
    
    rs.Fields("BillingRate") = BillingRate
    
    rs.Fields("GrossWage") = GrossWage
    rs.Fields("StateWage") = StateWage
    rs.Fields("StateTax") = StateTax
    rs.Fields("ManualStateTax") = ManualStateTax
    
    rs.Fields("CityWage") = CityWage
    rs.Fields("CityTax") = CityTax
    rs.Fields("CourtesyCityTax") = CourtesyCityTax
    rs.Fields("ManualCityTax") = ManualCityTax
    rs.Fields("ManualCourtesyCityTax") = ManualCourtesyCityTax
    
    rs.Fields("SUNWage") = SUNWage
    
    rs.Fields("HistFlag") = HistFlag
    
    rs.Fields("NotInNet") = NotInNet
    rs.Fields("SDTax") = SDTax
    
    rs.Fields("QBInvoiceID") = QBInvoiceID
    
    rs.Update
    
    DistID = rs!DistID
    
End Function

Public Sub Clear()
          
    DistID = 0
    EmployeeID = 0
    BatchID = 0
    HistID = 0
    StateID = 0
    CityID = 0
    CourtesyCityID = 0
    JobID = 0
    CustomerID = 0
    DepartmentID = 0

    YearMonth = 0
    PEDate = 0
    CheckDate = 0

    DistType = 0
    ItemID = 0
    EmployerItemID = 0
    ItemType = 0
    Hours = 0
    Rate = 0
    Amount = 0
    ManualAmount = 0
    BillingRate = 0
    GrossWage = 0

    StateWage = 0
    StateTax = 0
    ManualStateTax = 0

    CityWage = 0
    CityTax = 0
    CourtesyCityTax = 0
    ManualCityTax = 0
    ManualCourtesyCityTax = 0
      
    SUNWage = 0
      
    HistFlag = 0
    NotInNet = 0
    SDTax = 0
    Records = 0
      
    QBInvoiceID = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRDist", cn, rs
    
End Sub

Public Sub AssignFields()
            
    DistID = nNull(rs!DistID)
    EmployeeID = nNull(rs!EmployeeID)
    BatchID = nNull(rs!BatchID)
    HistID = nNull(rs!HistID)
    StateID = nNull(rs!StateID)
    CityID = nNull(rs!CityID)
    CourtesyCityID = nNull(rs!CourtesyCityID)
    JobID = nNull(rs!JobID)
    CustomerID = nNull(rs!CustomerID)
    DepartmentID = nNull(rs!DepartmentID)

    YearMonth = nNull(rs!YearMonth)
    PEDate = nNull(rs!PEDate)
    CheckDate = nNull(rs!CheckDate)

    DistType = nNull(rs!DistType)
    ItemID = nNull(rs!ItemID)
    EmployerItemID = nNull(rs!EmployerItemID)
    ItemType = nNull(rs!ItemType)
    Hours = nNull(rs!Hours)
    Rate = nNull(rs!Rate)
    Amount = nNull(rs!Amount)
    ManualAmount = nNull(rs!ManualAmount)
    
    BillingRate = nNull(rs!BillingRate)
    GrossWage = nNull(rs!GrossWage)

    StateWage = nNull(rs!StateWage)
    StateTax = nNull(rs!StateTax)
    ManualStateTax = nNull(rs!ManualStateTax)

    CityWage = nNull(rs!CityWage)
    CityTax = nNull(rs!CityTax)
    CourtesyCityTax = nNull(rs!CourtesyCityTax)
    ManualCourtesyCityTax = nNull(rs!ManualCourtesyCityTax)
        
    SUNWage = nNull(rs!SUNWage)
        
    HistFlag = nNull(rs!HistFlag)
    NotInNet = nNull(rs!NotInNet)
    SDTax = nNull(rs!SDTax)
        
    QBInvoiceID = rs!QBInvoiceID & ""
        
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
      Records = 0
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
   Records = rs.RecordCount
                             
End Function


Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRDist WHERE DistID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        Records = 0
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

    Records = 1

End Function

