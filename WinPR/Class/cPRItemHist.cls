VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRItemHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemHistID As Long
Public EmployeeID As Long
Public HistID As Long
Public BatchID As Long
Public ItemID As Long
Public DepartmentID As Long
Public EmployerItemID As Long
Public ItemType As Byte

Public YearMonth As Long
Public PEDate As Long
Public CheckDate As Date

Public Hours As Single
Public Rate As Currency
Public Amount As Currency
Public ManualAmount As Byte

Public Percent As Currency
Public WageBase As Currency
    
' wage excluded from basis for deduct by percent (401k match purposes)
Public WageExcluded As Currency

Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRItemHist"
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
    
    SQLString = "Delete * from PRItemHist where RecID = " & lngHistID
    
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
    rs.Fields("HistID") = HistID
    rs.Fields("BatchID") = BatchID
    rs.Fields("DepartmentID") = DepartmentID
    rs.Fields("ItemID") = ItemID
    rs.Fields("EmployerItemID") = EmployerItemID
    rs.Fields("ItemType") = ItemType
        
    rs.Fields("YearMonth") = YearMonth
    rs.Fields("PEDate") = PEDate
    rs.Fields("CheckDate") = CheckDate

    rs.Fields("Hours") = Hours
    rs.Fields("Rate") = Rate
    rs.Fields("Amount") = Amount
    rs.Fields("ManualAmount") = ManualAmount
    rs.Fields("Percent") = Percent
    rs.Fields("WageBase") = WageBase
    rs.Fields("WageExcluded") = WageExcluded

    rs.Update
    
    ItemHistID = rs!ItemHistID
    
End Function

Public Sub Clear()
      
    ItemHistID = 0
    EmployeeID = 0
    BatchID = 0
    HistID = 0
    ItemID = 0
    DepartmentID = 0
    EmployerItemID = 0
    ItemType = 0
    
    YearMonth = 0
    PEDate = 0
    CheckDate = 0
    
    Hours = 0
    Rate = 0
    Amount = 0
    ManualAmount = 0
    Percent = 0
    WageBase = 0
    WageExcluded = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRItemHist", cn, rs
    
End Sub

Public Sub AssignFields()
        
    ItemHistID = nNull(rs!ItemHistID)
    EmployeeID = nNull(rs!EmployeeID)
    HistID = nNull(rs!HistID)
    BatchID = nNull(rs!BatchID)
    ItemID = nNull(rs!ItemID)
    DepartmentID = nNull(rs!DepartmentID)
    ItemType = nNull(rs!ItemType)
    
    YearMonth = nNull(rs!YearMonth)
    PEDate = nNull(rs!PEDate)
    CheckDate = nNull(rs!CheckDate)
    
    Hours = nNull(rs!Hours)
    Rate = nNull(rs!Rate)
    Amount = nNull(rs!Amount)
    ManualAmount = nNull(rs!ManualAmount)
    Percent = nNull(rs!Percent)
    WageBase = nNull(rs!WageBase)
    WageExcluded = nNull(rs!WageExcluded)
    EmployerItemID = nNull(rs!EmployerItemID)
        
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

    SQLString = "SELECT * FROM PRItemHist WHERE ItemHistID = " & ID
    
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

