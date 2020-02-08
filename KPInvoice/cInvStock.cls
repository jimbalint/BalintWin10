VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StockID As Long
Public StockSelect As Boolean
Public QBID As String
Public QBName As String
Public JobID As Long
Public Show As Byte
Public Description As String
Public MasterPrice As Double
Public CustomerPrice As Double
Public Cost As Double
Public LastDate As Date
Public Active As Boolean
Public InventoryItem As Boolean

Public Records As Long

Dim rs As New ADODB.Recordset

Public Function GetBySQL(ByVal SQLString As String) As Boolean

   rsInit SQLString, cn, rs
   
   If rs.RecordCount = 0 Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True

End Function
Public Function GetByID(ByVal ID As Long) As Boolean
    SQLString = "SELECT * FROM InvStock WHERE StockID = " & ID
    GetByID = GetBySQL(SQLString)
End Function
Public Function OpenRS() As Boolean
    OpenRS = GetBySQL("SELECT * FROM InvStock")
End Function

Private Sub AssignFields()

    StockID = rs!StockID
    QBID = Trim(rs!QBID & "")
    StockSelect = nNull(rs!StockSelect)
    QBName = Trim(rs!QBName & "")
    JobID = nNull(rs!JobID)
    Show = nNull(rs!Show)
    Description = Trim(rs!Description & "")
    MasterPrice = nNull(rs!MasterPrice)
    CustomerPrice = nNull(rs!CustomerPrice)
    Cost = nNull(rs!Cost)
    LastDate = (rs!LastDate)
    Active = nNull(rs!Active)
    InventoryItem = nNull(rs!InventoryItem)

End Sub

Public Sub Clear()

    StockID = 0
    StockSelect = False
    QBID = ""
    QBName = ""
    JobID = 0
    Show = 0
    Description = ""
    MasterPrice = 0
    CustomerPrice = 0
    Cost = 0
    LastDate = 0
    Active = False
    InventoryItem = False

End Sub

Private Sub AssignDataSet()

    rs!QBID = QBID
    rs!StockSelect = StockSelect
    rs!QBName = QBName
    rs!JobID = JobID
    rs!Show = Show
    rs!Description = Description
    rs!MasterPrice = MasterPrice
    rs!CustomerPrice = CustomerPrice
    rs!Cost = Cost
    rs!LastDate = LastDate
    rs!Active = Active
    rs!InventoryItem = InventoryItem

End Sub
Public Function GetFirst() As Boolean

    If rs.RecordCount = 0 Then
        GetFirst = False
        Exit Function
    End If
    
    rs.MoveFirst
   
    AssignFields
    GetFirst = True

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
Public Function GetPrev() As Boolean

   rs.MovePrevious
   
   If rs.BOF Then
      GetPrev = False
      Exit Function
   Else
      AssignFields
      GetPrev = True
   End If

End Function

Public Sub rsAdd()
    
    rs.AddNew
    AssignDataSet
    rs.Update
    StockID = rs!StockID
    
End Sub

Public Sub rsPut()
    
    AssignDataSet
    rs.Update

End Sub

Public Sub rsClose()

    On Error Resume Next
    rs.Close
    On Error GoTo 0

End Sub

