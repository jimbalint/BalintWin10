VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HeaderID As Long

Public SoldAddr1 As String
Public SoldAddr2 As String
Public SoldAddr3 As String
Public SoldAddr4 As String
Public SoldCity As String
Public SoldState As String
Public SoldZip As String

Public ShipAddr1 As String
Public ShipAddr2 As String
Public ShipAddr3 As String
Public ShipAddr4 As String
Public ShipCity As String
Public ShipState As String
Public ShipZip As String

Public SoldJobID As Long
Public InvoiceNumber As Long
Public SaveFlag As Byte
Public OrderDate As Date
Public InvoiceDate As Date
Public PackageCount As Long
Public PalletCount As Long
Public ItemTotal As Currency
Public Freight As Currency
Public SalesTax As Currency
Public TotalAmount As Currency
Public PO1 As String
Public PO2 As String
Public TruckID1 As Long
Public TruckID2 As Long
Public TruckID3 As Long
Public TrailerID1 As Long
Public TrailerID2 As Long
Public TrailerID3 As Long
Public DriverID1 As Long
Public DriverID2 As Long
Public DriverID3 As Long
Public ApptDate As Date
Public ApptTime As String
Public Terms As String
Public QBInvoiceID As String

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
    SQLString = "SELECT * FROM InvHeader WHERE HeaderID = " & ID
    GetByID = GetBySQL(SQLString)
End Function

Public Function OpenRS() As Boolean
    OpenRS = GetBySQL("SELECT * FROM InvHeader")
End Function

Private Sub AssignFields()

    HeaderID = rs!HeaderID
    
    SoldAddr1 = Trim(rs!SoldAddr1 & "")
    SoldAddr2 = Trim(rs!SoldAddr2 & "")
    SoldAddr3 = Trim(rs!SoldAddr3 & "")
    SoldAddr4 = Trim(rs!SoldAddr4 & "")
    SoldCity = Trim(rs!SoldCity & "")
    SoldState = Trim(rs!SoldState & "")
    SoldZip = Trim(rs!SoldZip & "")
    
    ShipAddr1 = Trim(rs!ShipAddr1 & "")
    ShipAddr2 = Trim(rs!ShipAddr2 & "")
    ShipAddr3 = Trim(rs!ShipAddr3 & "")
    ShipAddr4 = Trim(rs!ShipAddr4 & "")
    ShipCity = Trim(rs!ShipCity & "")
    ShipState = Trim(rs!ShipState & "")
    ShipZip = Trim(rs!ShipZip & "")
    
    SoldJobID = nNull(rs!SoldJobID)
    InvoiceNumber = nNull(rs!InvoiceNumber)
    SaveFlag = nNull(rs!SaveFlag)
    InvoiceDate = nNull(rs!InvoiceDate)
    PackageCount = nNull(rs!PackageCount)
    PalletCount = nNull(rs!PalletCount)
    ItemTotal = nNull(rs!ItemTotal)
    SalesTax = nNull(rs!SalesTax)
    Freight = nNull(rs!Freight)
    TotalAmount = nNull(rs!TotalAmount)
    PO1 = Trim(rs!PO1 & "")
    PO2 = Trim(rs!PO2 & "")
    TruckID1 = (rs!TruckID1)
    TruckID2 = (rs!TruckID2)
    TruckID3 = nNull(rs!TruckID3)
    TrailerID1 = nNull(rs!TrailerID1)
    TrailerID2 = nNull(rs!TrailerID2)
    TrailerID3 = nNull(rs!TrailerID3)
    DriverID1 = nNull(rs!DriverID1)
    DriverID2 = nNull(rs!DriverID2)
    DriverID3 = nNull(rs!DriverID3)
    ApptDate = nNull(rs!ApptDate)
    ApptTime = Trim(rs!ApptTime & "")
    Terms = Trim(rs!Terms & "")
    OrderDate = nNull(rs!OrderDate)
    InvoiceDate = nNull(rs!InvoiceDate)
    QBInvoiceID = Trim(rs!QBInvoiceID & "")

End Sub

Public Sub Clear()

    HeaderID = 0
    
    SoldAddr1 = ""
    SoldAddr2 = ""
    SoldAddr3 = ""
    SoldAddr4 = ""
    SoldCity = ""
    SoldState = ""
    SoldZip = ""
    
    ShipAddr1 = ""
    ShipAddr2 = ""
    ShipAddr3 = ""
    ShipAddr4 = ""
    ShipCity = ""
    ShipState = ""
    ShipZip = ""
    
    SoldJobID = 0
    InvoiceNumber = 0
    SaveFlag = 0
    OrderDate = Now
    InvoiceDate = Now
    PackageCount = 0
    PalletCount = 0
    ItemTotal = 0
    SalesTax = 0
    Freight = 0
    TotalAmount = 0
    PO1 = ""
    PO2 = ""
    TruckID1 = 0
    TruckID2 = 0
    TruckID3 = 0
    TrailerID1 = 0
    TrailerID2 = 0
    TrailerID3 = 0
    DriverID1 = 0
    DriverID2 = 0
    DriverID3 = 0
    ApptDate = Now
    ApptTime = ""
    Terms = ""
    QBInvoiceID = ""

End Sub

Private Sub AssignDataSet()

    rs!SoldAddr1 = MaxLen(SoldAddr1, 40)
    rs!SoldAddr2 = MaxLen(SoldAddr2, 40)
    rs!SoldAddr3 = MaxLen(SoldAddr3, 40)
    rs!SoldAddr4 = MaxLen(SoldAddr4, 40)
    rs!SoldCity = MaxLen(SoldCity, 40)
    rs!SoldState = MaxLen(SoldState, 2)
    rs!SoldZip = MaxLen(SoldZip, 10)
    
    rs!ShipAddr1 = MaxLen(ShipAddr1, 40)
    rs!ShipAddr2 = MaxLen(ShipAddr2, 40)
    rs!ShipAddr3 = MaxLen(ShipAddr3, 40)
    rs!ShipAddr4 = MaxLen(ShipAddr4, 40)
    rs!ShipCity = MaxLen(ShipCity, 40)
    rs!ShipState = MaxLen(ShipState, 2)
    rs!ShipZip = MaxLen(ShipZip, 10)
    
    rs!SoldJobID = SoldJobID
    rs!InvoiceNumber = InvoiceNumber
    rs!SaveFlag = SaveFlag
    rs!OrderDate = OrderDate
    rs!InvoiceDate = InvoiceDate
    rs!PackageCount = PackageCount
    rs!PalletCount = PalletCount
    rs!ItemTotal = ItemTotal
    rs!SalesTax = SalesTax
    rs!Freight = Freight
    rs!TotalAmount = TotalAmount
    rs!PO1 = MaxLen(PO1, 50)
    rs!PO2 = MaxLen(PO2, 50)
    rs!TruckID1 = TruckID1
    rs!TruckID2 = TruckID2
    rs!TruckID3 = TruckID3
    rs!TrailerID1 = TrailerID1
    rs!TrailerID2 = TrailerID2
    rs!TrailerID3 = TrailerID3
    rs!DriverID1 = DriverID1
    rs!DriverID2 = DriverID2
    rs!DriverID3 = DriverID3
    rs!ApptDate = ApptDate
    rs!ApptTime = MaxLen(ApptTime, 10)
    rs!Terms = MaxLen(Terms, 50)
    rs!QBInvoiceID = MaxLen(QBInvoiceID, 50)

End Sub

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
    HeaderID = rs!HeaderID
    
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

