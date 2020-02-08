VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BodyID As Long
Public HeaderID As Long
Public LineNum As Long
Public QtyOrdered As Double
Public QtyShipped As Double
Public Description As String
Public StockID As Long
Public Price As Double
Public Amount As Double

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
    SQLString = "SELECT * FROM InvBody WHERE BodyID = " & ID
    GetByID = GetBySQL(SQLString)
End Function

Public Function OpenRS() As Boolean
    OpenRS = GetBySQL("SELECT * FROM InvBody")
End Function

Private Sub AssignFields()

    BodyID = rs!BodyID
    HeaderID = nNull(rs!HeaderID)
    LineNum = nNull(rs!LineNum)
    QtyOrdered = nNull(rs!QtyOrdered)
    QtyShipped = nNull(rs!QtyShipped)
    Description = Trim(rs!Description & "")
    StockID = nNull(rs!StockID)
    Price = nNull(rs!Price)
    Amount = nNull(rs!Amount)

End Sub

Public Sub Clear()

    BodyID = 0
    HeaderID = 0
    LineNum = 0
    QtyOrdered = 0
    QtyShipped = 0
    StockID = 0
    Price = 0
    Amount = 0
    Description = ""

End Sub

Private Sub AssignDataSet()

    rs!HeaderID = HeaderID
    rs!LineNum = LineNum
    rs!QtyOrdered = QtyOrdered
    rs!QtyShipped = QtyShipped
    rs!StockID = StockID
    rs!Price = Price
    rs!Amount = Amount
    rs!Description = MaxLen(Description, 255)

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
    BodyID = rs!BodyID
    
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

