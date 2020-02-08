VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetail99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DetailID As Long

Public PayeeID As Long
Public FormType As String
Public TaxYear As Long
Public BoxName As String
Public FieldValue As String

Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM Detail99"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngDetailID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngDetailID = 0 Then
       ErrMessage = "DetailID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from Detail99 where DetailID = " & lngDetailID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Find(ByVal lngDetailID As Long) As Boolean
   
   SQLString = "DetailID = " & lngDetailID
   
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
    
    rs.Fields("PayeeID") = PayeeID
    rs.Fields("TaxYear") = TaxYear
    rs.Fields("FormType") = FormType
    rs.Fields("BoxName") = BoxName
    rs.Fields("FieldValue") = FieldValue
    rs.Update
    
    DetailID = rs!DetailID
    
End Function

Public Sub Clear()
      
    DetailID = 0
    PayeeID = 0
    TaxYear = 0
    FormType = ""
    BoxName = ""
    FieldValue = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Detail99", cn, rs
    
End Sub

Public Sub AssignFields()
        
    DetailID = nNull(rs!DetailID)
            
    PayeeID = nNull(rs!PayeeID)
    TaxYear = nNull(rs!TaxYear)
    FormType = Trim(rs!FormType & "")
    BoxName = Trim(rs!BoxName & "")
    FieldValue = nNull(rs!FieldValue)
    
End Sub

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

    SQLString = "SELECT * FROM Detail99 WHERE DetailID = " & ID
    
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




