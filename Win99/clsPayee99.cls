VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayee99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PayeeID As Long
Public PayeeNumber As Long
Public PayeeName As String
Public Address As String
Public CSZ As String
Public FederalID As String
Public AccountNumber As String
Public Comment As String
Public Inactive As Byte

Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM Payee99"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngPayeeID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngPayeeID = 0 Then
       ErrMessage = "PayeeID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from Payee99 where PayeeID = " & lngPayeeID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Find(ByVal lngPayeeID As Long) As Boolean
   
   SQLString = "PayeeID = " & lngPayeeID
   
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
    
    rs.Fields("PayeeName") = Trim(PayeeName)
    rs.Fields("PayeeNumber") = Trim(PayeeNumber)
    rs.Fields("Address") = Trim(Address)
    rs.Fields("CSZ") = Trim(CSZ)
    rs.Fields("FederalID") = Trim(FederalID)
    rs.Fields("AccountNumber") = Trim(AccountNumber)
    rs.Fields("Comment") = Trim(Comment)
    rs.Fields("Inactive") = nNull(Inactive)
    rs.Update
    PayeeID = rs!PayeeID
    
End Function

Public Sub Clear()
      
    PayeeID = 0
    PayeeNumber = 0
    PayeeName = ""
    Address = ""
    CSZ = ""
    FederalID = ""
    AccountNumber = ""
    Comment = ""
    Inactive = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Payee99", cn, rs
    
End Sub

Public Sub AssignFields()
        
    PayeeID = nNull(rs!PayeeID)
        
    PayeeName = Trim(rs!PayeeName & "")
    PayeeNumber = nNull(rs!PayeeNumber)
    Address = Trim(rs!Address & "")
    CSZ = Trim(rs!CSZ & "")
    FederalID = Trim(rs!FederalID & "")
    AccountNumber = Trim(rs!AccountNumber & "")
    Comment = Trim(rs!Comment & "")
    Inactive = nNull(rs!Inactive)
    
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

    SQLString = "SELECT * FROM Payee99 WHERE PayeeID = " & ID
    
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



