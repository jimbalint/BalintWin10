VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ClientID As Long
Public ClientName As String
Public Prefix As String
Public Contact As String
Public Phone As String
Public Message1 As String
Public Message2 As String
Public Message3 As String

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM Client"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function FindFirst() As Byte
   
   rs.MoveFirst
   
   If rs.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
End Function

Public Function GetNext() As Byte

   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignFields
      GetNext = True
   End If

End Function

Public Function Save(ByVal booAddPut As Byte) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If

    rs!ClientName = ClientName
    rs!Prefix = Prefix
    rs!Contact = Contact
    rs!Phone = Phone
    rs!Message1 = Message1
    rs!Message2 = Message2
    rs!Message3 = Message3
    
    rs.Update
    
    ClientID = rs!ClientID
    
End Function

Public Sub Clear()
      
    ClientID = 0
    ClientName = ""
    Prefix = ""
    Contact = ""
    Phone = ""
    Message1 = ""
    Message2 = ""
    Message3 = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Client", cn, rs
    
End Sub

Public Sub AssignFields()
        
    ClientID = nNull(rs!ClientID)
    ClientName = rs!ClientName & ""
    Prefix = rs!Prefix & ""
    Contact = rs!Contact & ""
    Phone = rs!Phone & ""
    Message1 = rs!Message1 & ""
    Message2 = rs!Message2 & ""
    Message3 = rs!Message3 & ""
                                           
End Sub

Public Function GetByID(ByVal ClntID As Long) As Boolean

    rsInit "SELECT * FROM Client WHERE ClientID = " & CStr(ClntID), cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

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
                             
End Function

