VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQBAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public QBAccountID As Long
Public Name As String
Public Description As String
Public AccountType As String
Public AccountNumber As Long
Public QBID As String
Public RecCount As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM QBAccount"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   RecCount = rs.RecordCount

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngAccountNumber As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngAccountNumber = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from QBAccount where AccountNumber = " & lngAccountNumber
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT AccountNumber from QBAccount"
    
    rsInit SQLString, cn, rs
    
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
       Records = 0
    Else
       rs.MoveLast
       Records = rs.RecordCount
    End If

    rs.Close

End Function

Public Function Find(ByVal lngAccountNumber As Long) As Byte
   
   SQLString = "AccountNumber = " & lngAccountNumber
   
   rs.Find SQLString, 0, adSearchForward, 1
   
   If rs.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

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
    
    rs!Name = Name
    rs!Description = Description
    rs!AccountType = AccountType
    rs!AccountNumber = AccountNumber
    rs!QBID = QBID

    rs.Update
    
    QBAccountID = rs!QBAccountID
    
End Function

Public Sub Clear()
      
    Name = ""
    Description = ""
    AccountType = ""
    AccountNumber = 0
    QBID = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM QBAccount", cn, rs
    
End Sub

Public Sub AssignFields()
        
    QBAccountID = rs!QBAccountID
    Name = Trim(rs!Name) & ""
    Description = Trim(rs!Description) & ""
    AccountType = Trim(rs!AccountType) & ""
    AccountNumber = nNull(rs!AccountNumber)
    QBID = Trim(rs!QBID) & ""
        
End Sub

Public Function GetByID(ByVal lngQBAccountID As Long) As Boolean
    
    Clear
    rsInit "SELECT * FROM QBAccount WHERE QBAccountID = " & lngQBAccountID, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True
    RecCount = 1
End Function
Public Function GetByQBID(ByVal strQBID As String) As Boolean

    Clear
    rsInit "SELECT * FROM QBAccount WHERE QBID = '" & strQBID & "'", cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByQBID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByQBID = True
    RecCount = 1
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
   RecCount = rs.RecordCount

End Function

