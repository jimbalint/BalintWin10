VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRW2Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public W2BoxID As Long
Public W2BoxNumber As Long
Public W2BoxCode As String
Public W2BoxDescription As String

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRW2Box"
   rsInit SQLString, cnDes, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngID = 0 Then
       ErrMessage = "W2Box ID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRW2Box where W2BoxID = " & lngID
    
    rsInit SQLString, cnDes, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT W2BoxID from PRW2Box"
    
    rsInit SQLString, cnDes, rs
    
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

Public Function Find(ByVal lngID As Long) As Boolean
   
   SQLString = "W2BoxID = " & lngID
   
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
    
    rs.Fields("W2BoxNumber") = W2BoxNumber
    rs.Fields("W2BoxCode") = UCase(W2BoxCode)
    rs.Fields("W2BoxDescription") = UCase(W2BoxDescription)
    
    rs.Update
    
    W2BoxID = rs!W2BoxID
    
End Function

Public Sub Clear()
      
    W2BoxID = 0
    W2BoxNumber = 0
    W2BoxCode = ""
    W2BoxDescription = ""
    
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRW2Box", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    W2BoxID = nNull(rs!W2BoxID)
    W2BoxNumber = nNull(rs!W2BoxNumber)
    W2BoxCode = Trim(UCase(rs!W2BoxCode)) & ""
    W2BoxDescription = Trim(UCase(rs!W2BoxDescription)) & ""
    
End Sub


Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cnDes, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
End Function


Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRW2Box WHERE W2BoxID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


