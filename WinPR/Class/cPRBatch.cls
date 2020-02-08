VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BatchID As Long
Public UserID As Long
Public CreateDate As Date
Public PEDate As Date
Public CheckDate As Date
Public RecCount As Long
Public YearMonth As Long
Public JobDist As Byte

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRBatch"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngHistID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngHistID = 0 Then
       ErrMessage = "HistID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRBatch where RecID = " & lngHistID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT EmployeeID from PRBatch"
    
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

Public Function Find(ByVal lngHistID As Long) As Boolean
   
   SQLString = "RecID = " & lngHistID
   
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
    
    rs.Fields("UserID") = UserID
    rs.Fields("CreateDate") = CreateDate
    rs.Fields("PEDate") = PEDate
    rs.Fields("CheckDate") = CheckDate
    rs.Fields("RecCount") = RecCount
    rs.Fields("YearMonth") = YearMonth
    rs.Fields("JobDist") = JobDist
    
    rs.Update
    
    BatchID = rs!BatchID
    
End Function

Public Sub Clear()
      
    BatchID = 0
    UserID = 0
    CreateDate = 0
    PEDate = 0
    CheckDate = 0
    RecCount = 0
    YearMonth = 0
    JobDist = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRBatch", cn, rs
    
End Sub

Public Sub AssignFields()
        
    BatchID = nNull(rs!BatchID)
    UserID = nNull(rs!UserID)
    CreateDate = nNull(rs!CreateDate)
    PEDate = nNull(rs!PEDate)
    CheckDate = nNull(rs!CheckDate)
    RecCount = nNull(rs!RecCount)
    YearMonth = nNull(rs!YearMonth)
    JobDist = nNull(rs!JobDist)
        
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
                             
End Function

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRBatch WHERE PRBatch.BatchID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


