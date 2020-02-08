VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NoteID As Long
Public NoteType As Byte
Public NoteCat As Byte
Public RelatedID As Long
Public Subject As String
Public User As String
Public DateTm As Date
Public Notation As String

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM Notes"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngNoteID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngNoteID = 0 Then
       ErrMessage = "NoteID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from Notes where NoteID = " & lngNoteID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT NoteID from Notes"
    
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

Public Function Find(ByVal lngNoteID As Long) As Boolean
   
   SQLString = "NoteID = " & lngNoteID
   
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
    
    rs!NoteType = NoteType
    rs!NoteCat = NoteCat
    rs!RelatedID = RelatedID
    rs!Subject = Subject
    rs!User = User
    rs!DateTm = DateTm
    rs!Notation = Notation
    
    rs.Update
    
    NoteID = rs!NoteID
    
End Function

Public Sub Clear()
      
    NoteType = 0
    NoteCat = 0
    RelatedID = 0
    Subject = ""
    User = ""
    DateTm = 0
    Notation = ""

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Notes", cn, rs
    
End Sub

Public Sub AssignFields()
        
    NoteID = nNull(rs!NoteID)
    RelatedID = nNull(rs!RelatedID)
    NoteType = nNull(rs!NoteType)
    NoteCat = nNull(rs!NoteCat)
    Subject = rs!Subject & ""
    User = rs!User & ""
    DateTm = nNull(rs!DateTm)
    Notation = rs!Notation & ""
        
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
                             
End Function

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM Notes WHERE NoteID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function



