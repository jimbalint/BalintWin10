VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public ID As Long
Public JournalName As String
Public JournalSource As Integer

Dim SQLString As String
Dim rs As ADODB.Recordset


Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = True Then      ' Add
    
         SQLString = "GLJournal"
         rsInit SQLString, cn, rs
         rs.AddNew
    
    Else                           ' Put
    
         SQLString = "select * " & _
                     "from GLJournal " & _
                     "where JournalSource = " & JournalSource
         rsInit SQLString, cn, rs
    
    End If
    
    rs.Fields("JournalName") = JournalName
    rs.Fields("JournalSource") = JournalSource
    
    rs.Update
    rs.Close
    
    Set rs = Nothing
    
End Sub
    
Public Function GetData(ByVal lngJournalSource As Long) As Boolean

    If lngJournalSource <> 0 Then
        
        SQLString = "select * from GLJournal " & _
                    "where JournalSource = " & lngJournalSource
        rsInit SQLString, cn, rs
        
        ' Check to make sure a record
        ' actually came back
        If rs.EOF = True And rs.BOF = True Then
           GetData = False
           Exit Function
        End If
        
        ' Populate properties with
        ' results of the query.
        JournalName = rs!JournalName & ""
        JournalSource = rs!JournalSource & ""
        
        GetData = True
    
    Else
        
        GetData = False
        Exit Function
    
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function


Public Sub Clear()
      ID = 0
      JournalSource = 0
      JournalName = ""
End Sub

Public Sub DeleteAll()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Source = "Delete * from GLJournal"
    Set rs.ActiveConnection = cn
    rs.Open
    Set rs = Nothing
    
End Sub

Public Sub OpenRS()
   
   SQLString = "GLJournal"
   rsInit SQLString, cn, rs

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

Public Sub AssignFields()

    JournalSource = nNull(rs!JournalSource)
    JournalName = rs!JournalName & ""

End Sub

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

