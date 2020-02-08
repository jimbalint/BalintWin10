VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRGLUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GLUpdID As Long
Public GLType As Byte
Public RelatedID As Long
Public GLItemType As Byte
Public ItemID As Long
Public GLAccountNum As Long
Public Title As String

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRGLUpd"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngGLUpdID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngGLUpdID = 0 Then
       ErrMessage = "GLUpdID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRGLUpd where GLUpdID = " & lngGLUpdID
    cn.Execute SQLString
    
    ' rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    Records = rs.RecordCount

'    SQLString = "SELECT EmployeeID from PRGLUpd"
'
'    rsInit SQLString, cn, rs
'
'    ' Check to make sure a record
'    ' actually came back
'    If rs.EOF = True And rs.BOF = True Then
'       Records = 0
'    Else
'       rs.MoveLast
'       Records = rs.RecordCount
'    End If
'
'    rs.Close

End Function

Public Function Find(ByVal lngGLUpdID As Long) As Boolean
   
   SQLString = "GLUpdID = " & lngGLUpdID
   
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
    
    rs!GLType = GLType
    rs!RelatedID = RelatedID
    rs!GLItemType = GLItemType
    rs!ItemID = ItemID
    rs!GLAccountNum = GLAccountNum
    rs!Title = Title
    
    rs.Update
    
End Function

Public Sub Clear()
      
    GLType = 0
    RelatedID = 0
    GLItemType = 0
    ItemID = 0
    GLAccountNum = 0
    Title = ""

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRGLUpd", cn, rs
    
End Sub

Public Sub AssignFields()
        
    GLUpdID = rs!GLUpdID
    GLType = nNull(rs!GLType)
    RelatedID = nNull(rs!RelatedID)
    GLItemType = nNull(rs!GLItemType)
    ItemID = nNull(rs!ItemID)
    GLAccountNum = nNull(rs!GLAccountNum)
    Title = rs!Title & ""

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

Public Function GetByID(ByVal GLUpdID As Long) As Boolean

    rsInit "SELECT * FROM PRGLUpd WHERE GLUpdID = " & GLUpdID, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

