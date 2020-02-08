VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQBUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public QBUpdateID As Long
Public Category As Byte
Public Post As String
Public PerJob As Byte
Public Title As String
Public ItemType As Byte
Public RelatedID As Long
Public QBID As String
Public DebitAmount As Currency
Public CreditAmount As Currency

Public RecCount As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM QBUpdate"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   RecCount = rs.RecordCount

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


Public Function Save(ByVal booAddPut As Byte) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs!QBUpdateID = QBUpdateID
    rs!Category = Category
    rs!Post = Post
    rs!PerJob = PerJob
    rs!Title = Title
    rs!ItemType = ItemType
    rs!RelatedID = RelatedID
    rs!QBID = QBID
    rs.Update
    
    QBUpdateID = rs!QBUpdateID
    
End Function

Public Sub Clear()
      
    rs!QBUpdateID = 0
    rs!Category = 0
    rs!Post = ""
    rs!PerJob = 0
    rs!Title = ""
    rs!ItemType = 0
    rs!RelatedID = 0
    rs!QBID = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM QBUpdate", cn, rs
    
End Sub

Public Sub AssignFields()
        
    QBUpdateID = rs!QBUpdateID
    Category = nNull(rs!Category)
    Post = rs!Post & ""
    PerJob = nNull(rs!PerJob)
    Title = rs!Title & ""
    ItemType = nNull(rs!ItemType)
    RelatedID = nNull(rs!RelatedID)
    QBID = rs!QBID & ""
        
End Sub

Public Function GetByID(ByVal lngQBUpdateID As Long) As Boolean

    rsInit "SELECT * FROM QBUpdate WHERE QBUpdateID = " & lngQBUpdateID, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True
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


