VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLFFSched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private SQLString As String

Public FFSchedID As Long
Public GlobalID As Long
Public Account As Long
Public SortOrder As Long
Public PercentBase As Long
Public PrintTab As Byte
Public LineFeeds As Byte
Public AltDesc As String
Public Records As Long
Public ReportID As Byte
Public SignReverse As Byte

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM GLFFSched"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngFFSID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngFFSID = 0 Then
       ErrMessage = "ID number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from GLFFSched where FFSchedID = " & lngFFSID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
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

Public Function Save(ByVal booAddPut As Byte) As Boolean

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs!GlobalID = GlobalID
    rs!Account = Account
    rs!SortOrder = SortOrder
    rs!PercentBase = PercentBase
    rs!PrintTab = PrintTab
    rs!LineFeeds = LineFeeds
    rs!AltDesc = AltDesc & ""
    rs!ReportID = ReportID
    rs!SignReverse = SignReverse
    
    rs.Update
    
    FFSchedID = rs!FFSchedID
    
End Function

Public Sub Clear()
      
    GlobalID = 0
    Account = 0
    SortOrder = 0
    PercentBase = 0
    PrintTab = 0
    LineFeeds = 0
    AltDesc = ""
    ReportID = 0
    SignReverse = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM GLFFSched", cn, rs
    
End Sub

Public Sub AssignFields()
        
    FFSchedID = nNull(rs!FFSchedID)
    GlobalID = nNull(rs!GlobalID)
    Account = nNull(rs!Account)
    SortOrder = nNull(rs!SortOrder)
    PercentBase = nNull(rs!PercentBase)
    PrintTab = nNull(rs!PrintTab)
    LineFeeds = nNull(rs!LineFeeds)
    AltDesc = rs!AltDesc & ""
    Records = rs.RecordCount
    ReportID = nNull(rs!ReportID)
    SignReverse = nNull(rs!SignReverse)

End Sub

Public Function GetByID(ByVal FFSID As Long) As Boolean

    rsInit "SELECT * FROM GLFFSched WHERE FFSchedID = " & FFSID, cn, rs
    
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


