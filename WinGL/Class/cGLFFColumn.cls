VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLFFColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *** FROM GLSystem.MDB ***

Private SQLString As String

Public FFColumnID As Long
Public ColNum As Byte
Public Description As String
Public ColType As Byte
Public FiscalYear As Long
Public StartNum As Byte
Public EndNum As Byte
Public Budget As Byte
Public PrintTab As Byte
Public NonPrint As Byte
Public GlobalID As Long
Public Records As Long

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM GLFFColumn"
   rsInit SQLString, cnDes, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngFFCID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngFFCID = 0 Then
       ErrMessage = "ID number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from GLFFColumn where FFColumnID = " & lngFFCID
    
    rsInit SQLString, cnDes, rs
    
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
    
    rs!ColNum = ColNum
    rs!Description = Trim(Description) & ""
    rs!ColType = ColType
    rs!FiscalYear = FiscalYear
    rs!StartNum = StartNum
    rs!EndNum = EndNum
    rs!Budget = Budget
    rs!PrintTab = PrintTab
    rs!NonPrint = NonPrint
    rs!GlobalID = GlobalID
    
    rs.Update
    
    FFColumnID = rs!FFColumnID
    
End Function

Public Sub Clear()
      
    ColNum = 0
    Description = ""
    ColType = 0
    FiscalYear = 0
    StartNum = 0
    EndNum = 0
    Budget = 0
    PrintTab = 0
    NonPrint = 0
    GlobalID = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM GLFFColumn", cnDes, rs
    
End Sub

Public Sub AssignFields()
    
    FFColumnID = nNull(rs!FFColumnID)
    ColNum = nNull(rs!ColNum)
    Description = rs!Description & ""
    ColType = nNull(rs!ColType)
    FiscalYear = nNull(rs!FiscalYear)
    StartNum = nNull(rs!StartNum)
    EndNum = nNull(rs!EndNum)
    Budget = nNull(rs!Budget)
    PrintTab = nNull(rs!PrintTab)
    NonPrint = nNull(rs!NonPrint)
    GlobalID = nNull(rs!GlobalID)
    Records = rs.RecordCount

End Sub

Public Function GetByID(ByVal FFCID As Long) As Boolean

    rsInit "SELECT * FROM GLFFColumn WHERE FFColumnID = " & FFCID, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


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



