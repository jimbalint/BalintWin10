VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CityID As Long
Public CityNumber As Long
Public CityRate As Currency
Public CityName As String
Public ShortName As String
Public CityRecipRate As Currency
Public StateID As Long

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRCity"
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
       ErrMessage = "City ID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRCity where CityID = " & lngID
    
    rsInit SQLString, cnDes, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    Records = rs.RecordCount

'    SQLString = "SELECT CityID from PRCity"
'
'    rsInit SQLString, cnDes, rs
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

Public Function Find(ByVal lngID As Long) As Boolean
   
   SQLString = "CityID = " & lngID
   
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
    
    rs.Fields("CityNumber") = CityNumber
    rs.Fields("CityName") = UCase(CityName)
    rs.Fields("ShortName") = UCase(ShortName)
    rs.Fields("CityRate") = CityRate
    rs.Fields("CityRecipRate") = CityRecipRate
    rs.Fields("StateID") = StateID
    
    rs.Update
    
    CityID = rs!CityID
    
End Function

Public Sub Clear()
      
    CityID = 0
    CityNumber = 0
    CityName = ""
    ShortName = ""
    CityRate = 0
    CityRecipRate = 0
    StateID = 0
    
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRCity", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    CityID = nNull(rs!CityID)
    CityNumber = nNull(rs!CityNumber)
    CityName = Trim(UCase(rs!CityName)) & ""
    ShortName = Trim(UCase(rs!ShortName)) & ""
    CityRate = nNull(rs!CityRate)
    
    If ShortName = "" Then ShortName = CityName
    
    ' *** NOT USED ***
    'CityRecipRate = rs!CityRecipRate
    CityRecipRate = 0
    
    StateID = nNull(rs!StateID)
    
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

    SQLString = "SELECT * FROM PRCity WHERE CityID = " & CStr(ID)
    
    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function
