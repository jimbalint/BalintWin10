VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRCounty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CountyID As Long
Public CountyName As String
Public ShortName As String
Public StateID As Long
Public SalesTaxRate As Double

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRCounty"
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
       ErrMessage = "County ID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRCounty where CountyID = " & lngID
    
    rsInit SQLString, cnDes, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT CountyID from PRCounty"
    
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
   
   SQLString = "CountyID = " & lngID
   
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
    
    rs.Fields("CountyName") = UCase(CountyName)
    rs.Fields("ShortName") = UCase(ShortName)
    rs.Fields("StateID") = nNull(StateID)
    rs.Fields("SalesTaxRate") = nNull(SalesTaxRate)
    
    rs.Update
    
    CountyID = rs!CountyID
    
End Function

Public Sub Clear()
      
    CountyID = 0
    CountyName = ""
    ShortName = ""
    StateID = 0
    SalesTaxRate = 0
    
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRCounty", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    CountyID = nNull(rs!CountyID)
    CountyName = Trim(UCase(rs!CountyName)) & ""
    ShortName = Trim(UCase(rs!ShortName)) & ""
    If ShortName = "" Then ShortName = CountyName
    StateID = nNull(rs!StateID)
    SalesTaxRate = nNull(rs!SalesTaxRate)

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

    SQLString = "SELECT * FROM PRCounty WHERE CountyID = " & ID
    
    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


