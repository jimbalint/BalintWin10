VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRW2City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public W2CityID As Long
Public W2ID As Long
Public TaxYear As Long
Public CityID As Long
Public CityName As String
Public CityWage As Currency
Public CityTax As Currency
Public StateID As Long
Public SDTax As Byte
Public Courtesy As Byte

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRW2City"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngPRW2CityID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngPRW2CityID = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRW2City where PRW2CityID = " & lngPRW2CityID
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

'    SQLString = "SELECT PRW2CityID from PRW2City"
'
'    rsInit SQLString, cn, rs
'
    ' Check to make sure a record
    ' actually came back
        
    If rs.EOF = True And rs.BOF = True Then
       Records = 0
    Else
       Records = rs.RecordCount
    End If

'    rs.Close

End Function

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

Public Function Save(ByVal booAddPut As Byte) As String
    
    If booAddPut = True Then      ' Add
       rs.AddNew
    End If

    rs!W2ID = nNull(W2ID)
    rs!TaxYear = nNull(TaxYear)
    rs!CityID = nNull(CityID)
    rs!CityName = CityName & ""
    rs!CityWage = nNull(CityWage)
    rs!CityTax = nNull(CityTax)
    rs!StateID = nNull(StateID)
    rs!SDTax = nNull(SDTax)
    rs!Courtesy = nNull(Courtesy)
    rs.Update
    
    W2CityID = rs!W2CityID
    
End Function

Public Sub Clear()
      
    TaxYear = 0
    W2ID = 0
    CityID = 0
    CityName = ""
    CityWage = 0
    CityTax = 0
    StateID = 0
    SDTax = 0
    Courtesy = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRW2City", cn, rs
    
End Sub

Public Sub AssignFields()
        
    W2CityID = nNull(rs!W2CityID)
    TaxYear = nNull(rs!TaxYear)
    W2ID = nNull(rs!W2ID)
    CityID = nNull(rs!CityID)
    CityName = TextSet(rs!CityName)
    CityWage = nNull(rs!CityWage)
    CityTax = nNull(rs!CityTax)
    StateID = nNull(rs!StateID)
    SDTax = nNull(rs!SDTax)
    Courtesy = nNull(rs!Courtesy)
    
End Sub

Public Function GetByID(ByVal lngPRW2CityID As Long) As Boolean

    rsInit "SELECT * FROM PRW2City WHERE PRW2CityID = " & lngPRW2CityID, cn, rs
    
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


