VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRW2State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public W2StateID As Long
Public W2ID As Long
Public TaxYear As Long
Public ERStateID As String
Public StateID As Long
Public StateWage As Currency
Public StateTax As Currency

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRW2State"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngPRW2StateID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngPRW2StateID = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRW2State where PRW2StateID = " & lngPRW2StateID
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

'    SQLString = "SELECT PRW2StateID from PRW2State"
'
'    rsInit SQLString, cn, rs
    
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
    rs!StateID = nNull(StateID)
    rs!ERStateID = ERStateID & ""
    rs!StateWage = nNull(StateWage)
    rs!StateTax = nNull(StateTax)
    rs.Update
    
    W2StateID = rs!W2StateID
    
End Function

Public Sub Clear()
      
    TaxYear = 0
    W2ID = 0
    ERStateID = ""
    StateID = 0
    StateWage = 0
    StateTax = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRW2State", cn, rs
    
End Sub

Public Sub AssignFields()
        
    W2StateID = nNull(rs!W2StateID)
    TaxYear = nNull(rs!TaxYear)
    W2ID = nNull(rs!W2ID)
    StateID = nNull(rs!StateID)
    StateWage = nNull(rs!StateWage)
    StateTax = nNull(rs!StateTax)
    ERStateID = rs!ERStateID & ""

End Sub

Public Function GetByID(ByVal lngPRW2StateID As Long) As Boolean

    rsInit "SELECT * FROM PRW2State WHERE PRW2StateID = " & lngPRW2StateID, cn, rs
    
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



