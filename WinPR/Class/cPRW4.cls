VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRW4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public W4ID As Long
Public EmployeeID As Long
Public TwoJobs As Byte
Public FilingType As Byte
Public Dependents As Byte
Public DependentsOther As Byte
Public OtherIncome As Currency
Public Deductions As Currency
Public ExtraWH As Currency

Dim rs As New ADODB.Recordset

Public Sub OpenRS()
   
   SQLString = "SELECT * FROM PRW4"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngPRW4ID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngPRW4ID = 0 Then
       ErrMessage = "W4ID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRW4 where PRW4ID = " & lngPRW4ID
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

'    SQLString = "SELECT PRW4ID from PRW4"
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

    rs!EmployeeID = nNull(EmployeeID)
    rs!TwoJobs = nNull(TwoJobs)
    rs!FilingType = nNull(FilingType)
    rs!Dependents = nNull(Dependents)
    rs!DependentsOther = nNull(DependentsOther)
    rs!OtherIncome = nNull(OtherIncome)
    rs!Deductions = nNull(Deductions)
    rs!ExtraWH = nNull(ExtraWH)
    
    rs.Update
    W4ID = W4ID
    
End Function

Public Sub Clear()
      
    W4ID = 0
    EmployeeID = 0
    TwoJobs = 0
    FilingType = 0
    Dependents = 0
    DependentsOther = 0
    OtherIncome = 0
    Deductions = 0
    ExtraWH = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRW4", cn, rs
    
End Sub

Public Sub AssignFields()
        
    W4ID = nNull(rs!W4ID)
    EmployeeID = nNull(rs!EmployeeID)
    TwoJobs = nNull(rs!TwoJobs)
    FilingType = nNull(rs!FilingType)
    Dependents = nNull(rs!Dependents)
    DependentsOther = nNull(rs!DependentsOther)
    OtherIncome = nNull(rs!OtherIncome)
    Deductions = nNull(rs!Deductions)
    ExtraWH = nNull(rs!ExtraWH)
    
End Sub

Public Function GetByID(ByVal lngPRW4ID As Long) As Boolean

    rsInit "SELECT * FROM PRW4 WHERE PRW4ID = " & lngPRW4ID, cn, rs
    
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



