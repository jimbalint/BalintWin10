VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DepartmentID As Long
Public DepartmentNumber As Long
Public Name As String
Public WkcCat As Long

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRDepartment"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngDept As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngDept = 0 Then
       ErrMessage = "Department number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRDepartment where DepartmentNumber = " & lngDept
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    Records = rs.RecordCount

'    SQLString = "SELECT DepartmentNumber from PRDepartment"
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

Public Function Find(ByVal lngDept As Long) As Boolean
   
   SQLString = "DepartmentNumber = " & lngDept
   
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
    
    rs.Fields("DepartmentNumber") = DepartmentNumber
    rs.Fields("Name") = Trim(UCase(Name)) & ""
    rs.Fields("WkcCat") = WkcCat
    
    rs.Update
    
    DepartmentID = rs!DepartmentID
    
End Function

Public Sub Clear()
      
    DepartmentID = 0
    DepartmentNumber = 0
    Name = ""
    WkcCat = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRDepartment", cn, rs
    
End Sub

Public Sub AssignFields()
        
    DepartmentID = nNull(rs!DepartmentID)
    DepartmentNumber = nNull(rs!DepartmentNumber)
    Name = Trim(UCase(rs!Name)) & ""
    WkcCat = nNull(rs!WkcCat)
    
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

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRDepartment WHERE DepartmentID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

