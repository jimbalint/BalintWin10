VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, LastCompany, LastPRCompany As Long
Public Logon, Name, Password As String
Public LoadLastCompany As Boolean

Dim SQLString As String
Dim rs As ADODB.Recordset
Dim BatchNum As Long

Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = Equate.RecAdd Then
       rs.AddNew
    End If

    rs.Fields("LastCompany") = nNull(LastCompany)
    rs.Fields("LastPRCompany") = nNull(LastPRCompany)
    rs.Fields("Logon") = Logon & ""
    rs.Fields("Name") = Name & ""
    rs.Fields("Password") = Password & " "
    rs.Fields("LoadLastCompany") = nNull(LoadLastCompany)
    
    rs.Update
    
    ' rs.Close
    ' Set rs = Nothing
    
End Sub
    
Public Function GetBySQL(ByVal SQLString As String) As Boolean

    Set rs = New ADODB.Recordset
    rsInit SQLString, cnDes, rs
    
    ' not found
    If rs.EOF = True And rs.BOF = True Then
       GetBySQL = False
       Exit Function
    End If
       
    AssignFields
    GetBySQL = True
    
End Function

Public Sub Clear()
        
    LastCompany = 0
    Logon = ""
    Name = ""
    Password = ""
    LoadLastCompany = False
    LastPRCompany = 0
        
End Sub

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

Public Sub AssignFields()
    
    ID = rs!ID
    LastCompany = nNull(rs!LastCompany)
    Logon = Trim(rs!Logon & "")
    Name = Trim(rs!Name & "")
    Password = Trim(rs!Password & "")
    LoadLastCompany = nNull(rs!LoadLastCompany)
    LastPRCompany = nNull(rs!LastPRCompany)

End Sub

Public Function GetByID(ByVal UsrID As Long) As Boolean

    rsInit "SELECT * FROM Users WHERE ID = " & CStr(UsrID), cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

