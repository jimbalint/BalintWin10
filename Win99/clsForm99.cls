VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FormID As Long
Public FormType As String
Public TaxYear As Long
Public FormsPerPg As Byte
Public VersionNum As String
Public FormVert1 As Long
Public FormVert2 As Long
Public FormVert3 As Long
Public FormVert4 As Long

Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM Form99"
   rsInit SQLString, cn99, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngFormID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngFormID = 0 Then
       ErrMessage = "FormID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from Form99 where FormID = " & lngFormID
    
    rsInit SQLString, cn99, rs
    
    DeleteRecord = True
    
End Function

Public Function Find(ByVal lngFormID As Long) As Boolean
   
   SQLString = "FormID = " & lngFormID
   
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
    
    rs.Fields("FormType") = FormType
    rs.Fields("TaxYear") = TaxYear
    rs.Fields("FormsPerPg") = FormsPerPg
    rs.Fields("VersionNum") = VersionNum
    rs.Fields("FormVert1") = FormVert1
    rs.Fields("FormVert2") = FormVert2
    rs.Fields("FormVert3") = FormVert3
    rs.Fields("FormVert4") = FormVert4
    rs.Update

    FormID = rs!FormID
    
End Function

Public Sub Clear()
      
    FormID = 0
    FormType = ""
    TaxYear = 0
    FormsPerPg = 0
    VersionNum = ""
    FormVert1 = 0
    FormVert2 = 0
    FormVert3 = 0
    FormVert4 = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Form99", cn99, rs
    
End Sub

Public Sub AssignFields()
        
    FormID = nNull(rs!FormID)
    
    FormType = rs!FormType & ""
    TaxYear = nNull(rs!TaxYear)
    FormsPerPg = nNull(rs!FormsPerPg)
    VersionNum = rs!VersionNum & ""
    FormVert1 = nNull(rs!FormVert1)
    FormVert2 = nNull(rs!FormVert2)
    FormVert3 = nNull(rs!FormVert3)
    FormVert4 = nNull(rs!FormVert4)
    
End Sub

Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cn99, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
   Records = rs.RecordCount
                             
End Function

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM Form99 WHERE FormID = " & ID
    
    rsInit SQLString, cn99, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

    Records = 1

End Function
Public Function Count() As Long
    Count = rs.RecordCount
End Function


