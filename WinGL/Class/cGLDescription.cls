VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public ID As Long
Public Description As String
Public Number As Long

Dim rs As ADODB.Recordset
Dim SQLString As String

Dim x As String

Public Sub OpenRS()
   
   SQLString = "GLDescriptions"
   rsInit SQLString, cnDes, rs
   
   If Not rs.EOF Then
      rs.MoveFirst
      AssignFields
   End If

End Sub

Public Sub CloseRS()

   rs.Close
   Set rs = Nothing
   
End Sub


Public Function Find(ByVal DNumber As Long) As Boolean

   x = "Number = " & DNumber
   rs.Find x, 0, adSearchForward, 1
   
   If rs.EOF() Then
      Find = False
      Exit Function
   End If

   ID = rs!ID
   Description = rs!Description
   Number = rs!Number

   Find = True

End Function


Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("Description") = Description
    rs.Fields("Number") = Number
    
    rs.Update
    
End Sub
    
Public Sub GetData(lngNumber As Long)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If lngNumber <> 0 Then
        
        rs.Source = "select * from GLDescriptions " & _
                    "where Number = " & lngNumber
        rs.MaxRecords = 1
        
        Set rs.ActiveConnection = cnDes
        
        rs.Open
        
        ' Check to make sure a record
        ' actually came back
        If rs.EOF = True And rs.BOF = True Then
'            Err.Raise srCouldNotGetData, _
'                      "GLDescription - GetData", _
'                      "Description not found."
        
            Description = "Not Found GL#: " & GLAccount.Account & _
                          " " & GLAccount.DescNumber
            Number = 0
        
        
        Else
            
            ' Populate properties with
            ' results of the query.
            Description = rs!Description & ""
            Number = rs!Number & ""
        
        End If
        
    
    Else
            Err.Raise srCouldNotGetData, _
                      "GLDescriptions - GetData", _
                      "Description not found."
        Exit Sub
    
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub

Public Function GetNext() As Boolean
   rs.MoveNext
   If rs.EOF Then
      GetNext = False
   Else
      AssignFields
      GetNext = True
   End If
End Function

Public Sub Clear()
      ID = 0
      Number = 0
      Description = ""
End Sub

Public Sub DeleteAll()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Source = "Delete * from GLDescriptions"
    Set rs.ActiveConnection = cnDes
    rs.Open
    Set rs = Nothing
    
End Sub

Private Sub AssignFields()

   Number = rs!Number
   Description = rs!Description

End Sub
