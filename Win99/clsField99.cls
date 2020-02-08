VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsField99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FieldID As Long
Public FormType As String
Public TaxYear As Long
Public BoxName As String
Public FieldOrder As Long
Public FieldTitle As String
Public FieldFormat As Byte
Public HorzPosn As Long
Public VertPosn As Long
Public QuickEntry As Byte

Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM Field99"
   rsInit SQLString, cn99, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngFieldID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngFieldID = 0 Then
       ErrMessage = "FieldID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from Field99 where FieldID = " & lngFieldID
    
    rsInit SQLString, cn99, rs
    
    DeleteRecord = True
    
End Function

Public Function Find(ByVal lngFieldID As Long) As Boolean
   
   SQLString = "FieldID = " & lngFieldID
   
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
    
    rs.Fields("TaxYear") = TaxYear
    rs.Fields("FormType") = FormType
    rs.Fields("BoxName") = BoxName
    rs.Fields("FieldOrder") = FieldOrder
    rs.Fields("FieldTitle") = FieldTitle
    rs.Fields("FieldFormat") = FieldFormat
    rs.Fields("HorzPosn") = HorzPosn
    rs.Fields("VertPosn") = VertPosn
    rs.Fields("QuickEntry") = QuickEntry
    rs.Update
    FieldID = rs!FieldID
    
End Function

Public Sub Clear()
      
    TaxYear = 0
    FormType = ""
    BoxName = ""
    FieldTitle = ""
    FieldFormat = 0
    HorzPosn = 0
    VertPosn = 0
    QuickEntry = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Field99", cn99, rs
    
End Sub

Public Sub AssignFields()
        
    FieldID = nNull(rs!FieldID)
    TaxYear = nNull(rs!TaxYear)
    FormType = Trim(rs!FormType & "")
    BoxName = Trim(rs!BoxName & "")
    FieldOrder = nNull(rs!FieldOrder)
    FieldTitle = Trim(rs!FieldTitle & "")
    FieldFormat = nNull(rs!FieldFormat)
    HorzPosn = nNull(rs!HorzPosn)
    VertPosn = nNull(rs!VertPosn)
    QuickEntry = nNull(rs!QuickEntry)

End Sub

Public Function BTitle() As String
    BTitle = Trim(BoxName) & " " & Mid(Trim(FieldTitle), 1, 10)
End Function


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

    SQLString = "SELECT * FROM Field99 WHERE FieldID = " & ID
    
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



