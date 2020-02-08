Attribute VB_Name = "modGlobal"
Option Explicit
Public cn As ADODB.Connection
Public SQLString As String
Public Client As clsClient
Public Customer As clsCustomer
Public RecAdd As Boolean
Public RecPut As Boolean

Public Function AddField(ByVal TableName As String, _
                         ByVal ColumnName As String, _
                         ByVal ColumnType As String, _
                         ByRef adoConn As ADODB.Connection) _
                         As Byte
                         
Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim fString As String
Dim TblExists As Boolean
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
   
    FldFlag = False
    TblExists = False
    Do Until frs.EOF = True
      
       If UCase(frs!Table_Name) = UCase(TableName) Then
          TblExists = True
       End If
      
       If UCase(frs!Table_Name) = UCase(TableName) And UCase(frs!Column_Name) = UCase(ColumnName) Then
         FldFlag = True
         Exit Do
       End If
      
       frs.MoveNext
   
   Loop

   ' the table was not found
   If TblExists = False Then Exit Function

   ' the field already exists - no need to add it
   If FldFlag = True Then
      AddField = 0
      Exit Function
   End If
   
   frs.Close
   Set frs = Nothing
   
   ' add it
   On Error Resume Next
   fString = "ALTER TABLE " & TableName & _
             " ADD COLUMN " & ColumnName & _
             " " & ColumnType
   adoConn.Execute fString
   
   If Err.Number <> 0 Then
      GoTo AddFieldError
   End If
   
   AddField = 1
   Exit Function
   
AddFieldError:
   
   ' problem with field add !?
   MsgBox TableName & "/" & ColumnName & " " & ColumnType & _
       vbCrLf & vbCrLf & "Field Add Error" & Err.Description, _
       vbOKOnly + vbCritical
   AddField = 0
   End

End Function


Public Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim fString As String
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
           
    TableExists = False
           
    Do Until frs.EOF = True
                  
        If frs!Table_Name = TableName Then
            TableExists = True
            Exit Do
        End If
        
       frs.MoveNext
   
   Loop

End Function

Public Sub DropTable(ByVal TableName As String, _
                      ByVal adoCn As ADODB.Connection)

' *** Drop a table if it exists ***

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim TableFlag As Boolean
Dim fString As String
                         
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoCn.OpenSchema(adSchemaColumns)
       
    TableFlag = False
       
    Do Until frs.EOF = True
              
        If frs!Table_Name = TableName Then
            TableFlag = True
            Exit Do
        End If
      
        frs.MoveNext
   
    Loop

    frs.Close
    
    ' table does not exist
    If TableFlag = False Then Exit Sub

    fString = "DROP TABLE " & TableName
    adoCn.Execute fString

End Sub

Public Sub rsInit(ByVal SQLString As String, _
                  ByRef cni As ADODB.Connection, _
                  ByRef rsi As ADODB.Recordset)
   
    Set rsi = New ADODB.Recordset
    rsi.Source = SQLString
    rsi.ActiveConnection = cni
   
    rsi.CursorLocation = adUseServer
    rsi.CursorType = adOpenKeyset
    rsi.LockType = adLockOptimistic
    rsi.Open

End Sub

Public Function nNull(ByVal InVal As Variant) As Variant

    If IsNull(InVal) Then
        nNull = 0
    Else
        nNull = InVal
    End If

End Function

