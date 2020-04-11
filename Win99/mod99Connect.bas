Attribute VB_Name = "mod99Connect"
Option Explicit

Public cn As New ADODB.Connection
Public cnDes As New ADODB.Connection
Public cn99 As New ADODB.Connection
   
Public Function CNOpen(ByVal FName As String, ByVal Password As String) As Boolean
   
   On Error Resume Next
   
   Set cn = New ADODB.Connection
   
    If Right(LCase(FName), 6) = ".accdb" Then
        cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    Else
        cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    End If
   
   cn.ConnectionString = FName
   
   If Password <> "" Then
      cn.Properties("Jet OLEDB:Database Password") = Password
   End If
   
   ' getting mode error @ eaglowski
   cn.Mode = adModeReadWrite
   
   cn.Open

   If Err.Number <> 0 Then
      If Err.Description = "Not a valid password." Then
         CNOpen = False
         Do
            X = InputBox("Enter Password:", "Open DB")
            If X = "" Then Exit Do
            Err.Clear
            cn.Properties("Jet OLEDB:Database Password") = X
            cn.Open
            If Err Then
               If Err.Description = "Not a valid password." Then
                  MsgBox "Incorrect password !!!", vbExclamation + vbOKOnly, "Windows GL"
               Else
                  Exit Do
               End If
            Else
               CNOpen = True
               Exit Do
            End If
         Loop
      Else
         MsgBox "Error connecting to: " & FName & " " & Err.Description & " " & Err.Number, _
                vbExclamation + vbOKOnly, "Windows PR"
      End If
   Else
      CNOpen = True
   End If
   
   On Error GoTo 0
   
   If CNOpen = False Then
      MsgBox "File Open Error: " & Err.Description & " " & Err.Number, vbExclamation + vbOKOnly, "Windows GL"
   Else
      ' check for field modifications
      UpdateCheck False, cn
   End If
   
End Function

Public Function SysOpen(ByVal FName As String) As Boolean

    Set cnDes = New ADODB.Connection
   
    If Right(LCase(FName), 6) = ".accdb" Then
        cnDes.Provider = "Microsoft.ACE.OLEDB.12.0"
    Else
        cnDes.Provider = "Microsoft.Jet.OLEDB.4.0"
    End If
   
    cnDes.ConnectionString = FName
   
    On Error Resume Next
    cnDes.Mode = adModeReadWrite
    On Error GoTo 0
    
    cnDes.Open
    SysOpen = True
    UpdateCheck True, cnDes

End Function

Public Function CN99Open(ByVal FName As String) As Boolean
   
   Set cn99 = New ADODB.Connection
   
    If Right(LCase(FName), 6) = ".accdb" Then
        cn99.Provider = "Microsoft.ACE.OLEDB.12.0"
    Else
        cn99.Provider = "Microsoft.Jet.OLEDB.4.0"
    End If
   
   cn99.ConnectionString = FName
   
   On Error Resume Next
   cn99.Mode = adModeReadWrite
   On Error GoTo 0
   
   cn99.Open

End Function

Public Sub rsInit(ByVal SQLString As String, _
                  ByRef cni As ADODB.Connection, _
                  ByRef rsi As ADODB.Recordset)
   
    On Error Resume Next
    rsi.Close
    On Error GoTo 0
   
    Set rsi = New ADODB.Recordset
    rsi.Source = SQLString
    rsi.ActiveConnection = cni
   
    ' open as disconnected - client side
    If DisConn Then
        rsi.CursorLocation = adUseClient
        rsi.CursorType = adOpenKeyset
        rsi.LockType = adLockBatchOptimistic
    Else
        rsi.CursorLocation = adUseServer
        rsi.CursorType = adOpenKeyset
        rsi.LockType = adLockOptimistic
    End If

    rsi.Open

    If DisConn Then
        Set rsi.ActiveConnection = Nothing
        DisConn = False
    End If

End Sub

Public Sub UpdateCheck(ByVal GLSys As Boolean, _
                       ByRef adoConn As ADODB.Connection)

    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim urs As ADODB.Recordset
    Dim uCount As Long
    Dim ucmd As ADODB.Command

    If GLSys = True Then
        If AddField("GLCompany", "FederalID", "String", adoConn) Then
        End If
        If AddField("GLCompany", "SSN", "String", adoConn) Then
        End If
    End If


End Sub

Public Function AddField(ByVal TableName As String, _
                         ByVal ColumnName As String, _
                         ByVal ColumnType As String, _
                         ByRef adoConn As ADODB.Connection) _
                         As Byte
                         
Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String
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
   
    ' add it - with retry
    Do
        
        On Error Resume Next
        
        FString = "ALTER TABLE " & TableName & _
                  " ADD COLUMN [" & ColumnName & "]" & _
                  " " & ColumnType
        adoConn.Execute FString
        
        If Err.Number = 0 Then
            AddField = 1
            Exit Do
        Else
            If InStr(1, LCase(Err.Description), "could not lock", vbTextCompare) Then
                MsgResponse = MsgBox("Database update not complete" & vbCr & _
                              "ALL other users must exit to proceed!", vbRetryCancel + vbExclamation)
                If MsgResponse = vbCancel Then
                    MsgBox "Update not complete - aborting ...", vbExclamation
                    End
                End If
            Else
                MsgBox TableName & "/" & ColumnName & " " & ColumnType & _
                     vbCrLf & vbCrLf & "Field Add Error" & Err.Description, _
                     vbOKOnly + vbCritical
                AddField = 0
                End
            End If
        End If
    
    Loop
    
End Function

