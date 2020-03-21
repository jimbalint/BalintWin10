Attribute VB_Name = "modPRConnect"
Option Explicit
Public cn As New ADODB.Connection
Public cnDes As New ADODB.Connection
Public cnPRCK As New ADODB.Connection
   
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
            x = InputBox("Enter Password:", "Open DB")
            If x = "" Then Exit Do
            Err.Clear
            cn.Properties("Jet OLEDB:Database Password") = x
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

Public Function CNPRCKOpen(ByVal FName As String, ByVal Password As String) As Boolean
   
   On Error Resume Next
   
   Set cnPRCK = New ADODB.Connection
   cnPRCK.Provider = "Microsoft.Jet.OLEDB.4.0"
   cnPRCK.ConnectionString = FName
   
   If Password <> "" Then
      cnPRCK.Properties("Jet OLEDB:Database Password") = Password
   End If
   
   cnPRCK.Open

   If Err.Number <> 0 Then
      If Err.Description = "Not a valid password." Then
         CNPRCKOpen = False
         Do
            x = InputBox("Enter Password:", "Open DB")
            If x = "" Then Exit Do
            Err.Clear
            cnPRCK.Properties("Jet OLEDB:Database Password") = x
            cnPRCK.Open
            If Err Then
               If Err.Description = "Not a valid password." Then
                  MsgBox "Incorrect password !!!", vbExclamation + vbOKOnly, "Windows GL"
               Else
                  Exit Do
               End If
            Else
               CNPRCKOpen = True
               Exit Do
            End If
         Loop
      Else
         MsgBox "Error connecting to: " & FName & " " & Err.Description & " " & Err.Number, _
                vbExclamation + vbOKOnly, "Windows PR"
      End If
   Else
      CNPRCKOpen = True
   End If
   
   On Error GoTo 0
   
   If CNPRCKOpen = False Then
      MsgBox "File Open Error: " & Err.Description & " " & Err.Number, vbExclamation + vbOKOnly, "Windows PR"
   End If
   
End Function

