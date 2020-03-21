Attribute VB_Name = "modGLConnect"
Option Explicit
Public cnDes As ADODB.Connection
Public cn As New ADODB.Connection
   
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
    
    ' getting mode error @ Eaglowsk
    cn.Mode = adModeReadWrite
  
    cn.Open
 
    If Err.Number <> 0 Then
        
        If Err.Description = "Not a valid password." Then
            frmEnterDBPassword.lblCompanyName = GLCompany.Name
            frmEnterDBPassword.lblFileName = GLCompany.FileName
            
            Do
                frmEnterDBPassword.tdbPassword = ""
                frmEnterDBPassword.Show vbModal
                If frmEnterDBPassword.tdbPassword = "" Then Exit Do
                Err.Clear
                cn.Properties("Jet OLEDB:Database Password") = frmEnterDBPassword.tdbPassword
                cn.Open
                If Err Then
                    If Err.Description = "Not a valid password." Then
                        MsgBox "Incorrect password !!!", vbExclamation + vbOKOnly, "Windows GL"
                    Else
                        Exit Do
                    End If
                Else
                    dbPwd = frmEnterDBPassword.tdbPassword
                    CNOpen = True
                    Exit Do
                End If
            Loop
         
'         CNOpen = False
'         Do
'            x = InputBox("Enter Password:", "Open DB")
'            If x = "" Then Exit Do
'            Err.Clear
'            cn.Properties("Jet OLEDB:Database Password") = x
'            cn.Open
'            If Err Then
'               If Err.Description = "Not a valid password." Then
'                  MsgBox "Incorrect password !!!", vbExclamation + vbOKOnly, "Windows GL"
'               Else
'                  Exit Do
'               End If
'            Else
'               CNOpen = True
'               Exit Do
'            End If
'         Loop
      
      Else
         MsgBox "Error connecting to: " & FName & " " & Err.Description & " " & Err.Number, _
                vbExclamation + vbOKOnly, "Windows GL"
      End If
   Else
      CNOpen = True
   End If
   
   On Error GoTo 0
   
   If CNOpen = False Then
      MsgBox "File Open Error: " & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "Windows GL"
      ' End
   Else
      ' check for field modifications
      If Not NoFieldCheck Then UpdateCheck False, cn
   End If
   
End Function

Public Function CNDesOpen(ByVal FName As String) As Boolean
   
    On Error Resume Next
    cnDes.Close
    Set cnDes = Nothing
    On Error GoTo 0
   
    Set cnDes = New ADODB.Connection
    If Right(LCase(FName), 6) = ".accdb" Then
        cnDes.Provider = "Microsoft.ACE.OLEDB.12.0"
    Else
        cnDes.Provider = "Microsoft.Jet.OLEDB.4.0"
    End If
    
    cnDes.ConnectionString = FName
   
    On Error Resume Next
    cn.Mode = adModeReadWrite
    On Error GoTo 0
   
    cnDes.Open
    CNDesOpen = True
    UpdateCheck True, cnDes

End Function

Public Sub rsInit(ByVal SQLString As String, _
                  ByRef cni As ADODB.Connection, _
                  ByRef rsi As ADODB.Recordset)
   
   Set rsi = New ADODB.Recordset
   rsi.Source = SQLString
   rsi.ActiveConnection = cni
    
       
       
           rsi.CursorType = adOpenKeyset
   rsi.LockType = adLockOptimistic
   rsi.Open

End Sub
