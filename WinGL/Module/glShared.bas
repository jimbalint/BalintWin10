Attribute VB_Name = "glShared"
Public FileName As String           ' storage .mdb for company
Public curUser As Long              ' ID for the user record
Public curCompany As Long           ' ID for the company record
Public xFactory As New cFactory
Public com As New rCompany
Public use As New rUser
Public bat As New rBatch
Public Response As Boolean
Public Password As String
Public CompanyID As Long
Public DriveLetter As String
Public TestMode As Boolean

Public BackName As String

'Public bat As New rBatch

Public Cn As ADODB.Connection
Public rs As ADODB.Recordset
Public BalintFolder As String

Public PRGlobal As cPRGlobal
Public SQLString As String
Public cnDES As New ADODB.Connection
Public ErrMessage As String

Sub Main()
    
Dim i As Long
Dim x As String
Dim y As String
Dim cmdline As String
    
    GLDE = True
    
    On Error GoTo glErr
    cmdline = Command()
    
    DriveLetter = Left(App.Path, 2)
    
    OpenTab = 1
    
    If cmdline = "" Then
       
       curUser = 2
       Password = ""
       BackName = "\Balint\GLMenu.exe"
       BackName = ""
       TestMode = True
       BalintFolder = "c:\Balint"
    
    Else

       curUser = GetCmd(cmdline, "UserID", "Num")
       Password = GetCmd(cmdline, "dbPwd", "Str")
       BackName = GetCmd(cmdline, "BackName", "Str")
       BalintFolder = GetCmd(cmdline, "BalintFolder", "Str")
       TestMode = False

    End If

    
    ' =============================================================
    ' ADO connection to GLSystem
    ' used for PRGlobal Class
    '
    Set cnDES = New ADODB.Connection
    cnDES.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnDES.ConnectionString = FName
    If BalintFolder = "" Then
        cnDES.ConnectionString = "\balint\data\glSystem.mdb"
    Else
        cnDES.ConnectionString = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If
    cnDES.Mode = adModeReadWrite
    cnDES.Open
    ' =============================================================
    
    frmProgress.lblMsg2 = "Now loading information ... "
    frmProgress.Show
    
    use.GetRecord (curUser)
    
    frmProgress.lblMsg1 = com.Name
    frmProgress.lblMsg1.Refresh
    
    curCompany = use.LastCompany
    UserID = use.ID
    CompanyID = curCompany
    Set PRGlobal = New cPRGlobal
    
    If com.GetRecord(curCompany) Then
    
       ' open the company database
       If BalintFolder = "" Then
            x = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
       Else
            x = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(com.FileName)
       End If
       
       If Not CNOpen(x, Password) Then End
       UpdateCheck False, Cn
       Cn.Close
    
       MainMenu.Show
    
    End If
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)
    End
End Sub

Public Function UserName(ByVal ID As Long) As String
    On Error GoTo glErr
    UserName = "USER"
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    Set rs = db.OpenRecordset("SELECT * FROM users WHERE ID=" & ID)
    UserName = rs!Name
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function ShowValue(ByVal Amount As Currency) As String
    ShowValue = FormatCurrency(Amount, 2)
End Function

Public Function ShowDate(ByVal thisDate As Date) As String
    ShowDate = Format(thisDate, "mm/dd/yyyy")
End Function

Private Sub txtSource_GotFocus()    ' Each Box (Field) sets up this way
    cmdSave.Enabled = True          ' Save on Edit
    txtSource.SelStart = 0          ' select data on entry
    txtSource.SelLength = Len(txtSource.Text)
End Sub

Public Function GetCommandLine(Optional MaxArgs)
    Dim C, cmdline, CmdLnLen, InArg, i, NumArgs
    If IsMissing(MaxArgs) Then MaxArgs = 10
    ReDim argarray(MaxArgs)
    NumArgs = 0
    InArg = False
    cmdline = Command()
    CmdLnLen = Len(cmdline)
    For i = 1 To CmdLnLen
        C = Mid(cmdline, i, 1)
        If (C <> " " And C <> vbTab) Then
            If Not InArg Then
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                InArg = True
            End If
            argarray(NumArgs) = argarray(NumArgs) & C
        Else
            InArg = False
        End If
    Next i
    ReDim Preserve argarray(NumArgs)
    GetCommandLine = argarray()
End Function

Public Function glTypeByte(ByVal AcctType As String) As Integer
    glTypeByte = 0
    If AcctType = "0" Then glTypeByte = 1
    If AcctType = "T" Then glTypeByte = 2
    If AcctType = "H" Then glTypeByte = 3
    If AcctType = "D" Then glTypeByte = 4
    If AcctType = "I" Then glTypeByte = 5
    If AcctType = "L" Then glTypeByte = 6
    If AcctType = "A" Then glTypeByte = 7
    If AcctType = "E" Then glTypeByte = 8
    If AcctType = "U" Then glTypeByte = 9
    If AcctType = "S" Then glTypeByte = 10
    If AcctType = "." Then glTypeByte = 11
    If AcctType = "M" Then glTypeByte = 12
    If AcctType = "B" Then glTypeByte = 13
    If AcctType = "P" Then glTypeByte = 14
    If AcctType = "C" Then glTypeByte = 15
End Function

Public Function glTypeChar(ByVal ndx As Byte) As String
    glTypeChar = " "
    If ndx = 1 Then glTypeChar = "0"
    If ndx = 2 Then glTypeChar = "T"
    If ndx = 3 Then glTypeChar = "H"
    If ndx = 4 Then glTypeChar = "D"
    If ndx = 5 Then glTypeChar = "I"
    If ndx = 6 Then glTypeChar = "L"
    If ndx = 7 Then glTypeChar = "A"
    If ndx = 8 Then glTypeChar = "E"
    If ndx = 9 Then glTypeChar = "U"
    If ndx = 10 Then glTypeChar = "S"
    If ndx = 11 Then glTypeChar = "."
    If ndx = 12 Then glTypeChar = "M"
    If ndx = 13 Then glTypeChar = "B"
    If ndx = 14 Then glTypeChar = "P"
    If ndx = 15 Then glTypeChar = "C"
End Function

Public Function glTypeName(ByVal ndx As Byte) As String
    glTypeName = "ERROR"
    If ndx = 0 Then glTypeName = "BLANK"
    If ndx = 1 Then glTypeName = "ZERO ACCOUNT POSTABLE"
    If ndx = 2 Then glTypeName = "TOTAL RECORD"
    If ndx = 3 Then glTypeName = "HEADING OR DESCRIPTIVE RECORD"
    If ndx = 4 Then glTypeName = "DATE ROUTINE"
    If ndx = 5 Then glTypeName = "INCOME CATEGORY"
    If ndx = 6 Then glTypeName = "LIABILITY OR CAPITAL CATEGORY"
    If ndx = 7 Then glTypeName = "ASSET CATEGORY"
    If ndx = 8 Then glTypeName = "EXPENSE CATEGORY"
    If ndx = 9 Then glTypeName = "UNDERLINE RECORD"
    If ndx = 10 Then glTypeName = "SIGN RECORD"
    If ndx = 11 Then glTypeName = "PERCENT BASE"
    If ndx = 12 Then glTypeName = "MATH RECORD"
    If ndx = 13 Then glTypeName = "BALANCE SHEET"
    If ndx = 14 Then glTypeName = "PROFIT AND LOSS"
    If ndx = 15 Then glTypeName = "CLEARING RECORD"
End Function

Public Function glPeriodYear(ByVal fy As Integer, ByVal fp As Byte, ByVal mon As Byte) As String
    Dim v As Variant
    If mon < fp Then
        v = DateSerial(fy - 1, mon, 1)
    Else
        v = DateSerial(fy, mon, 1)
    End If
    glPeriodYear = Format(v, "mmm yyyy")
End Function

Public Function glAccountName(ByVal Acct As Long, ByRef AcctType As String) As String
    glAccountName = "Acct #" & CStr(Acct)
    On Error GoTo glErr
    Dim des As New rDescriptions
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    AcctType = ""
    
    Set db = OpenDatabase(Name:=FileName, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)
    
    Set rs = db.OpenRecordset("SELECT * FROM GLAccount WHERE Account=" & CStr(Acct))
    
    If rs.BOF And rs.EOF Then
       glAccountName = "NOT FOUND"
       Exit Function
    End If
    
    rs.MoveLast
    If rs!DescNumber > 0 Then
        glAccountName = des.GetDescription(rs!DescNumber)
    Else
        glAccountName = rs!Description
    End If
    
    AcctType = rs!AcctType
    
    rs.Close
    Exit Function
glErr:
    rs.Close
End Function

Public Function GetCmd(ByVal cmdline As String, ByVal Argument As String, ByVal StrNum As String) As Variant

' return xxxx - Argument=xxxx

Dim i As Long
Dim cmd As String

    StrNum = LCase(StrNum)
    If StrNum <> "str" And StrNum <> "num" Then
       MsgBox "StrNum not assigned !"
       GetCmd = ""
       Exit Function
    End If
    
    If StrNum = "str" Then
       GetCmd = ""
    Else
       GetCmd = 0
    End If

    ' bad value traps
    If IsNull(cmdline) Then Exit Function
    If IsNull(Argument) Then Exit Function
    If cmdline = "" Then Exit Function
    If Argument = "" Then Exit Function

    ' ignore case for argument type but keep it for the return string
    cmd = LCase(cmdline)
    Argument = LCase(Argument)

    ' search for Argument=xxxxx
    i = InStr(1, cmd, Argument, vbTextCompare)
    If i = 0 Then Exit Function
    
    ' now look for the "=" sign
    If Mid(cmdline, i + Len(Argument), 1) <> "=" Then Exit Function
    
    ' append to make return string until a space or end of line
    i = i + Len(Argument) + 1
    Do
       If i > Len(cmdline) Then Exit Do
       If Mid(cmdline, i, 1) = " " Then Exit Do
       GetCmd = GetCmd & Mid(cmdline, i, 1)
       i = i + 1
    Loop

End Function


Public Function CNOpen(ByVal FName As String, ByVal Password As String) As Boolean
   
   On Error Resume Next
   
   Set Cn = New ADODB.Connection
   Cn.Provider = "Microsoft.Jet.OLEDB.4.0"
   Cn.ConnectionString = FName
   
   If Password <> "" Then
      Cn.Properties("Jet OLEDB:Database Password") = Password
   End If
   
   Cn.Mode = adModeReadWrite
   
   Cn.Open

   If Err.Number <> 0 Then
      If Err.Description = "Not a valid password." Then
         CNOpen = False
         Do
            x = InputBox("Enter Password:", "Open DB")
            If x = "" Then Exit Do
            Err.Clear
            Cn.Properties("Jet OLEDB:Database Password") = x
            Cn.Open
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
                vbExclamation + vbOKOnly, "Windows GL"
      End If
   Else
      CNOpen = True
   End If
   
   On Error GoTo 0
   
   If CNOpen = False Then
      MsgBox "File Open Error: " & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "Windows GL"
   Else
      UpdateCheck False, Cn   ' check for new Field
   End If
   
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

Public Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String
                         
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

Public Function mdbName(ByVal str As String) As String

Dim mdbI, mdbJ, mdbK As Long

    mdbName = ""
    If str = "" Then Exit Function
    If InStr(1, str, "\", vbTextCompare) = 0 Then Exit Function
    
    mdbK = Len(str)
    For mdbI = mdbK To 1 Step -1
        If Mid(str, mdbI, 1) = "\" Then
            Exit For
        End If
    Next mdbI
    If mdbI = 0 Then Exit Function
    mdbName = Trim(Mid(str, mdbI + 1, mdbK))

End Function

Public Sub FFColumnCreate()
    ' for compatibility only
End Sub

Public Function nNull(ByVal InVal As Variant) As Variant

    If IsNull(InVal) Then
        nNull = 0
    ElseIf InVal = "" Then
        nNull = 0
    Else
        nNull = InVal
    End If

End Function

