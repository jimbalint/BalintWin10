Attribute VB_Name = "modConvertADO"
Option Explicit

Dim cnSys As ADODB.Connection
Dim cnNew As ADODB.Connection
Dim cnOld As ADODB.Connection

Dim strSQL As String

Dim rs As New ADODB.Recordset
Dim rsNew As New ADODB.Recordset
Dim frs As New ADODB.Recordset
Dim log As Integer

Dim x, Y, z As String
Dim I, J, K As Long

Dim rsNewSchema As ADODB.Recordset
Dim frm As New frmProgress

Dim rc4Key As String
Dim BalintFolder As String
Dim NewFolder As String
Dim dbBlank As String


Public Sub RunADO_Conversion(ByVal BalintFolder As String)

    Dim resp As Integer
    x = "Convert ALL DBs to New ADO?" & vbCr & "Make sure ALL users are out of the software!"
    resp = MsgBox(x, vbExclamation + vbYesNo, "Balint Windows Acctg")
    If resp = vbNo Then End

    On Error Resume Next
    cn.Close
    cnDes.Close
    On Error GoTo 0

    Dim fnm As String
    Dim nfnm As String
    Dim cnm As String
    Dim rsC As New ADODB.Recordset
    
    rc4Key = "B@lint19742101!@#$%^&*"
    
'    BalintFolder = "\\vboxsrv\vm-share\Balint"
'    BalintFolder = "c:\Balint"
    
    BalintFolder = Replace(BalintFolder, "^", " ")
    NewFolder = BalintFolder & "\Data_New"
    
    If Len(Dir(NewFolder, vbDirectory)) > 0 Then
        MsgBox "Folder already exists: " & NewFolder, vbExclamation, "Data Conversion"
        On Error Resume Next
        Kill NewFolder & "\*.*"
        On Error GoTo 0
    Else
        MkDir NewFolder
    End If

    'dbBlank = "c:\Balint\Data\BlankAccdb.accdb"
    ' FileCopy dbBlank, BalintFolder & "\Blank\BlankAccdb.accdb"
    
    log = FreeFile
    Open NewFolder & "\ConvertLog.txt" For Output As #log
    
    Set cnSys = SQLConnect(BalintFolder & "\Data\GLSystem.mdb")
    
    x = BalintFolder & "\Data_New\GLSystem.accdb"
    dbBlank = BalintFolder & "\Blank\BlankAccdb.accdb"
    If Len(Dir(dbBlank, vbNormal)) = 0 Then
        MsgBox "Blank DB not found: " & Y, vbExclamation, "Data Conversion"
        End
    End If
    FileCopy dbBlank, x
    Set cnNew = SQLConnect(x)

    frm.Show
    frm.lblMsg1 = "Now converting GLSystem.mdb"
    frm.lblMsg2 = ""
    frm.lblMsg3 = ""
    frm.Refresh
    
    ' convert GLSystem.mdb
    InitSchemaRS
    PopSchemaRS cnSys
    CreateTables cnSys, cnNew
    CreateFields cnNew
    CopyData cnSys, cnNew
    strSQL = "ALTER TABLE glDescriptions" & _
            " ADD CONSTRAINT [Number] UNIQUE ([Number])"
    cnNew.Execute strSQL
    cnNew.Close
    Print #log, "GLSystem.mdb converted" & vbCrLf
    
    ' convert company files
    frm.lblMsg2 = ""
    frm.lblMsg3 = ""
    frm.Refresh
    
    strSQL = "select * from GLCompany"
    rsC.Source = strSQL
    rsC.LockType = adLockOptimistic
    rsC.CursorType = adOpenKeyset
    rsC.CursorLocation = adUseServer
    Set rsC.ActiveConnection = cnSys
    rsC.Open
    Do While Not rsC.EOF
        
        fnm = rsC!FileName
        cnm = rsC!Name
        frm.Hide
        frm.lblMsg1 = "Now converting: " & cnm
        frm.lblMsg2 = ""
        frm.lblMsg3 = ""
        frm.Show
        frm.Refresh
        
        If BalintFolder = "" Then
            fnm = Mid(App.Path, 1, 2) & Mid(fnm, 3, Len(fnm) - 2)
        Else
            fnm = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(fnm)
        End If
        If Len(Dir(fnm, vbNormal)) = 0 Then
            Print #log, fnm & " not found for: " & cnm & vbCrLf
        Else
            Print #log, vbCrLf & "Converting: " & fnm & " for: " & cnm
            nfnm = NewFolder & "\" & Replace(mdbName(fnm), ".mdb", ".accdb")
            FileCopy dbBlank, nfnm
            
            Set cnOld = SQLConnect(fnm)
            Set cnNew = SQLConnect(nfnm)
            InitSchemaRS
            PopSchemaRS cnOld
            CreateTables cnOld, cnNew
            CreateFields cnNew
            CopyData cnOld, cnNew
            
            On Error Resume Next
            strSQL = "DELETE * FROM PRDepartment WHERE DepartmentNumber = 0"
            cnNew.Execute strSQL
            strSQL = "ALTER TABLE PRDepartment" & _
                    " ADD CONSTRAINT [dptNumberKey] UNIQUE ([DepartmentNumber])"
            cnNew.Execute strSQL
            
            strSQL = "ALTER TABLE PREmployee" & _
                    " ADD CONSTRAINT [empNumberKey] UNIQUE ([EmployeeNumber])"
            cnNew.Execute strSQL
            On Error GoTo 0
            
            cnOld.Close
            cnNew.Close
            Print #log, ""
            Print #log, ""
        
        End If
        
        ' CNOpen x, dbPwd
        rsC.MoveNext
    Loop
    rsC.Close
           
    ' Win1099.mdb
    fnm = "Win1099.mdb"
    frm.lblMsg1 = "Now converting: 1099 Data"
    frm.lblMsg2 = ""
    frm.lblMsg3 = ""
    frm.Refresh
    If BalintFolder = "" Then
        fnm = Mid(App.Path, 1, 2) & Mid(fnm, 3, Len(fnm) - 2)
    Else
        fnm = Replace(BalintFolder, "^", " ") & "\Data\" & fnm
    End If
    
    If Len(Dir(fnm, vbNormal)) > 0 Then
        Print #log, vbCrLf & "Converting: " & fnm
        nfnm = NewFolder & "\" & Replace(mdbName(fnm), ".mdb", ".accdb")
        
        FileCopy dbBlank, nfnm
        
        Set cnOld = SQLConnect(fnm)
        Set cnNew = SQLConnect(nfnm)
        InitSchemaRS
        PopSchemaRS cnOld
        CreateTables cnOld, cnNew
        CreateFields cnNew
        CopyData cnOld, cnNew
        
        cnOld.Close
        cnNew.Close
        Print #log, ""
        Print #log, ""
    End If
           
'    If op = 0 Then
'    Else
'    End If
   
    ' ==================
    On Error Resume Next
    cnSys.Close
    Set cnSys = Nothing
    cnNew.Close
    Set cnNew = Nothing
    cnOld.Close
    Set cnOld = Nothing
    On Error GoTo 0
    
    Close #log
    frm.Hide
    Set frm = Nothing
    
    ' copy the files for PR Check setup
    Copy2 BalintFolder & "\Data", BalintFolder & "\Data_New", "PRCK*.mdb"
    Copy2 BalintFolder & "\Data", BalintFolder & "\Data_New", "*.jpg"
    
    ' =======================================================
    MsgBox "Conversion complete, hit OK to ReName and complete!", vbInformation
    Name BalintFolder & "\Data" As BalintFolder & "\Data_Old"
    Name BalintFolder & "\Data_New" As BalintFolder & "\Data"
    End
    ' =======================================================

End Sub

Private Sub CopyData(ByRef cnFrom As ADODB.Connection, ByRef cnTo As ADODB.Connection)
    Set frs = cnFrom.OpenSchema(adSchemaTables)
    Do Until frs.EOF = True
        x = frs!Table_Name
        If Left(x, 4) <> "MSys" And Left(x, 5) <> "Paste" Then
            frm.lblMsg2 = "Now converting table: " & x
            frm.Refresh
            CopyDataProcess x, cnFrom, cnTo
        End If
        frs.MoveNext
    Loop
End Sub


Private Sub CopyDataProcess(ByVal TblName As String, ByRef cnFrom As ADODB.Connection, ByRef cnTo As ADODB.Connection)

    Dim Ct1, Ct2 As Long
    Dim eFlag As Boolean
    
    Dim io As Integer
    io = FreeFile
    ' Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_SQL.txt" For Output As #io
    
    strSQL = "delete * from " & TblName
    cnTo.Execute strSQL
    
    Dim fld As ADODB.Field
    strSQL = "select * from " & TblName
    
'    If TblName = "InvBody" Then
'        strSQL = "select * from " & TblName & _
'                " where BodyID <= 20480 or BodyID >= 20496"
'    End If
    
    Set rs = New ADODB.Recordset
    rs.Source = strSQL
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseServer
    Set rs.ActiveConnection = cnFrom
    rs.Open
    Ct1 = rs.RecordCount
    
    Set rsNew = New ADODB.Recordset
    rsNew.Source = strSQL
    rsNew.LockType = adLockOptimistic
    rsNew.CursorType = adOpenKeyset
    rsNew.CursorLocation = adUseServer
    Set rsNew.ActiveConnection = cnTo
    rsNew.Open
    
    Dim ffield As String
    
    Ct2 = 0
    Do While Not rs.EOF
    
        ' If TblName = "InvBody" And Ct2 = 20400 Then GoTo NxtRec
    
        ffield = ""
        rsNew.AddNew
        For Each fld In rs.Fields

            If ffield = "" Then ffield = rs.Fields(fld.Name)
        
            x = fld.Name & vbTab & rs.Fields(fld.Name)
            eFlag = False
            If TblName = "PREmployee" And fld.Name = "SSN" Then eFlag = True
            ' If TblName = "Detail99" And fld.Name = "PayeeID" Then eFlag = True
            If TblName = "Payee99" And fld.Name = "FederalID" Then eFlag = True
            If eFlag Then
                Y = RC4Encrypt(rs.Fields(fld.Name), rc4Key)
            Else
                Y = rs.Fields(fld.Name)
            End If
            Select Case Y
                Case "True": rsNew.Fields(fld.Name) = 1
                Case "False": rsNew.Fields(fld.Name) = 0
                
                ' use nNull for numeric fields only!!!
                ' PRGlobal.Var fields set to "0"
                ' Case Else: rsNew.Fields(fld.Name) = nNull(Y)
                Case Else: rsNew.Fields(fld.Name) = Y
                
            End Select
        Next fld
            
        If ffield = "" Then
            x = "Skipping " & TblName & " Rec#: " & Ct2 + 1
            MsgBox x
            Print #log, x
            Print #log, "----"
            GoTo NxtRec
        End If
            
        On Error Resume Next
        rsNew.Update
        If Err.Number <> 0 Then
            ' x = "Error adding: " & TblName & " " & fld.Name & " " & Y & " " & Ct2
            x = "Error adding: " & TblName & " " & Y & " " & Ct2 & " >>>" & ffield & "<<<"
            x = x & vbCr & ">>> " & Err.Description
            MsgBox x
            Print #log, x
            Print #log, "----"
            ' K = MsgBox("Stop???", vbExclamation + vbYesNo, "Data Conversion")
            K = vbNo
            If K = vbYes Then
                rs.Close
                rsNew.Close
                On Error Resume Next
                cnSys.Close
                Set cnSys = Nothing
                cnNew.Close
                Set cnNew = Nothing
                cnOld.Close
                Set cnOld = Nothing
                On Error GoTo 0
                Close #log
                frm.Hide
                Set frm = Nothing
                End
            Else
                rsNew.Cancel
            End If
        End If
        On Error GoTo 0
        
NxtRec: Ct2 = Ct2 + 1
        If Ct2 Mod 100 = 1 Then
            frm.lblMsg3 = TblName & " " & Ct2 & " of: " & Ct1
            frm.Refresh
        End If
        
        On Error Resume Next
        rs.MoveNext
        If Err.Number <> 0 Then
            MsgBox "MoveNext err: " & Err.Description
            End
        End If
        
    Loop
    rs.Close
    rsNew.Close
    frm.lblMsg3 = TblName & " " & Ct2 & " of: " & Ct1
    frm.Refresh
    Print #log, "--- " & TblName & " Records Converted: " & Ct2
    
End Sub


Private Sub InitSchemaRS()

    Set rsNewSchema = New ADODB.Recordset
    rsNewSchema.CursorLocation = adUseClient
    rsNewSchema.Fields.Append "TableName", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "FieldName", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "FieldNum", adInteger
    rsNewSchema.Fields.Append "ConstraintName", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "FieldType", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "FieldType2", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "MaxLength", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "Precision", adVarChar, 100, adFldIsNullable
    rsNewSchema.Fields.Append "Scale", adVarChar, 100, adFldIsNullable
    rsNewSchema.Open , , adOpenDynamic, adLockOptimistic

End Sub

Private Sub PopSchemaRS(ByRef cn As ADODB.Connection)
    
    ' fields
    Dim FldNum As Integer
    FldNum = 0
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cn.OpenSchema(adSchemaColumns)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" And frs!Table_Name <> "Paste Errors" Then
            FldNum = FldNum + 1
            rsNewSchema.AddNew
            rsNewSchema!TableName = frs!Table_Name
            rsNewSchema!FieldName = frs!Column_Name
            rsNewSchema!FieldNum = frs!ORDINAL_POSITION
            rsNewSchema!ConstraintName = ""
            rsNewSchema!MaxLength = frs!CHARACTER_MAXIMUM_LENGTH
            rsNewSchema!Precision = frs!NUMERIC_PRECISION
            rsNewSchema!Scale = frs!NUMERIC_SCALE
            
            rsNewSchema!FieldType = frs!Data_Type
            
            x = ""
            I = frs!Data_Type
            ' *** 4 = double
            Select Case I
                Case 2: x = "Short"
                Case 3: x = "Long"
                Case 4: x = "Double"
                Case 5: x = "Double"
                Case 6: x = "Currency"
                Case 7: x = "DateTime"
                Case 11: x = "Logical"
                Case 17: x = "Byte"
                Case 130
                    If rsNewSchema!MaxLength = 255 Then
                        x = "LongText"
                    Else
                        x = "Char(" & rsNewSchema!MaxLength & ")"
                    End If
                Case Else
                    MsgBox "Data Type NF: " & I
                    End
            End Select
            rsNewSchema!FieldType2 = x
            rsNewSchema.Update
        End If
        frs.MoveNext
    Loop
    
    
    
    ' dump it
'    Dim io As Integer
'    Dim fld As ADODB.Field
'    io = FreeFile
'    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Schema.txt" For Output As #io
'    rsNewSchema.Sort = "TableName ASC, FieldNum ASC"
'    rsNewSchema.MoveFirst
'    Do While Not rsNewSchema.EOF
'        x = rsNewSchema!TableName & vbTab
'        x = x & rsNewSchema!FieldName & vbTab
'        x = x & rsNewSchema!FieldNum & vbTab
'        x = x & rsNewSchema!ConstraintName & vbTab
'        x = x & rsNewSchema!FieldType & vbTab
'        x = x & rsNewSchema!MaxLength & vbTab
'        x = x & rsNewSchema!Precision & vbTab
'        x = x & rsNewSchema!Scale & vbTab
'        Print #io, x
'        rsNewSchema.MoveNext
'    Loop
'    Close #io

End Sub

Private Sub CreateTables(ByRef cnFrom As ADODB.Connection, ByRef cnTo As ADODB.Connection)
    
    ' clear
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnTo.OpenSchema(adSchemaTables)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" And frs!Table_Name <> "Paste" And Left(frs!Table_Name, 1) <> "~" Then
            cnTo.Execute "drop table " & frs!Table_Name
        End If
        frs.MoveNext
    Loop
    frs.Close
    
    ' add
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnFrom.OpenSchema(adSchemaPrimaryKeys)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" Then
            cnTo.Execute "create table " & frs!Table_Name
        End If
        frs.MoveNext
    Loop
    frs.Close
    
End Sub

Private Sub CreateFields(ByRef cn As ADODB.Connection)
    Dim FString As String
    Dim LastTblName As String
    Dim eFlag As Boolean
    rsNewSchema.Sort = "TableName ASC, FieldNum ASC"
    rsNewSchema.MoveFirst
    Dim LastTable As String
    LastTable = ""
    Do While Not rsNewSchema.EOF
        If LastTable = "" Or LastTable <> rsNewSchema!TableName Then
            FString = "ALTER TABLE " & rsNewSchema!TableName & _
                      " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                      " COUNTER PRIMARY KEY"
        Else
            eFlag = False
            If rsNewSchema!TableName = "PREmployee" And rsNewSchema!FieldName = "SSN" Then eFlag = True
            If rsNewSchema!TableName = "Payee99" And rsNewSchema!FieldName = "FederalID" Then eFlag = True
            ' If rsNewSchema!TableName = "Detail99" And rsNewSchema!FieldName = "PayeeID" Then eFlag = True
            If eFlag Then
                FString = "ALTER TABLE " & rsNewSchema!TableName & _
                          " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                          " String"
            Else
                FString = "ALTER TABLE " & rsNewSchema!TableName & _
                          " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                          " " & rsNewSchema!FieldType2
            End If
        End If
        LastTable = rsNewSchema!TableName
        cn.Execute FString
        rsNewSchema.MoveNext
    Loop
End Sub


Private Sub GetTables()
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Tables.txt" For Output As #io
    
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    ' Set frs = cnOld.OpenSchema(adSchemaColumns)
    ' Set frs = cnOld.OpenSchema(adSchemaTables)
    Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    
    ' Table names only
    Dim dc As ADODB.Field
    For Each dc In frs.Fields
        Print #io, (dc.Name)
    Next
    Exit Sub

    
    ' tables PK info
    Do Until frs.EOF = True
        x = frs!Table_Name
        x = frs!Table_Name & vbTab & frs!Column_Name & vbTab & frs!PK_Name
        If Left(x, 4) <> "MSys" Then
            Print #io, x
        End If
        frs.MoveNext
    Loop
    Close #io



End Sub

Private Function SQLConnect(ByVal dbName As String) As ADODB.Connection
        
    Dim pwd As String
    Set SQLConnect = New ADODB.Connection
    If Right(LCase(dbName), 6) = ".accdb" Then
        SQLConnect.Provider = "Microsoft.ACE.OLEDB.12.0"
    Else
        SQLConnect.Provider = "Microsoft.Jet.OLEDB.4.0"
    End If
    SQLConnect.Mode = adModeReadWrite
    SQLConnect.ConnectionString = dbName
    On Error Resume Next
    SQLConnect.Open
    If Err.Number <> 0 Then
        If Err.Description = "Not a valid password." Then
            If InStr(1, LCase(dbName), "balint") Then
                pwd = "OLDBB35"
            Else
                pwd = InputBox("Please enter the password for: " & dbName)
            End If
            SQLConnect.Properties("Jet OLEDB:Database Password") = pwd
            On Error GoTo 0
            SQLConnect.Open
        Else
            x = "Error opening: " & dbName & vbCr & Err.Description
            MsgBox x, vbExclamation, "Data Conversion"
            End
        End If
    End If
    
    On Error GoTo 0
    
End Function

Private Sub Copy2(ByVal Folder1 As String, ByVal Folder2 As String, ByVal FileSpec As String)
    Dim fnm As String
    fnm = Dir$(Folder1 & "\" & FileSpec)
    While fnm <> ""
        FileCopy Folder1 & "\" & fnm, Folder2 & "\" & fnm
        fnm = Dir$
    Wend
End Sub
