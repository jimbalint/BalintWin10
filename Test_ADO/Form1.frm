VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnSys As adodb.Connection
Dim cnNew As adodb.Connection
Dim cnOld As adodb.Connection

Dim strSQL As String

Dim rs As New adodb.Recordset
Dim rsNew As New adodb.Recordset
Dim frs As New adodb.Recordset
Dim log As Integer

Dim x, y, z As String
Dim i, j, k As Long

Dim rsNewSchema As adodb.Recordset
Dim frm As New frmProgress

Dim rc4Key As String
Dim BalintFolder As String
Dim NewFolder As String
Dim dbBlank As String

' todo
' byte/boolean ...
' add by using ADO record set

Private Sub Form_Load()
' https://stackoverflow.com/questions/9408245/is-it-possible-to-use-vba-to-change-the-current-accdb-e-database-password
Dim strAlterPassword As String
' new / old
strAlterPassword = "ALTER DATABASE PASSWORD [abc123] [abc1234];"

Dim ADO_Cnnct
Set ADO_Cnnct = New adodb.Connection
With ADO_Cnnct
    .Mode = adModeShareExclusive

    .Provider = "Microsoft.ACE.OLEDB.12.0"
    '  Use old password to establish connection
    .Properties("Jet OLEDB:Database Password") = "abc1234"

    'name  current DB

    ' DBPath = [CurrentProject].[FullName]  <- this does not work: get a file already in use error

    .Open "Data Source= " & "c:\Balint\Data\NewDB3.accdb" & ";"
    ' Execute the SQL statement to change the password.
    .Execute (strAlterPassword)
End With

'Clean up objects.
ADO_Cnnct.Close
Set ADO_Cnnct = Nothing

End


'Dim cn As New ADODB.Connection
'Set cn = SQLConnect("c:\Balint\Data\NewDB3.accdb")
'cn.Execute "alter database password null abc123"
'MsgBox ("OK")
'End




    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fnm As String
    Dim nfnm As String
    Dim cnm As String
    Dim rsC As New adodb.Recordset
    
    rc4Key = "B@lint19742101!@#$%^&*"
    
    BalintFolder = "\\vboxsrv\vm-share\Balint"
    NewFolder = BalintFolder & "\Data_New"
    
    dbBlank = BalintFolder & "\Blank\BlankAccdb.accdb"
    If Len(Dir(NewFolder, vbDirectory)) > 0 Then
        MsgBox "Folder already exists: " & NewFolder, vbExclamation, "Data Conversion"
        End
        ' On Error Resume Next
        Kill NewFolder & "\*.*"
        ' On Error GoTo 0
    Else
        MkDir NewFolder
    End If

    log = FreeFile
    Open NewFolder & "\ConvertLog.txt" For Output As #log
    
    Set cnSys = SQLConnect(BalintFolder & "\Data\GLSystem.mdb")
    
    x = BalintFolder & "\Data_New\GLSystem.accdb"
    dbBlank = BalintFolder & "\Blank\BlankAccdb.accdb"
    If Len(Dir(dbBlank, vbNormal)) = 0 Then
        MsgBox "Blank DB not found: " & y, vbExclamation, "Data Conversion"
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
        frm.lblMsg1 = "Now converting: " & cnm
        frm.lblMsg2 = ""
        frm.lblMsg3 = ""
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
            
            strSQL = "ALTER TABLE PRDepartment" & _
                    " ADD CONSTRAINT [dptNumberKey] UNIQUE ([DepartmentNumber])"
            cnNew.Execute strSQL
            strSQL = "ALTER TABLE PREmployee" & _
                    " ADD CONSTRAINT [empNumberKey] UNIQUE ([EmployeeNumber])"
            cnNew.Execute strSQL
            
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
    Close #log
    cnSys.Close
    cnNew.Close
    On Error GoTo 0
    
    MsgBox ("OK..")
    frm.Hide
    End
    ' ==================
    
    GetSchema
    GetTables
    GetConstraints
    
    ' CopySchema
    
    x = RC4Encrypt("Blaze3215", rc4Key)
    MsgBox (x)
    x = RC4Decrypt(x, rc4Key)
    MsgBox (x)
    End
    
    
    End
    
    
End Sub

Private Sub CopyData(ByVal cnFrom As adodb.Connection, ByVal cnTo As adodb.Connection)
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


Private Sub CopyDataProcess(ByVal TblName As String, ByVal cnFrom As adodb.Connection, ByVal cnTo As adodb.Connection)

    Dim ct1, ct2 As Long
    Dim eFlag As Boolean
    
    Dim io As Integer
    io = FreeFile
    ' Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_SQL.txt" For Output As #io
    
    strSQL = "delete * from " & TblName
    cnTo.Execute strSQL
    
    Dim fld As adodb.Field
    strSQL = "select * from " & TblName
    
    Set rs = New adodb.Recordset
    rs.Source = strSQL
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseServer
    Set rs.ActiveConnection = cnFrom
    rs.Open
    ct1 = rs.RecordCount
    
    Set rsNew = New adodb.Recordset
    rsNew.Source = strSQL
    rsNew.LockType = adLockOptimistic
    rsNew.CursorType = adOpenKeyset
    rsNew.CursorLocation = adUseServer
    Set rsNew.ActiveConnection = cnTo
    rsNew.Open
    
    ct2 = 0
    Do While Not rs.EOF
        rsNew.AddNew
        For Each fld In rs.Fields
            x = fld.Name & vbTab & rs.Fields(fld.Name)
            eFlag = False
            If TblName = "PREmployee" And fld.Name = "SSN" Then eFlag = True
            ' If TblName = "Detail99" And fld.Name = "PayeeID" Then eFlag = True
            If TblName = "Payee99" And fld.Name = "FederalID" Then eFlag = True
            If eFlag Then
                y = RC4Encrypt(rs.Fields(fld.Name), rc4Key)
            Else
                y = rs.Fields(fld.Name)
            End If
            Select Case y
                Case "True": rsNew.Fields(fld.Name) = 1
                Case "False": rsNew.Fields(fld.Name) = 0
                Case Else: rsNew.Fields(fld.Name) = y
            End Select
        Next fld
        rsNew.Update
        ct2 = ct2 + 1
        If ct2 Mod 100 = 1 Then
            frm.lblMsg3 = TblName & " " & ct2 & " of: " & ct1
            frm.Refresh
        End If
        rs.MoveNext
    Loop
    rs.Close
    rsNew.Close
    frm.lblMsg3 = TblName & " " & ct2 & " of: " & ct1
    frm.Refresh
    Print #log, "--- " & TblName & " Records Converted: " & ct2
    
End Sub


Private Sub CopyData2()
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_SQL.txt" For Output As #io
    
    Dim fld As adodb.Field
    strSQL = "select * from Users"
    Set rs = New adodb.Recordset
    rs.Source = strSQL
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseServer
    Set rs.ActiveConnection = cnOld
    rs.Open
    Do While Not rs.EOF
        strSQL = "insert into Users ("
        For Each fld In rs.Fields
            strSQL = strSQL & fld.Name & ", "
        Next fld
        strSQL = Left(strSQL, Len(strSQL) - 2)
        strSQL = strSQL & ") values ("
        For Each fld In rs.Fields
            strSQL = strSQL & SQLFormat("", fld.Name, rs.Fields(fld.Name)) & ", "
        Next fld
        strSQL = Left(strSQL, Len(strSQL) - 2)
        strSQL = strSQL & ")"
    Print #io, strSQL
    ' MsgBox (strSQL)
        cnNew.Execute strSQL
        rs.MoveNext
    Loop
    Close #io
    
End Sub

Function SQLFormat(ByVal TblName As String, ByVal FldName As String, ByVal FldVal As String) As String

    SQLFormat = FldVal
    If FldName = "LoadLastCompany" Then
        SQLFormat = "1"
        Exit Function
    End If
    If FldName = "LastCompany" Or FldName = "LoadLastCompany" Or FldName = "LastPRCompany" Or FldName = "ID" Then
        Exit Function
    End If
    SQLFormat = "'" & FldVal & "'"

End Function

Private Sub InitSchemaRS()

    Set rsNewSchema = New adodb.Recordset
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

Private Sub PopSchemaRS(ByVal cn As adodb.Connection)
    
    ' fields
    Dim FldNum As Integer
    FldNum = 0
    Set frs = New adodb.Recordset
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
            i = frs!Data_Type
            Select Case i
                Case 2: x = "Short"
                Case 3: x = "Long"
                Case 4: x = "Short"
                Case 5: x = "Double"
                Case 6: x = "Currency"
                Case 7: x = "DateTime"
                Case 11: x = "Logical"
                Case 17: x = "Byte"
                Case 130: x = "LongText"
                    If rsNewSchema!MaxLength = 255 Then
                        x = "LongText"
                    Else
                        x = "Char(" & rsNewSchema!MaxLength & ")"
                    End If
                Case Else
                    MsgBox "Data Type NF: " & i
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

Private Sub CreateTables(ByVal cnFrom As adodb.Connection, ByVal cnTo As adodb.Connection)
    
    ' clear
    Set frs = New adodb.Recordset
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
    Set frs = New adodb.Recordset
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

Private Sub CreateFields(ByVal cn As adodb.Connection)
    Dim fString As String
    Dim LastTblName As String
    Dim eFlag As Boolean
    rsNewSchema.Sort = "TableName ASC, FieldNum ASC"
    rsNewSchema.MoveFirst
    Dim LastTable As String
    LastTable = ""
    Do While Not rsNewSchema.EOF
        If LastTable = "" Or LastTable <> rsNewSchema!TableName Then
            fString = "ALTER TABLE " & rsNewSchema!TableName & _
                      " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                      " COUNTER PRIMARY KEY"
        Else
            eFlag = False
            If rsNewSchema!TableName = "PREmployee" And rsNewSchema!FieldName = "SSN" Then eFlag = True
            If rsNewSchema!TableName = "Payee99" And rsNewSchema!FieldName = "FederalID" Then eFlag = True
            ' If rsNewSchema!TableName = "Detail99" And rsNewSchema!FieldName = "PayeeID" Then eFlag = True
            If eFlag Then
                fString = "ALTER TABLE " & rsNewSchema!TableName & _
                          " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                          " String"
            Else
                fString = "ALTER TABLE " & rsNewSchema!TableName & _
                          " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                          " " & rsNewSchema!FieldType2
            End If
        End If
        LastTable = rsNewSchema!TableName
        cn.Execute fString
        rsNewSchema.MoveNext
    Loop
End Sub

Private Sub CopySchema()
    ' http://www.devx.com/vb2themax/Tip/19114
    
    Set frs = New adodb.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaColumns)
    Set frs = cnOld.OpenSchema(adSchemaTables)
    
    Do Until frs.EOF = True
        x = frs!Table_Name
        If Left(x, 4) <> "MSys" Then
            MsgBox (x)
        End If
        frs.MoveNext
    Loop

End Sub

Private Sub GetTables()
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Tables.txt" For Output As #io
    
    Set frs = New adodb.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    ' Set frs = cnOld.OpenSchema(adSchemaColumns)
    ' Set frs = cnOld.OpenSchema(adSchemaTables)
    Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    
    ' Table names only
    Dim dc As adodb.Field
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

Private Sub GetConstraints()
    ' https://docs.microsoft.com/en-us/sql/relational-databases/system-information-schema-views/columns-transact-sql?view=sql-server-ver15
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Constraints.txt" For Output As #io
    
    Set frs = New adodb.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaConstraintColumnUsage)
    ' Set frs = cnOld.OpenSchema(adSchemaTableConstraints)
    ' Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    
'    Dim fld As ADODB.Field
'    For Each fld In frs.Fields
'        x = fld.Name
'        Print #io, x
'    Next fld
'    Close #io
'    Exit Sub
    

    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" Then
             x = frs!Table_Name & vbTab & _
                 frs!Column_Name & vbTab & _
                 frs!Constraint_name
            
             Print #io, x
        End If
        frs.MoveNext
    Loop
    Close #io

End Sub


Private Sub GetSchema()
    ' https://docs.microsoft.com/en-us/sql/relational-databases/system-information-schema-views/columns-transact-sql?view=sql-server-ver15
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Schema.txt" For Output As #io
    
    Set frs = New adodb.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaColumns)
    ' Set frs = cnOld.OpenSchema(adSchemaTableConstraints)
    ' Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    ' frs.Sort = "ORDINAL_POSITION asc"
    Dim fld As adodb.Field
    For Each fld In frs.Fields
        Print #io, fld.Name
    Next fld
    Print #io, "====================="
    Do Until frs.EOF = True
        x = frs!Table_Name & vbTab & _
            frs!Column_Name & vbTab & _
            frs!ORDINAL_POSITION & vbTab & _
            frs!Data_Type & vbTab & _
            frs!is_nullable & vbTab & _
            frs!CHARACTER_MAXIMUM_LENGTH & vbTab & _
            frs!NUMERIC_PRECISION & vbTab & _
            frs!NUMERIC_SCALE & vbTab & _
            "|" & vbTab & _
            frs!ORDINAL_POSITION
            
            
            ' frs!NUMERIC_PRECISION_RADIX
        
        Print #io, x
        frs.MoveNext
    Loop
    Close #io

End Sub

Public Function AddField(ByVal TableName As String, _
                         ByVal ColumnName As String, _
                         ByVal ColumnType As String, _
                         ByRef adoConn As adodb.Connection) _
                         As Byte
                         
Dim cm As adodb.Command
Dim frs As adodb.Recordset
Dim FldFlag As Boolean
Dim fString As String
Dim TblExists As Boolean
                         
    ' see if the field is already in the Table
    Set frs = New adodb.Recordset
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
        
        fString = "ALTER TABLE " & TableName & _
                  " ADD COLUMN [" & ColumnName & "]" & _
                  " " & ColumnType
        adoConn.Execute fString
        
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

Private Function SQLConnect(ByVal dbName As String) As adodb.Connection
        
    Dim pwd As String
    Set SQLConnect = New adodb.Connection
    If InStr(1, ".mdb", dbName) Then
        SQLConnect.Provider = "Microsoft.Jet.OLEDB.4.0"
    Else
        SQLConnect.Provider = "Microsoft.ACE.OLEDB.12.0"
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

Private Sub Test1()

    rs.CursorLocation = adUseClient
    rs.Open "select * from Users", cnOld, adOpenDynamic, adLockBatchOptimistic
    
    MsgBox (rs.RecordCount)
    Set rs.ActiveConnection = Nothing
    MsgBox (rs.RecordCount)
    
    rs.ActiveConnection = cnNew
    rs.UpdateBatch
    rs.Close
    
    rs.Open "select * from Users", cnNew, adOpenDynamic, adLockBatchOptimistic
    MsgBox (rs.RecordCount)
    rs.Close
    
End Sub

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

