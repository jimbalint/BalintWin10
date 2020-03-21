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

Dim cnNew As ADODB.Connection
Dim cnOld As ADODB.Connection
Dim strSQL As String

Dim rs As New ADODB.Recordset
Dim rsNew As New ADODB.Recordset
Dim frs As New ADODB.Recordset

Dim x, y, z As String
Dim i, j, k As Long

Dim rsNewSchema As ADODB.Recordset

' todo
' byte/boolean ...
' add by using ADO record set

Private Sub Form_Load()
    
    
    x = "c:\balint\data\glsystem.accdb"
    y = Len(Dir(x, vbNormal))
    MsgBox (y)
    End
    
    
    
    
    
    Dim dbName1, dbName2 As String
    Dim op As Integer
    op = 0
    If op = 0 Then
        dbName1 = "\\vboxsrv\vm-share\balint\Test_ADO\glSystem.mdb"
    Else
        dbName1 = "\\vboxsrv\vm-share\balint\Test_ADO\A CRANO EXCAVATING INC.mdb"
    End If
    dbName2 = "\\vboxsrv\vm-share\balint\Test_ADO\Database81.accdb"
    ' ==================
    SQLConnect dbName1, dbName2
    ' ==================
    
    InitSchemaRS
    PopSchemaRS
    CreateTables
    CreateFields

    Set frs = cnOld.OpenSchema(adSchemaTables)
    Do Until frs.EOF = True
        x = frs!Table_Name
        If Left(x, 4) <> "MSys" And Left(x, 5) <> "Paste" Then
            CopyData (x)
        End If
        frs.MoveNext
    Loop
    
    If op = 0 Then
        strSQL = "ALTER TABLE glDescriptions" & _
                " ADD CONSTRAINT [Number] UNIQUE ([Number])"
        cnNew.Execute strSQL
    Else
        strSQL = "ALTER TABLE PRDepartment" & _
                " ADD CONSTRAINT [dptNumberKey] UNIQUE ([DepartmentNumber])"
        cnNew.Execute strSQL
    
        strSQL = "ALTER TABLE PREmployee" & _
                " ADD CONSTRAINT [empNumberKey] UNIQUE ([EmployeeNumber])"
        cnNew.Execute strSQL
    End If
   
    ' ==================
    cnOld.Close
    cnNew.Close
    MsgBox ("OK..")
    End
    ' ==================
    
    GetSchema
    GetTables
    GetConstraints
    
    ' CopySchema
    
    End
    
    
End Sub

Private Sub CopyData(ByVal TblName As String)

    Dim io As Integer
    io = FreeFile
    ' Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_SQL.txt" For Output As #io
    
    strSQL = "delete * from " & TblName
    cnNew.Execute strSQL
    
    Dim fld As ADODB.Field
    strSQL = "select * from " & TblName
    
    Set rs = New ADODB.Recordset
    rs.Source = strSQL
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseServer
    Set rs.ActiveConnection = cnOld
    rs.Open
    
    Set rsNew = New ADODB.Recordset
    rsNew.Source = strSQL
    rsNew.LockType = adLockOptimistic
    rsNew.CursorType = adOpenKeyset
    rsNew.CursorLocation = adUseServer
    Set rsNew.ActiveConnection = cnNew
    rsNew.Open
    
    Do While Not rs.EOF
        rsNew.AddNew
        For Each fld In rs.Fields
            x = fld.Name & vbTab & rs.Fields(fld.Name)
            ' MsgBox (x)
            y = rs.Fields(fld.Name)
            Select Case y
                Case "True": rsNew.Fields(fld.Name) = 1
                Case "False": rsNew.Fields(fld.Name) = 0
                Case Else: rsNew.Fields(fld.Name) = rs.Fields(fld.Name)
            End Select
        Next fld
        rsNew.Update
        rs.MoveNext
    Loop
    rs.Close
    rsNew.Close
    
End Sub


Private Sub CopyData2()
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_SQL.txt" For Output As #io
    
    Dim fld As ADODB.Field
    strSQL = "select * from Users"
    Set rs = New ADODB.Recordset
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

Private Sub PopSchemaRS()

    ' constraints
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaConstraintColumnUsage)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" Then
'            rsNewSchema.AddNew
'            rsNewSchema!TableName = frs!Table_Name
'            rsNewSchema!FieldName = frs!Column_Name
'            rsNewSchema!FieldNum = 0
'            rsNewSchema!ConstraintName = frs!Constraint_name
'            rsNewSchema!FieldType = ""
'            rsNewSchema!FieldType2 = "Long"
'            rsNewSchema!MaxLength = ""
'            rsNewSchema!Precision = ""
'            rsNewSchema!Scale = ""
'            rsNewSchema.Update
        End If
        frs.MoveNext
    Loop
    
    ' fields
    Dim FldNum As Integer
    FldNum = 0
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaColumns)
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
                Case 11: x = "Byte"
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
    Dim io As Integer
    Dim fld As ADODB.Field
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Schema.txt" For Output As #io
    rsNewSchema.Sort = "TableName ASC, FieldNum ASC"
    rsNewSchema.MoveFirst
    Do While Not rsNewSchema.EOF
        x = rsNewSchema!TableName & vbTab
        x = x & rsNewSchema!FieldName & vbTab
        x = x & rsNewSchema!FieldNum & vbTab
        x = x & rsNewSchema!ConstraintName & vbTab
        x = x & rsNewSchema!FieldType & vbTab
        x = x & rsNewSchema!MaxLength & vbTab
        x = x & rsNewSchema!Precision & vbTab
        x = x & rsNewSchema!Scale & vbTab
        Print #io, x
        rsNewSchema.MoveNext
    Loop
    Close #io

End Sub

Private Sub CreateTables()
    
    ' clear
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnNew.OpenSchema(adSchemaTables)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" And frs!Table_Name <> "Paste" And Left(frs!Table_Name, 1) <> "~" Then
            cnNew.Execute "drop table " & frs!Table_Name
        End If
        frs.MoveNext
    Loop
    frs.Close
    
    ' add
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    Do Until frs.EOF = True
        If Left(frs!Table_Name, 4) <> "MSys" Then
            cnNew.Execute "create table " & frs!Table_Name
        End If
        frs.MoveNext
    Loop
    frs.Close
    
End Sub

Private Sub CreateFields()
    Dim fString As String
    Dim LastTblName As String
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
            fString = "ALTER TABLE " & rsNewSchema!TableName & _
                      " ADD COLUMN [" & rsNewSchema!FieldName & "]" & _
                      " " & rsNewSchema!FieldType2
        End If
        LastTable = rsNewSchema!TableName
        cnNew.Execute fString
        rsNewSchema.MoveNext
    Loop
End Sub

Private Sub CopySchema()
    ' http://www.devx.com/vb2themax/Tip/19114
    
    Set frs = New ADODB.Recordset
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

Private Sub GetConstraints()
    ' https://docs.microsoft.com/en-us/sql/relational-databases/system-information-schema-views/columns-transact-sql?view=sql-server-ver15
    
    Dim io As Integer
    io = FreeFile
    Open "\\vboxsrv\vm-share\balint\Test_ADO\Balint_Constraints.txt" For Output As #io
    
    Set frs = New ADODB.Recordset
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
    
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = cnOld.OpenSchema(adSchemaColumns)
    ' Set frs = cnOld.OpenSchema(adSchemaTableConstraints)
    ' Set frs = cnOld.OpenSchema(adSchemaPrimaryKeys)
    ' frs.Sort = "ORDINAL_POSITION asc"
    Dim fld As ADODB.Field
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



Private Sub SQLConnect(ByVal dbName1 As String, ByVal dbName2 As String)

    Set cnOld = New ADODB.Connection
    cnOld.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnOld.ConnectionString = dbName1
    cnOld.Open
    
    Set cnNew = New ADODB.Connection
    cnNew.Provider = "Microsoft.ACE.OLEDB.12.0"
    cnNew.ConnectionString = dbName2
    cnNew.Open

End Sub

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

