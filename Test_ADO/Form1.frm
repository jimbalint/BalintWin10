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
Dim frs As New ADODB.Recordset

Dim x, y, z As String
Dim i, j, k As Long

Private Sub Form_Load()
    
    SQLConnect
    GetSchema
    ' CopySchema
    ' GetTables
    
    cnOld.Close
    cnNew.Close
    MsgBox ("OK..")
    End
    
    
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
    
'    Table names only
'    Dim dc As ADODB.Field
'    For Each dc In frs.Fields
'        MsgBox (dc.Name)
'    Next
'    Exit Sub
'
    
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
    

    Do Until frs.EOF = True
        x = frs!Table_Name & vbTab & _
            frs!Column_Name & vbTab & _
            frs!Data_Type & vbTab & _
            frs!is_nullable & vbTab & _
            frs!CHARACTER_MAXIMUM_LENGTH & vbTab & _
            frs!NUMERIC_PRECISION & vbTab & _
            frs!NUMERIC_SCALE
            
            
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



Private Sub SQLConnect()

    Set cnOld = New ADODB.Connection
    cnOld.Provider = "Microsoft.Jet.OLEDB.4.0"
    ' cnOld.ConnectionString = "\\vboxsrv\vm-share\balint\Test_ADO\glSystem.mdb"
    cnOld.ConnectionString = "\\vboxsrv\vm-share\balint\Test_ADO\A CRANO EXCAVATING INC.mdb"
    cnOld.Open
    
    Set cnNew = New ADODB.Connection
    cnNew.Provider = "Microsoft.ACE.OLEDB.12.0"
    cnNew.ConnectionString = "\\vboxsrv\vm-share\balint\Test_ADO\Database81.accdb"
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
