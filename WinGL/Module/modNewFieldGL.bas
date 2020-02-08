Attribute VB_Name = "modNewFieldGL"
Option Explicit

Dim Lvl As Integer
Dim FWTRange(9), FWTAmount(9), FWTPct(9) As Currency
Dim SnglMarr As Byte
Dim MsgResponse As Variant

Public Sub UpdateCheck(ByVal GLSys As Boolean, _
                       ByRef adoConn As ADODB.Connection)

Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim urs As ADODB.Recordset
Dim uCount As Long
Dim ucmd As ADODB.Command

    ' 06/01/2010
    If GLSys = True Then
        If TableExists("GLFFColumn", adoConn) = False Then
            FFColumnCreate
        End If
    End If

    ' 10/23/2009 FiscalYear to GLPrint (for old installations)
    If GLSys = False Then
        If AddField("GLPrint", "FiscalYear", "Long", adoConn) Then
        End If
    End If

    ' 01/22/08 - add LastPRCompany to GLCompany
    If GLSys = True Then
        If AddField("Users", "LastPRCompany", "Long", adoConn) Then
        End If
    End If
    
    ' 12/13/06 - add date/time posted field to GLHistory
'    If GLSys = False Then
'        If AddField("GLHistory", "PostDate", "DateTime", adoConn) <> 0 Then
'
'           ' add a key also
'           cn.Execute "CREATE INDEX PostKey ON GLHistory ([PostDate])"
'
'           ' sweep in initial values
'           rsInit "SELECT * FROM GLHistory ORDER BY ID", cn, urs
'
'           ' display screen
'           frmProgress.Show
'           frmProgress.lblMsg1 = "Adding PostDate Field to GLHistory"
'
'           Do Until urs.EOF
'
'              urs.Fields("PostDate") = DateSerial(Year(Now()), Month(Now()), Day(Now())) + _
'                                       TimeSerial(0, 0, urs!ID)
'              urs.Update
'
'              uCount = uCount + 1
'              If uCount Mod 100 = 1 Then
'                 frmProgress.lblMsg2 = "On Record: " & Format(uCount, "###,###,##0")
'                 frmProgress.Refresh
'              End If
'
'              urs.MoveNext
'           Loop
'
'        End If
'
'    End If

'    Unload frmProgress

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
                  " ADD COLUMN " & ColumnName & _
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

Public Sub FieldSweep()

Dim DfltStateID As Long

End Sub



