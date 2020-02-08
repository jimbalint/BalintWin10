Attribute VB_Name = "modNewField"
Option Explicit
Public Sub UpdateCheck(ByVal GLSys As Boolean, _
                       ByRef adoConn As ADODB.Connection)

'Dim urs As ADODB.Recordset
'Dim uCount As Long
'Dim ucmd As ADODB.Command
'
'    ' GLSys = false - check for actual company data file
'    ' GLSys = true - check for GLSystem File
'
'    ' 10/23/2009 FiscalYear to GLPrint (for old installations)
'    If GLSys = False Then
'        If AddField("GLPrint", "FiscalYear", "Long", adoConn) Then
'        End If
'    End If
'
'    ' 10/22/2009 Wkc Comp to PRHist
'    If GLSys = False Then
'        If AddField("PRHist", "WkcAmount", "Currency", adoConn) Then
'        End If
'    End If
'
'    ' 10/20/09 - StateUnempID
'    If GLSys = True Then
'        If AddField("PRCompany", "StateUnempID", "String", adoConn) Then
'        End If
'    End If
'
'    ' 09/15/09 - flag for Courtesy CWT add
'    If GLSys = False Then
'        If AddField("PREmployee", "CourtesyAdd", "Byte", adoConn) Then
'        End If
'    End If
'
'    ' company phn #
'    If GLSys = True Then
'        If AddField("PRCompany", "PhoneNumber", "String", adoConn) Then
'        End If
'    End If
'
'    ' 09/12/09 - Var fields in PRGlobal
'    If GLSys = True Then
'        If AddField("PRGlobal", "UserID", "Long", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var1", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var2", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var3", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var4", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var5", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var6", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var7", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var8", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var9", "String", adoConn) Then
'        End If
'        If AddField("PRGlobal", "Var10", "String", adoConn) Then
'        End If
'    End If
'
'    ' 08/13/09 - Courtesy CWT
'    If GLSys = False Then
'        If AddField("PREmployee", "CourtesyCityID", "Long", adoConn) Then
'        End If
'        If AddField("PRDist", "CourtesyCityID", "Long", adoConn) Then
'        End If
'        If AddField("PRDist", "CourtesyCityTax", "Currency", adoConn) Then
'        End If
'        If AddField("PRDist", "ManualCourtesyCityTax", "Byte", adoConn) Then
'        End If
'    End If
'
'    ' 08/03/09 - wage base to PRHist
'    If GLSys = False Then
'        If AddField("PRHist", "SSWageBase", "Currency", adoConn) Then
'        End If
'        If AddField("PRHist", "FUNWageBase", "Currency", adoConn) Then
'        End If
'        If AddField("PRHist", "SUNWageBase", "Currency", adoConn) Then
'        End If
'    End If
'
'    ' 07/25/09 - add Unemployment max to PRState
'    If GLSys = True Then
'        If AddField("PRState", "UnempMax", "Currency", adoConn) Then
'        End If
'    End If
'
'    ' 7/25/09 - add StateID to PRHist
'    If GLSys = False Then
'        If AddField("PRHist", "StateID", "Long", adoConn) Then
'        End If
'        If AddField("PRHist", "SUNWage", "Currency", adoConn) Then
'        End If
'    End If
'
'    ' 01/22/08 - add LastPRCompany to GLCompany
'    If GLSys = True Then
'        If AddField("Users", "LastPRCompany", "Long", adoConn) Then
'        End If
'    End If
'
'    ' 12/13/06 - add date/time posted field to GLHistory
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
'
'    Unload frmProgress

End Sub

Public Function AddField(ByVal TableName As String, _
                         ByVal ColumnName As String, _
                         ByVal ColumnType As String, _
                         ByRef adoConn As ADODB.Connection) _
                         As Byte
                         
'Dim cm As ADODB.Command
'Dim frs As ADODB.Recordset
'Dim FldFlag As Boolean
'Dim fString As String
'Dim TblExists As Boolean
'
'    ' see if the field is already in the Table
'    Set frs = New ADODB.Recordset
'    frs.CursorLocation = adUseClient
'    frs.CursorType = adOpenStatic
'    frs.LockType = adLockBatchOptimistic
'    Set frs = adoConn.OpenSchema(adSchemaColumns)
'
'    FldFlag = False
'    TblExists = False
'    Do Until frs.EOF = True
'
'       If UCase(frs!Table_Name) = UCase(TableName) Then
'          TblExists = True
'       End If
'
'       If UCase(frs!Table_Name) = UCase(TableName) And UCase(frs!Column_Name) = UCase(ColumnName) Then
'         FldFlag = True
'         Exit Do
'       End If
'
'       frs.MoveNext
'
'   Loop
'
'   ' the table was not found
'   If TblExists = False Then Exit Function
'
'   ' the field already exists - no need to add it
'   If FldFlag = True Then
'      AddField = 0
'      Exit Function
'   End If
'
'   frs.Close
'   Set frs = Nothing
'
'   ' add it
'   On Error Resume Next
'   fString = "ALTER TABLE " & TableName & _
'             " ADD COLUMN " & ColumnName & _
'             " " & ColumnType
'   adoConn.Execute fString
'
'   If Err.Number <> 0 Then
'      GoTo AddFieldError
'   End If
'
'   AddField = 1
'   Exit Function
'
'AddFieldError:
'
'   ' problem with field add !?
'   MsgBox TableName & "/" & ColumnName & " " & ColumnType & _
'       vbCrLf & vbCrLf & "Field Add Error" & Err.Description, _
'       vbOKOnly + vbCritical
'   AddField = 0
'   End

End Function

