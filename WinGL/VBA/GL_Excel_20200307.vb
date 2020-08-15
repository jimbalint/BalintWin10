Option Explicit

' Balint - Windows GL VBA code
' 2020-03-07 Usd ADO recordset instead of XArrayDB

Dim strProvider As String

Dim cn As ADODB.Connection
Dim rsAmount As ADODB.Recordset
Dim rsAccount As ADODB.Recordset
Dim rsDesc As ADODB.Recordset
Dim rsCompany As ADODB.Recordset

Dim PassWord As String
Dim ErrMsg As String

' Dim xdbAmount As New XArrayDB
' Dim xdbAccount As New XArrayDB
' Dim xdbDesc As New XArrayDB

Dim RecCount As Long
Dim xRow As Long
Dim xRow2 As Long
Dim SearchString As String

Dim cnSys As ADODB.Connection
Dim AcctSub As String
Dim AcctBase As String

Dim FName As String
Dim Flg As Boolean

Dim xx, yy As String
Dim i, j As Long
Dim b As Byte

Dim SQLString As String

Dim bytConnected As Byte

Dim CmdBar As CommandBar
Dim CmdBarBut As CommandBarButton

Dim WkSht As Worksheet
Dim jCell As Range
Dim DriveLetter As String

Sub Connect()
   
Dim ErrMsg As String

    strProvider = "Microsoft.Jet.OLEDB.4.0"

    ' data disconnect if already connected
    If bytConnected Then Disconnect
    bytConnected = 1
   
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!
'   Application.Calculation = xlCalculationManual
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!
   
    FName = Range("GLFile")
   
    If FName = "" Then
        MsgBox "Please enter a dBase file name", vbCritical + vbOKOnly, "File Open"
        Range("GLFile").Select
        Exit Sub
    ElseIf Exists(FName) = False Then
        MsgBox FName & vbCrLf & "dBase file name NOT FOUND !!!", vbCritical + vbOKOnly, "File Open"
        Range("GLFile").Select
        Exit Sub
    End If
   
    ' store drive letter and colon
    DriveLetter = Left(FName, 2)

'   On Error GoTo ConnectErr
    ErrMsg = "GLSystem Connect"
   
    ' ====================================================================
   
    ' open the GLDesc file
    Set cnSys = New ADODB.Connection
    cnSys.Provider = strProvider
    cnSys.ConnectionString = DriveLetter & "\balint\data\GLSystem.mdb"
    cnSys.Open
   
    GetCompany
   
    ErrMsg = "GLDescription Connect"
   
    Set rsDesc = New ADODB.Recordset
    rsDesc.Source = "SELECT Number, Description FROM GLDescriptions"
    rsDesc.LockType = adLockOptimistic
    rsDesc.CursorType = adOpenKeyset
    Set rsDesc.ActiveConnection = cnSys
    rsDesc.Open
   
    ' ========================================================================
   
    ErrMsg = "Data Base Connect"
   
    On Error Resume Next
   
    Set cn = New ADODB.Connection
    cn.Provider = strProvider
    cn.ConnectionString = FName
       
    If PassWord <> "" Then
        cn.Properties("Jet OLEDB:Database Password") = PassWord
    End If
       
    cn.Open
   
    If Err.Number <> 0 Then
        If Err.Description = "Not a valid password." Then
            On Error GoTo 0
            PassWord = InputBox("Enter Database Password:", "Windows GL")
            cn.Properties("Jet OLEDB:Database Password") = PassWord
            cn.Open
        Else
            MsgBox "Error connecting to: " & FName & " " & _
                Err.Description & " " & Err.Number, _
                vbExclamation + vbOKOnly, "Windows GL"
        End If
    End If
   
    ErrMsg = "GLAmount Get"
    
    On Error GoTo 0
    
    GetChartOfAccounts
       
    ' force recalc of all cells
    For i = 1 To Sheets.Count
        Sheets(i).Select
        Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell)).Dirty
        Range("A1", ActiveCell.SpecialCells(xlCellTypeLastCell)).Calculate
    Next i
   
    Application.CalculateFull

   ' MANUAL DIRTY - force it ???
'   For Each WkSht In Application.Worksheets
'       WkSht.Select
'       Range("$A$1", ActiveCell.SpecialCells(xlCellTypeLastCell)).Select
'
'       For Each jCell In Selection
'           If jCell.Formula <> "" Then
'              yy = "YY"
'              xx = jCell.Formula
'              jCell.Formula = yy
'              jCell.Formula = xx
'           End If
'       Next jCell
'       Range("$A$1").Select
'   Next WkSht
   Application.Calculate
   
'   ' home the cursor
'   Sheets(1).Select
'   Range("A2").Select
   
   MsgBox "Connect to " & FName & " complete" & vbCrLf & vbCrLf & _
          "Press Alt-G to refresh or if the data base changes", vbInformation
   
   
   ' close connections / release record sets
'   rsDesc.Close
'   Set rsDesc = Nothing
'
'   rsCompany.Close
'   Set rsCompany = Nothing
'
'   cnSys.Close
'   Set cnSys = Nothing
   
   rsAccount.Close
   Set rsAccount = Nothing
   
   cn.Close
   Set cn = Nothing
   
   cnSys.Close
   Set cnSys = Nothing
       
   Exit Sub
   
ConnectErr:
   
   MsgBox ErrMsg & " - " & Err.Description & " " & Str$(Err), vbCritical + vbOKOnly, _
          "Init Error!"
   Range("GLFile").Select
          
End Sub

Sub Disconnect()
    
'    rsAmount.Close
'    Set rsAmount = Nothing

'    rsAccount.Close
'    Set rsAccount = Nothing

'    cn.Close
'    Set cn = Nothing

    On Error Resume Next
    
    rsDesc.Close
    Set rsDesc = Nothing

    rsCompany.Close
    Set rsCompany = Nothing

    cnSys.Close
    Set cnSys = Nothing
    
    On Error GoTo 0

End Sub

Function GetAmount(Acct As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency
    GetAmount = GetGLAmount(Acct, FY1, FY2, Pd1, Pd2, "Amount")
End Function
Function GetBudget(Acct As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte) As Currency
    GetBudget = GetGLAmount(Acct, FY1, FY2, Pd1, Pd2, "Budget")
End Function

Function GetGLAmount(Acct As Long, FY1 As Long, FY2 As Long, Pd1 As Byte, Pd2 As Byte, AmtBudg As String) As Currency
    On Error GoTo 0

    Dim x, fldname As String
    GetGLAmount = 0
    Dim ThisAcct As Long
    ' consolidated
    If Range("HiConsolidated").Value <> 0 And Acct Mod 10 ^ rsCompany!SubDigits = Range("HiConsolidated").Value Then
        ThisAcct = Acct / 10 ^ rsCompany!SubDigits
    Else
        ThisAcct = Acct
    End If
        
    SQLString = "select * from GLAmount" & _
                " where Account = " & ThisAcct & _
                " and FiscalYear between " & FY1 & " and " & FY2
    Set rsAmount = New ADODB.Recordset
    rsAmount.Source = SQLString
    rsAmount.LockType = adLockOptimistic
    rsAmount.CursorType = adOpenKeyset
    rsAmount.CursorLocation = adUseServer
    Set rsAmount.ActiveConnection = cn
    rsAmount.Open
    If rsAmount.RecordCount() = 0 Then Exit Function
    Do
        For b = Pd1 To Pd2
            fldname = AmtBudg & Format$(b, "00")
            GetGLAmount = GetGLAmount + rsAmount.Fields(fldname)
        Next b
        rsAmount.MoveNext
        If rsAmount.EOF() Then Exit Do
    Loop
    rsAmount.Close
    Set rsAmount = Nothing
    
End Function

Function GetGLDesc(ByVal DescNum As Long) As String
    rsDesc.MoveFirst
    rsDesc.Find "Number = " & DescNum
    If Not rsDesc.EOF Then
        GetGLDesc = rsDesc!Description
    Else
        GetGLDesc = "Description #" & DescNum & " not found"
    End If
End Function
Function GetDesc(ByVal AcctNum As String) As String
    rsAccount.MoveFirst
    rsAccount.Find "Account = " & AcctNum
    If Not rsAccount.EOF Then
        GetDesc = rsAccount!Description
    Else
        GetDesc = ""
    End If
End Function


Function GetDate(ByVal FormatNum As Byte, _
                 ByVal FY As Long, _
                 ByVal Pd1 As Byte, _
                 ByVal Pd2 As Byte) As String

Dim CurrYrPdEnd As String
Dim CurrYrCurrPdBeg As String
Dim CurrYrFYBeg As String
Dim CurrYrFYEnd As String
Dim PrevYrPdEnd As String
Dim PrevYrPDBeg As String
Dim PrevYrFYBeg As String
Dim ReportDate As String
Dim x As String
Dim y As String

Dim Year1 As Long
Dim Month1 As Byte
Dim Year2 As Long
Dim Month2 As Byte
      
Dim DString As String
Dim Months As String
Dim Digit(13) As String
Dim fmt As String
Dim ii As Double
      
    ' get the calendar year and month for the start/end periods
    If rsCompany!FirstPeriod = 1 Then
       
       Month1 = Pd1
       Year1 = FY
       
       Month2 = Pd2
       Year2 = FY
    
    Else
       
       x = DateSerial(FY - 1, rsCompany!FirstPeriod, 1)
       
       y = DateAdd("m", Pd1 - 1, x)
       Month1 = Month(y)
       Year1 = Year(y)
       
       y = DateAdd("m", Pd2 - 1, x)
       Month2 = Month(y)
       Year2 = Year(y)
       
    End If
    
    ' curr yr pd ending      1
    CurrYrPdEnd = DateSerial(Year2, Month2, 1)
    CurrYrPdEnd = DateAdd("m", 1, CurrYrPdEnd)
    CurrYrPdEnd = DateAdd("d", -1, CurrYrPdEnd)
   
    ' curr yr pd beg        2
    CurrYrCurrPdBeg = DateSerial(Year1, Month1, 1)
                               
    ' curr yr fy beg        3
    CurrYrFYBeg = DateSerial(Year1, rsCompany!FirstPeriod, 1)
    If rsCompany!FirstPeriod > Month1 Then
       CurrYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)
    End If
      
    ' FY end date
    CurrYrFYEnd = DateAdd("yyyy", 1, CurrYrFYBeg)
    CurrYrFYEnd = DateAdd("d", -1, CurrYrFYEnd)
      
    ' prev yr pd end        4
    PrevYrPdEnd = DateAdd("yyyy", -1, CurrYrPdEnd)
      
    ' prev yr pd beg        5
    PrevYrPDBeg = DateAdd("yyyy", -1, CurrYrCurrPdBeg)
      
    ' prev yr fy beg        6
    PrevYrFYBeg = DateAdd("yyyy", -1, CurrYrFYBeg)

    fmt = "mmmm dd, yyyy"
   
    Digit(0) = " Zero "
    Digit(1) = " One "
    Digit(2) = " Two "
    Digit(3) = " Three "
    Digit(4) = " Four "
    Digit(5) = " Five "
    Digit(6) = " Six "
    Digit(7) = " Seven "
    Digit(8) = " Eight "
    Digit(9) = " Nine "
    Digit(10) = " Ten "
    Digit(11) = " Eleven "
    Digit(12) = " Twelve "
    Digit(13) = " Thirteen "
   
   
    Select Case FormatNum
    
       Case 0
          GetDate = Format(CurrYrPdEnd, fmt)
   
       Case 1
          If Pd2 = 1 Then
             GetDate = CStr(Digit(Pd2)) & " Month Ended " & Format(CurrYrPdEnd, fmt)
          Else
             GetDate = CStr(Digit(Pd2)) & " Months Ended " & Format(CurrYrPdEnd, fmt)
          End If
   
       Case 2
          GetDate = Format(CurrYrCurrPdBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
       Case 3
          GetDate = Format(CurrYrFYBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
       Case 4
          If Pd2 = 1 Then
             GetDate = CStr(Digit(Pd2)) & " Month Ended " & Format(PrevYrPdEnd, fmt)
          Else
             GetDate = CStr(Digit(Pd2)) & " Months Ended " & Format(PrevYrPdEnd, fmt)
          End If
   
       Case 5
          GetDate = Format(PrevYrPDBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
       Case 6
          GetDate = Format(PrevYrFYBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
       Case 7
          GetDate = Format(Now, fmt)
   
       Case 8
          GetDate = CStr(Pd2 * 4) & " Weeks Ended " & Format(CurrYrPdEnd, fmt)
   
       Case 9
          GetDate = CStr(Pd2 * 4) & " Weeks Ended " & Format(PrevYrPdEnd, fmt)
   
       Case 10
          GetDate = Space(18 - Len(Format(CurrYrPdEnd, fmt))) & Format(CurrYrPdEnd, fmt)
   
       Case 11
          GetDate = Space(18 - Len(Format(CurrYrCurrPdBeg, fmt))) & Format(CurrYrCurrPdBeg, fmt)
   
       Case 12
          GetDate = Space(18 - Len(Format(CurrYrFYBeg, fmt))) & Format(CurrYrFYBeg, fmt)
   
       Case 13
          GetDate = Space(18 - Len(Format(PrevYrPdEnd, fmt))) & Format(PrevYrPdEnd, fmt)
   
       Case 14
          GetDate = Space(18 - Len(Format(PrevYrPDBeg, fmt))) & Format(PrevYrPDBeg, fmt)
   
       Case 15
          GetDate = Space(18 - Len(Format(PrevYrFYBeg, fmt))) & Format(PrevYrFYBeg, fmt)
   
       Case 16
          If Pd2 - Pd1 + 1 = 1 Then
             GetDate = CStr(Digit(Pd2 - Pd1 + 1)) & " Month"
          Else
             GetDate = CStr(Digit(Pd2 - Pd1 + 1)) & " Months"
          End If

       Case 17
          If Pd2 = 1 Then
             GetDate = CStr(Pd2) & " Month"
          Else
             GetDate = CStr(Pd2) & " Months"
          End If
   
       Case 18
          GetDate = " And " & Year(CurrYrFYEnd)
   
       Case 19
          GetDate = " And " & Year(CurrYrFYEnd) - 1
   
    End Select

End Function

Private Sub GetChartOfAccounts()

    ' ========================================================================
    ' get chart of accounts - for descriptions
    Set rsAccount = New ADODB.Recordset
    rsAccount.CursorLocation = adUseClient
    rsAccount.Fields.Append "Account", adVarChar, 100, adFldIsNullable
    rsAccount.Fields.Append "Description", adVarChar, 100, adFldIsNullable
    rsAccount.Open , , adOpenDynamic, adLockOptimistic
   
    ErrMsg = "GLAccount Get"
   
    Dim rsAcct As ADODB.Recordset
    Set rsAcct = New ADODB.Recordset
    rsAcct.Source = "SELECT Account, Description, DescNumber FROM GLAccount ORDER BY Account"
    rsAcct.LockType = adLockOptimistic
    rsAcct.CursorType = adOpenKeyset
    Set rsAcct.ActiveConnection = cn
    rsAcct.Open
   
    If rsAcct.BOF And rsAcct.EOF Then
       ErrMsg = "No GLAccount records found !"
       ' GoTo C onnectErr
       Exit Sub
    End If
      
    ' copy into local recordset w/ account descriptions
    Do Until rsAcct.EOF
       
        rsAccount.AddNew
        rsAccount!Account = CStr(rsAcct!Account)
        If rsAcct!DescNumber = 0 Or IsNull(rsAcct!DescNumber) Then
           rsAccount!Description = rsAcct!Description
        ElseIf rsAccount!DescNumber = 1 Then
           rsAccount!Description = CStr(rsCompany!Name)
        Else
           rsAccount!Description = rsAcct!Description & GetGLDesc(rsAcct!DescNumber)
        End If
        
        rsAccount.Update
        
        rsAcct.MoveNext
        
    Loop
    rsAcct.Close
    Set rsAcct = Nothing
    
    ' ========================================================================

End Sub

Private Sub GetCompany()

    ErrMsg = "GLCompany Get"
   
    ' open the company file and get the first record
    Set rsCompany = New ADODB.Recordset
    rsCompany.Source = "SELECT * FROM GLCompany"
    rsCompany.LockType = adLockOptimistic
    rsCompany.CursorType = adOpenKeyset
    Set rsCompany.ActiveConnection = cnSys
    rsCompany.Open
   
    If rsCompany.BOF And rsCompany.EOF Then
        ErrMsg = "No Company records found!"
        ' GoTo ConnectErr
        Exit Sub
    End If
    
    ' loop through the company file - stop on the first one that has the same file name
    Flg = False
    rsCompany.MoveFirst
    Do Until rsCompany.EOF
      
        ' don't include drive letter
        If Mid(StrConv(rsCompany!FileName, vbLowerCase), 3, Len(rsCompany!FileName) - 2) = _
            Mid(StrConv(FName, vbLowerCase), 3, Len(FName) - 2) Then
         
            Range("LowBranch").Value = rsCompany!LowBranch
            Range("HiBranch").Value = rsCompany!HiBranch
            Range("LowConsolidated").Value = rsCompany!LowConsolidated
            Range("HiConsolidated").Value = rsCompany!HiConsolidated
            Flg = True
         
            Exit Do
        End If
        rsCompany.MoveNext
    Loop
   
    If Not Flg Then
        MsgBox "Warning - Company record not found for: " & vbCrLf & vbCrLf & _
             FName, vbInformation + vbOKOnly
    End If

End Sub

Private Function Exists(ByVal FName As String) As Boolean
   
   On Error GoTo NotFound
   GetAttr (FName)
   Exists = True
   On Error GoTo 0
   Exit Function
   
NotFound:
   On Error GoTo 0
   Exists = False

End Function



