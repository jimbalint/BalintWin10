Attribute VB_Name = "modFreeFormat"
Option Explicit

Dim SchedGlobID, ColGlobID As Long
Dim flg As Boolean
Dim rsFY As New ADODB.Recordset
Dim I, J, K As Long
Dim X, Y, Z As String
Dim Val1, Val2 As Currency
Dim boo As Boolean
Dim ct As Long

    ' jb 10/28/2010 (20) changed to (99)

Dim ColVal(20) As Currency

Dim GLPStartPd, GLPEndPd As Byte
Dim PrintCols As Byte
Dim PrintString As String
Dim PrintTab As Byte
Dim GLDate(20) As String
Dim FullDesc As String

Dim Total(5, 20) As Currency    ' *** changed from 5, 20
Dim ConsAmt(20) As Currency
Dim TotalLvl As Byte

Dim tLVL, tCOL As Byte

Dim PctBaseAmt(20) As Currency
Dim LineFeeds As Long

' *** variables for the 20 possible columns ***
Dim ColDesc(20) As String
Dim ColColType(20) As Byte
Dim ColColCat(20) As String
Dim ColFiscalYear(20) As Byte
Dim ColStartNum(20) As Byte
Dim ColEndNum(20) As Byte
Dim ColBudget(20) As Byte
Dim ColPrintTab(20) As Byte
Dim ColNonPrint(20) As Byte
' **********************************************

Dim FntSize As Integer
Dim byteOrient As Byte
Dim strOrient As String
Dim FFSchedID2 As Long
Dim SubAcct As Long

Dim fgPrintFlag As Boolean
Dim fgSignFlag As Integer
Dim MathRecFlag As Boolean
Dim PFlag As Boolean
Dim PAcct As Long
Dim AmtFlag As Boolean

Public Sub FreeFormatPrint(ByVal FFColID As Long, _
                           ByVal FFSchedID As Long, _
                           ByVal FFCount As Byte)
                           
    ' -------------------------------------------------------------------
    '   fgProcessLine and related subs
    '   used for GLFG2.RLH - order by acct# - not using GLFFSched
    '   has logic for consolidated
    ' -------------------------------------------------------------------
    
    frmProgress.Show
    
    If PRGlobal.GetByID(FFColID) = False Then
        MsgBox "GL Free Format column definition not found!", vbExclamation
        GoBack
    End If
    ColGlobID = PRGlobal.GlobalID
    
    FFSchedID2 = FFSchedID
    
    SchedGlobID = 0
    If FFSchedID <> 0 Then
        If PRGlobal.GetByID(FFSchedID) = False Then
            MsgBox "GL Free Format Schedule not found!", vbExclamation
            GoBack
        End If
        SchedGlobID = PRGlobal.GlobalID
    End If
    
    GLPStartPd = GetPeriod(GLCompany.NumberPds, GLPrint.BeginDate Mod 100, GLCompany.FirstPeriod)
    GLPEndPd = GetPeriod(GLCompany.NumberPds, GLPrint.EndDate Mod 100, GLCompany.FirstPeriod)
    
    GetDates GLCompany.ID
    DateSetup

    ' loop thru columns
    ' determine how many years of data is needed
    '
    ' also - store data from GLColumn in dim array
    On Error Resume Next
    rsFY.Close
    On Error GoTo 0
    Set rsFY = Nothing
    rsFY.CursorLocation = adUseClient
    rsFY.Fields.Append "FY", adDouble
    rsFY.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM GLFFColumn WHERE GlobalID = " & ColGlobID & _
                " ORDER BY ColNum"
    If GLFFColumn.GetBySQL(SQLString) = False Then
        MsgBox "No column definitions found for: " & ColGlobID, vbExclamation
        GoBack
    End If
    
    I = 0
    PrintCols = 0
    ct = 0
    
    Do
        
        ' store the different FY values needed
        If ColCat(GLFFColumn.ColType) = "Mth" Then
            rsFY.Find "FY = " & GLPrint.FiscalYear - GLFFColumn.FiscalYear, 0, adSearchForward, 1
            If rsFY.EOF = True Then
                rsFY.AddNew
                rsFY!FY = GLPrint.FiscalYear - GLFFColumn.FiscalYear
                rsFY.Update
            End If
        End If
        
        I = I + 1
        
        ' store data in dim array variables also
        ColDesc(I) = GLFFColumn.Description
        ColColType(I) = GLFFColumn.ColType
        ColColCat(I) = ColCat(GLFFColumn.ColType)
        ColFiscalYear(I) = GLFFColumn.FiscalYear
        ColStartNum(I) = GLFFColumn.StartNum
        ColEndNum(I) = GLFFColumn.EndNum
        ColBudget(I) = GLFFColumn.Budget
        ColPrintTab(I) = GLFFColumn.PrintTab
        ColNonPrint(I) = GLFFColumn.NonPrint
        
        ' number of columns to print
        If GLFFColumn.NonPrint = 0 And GLFFColumn.ColType <> 0 Then
            PrintCols = PrintCols + 1
        End If
        
        If GLFFColumn.GetNext = False Then Exit Do
    
    Loop
    
    If GLPrint.WidePrint = False Then
        ' force to landscape
        byteOrient = Equate.LandScape
        strOrient = "Land"
        If PrintCols <= 6 Then
            FntSize = 9
        ElseIf PrintCols = 7 Then
            FntSize = 8
        Else
            FntSize = 7
        End If
    ElseIf (GLPrint.WidePrint = True And PrintCols <= 6) Or PrintCols < 6 Then
        ' portrait if it can fit
        byteOrient = Equate.Portrait
        strOrient = "Port"
        If PrintCols <= 4 Then
            FntSize = 9
        ElseIf PrintCols = 5 Then
            FntSize = 8
        Else
            FntSize = 7
        End If
    Else
        ' landscape
        ' force to landscape
        byteOrient = Equate.LandScape
        strOrient = "Land"
        If PrintCols <= 6 Then
            FntSize = 9
        ElseIf PrintCols = 7 Then
            FntSize = 8
        Else
            FntSize = 7
        End If
    End If
    
    ' output formatting
    If FFCount = 1 Then
        PrtInit strOrient
        SetFont FntSize, byteOrient
    Else
        SetFont FntSize, byteOrient
        Prvw.vsp.NewPage
    End If
    Ln = 0
    
    ' clear variables
    For I = 1 To 20
        ColVal(I) = 0
    Next I
    For tLVL = 1 To 5
        For tCOL = 1 To 20
            Total(tLVL, tCOL) = 0
        Next tCOL
    Next tLVL
    
    ' loop thru the accounts in account # order
    ' for the range in GLPrint
    frmProgress.MousePointer = vbHourglass
    If SchedGlobID = 0 Then
        LoopAccts
    Else        ' use GLFFSched
        LoopFFSched
    End If
    frmProgress.MousePointer = vbArrow
    
End Sub

Private Sub LoopFFSched()
    
    GLAccount.OpenRS
    GLAmount.OpenRS
    
    SQLString = "SELECT * FROM GLFFSched WHERE GlobalID = " & SchedGlobID & _
                " ORDER BY SortOrder"
    If GLFFSched.GetBySQL(SQLString) = False Then
        boo = PRGlobal.GetByID(SchedGlobID)
        MsgBox "No schedule records found for: " & PRGlobal.Description, vbExclamation
        GoBack
    End If

    Do
        
        ct = ct + 1
        If ct Mod 20 = 1 Then
            frmProgress.lblMsg2 = "Processing record: " & Format(ct, "#,###,##0") & " Of: " & _
                                  Format(GLFFSched.Records, "#,###,##0")
            frmProgress.Refresh
        End If
        
        If GLFFSched.Account = 0 Then GoTo NxtFF
                                    
        If GLAccount.GetAccount(GLFFSched.Account) = False Then
            MsgBox "GL Account not found: " & GLFFSched.Account, vbExclamation
            GoBack
        End If
    
        For I = 1 To 20
            ColVal(I) = 0
        Next I
    
        PrintTab = GLFFSched.PrintTab
        
        If GLFFSched.PercentBase <> 0 Then
            ' clear and set the amounts
            For I = 1 To 20
                PctBaseAmt(I) = 0
                If ColColCat(I) = "Mth" Then
                    PctBaseAmt(I) = Abs(GetMthVal(GLFFSched.PercentBase, I))
                End If
            Next I
        End If
        
        ProcessLine
    
NxtFF:
        If GLFFSched.GetNext = False Then Exit Do

    Loop

    frmProgress.MousePointer = vbArrow
    frmProgress.Hide

End Sub

Private Sub LoopAccts()

    ' *****************************************************
    ' from GLFG - use accounts straight from GLAccount
    ' uses fg subs
    ' *****************************************************
    
    GLAccount.GetAllAccounts
    
    PFlag = False
    PAcct = 0
    fgSignFlag = 1
    
    ' ************
    ' modify screen to allow Lo/Hi cons for cons option
    If GLPrint.RegBraCon = Equate.Consol Then
        GLPrint.LowBranchAcct = 0       ' lo cons
        GLPrint.HiBranchAcct = 999      ' hi cons
        GLPrint.LowConsAcct = 0
        GLPrint.HiConsAcct = 999
    End If
    If GLPrint.RegBraCon = Equate.Branch Then
        GLPrint.HiBranchAcct = GLPrint.LowBranchAcct
    End If
    ' ************
    
    Do
    
        ct = ct + 1
        If ct Mod 20 = 1 Then
            frmProgress.lblMsg2 = "Processing record: " & Format(ct, "#,###,##0") & " " & _
                                  GLAccount.Account
            frmProgress.Refresh
        End If
        
        fgPrintFlag = False
        
        If GLAccount.Account < GLPrint.LowAccount Then GoTo NxtAcct
        If GLAccount.Account > GLPrint.HiAccount Then Exit Do
        
        SubAcct = GLAccount.Account Mod 10 ^ GLCompany.SubDigits

        If GLPrint.RegBraCon = Equate.Branch Or GLPrint.RegBraCon = Equate.Consol Then
            If SubAcct < GLPrint.LowBranchAcct Then GoTo NxtAcct
            If SubAcct > GLPrint.HiBranchAcct Then GoTo NxtAcct
        End If
        
        ' ------------------------------------------------------------------------------
        ' --- 1358 thru 1366 ------------------
        If GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements = True Then
        ElseIf GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules = True Then
        ElseIf GLPrint.RegBraCon = Equate.Branch And GLAccount.BranchAcct = True Then
        ElseIf GLPrint.RegBraCon = Equate.Consol And GLAccount.ConsAcct = True Then
        ElseIf InStr(1, "N0TM", GLAccount.AcctType, vbTextCompare) = 0 Then
            GoTo NxtAcct
        End If
        fgPrintFlag = True
        ' ------------------------------------------------------------------------------
        
        ' 1346 - suppress print tab
        If InStr(1, "IET0", GLAccount.AcctType, vbTextCompare) Then
            GLAccount.PrintTab = 0
        End If
    
        FullDesc = GLAccount.FullDesc
        
        ' GLFG2.RLH - default tabs
        PrintTab = GLAccount.PrintTab
        If GLAccount.PrintTab = 0 Then
            If InStr(1, "ALIE", GLAccount.AcctType, vbTextCompare) Then PrintTab = 1
            If InStr(1, "0NM", GLAccount.AcctType, vbTextCompare) Then PrintTab = 3
            If GLAccount.AcctType = "T" Then PrintTab = 5
        End If
        
        TotalLvl = GLAccount.TotalLevel
        If InStr(1, "CMT", GLAccount.AcctType, vbTextCompare) Then
            If GLAccount.TotalLevel < 1 Then TotalLvl = 1
            If GLAccount.TotalLevel > 5 Then TotalLvl = 5
        End If
        
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
        Else        ' 2085 - 2100
                    
            If GLAccount.AcctType = "D" And GLAccount.TotalLevel < 10 Then fgProcessLine
            
            If GLAccount.AcctType = "H" Then
                ' *** total level - link to header select ***
                If GLAccount.TotalLevel = 0 Then
                    If GLAccount.PrintTab = 0 Then
                        fgProcessLine
                    End If
                End If
                If GLAccount.TotalLevel = 10 Then
                    fgProcessLine
                End If
                If GLAccount.TotalLevel = 5 Then    ' ******
                    fgProcessLine
                End If
            End If
                
            If InStr(1, "ALIE", GLAccount.AcctType, vbTextCompare) Then
                If InStr(1, "AE", GLAccount.AcctType, vbTextCompare) Then
                    fgSignFlag = 1
                Else
                    fgSignFlag = -1
                End If
                fgProcessLine
            End If
        
        End If
                
        If InStr(1, "0N", GLAccount.AcctType, vbTextCompare) Then
            fgType0MN
        End If
        
        If GLAccount.AcctType = "M" And GLPrint.UseMathRec = True Then
            fgType0MN
        End If
        
        If GLAccount.AcctType = "T" Then
            If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
            Else
                boo = False
                If GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements = True Then boo = True
                If GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules = True Then boo = True
                If boo Then fgTypeT
                fgClearTotals
            End If
        End If
        
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then GoTo NxtAcct
        
        If GLAccount.AcctType = "C" Then fgClearTotals
        
        If GLAccount.AcctType = "U" Then TypeU
        
        If InStr(1, "BP", GLAccount.AcctType, vbTextCompare) Then
            If GLAccount.AcctType = "P" Then
                PFlag = True
                PAcct = GLAccount.Account
            End If
            FormFeed
            Ln = 0
        End If

NxtAcct:
        If GLAccount.GetNext = False Then Exit Do
    
    Loop

End Sub

Private Sub fgProcessLine()
 
    If InStr(1, "BP", GLAccount.AcctType, vbTextCompare) Then
        
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
            Exit Sub
        End If
        
        FormFeed
        Ln = 0
        Exit Sub
    End If
    
    If GLAccount.AcctType = "C" Then
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
            Exit Sub
        End If
    End If
    
    FullDesc = GLAccount.FullDesc
    
    TotalLvl = GLAccount.TotalLevel
    If InStr(1, "CMT", GLAccount.AcctType, vbTextCompare) Then
        If GLAccount.TotalLevel < 1 Then TotalLvl = 1
        If GLAccount.TotalLevel > 5 Then TotalLvl = 5
    End If
    
    ' alt desc from the GLFFSched record
    If SchedGlobID <> 0 Then
        If Trim(GLFFSched.AltDesc) <> "" Then
            FullDesc = GLFFSched.AltDesc
        End If
    End If
    
    ' date string
    If GLAccount.AcctType = "D" Then
        If GLAccount.TotalLevel <= 17 Or GLAccount.TotalLevel = 20 Then
            FullDesc = FullDesc & GLDate(GLAccount.TotalLevel)
        End If
    End If
    
    If GLAccount.DescNumber = 1 And GLAccount.AcctType = "H" Then
        FullDesc = GLCompany.Name
    End If
    
    If InStr(1, "0MN", GLAccount.AcctType, vbTextCompare) Then fgType0MN
        
    If GLAccount.AcctType = "T" Then
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
            Exit Sub
        End If
        fgTypeT
    End If
        
    If InStr(1, "ALIEHD", GLAccount.AcctType, vbTextCompare) Then
        
        If PrintTab <> 0 Then
            X = Space(PrintTab) & FullDesc
        Else        ' center it
            If GLAccount.AcctType = "H" Or GLAccount.AcctType = "D" Then
                X = Space((Columns - Len(FullDesc)) / 2) & FullDesc
            Else
                X = Space(PrintTab) & FullDesc
            End If
        End If
        
        If GLPrint.PrtAcctNum = True Then
            Y = CStr(GLAccount.Account)
            If Len(Trim(X)) <= Len(Trim(Y)) Then
                X = Trim(Y) & " " & Trim(X)
            Else
                X = Trim(Y) & Mid(X, Len(Y) + 1, Len(X) + Len(Y))
            End If
        End If
   
   
        PrintValue(1) = X:          FormatString(1) = "a" & Columns
        PrintValue(2) = " ":        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
    
    End If

    ' line feeds
    ' bottom of page ???
    ' If Ln >= MaxLines - 3 Then
    If Ln >= MaxLines Then
        FormFeed
        Ln = 0
    End If
    
    PrintLineFeeds

End Sub

Private Sub fgTypeT()

    If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
        Exit Sub
    End If

    AmtFlag = False
    
    For K = 1 To 20
        
        If ColColType(K) = 0 Then GoTo NxtT
        
        If ColStartNum(K) = 99 Then
            ' used for pct base calc
            ' only valid if using FFSched
            Val1 = 0
        ElseIf ColStartNum(K) <> 0 Then
            Val1 = Total(TotalLvl, ColStartNum(K))
            If ColBudget(ColStartNum(K)) Then Val1 = -Val1
        Else
            Val1 = 0
        End If
        
                
        If ColEndNum(K) = 99 Then
            ' used for pct base calc
            ' only valid if using FFSched
            Val2 = 0
        ElseIf ColEndNum(K) <> 0 Then
            Val2 = Total(TotalLvl, ColEndNum(K))
            If ColBudget(ColEndNum(K)) Then Val2 = -Val2
        Else
            Val2 = 0
        End If
        
        If ColColCat(K) = "Col" Then
            Select Case ColColType(K)
                Case Equate.ColAdd
                    ColVal(K) = Val1 + Val2
                Case Equate.ColAvg
                    ColVal(K) = Val1 / GLPEndPd
                Case Equate.ColDivide
                    If ColEndNum(K) = 99 Then
                        ' percent base calc
                        ' *** only valid if using FFSched ***
                        ColVal(K) = 0
                        'If PctBaseAmt(ColStartNum(K)) <> 0 Then
                        '    ColVal(K) = Div0(Val1, PctBaseAmt(ColStartNum(K)))
                        'End If
                    Else
                        ' divided column values on same row
                        If ColStartNum(K) <> 0 And ColEndNum(K) <> 0 Then
                            ColVal(K) = Div0(Val1, Val2)
                        End If
                    End If
                Case Equate.ColMultiply
                    ColVal(K) = Val1 * Val2
                Case Equate.ColProj
                    ColVal(K) = Val1 / GLPEndPd * GLCompany.NumberPds
                Case Equate.ColSubtract
                    ColVal(K) = Val1 - Val2
            End Select
        
        Else
        
            ColVal(K) = Total(TotalLvl, K)
        
        End If
        
        ' --- 2980 to 2910 ----------------------
        I = fgSignFlag
        If GLAccount.AcctType = "T" And TotalLvl = 5 And PFlag = True And GLAccount.Account > PAcct Then
            I = -1
        End If
        If GLPrint.StaSch = Equate.Stmt And GLAccount.SignRevStmt = True Then
            I = -I
        End If
        If GLPrint.StaSch = Equate.Sched And GLAccount.SignRevSched = True Then
            I = -I
        End If
        ColVal(K) = ColVal(K) * I
        
        ' 2910
        If (ColColCat(K) = "Col" And ColColType(K) = Equate.ColDivide) = False Then
            If ColBudget(K) = 1 Then
                ColVal(K) = -ColVal(K)
            End If
        End If
        ' --------------------------------------

        If ColVal(K) <> 0 Then AmtFlag = True

NxtT:
    
    Next K

    If AmtFlag = False And GLPrint.PrtZeroBal = False Then
    ElseIf GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements = False Then
    ElseIf GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules = False Then
    ElseIf fgPrintFlag = False Then
    Else
        PrintAmtLine
    End If

    fgClearTotals
    
End Sub

Private Sub fgType0MN()

Dim FFCol As Byte

    If GLAccount.AcctType = "M" Then
        fgClearTotals
    End If
    
    ' get the lookup values for the line
    For FFCol = 1 To 20
        
        ' clear consolidated accums 2460 / 2465
        If GLPrint.RegBraCon <> Equate.Consol Then
            ColVal(FFCol) = 0
        ElseIf GLPrint.RegBraCon = Equate.Consol And SubAcct = GLPrint.LowConsAcct Then
            ColVal(FFCol) = 0
        End If
        
        If ColColType(FFCol) <> 0 Then
            If ColColCat(FFCol) = "Mth" Then
                ColVal(FFCol) = ColVal(FFCol) + GetMthVal(GLAccount.Account, FFCol)
            End If
        End If
    
    Next FFCol
    
    ' column calcs
    For FFCol = 1 To 20
        
        If ColColCat(FFCol) = "Col" Then
                
            ' start/end num is "99" for percent base in use with FFSched
            If ColStartNum(FFCol) <> 0 And ColStartNum(FFCol) <= 20 Then
                Val1 = ColVal(ColStartNum(FFCol))
                If ColBudget(ColStartNum(FFCol)) Then Val1 = -Val1
            Else
                Val1 = 0
            End If
            
            If ColEndNum(FFCol) <> 0 And ColEndNum(FFCol) <= 20 Then
                Val2 = ColVal(ColEndNum(FFCol))
                If ColBudget(ColEndNum(FFCol)) Then Val2 = -Val2
            Else
                Val2 = 0
            End If
            
            Select Case ColColType(FFCol)
                Case Equate.ColAdd
                    ColVal(FFCol) = Val1 + Val2
                Case Equate.ColAvg
                    ColVal(FFCol) = Val1 / GLPEndPd
                Case Equate.ColDivide
                    If ColEndNum(FFCol) = 99 Then
                        ' percent base calc
                        ' *** only applies if using FFSched ***
                        ColVal(FFCol) = 0
'                        If PctBaseAmt(ColStartNum(FFCol)) <> 0 Then
'                            ColVal(FFCol) = Div0(Val1, PctBaseAmt(ColStartNum(FFCol)))
'                        End If
                    Else
                        ' divided column values on same row
                        If ColStartNum(FFCol) <> 0 And ColEndNum(FFCol) <> 0 Then
                            ColVal(FFCol) = Div0(Val1, Val2)
                        End If
                    End If
                Case Equate.ColMultiply
                    ColVal(FFCol) = Val1 * Val2
                Case Equate.ColProj
                    ColVal(FFCol) = Val1 / GLPEndPd * GLCompany.NumberPds
                Case Equate.ColSubtract
                    ColVal(FFCol) = Val1 - Val2
                                
            End Select
        
        End If
    
        ' update totals
        For K = 1 To 5
            
            ' 2610
            If GLPrint.RegBraCon <> Equate.Consol And GLAccount.AcctType <> "T" Then
                Total(K, FFCol) = Total(K, FFCol) + ColVal(FFCol)
            End If
            
            ' 2615
            If GLPrint.RegBraCon = Equate.Consol And GLAccount.AcctType <> "T" And SubAcct = GLPrint.HiConsAcct Then
                Total(K, FFCol) = Total(K, FFCol) + ColVal(FFCol)
            End If
            
            ' 2620
            If GLAccount.AcctType = "M" And K >= TotalLvl Then Exit For
        
        Next K
        
        ' round ? - 2630 thru 2650
    
    Next FFCol
 
    ' 2665-2670
    If GLAccount.AcctType = "M" Then
        MathRecFlag = True
        Exit Sub
    End If
    
    ' 2675
    If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
        Exit Sub
    End If
    
    AmtFlag = False
    
    For FFCol = 1 To 20

        ' 2688
        ColVal(FFCol) = ColVal(FFCol) * fgSignFlag

        ' 2690
        If ColBudget(FFCol) = 1 Then ColVal(FFCol) = -ColVal(FFCol)

        ' 2692
        If GLPrint.StaSch = Equate.Stmt And GLAccount.SignRevStmt = True Then
            ColVal(FFCol) = -ColVal(FFCol)
        End If
        If GLPrint.StaSch = Equate.Sched And GLAccount.SignRevSched = True Then
            ColVal(FFCol) = -ColVal(FFCol)
        End If

        If ColVal(FFCol) <> 0 Then AmtFlag = True

    Next FFCol
    
    ' 2710
    If AmtFlag = False And GLPrint.PrtZeroBal = False Then Exit Sub
    
    ' 2715
    If GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements = False Then
        Exit Sub
    End If
    If GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules = False Then
        Exit Sub
    End If
    
    ' 2720
    If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
        Exit Sub
    End If
    
    ' 2725
    If fgPrintFlag = False Then Exit Sub
    
    PrintAmtLine
    
    ' 2745
    If GLAccount.AcctType = "T" Then
        MathRecFlag = False
    End If
    
    
End Sub

Private Sub fgClearTotals()

Dim Clr As Byte
Dim ClrLvl As Byte
Dim fgSign As Integer
Dim SignFlag As Boolean
    
    
    For Clr = 1 To 20
            
        fgSign = fgSignFlag
        If GLAccount.AcctType = "T" And TotalLvl = 5 And GLAccount.Account > PAcct Then fgSign = -1
        Total(TotalLvl, Clr) = Total(TotalLvl, Clr) * fgSign
        If GLAccount.AcctType = "T" And TotalLvl = 5 And ColBudget(Clr) = 1 Then
            Total(TotalLvl, Clr) = -Total(TotalLvl, Clr)
        End If
    
        ' 3020
        If InStr(1, "MT", GLAccount.AcctType, vbTextCompare) And TotalLvl <= 4 Then
            For ClrLvl = 1 To TotalLvl
                Total(ClrLvl, Clr) = 0
            Next ClrLvl
        End If
    
    Next Clr

End Sub


Private Sub ProcessLine()
    
    If InStr(1, "BP", GLAccount.AcctType, vbTextCompare) Then
        FormFeed
        Ln = 0
        Exit Sub
    End If
    
    FullDesc = GLAccount.FullDesc
    
    TotalLvl = GLAccount.TotalLevel
    If InStr(1, "CMT", GLAccount.AcctType, vbTextCompare) Then
        If GLAccount.TotalLevel < 1 Then TotalLvl = 1
        If GLAccount.TotalLevel > 5 Then TotalLvl = 5
    End If
    
    ' alt desc from the GLFFSched record
    If Trim(GLFFSched.AltDesc) <> "" Then
        FullDesc = GLFFSched.AltDesc
    End If
    
    ' date string
    If GLAccount.AcctType = "D" Then
        If GLAccount.TotalLevel <= 17 Or GLAccount.TotalLevel = 20 Then
            FullDesc = FullDesc & GLDate(GLAccount.TotalLevel)
        End If
    End If
    
    If GLAccount.AcctType = "U" Then
        TypeU True
    End If
    
    If GLAccount.DescNumber = 1 And GLAccount.AcctType = "H" Then
        FullDesc = GLCompany.Name
    End If
    
    If InStr(1, "0MN", GLAccount.AcctType, vbTextCompare) Then Type0MN
        
    If GLAccount.AcctType = "T" Then
        If GLPrint.RegBraCon = Equate.Consol And SubAcct <> GLPrint.HiConsAcct Then
            Exit Sub
        End If
        TypeT
    End If
        
    If InStr(1, "ALIEHD", GLAccount.AcctType, vbTextCompare) Then
        
        If PrintTab <> 0 Then
                X = Space(PrintTab) & FullDesc
        Else        ' center it
            If GLAccount.AcctType = "H" Or GLAccount.AcctType = "D" Then
                X = Space((Columns - Len(FullDesc)) / 2) & FullDesc
            Else
                X = Space(PrintTab) & FullDesc
            End If
        End If
        
        If GLPrint.PrtAcctNum = True Then
            Y = CStr(GLAccount.Account)
            If Len(Trim(X)) <= Len(Trim(Y)) Then
                X = Trim(Y) & " " & Trim(X)
            Else
                X = Trim(Y) & Mid(X, Len(Y) + 1, Len(X) + Len(Y))
            End If
        End If
   
        PrintValue(1) = X:          FormatString(1) = "a" & Columns
        PrintValue(2) = " ":        FormatString(2) = "~"
        FormatPrint
        Ln = Ln + 1
    
    End If

    ' line feeds
    ' bottom of page ???
    ' If Ln >= MaxLines - 3 Then
    If Ln >= MaxLines Then
        FormFeed
        Ln = 0
    End If
    
    PrintLineFeeds
    
End Sub

Private Sub PrintLineFeeds()
    
Dim lfFlag As Boolean
    
    If SchedGlobID <> 0 Then
        LineFeeds = GLFFSched.LineFeeds
        
        ' if > 100 last digit designates lines from bottom of page
        ' ex.  LineFeeds = 102 ==> go to 2 lines from bottom of page
        If LineFeeds > 100 Then
            LineFeeds = MaxLines - Ln - (LineFeeds Mod 100) - 1
        End If
        
        lfFlag = True
        
    Else
        LineFeeds = GLAccount.LineFeeds
        If LineFeeds = 255 Then LineFeeds = -1
    End If
    
    If LineFeeds >= 40 And Ln < LineFeeds And lfFlag = False Then
        LineFeeds = LineFeeds - Ln
    End If
    
    For I = 1 To LineFeeds
        Ln = Ln + 1
        If Ln >= MaxLines Then
            FormFeed
            Ln = 0
        End If
    Next I

End Sub

Private Sub PrintAmtLine()

Dim DollarSign As String

    ' print audit line?
    If GLUser.Logon = "jim" And GLPrint.PrtAcctNum = True Then
        Dim ff As Integer
        For ff = 1 To 5
            PrintValue(ff) = Total(ff, 1)
            FormatString(ff) = "d12"
        Next ff
        PrintValue(6) = ""
        FormatString(6) = "~"
        FormatPrint
        Ln = Ln + 1
    End If

    If GLPrint.PrtAcctNum = True Then
        X = GLAccount.Account & " " & FullDesc
        ' x = GLAccount.Account & " " & GLAccount.AcctType & " " & FullDesc
    Else
        X = FullDesc
    End If
    
    If FFSchedID2 <> 0 Then
        PrintValue(1) = X:         FormatString(1) = "a40"  ' 41?
    Else
        PrintValue(1) = X:         FormatString(1) = "a36"  ' 41?
    End If
    
    If GLAccount.DollarSign = True Then
        DollarSign = "$"
    Else
        DollarSign = " "
    End If
  
  ' GLPrint.RoundDollars = True
    AmtFlag = False
    J = 2
    For I = 1 To 20
        
        If ColColType(I) <> 0 Then
            
            If ColNonPrint(I) <> 1 Then
                
                If ColVal(I) <> 0 Then AmtFlag = True
                
                If ColColType(I) = Equate.ColDivide Then
                    
                    ' for Free Format - fix the percent sign
                    If SchedGlobID <> 0 And ColEndNum(I) = 99 Then
                        If ColVal(I - 1) > 0 Then
                            If ColVal(I) < 0 Then ColVal(I) = -ColVal(I)
                        Else
                            If ColVal(I) > 0 Then ColVal(I) = -ColVal(I)
                        End If
                    End If
                
                    ' percent format
                    PrintValue(J) = "":                     FormatString(J) = "a1"
                    PrintValue(J + 1) = ColVal(I):          FormatString(J + 1) = "p6"
                
                ElseIf GLPrint.RoundDollars = True Then
                    ' round to nearest dollar
                    PrintValue(J) = DollarSign:             FormatString(J) = "a1"
                    PrintValue(J + 1) = Round(ColVal(I), 0): FormatString(J + 1) = "i14"
                Else
                    ' normal amount print
                    PrintValue(J) = DollarSign:             FormatString(J) = "a1"
                    PrintValue(J + 1) = ColVal(I):          FormatString(J + 1) = "d14"
                End If
                
                J = J + 2
            
            End If
        
        End If
    
    Next I
 
    ' suppress zero amount lines
    If GLPrint.PrtZeroBal = False And AmtFlag = False Then Exit Sub
    
    PrintValue(J) = " ":        FormatString(J) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub PrintHdrLine()

    If GLPrint.PrtAcctNum = True Then
        X = GLAccount.Account & " " & GLAccount.FullDesc
        X = GLAccount.Account & " " & GLAccount.AcctType & " " & GLAccount.FullDesc
    Else
        X = GLAccount.FullDesc
    End If
    PrintValue(1) = X:          FormatString(1) = "a30"
    PrintValue(2) = " ":        FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

End Sub

Private Sub TypeT()

    For K = 1 To 20
        
        ColVal(K) = 0
        If ColColType(K) = 0 Then GoTo NxtT
        
        ' 2830 thru 2860 - Pct / Div
        
        ' 2890
        J = 20 * TotalLvl - 20
        
        ' 2900 round dollars
        
        ' 2910
        If Not (ColColCat(K) = "Col" And ColColType(K) = Equate.ColDivide) Then
            ColVal(K) = Total(TotalLvl, K) * GLSign(ColBudget(K))
        Else
            If ColEndNum(K) = 99 Then
                
                ' pct base
                ' ColVal(K) = Div0(Total(TotalLvl, ColStartNum(K)), PctBaseAmt(ColStartNum(K)))
            
                ' *** jb 10/30/2010 - always use the column to the left for pct base calcs
                ColVal(K) = Div0(Total(TotalLvl, K - 1), PctBaseAmt(K - 1))
            
            Else
                ' pct between columns
                ColVal(K) = Div0(Total(TotalLvl, ColStartNum(K)), Total(TotalLvl, ColEndNum(K)))
            End If
        End If

        ' 2910
        ' **** if ~[pct and column val] ****
        ' ColVal(k) = ColVal(k) * GLSign(ColBudget(k))

NxtT:
    
    Next K

    If GLAccount.AllStatements = True Then
        PrintAmtLine
    End If
    
    ClearTotals
    
End Sub

Private Sub Type0MN()

Dim FFCol As Byte

    ' MsgBox "Type 0MN" & vbCr & GLAccount.Account
    
    ' ????
    ' GLAccount.AcctType = "M" Then ClearTotals
    
    ' get the lookup values for the line
    For FFCol = 1 To 20
        
        If ColColType(FFCol) <> 0 Then
        
            If ColColCat(FFCol) = "Mth" Then
                ColVal(FFCol) = GetMthVal(GLAccount.Account, FFCol)
            End If
        
        End If
    
    Next FFCol
    
    ' loop back thru and do the "Col" columns
    '   and update totals for all columns
    For FFCol = 1 To 20
        
        If ColColType(FFCol) <> 0 Then
        
            If ColColCat(FFCol) = "Col" Then
                
                Select Case ColColType(FFCol)
                    Case Equate.ColAdd
                        ColVal(FFCol) = ColVal(ColStartNum(FFCol)) + ColVal(ColEndNum(FFCol))
                    Case Equate.ColAvg
                        ColVal(FFCol) = ColVal(ColStartNum(FFCol)) / GLPEndPd
                    Case Equate.ColDivide
                        If ColEndNum(FFCol) = 99 Then
                            
'                            ' percent base calc
'                            If PctBaseAmt(ColStartNum(FFCol)) <> 0 Then
'                                ColVal(FFCol) = Div0(ColVal(ColStartNum(FFCol)), PctBaseAmt(ColStartNum(FFCol)))
'                            End If
                        
                            ' percent base calc
                            ' *** jb 10/30/2010 - always use the column to the left for pct base calcs
                            If PctBaseAmt(FFCol - 1) <> 0 Then
                                
                                ColVal(FFCol) = Div0(ColVal(FFCol - 1), PctBaseAmt(FFCol - 1))
                                
'                                ' ??? always assume the pct base is positive
'                                ' match the sign of the value of the column to the left
'                                If ColVal(FFCol - 1) > 0 Then
'                                    ColVal(FFCol) = Abs(ColVal(FFCol))
'                                Else
'                                    If ColVal(FFCol) > 0 Then ColVal(FFCol) = -ColVal(FFCol)
'                                End If
                            
                            End If
                        
                        Else
                            ' divided column values on same row
                            If ColStartNum(FFCol) <> 0 And ColEndNum(FFCol) <> 0 Then
                                ColVal(FFCol) = Div0(ColVal(ColStartNum(FFCol)), ColVal(ColEndNum(FFCol)))
                            End If
                        End If
                    Case Equate.ColMultiply
                    Case Equate.ColProj
                        ColVal(FFCol) = ColVal(ColStartNum(FFCol)) / GLPEndPd * GLCompany.NumberPds
                    Case Equate.ColSubtract
                        ColVal(FFCol) = ColVal(ColStartNum(FFCol)) - ColVal(ColEndNum(FFCol))
                End Select

                ColVal(FFCol) = ColVal(FFCol)

            End If
            
            ' update totals
            For K = 1 To 5
                Total(K, FFCol) = Total(K, FFCol) + ColVal(FFCol)
            Next K
        
        End If
        
    Next FFCol

    ' loop for sign rev before printing
    If GLAccount.AcctType <> "M" Then
        For FFCol = 1 To 20
            If ColEndNum(FFCol) <> 99 Then ' leave percentages alone ???
                ColVal(FFCol) = ColVal(FFCol) * GLSign(ColBudget(FFCol))
            End If
        Next FFCol
    End If
    
    If GLAccount.AcctType = "M" Then Exit Sub       ' 2670
      
    ' don't print the line for FF statements
    If GLAccount.AllStatements = False Then Exit Sub
    
    PrintAmtLine
    
End Sub

Private Function GetMthVal(ByVal GLAcct As Long, ByVal ColNum As Byte)

Dim StartPD, EndPd As Byte
                
   Select Case ColColType(ColNum)
       Case Equate.ColAllPd
           StartPD = 1
           EndPd = GLCompany.NumberPds
       Case Equate.ColCurrPd
           StartPD = GLPStartPd
           EndPd = GLPEndPd
       Case Equate.ColCustom
           StartPD = ColStartNum(ColNum)
           EndPd = ColEndNum(ColNum)
       Case Equate.ColPriorPd
           StartPD = GLPStartPd - 1
           EndPd = GLPEndPd - 1
       Case Equate.ColYTD
           StartPD = 1
           EndPd = GLPEndPd
       Case Else
           MsgBox "Col Type?: " & ColColType(ColNum) & " " & ColNum, vbExclamation
           GoBack
   End Select

   If ColBudget(ColNum) = 0 Then
       ' get amount for the Fy/Pd range
       GetMthVal = GLAmount.GetAmount(GLAcct, _
                                          GLPrint.FiscalYear - ColFiscalYear(ColNum), _
                                          StartPD, EndPd)
   
   Else
       ' get BUDGET amount for the Fy/Pd range
       GetMthVal = GLAmount.GetBudget(GLAcct, _
                                          GLPrint.FiscalYear - ColFiscalYear(ColNum), _
                                          StartPD, EndPd)
   End If

End Function

Private Function ColCat(ByVal ColType As Byte) As String

    ColCat = ""
    Select Case ColType
        
        Case Equate.ColAllPd:       ColCat = "Mth"
        Case Equate.ColCurrPd:      ColCat = "Mth"
        Case Equate.ColCustom:      ColCat = "Mth"
        Case Equate.ColPriorPd:     ColCat = "Mth"
        Case Equate.ColYTD:         ColCat = "Mth"
        
        Case Equate.ColAdd:         ColCat = "Col"
        Case Equate.ColAvg:         ColCat = "Col"
        Case Equate.ColDivide:      ColCat = "Col"
        Case Equate.ColMultiply:    ColCat = "Col"
        Case Equate.ColProj:        ColCat = "Col"
        Case Equate.ColSubtract:    ColCat = "Col"
    
    End Select

End Function

Private Sub DateSetup()       ' 0660
   
Dim DString As String
Dim Months As String
Dim Digit(13) As String
Dim fmt As String
Dim ii As Double
Dim Mths As Byte

    fmt = "mmmm dd, yyyy"
   
    Digit(0) = " Zero"
    Digit(1) = " One"
    Digit(2) = " Two"
    Digit(3) = " Three"
    Digit(4) = " Four"
    Digit(5) = " Five"
    Digit(6) = " Six"
    Digit(7) = " Seven"
    Digit(8) = " Eight"
    Digit(9) = " Nine"
    Digit(10) = " Ten"
    Digit(11) = " Eleven"
    Digit(12) = " Twelve"
    Digit(13) = " Thirteen"
   
    GLDate(0) = Format(CurrYrPdEnd, fmt)
       
    If GLPEndPd = 1 Then
        GLDate(1) = CStr(Digit(GLPEndPd)) & " Month Ended " & Format(CurrYrPdEnd, fmt)
    Else
        GLDate(1) = CStr(Digit(GLPEndPd)) & " Months Ended " & Format(CurrYrPdEnd, fmt)
    End If
   
    GLDate(2) = Format(CurrYrCurrPdBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
    GLDate(3) = Format(CurrYrFYBeg, fmt) & " To " & Format(CurrYrPdEnd, fmt)
   
    If GLPEndPd = 1 Then
        GLDate(4) = CStr(Digit(GLPEndPd)) & " Month Ended " & Format(PrevYrPdEnd, fmt)
    Else
        GLDate(4) = CStr(Digit(GLPEndPd)) & " Months Ended " & Format(PrevYrPdEnd, fmt)
    End If
   
    GLDate(5) = Format(PrevYrPDBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
    GLDate(6) = Format(PrevYrFYBeg, fmt) & " To " & Format(PrevYrPdEnd, fmt)
   
    GLDate(7) = Format(Now, fmt)
   
    GLDate(8) = CStr(GLPEndPd * 4) & " Weeks Ended " & Format(CurrYrPdEnd, fmt)
   
    GLDate(9) = CStr(GLPEndPd * 4) & " Weeks Ended " & Format(PrevYrPdEnd, fmt)
   
    X = Trim(Format(CurrYrPdEnd, fmt))
    GLDate(10) = Space(18 - Len(X)) & X
   
    GLDate(11) = Space(18 - Len(Format(CurrYrCurrPdBeg, fmt))) & Format(CurrYrCurrPdBeg, fmt)
   
    GLDate(12) = Space(18 - Len(Format(CurrYrFYBeg, fmt))) & Format(CurrYrFYBeg, fmt)
   
    GLDate(13) = Space(18 - Len(Format(PrevYrPdEnd, fmt))) & Format(PrevYrPdEnd, fmt)
   
    GLDate(14) = Space(18 - Len(Format(PrevYrPDBeg, fmt))) & Format(PrevYrPDBeg, fmt)
   
    GLDate(15) = Space(18 - Len(Format(PrevYrFYBeg, fmt))) & Format(PrevYrFYBeg, fmt)
   
    If GLPEndPd - GLPStartPd + 1 = 1 Then
        DString = CStr(Digit(GLPEndPd - GLPStartPd + 1)) & " Month"
    Else
        DString = CStr(Digit(GLPEndPd - GLPStartPd + 1)) & " Months"
    End If
    X = Trim(DString)
    GLDate(16) = Space(18 - Len(X)) & X

    If GLPEndPd = 1 Then
        DString = CStr(Digit(GLPEndPd)) & " Month"
    Else
        DString = CStr(Digit(GLPEndPd)) & " Months"
    End If
    X = Trim(DString)
    GLDate(17) = Space(18 - Len(X)) & X
   
    GLDate(18) = " And " & Year(CurrYrFYEnd)
   
    GLDate(19) = " And " & Year(CurrYrFYEnd) - 1
   
    ' date type 1   - ex. pd 7 - 9
    Mths = GLPEndPd - GLPStartPd + 1
    GLDate(20) = CStr(Digit(Mths)) & " Month"                              ' Three Month
    If Mths > 1 Then GLDate(20) = GLDate(20) & "s"                          ' Three Months
    GLDate(20) = GLDate(20) & " And" & CStr(Digit(GLPEndPd)) & " Month"       ' Three Months And Nine Month
    If GLPEndPd > 1 Then GLDate(20) = GLDate(20) & "s"                         ' Three Months And Nine Months
    GLDate(20) = GLDate(20) & " Ended " & Format(CurrYrPdEnd, fmt)
   
    ' shift to upper case 0980
    If GLPrint.LowerCaseDate = False Then
        For ii = 0 To 20
            GLDate(ii) = UCase(GLDate(ii))
        Next ii
    End If

End Sub

Private Sub ClearTotals()

Dim Clr As Byte
Dim SignFlag As Boolean

    For tLVL = 1 To TotalLvl
        
        For tCOL = 1 To 20
            
            If GLAccount.AcctType = "C" Then
                Total(tLVL, tCOL) = 0
            End If

            If InStr(1, "MT", GLAccount.AcctType, vbTextCompare) And tLVL <= 4 Then
                Total(tLVL, tCOL) = 0
            End If

        Next tCOL
    
    Next tLVL

End Sub

Private Sub TypeU(Optional FFSched As Boolean)
        
Dim uLine As String
Dim uFlag As Boolean
    
    If FFSched = False Then
        uFlag = False
        If GLPrint.StaSch = Equate.Stmt And GLAccount.AllStatements = True Then uFlag = True
        If GLPrint.StaSch = Equate.Sched And GLAccount.AllSchedules = True Then uFlag = True
        If GLPrint.RegBraCon = Equate.Branch And GLAccount.BranchAcct = True Then uFlag = True
        If GLPrint.RegBraCon = Equate.Consol And GLAccount.ConsAcct = True Then uFlag = True
        If uFlag = False Then Exit Sub
    End If

    If FFSchedID2 <> 0 Then
        uLine = String(40, " ")
    Else
        uLine = String(36, " ")
    End If

    For I = 1 To 20
        
        If ColColType(I) <> 0 And ColNonPrint(I) <> 1 Then
            
            J = 0
            If ColColType(I) = Equate.ColDivide Then
                ' percent format
                J = 5
            
                ' jb 11/02/2010 - richlak fix
                J = 6
            
            ElseIf GLPrint.RoundDollars = True Then
                ' round to nearest dollar
                J = 11
            Else
                ' normat amount print
                J = 14
            End If
            
            If J <> 0 Then
                If GLAccount.TotalLevel >= 1 Then
                    uLine = uLine & String(J, "-") & " "
                Else
                    uLine = uLine & String(J, "=") & " "
                End If
            End If
            
        End If
    
    Next I

    PrintValue(1) = uLine:          FormatString(1) = "a" & Len(uLine)
    PrintValue(2) = " ":            FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintLineFeeds

End Sub

Private Function GLSign(ByVal BudgFlag As Byte) As Integer

Dim LineFlag As Byte
    
    ' changes sign if reverse flag set or a budget column - but not both
    
    ' LineFlag - see if line marked as sign reversal from GLFFSched
    If SchedGlobID = 0 Then
        LineFlag = 0
    Else
        LineFlag = GLFFSched.SignReverse
    End If
    
    ' sign reversal
    
    ' LineFlag from GLFFSched - sign reversal flag
    ' BudgFlag from Column is a budget column
    If LineFlag = 0 And BudgFlag = 0 Then GLSign = 1
    If LineFlag = 1 And BudgFlag = 0 Then GLSign = -1
    If LineFlag = 0 And BudgFlag = 1 Then GLSign = -1
    If LineFlag = 1 And BudgFlag = 1 Then GLSign = 1

End Function


