VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
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
   ScaleHeight     =   5040
   ScaleWidth      =   7995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SchedGlobID, ColGlobID As Long
Dim flg As Boolean
Dim rsFY As New ADODB.Recordset
Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim ct As Long

Dim ColVal(20) As Currency

Dim GLPStartPd, GLPEndPd As Byte
Dim PrintCols As Byte
Dim PrintString As String
Dim PrintTab As Byte
Dim GLDate(20) As String
Dim FullDesc As String

Dim Total(100) As Currency
Dim TotalLvl As Byte
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

Private Sub Form_Load()

'    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSched & _
'                " AND UserID = " & GLCompany.ID & _
'                " AND Description = 'KH'"
'    If PRGlobal.GetBySQL(SQLString) = False Then
        
        
    If PRGlobal.GetByID(328) = False Then
        MsgBox "Global NF"
        End
    End If
    SchedGlobID = PRGlobal.GlobalID
    
'    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFColumn & _
'                " AND UserID = " & GLCompany.ID & _
'                " AND Description = 'KH'"
'    If PRGlobal.GetBySQL(SQLString) = False Then
        
    If PRGlobal.GetByID(325) = False Then
        MsgBox "Global NF"
        End
    End If
    ColGlobID = PRGlobal.GlobalID
    
    GLPrint.GetData "Default", flg
    
    ' new GLPrint record was created for the user
    ' load defaults from the GLCompany file
    GLPrint.LowAccount = 1
    GLPrint.HiAccount = 999999999
    GLPrint.LowBranchAcct = GLCompany.LowBranch
    GLPrint.HiBranchAcct = GLCompany.HiBranch
    GLPrint.LowConsAcct = GLCompany.LowConsolidated
    GLPrint.HiConsAcct = GLCompany.HiConsolidated
    GLPrint.FiscalYear = 2008
    GLPrint.BeginDate = 200809
    GLPrint.EndDate = 200809
    GLPrint.Save (Equate.RecPut)
    
    ' start / end period #
    GLPStartPd = GetPeriod(GLCompany.NumberPds, GLPrint.BeginDate Mod 100, GLCompany.FirstPeriod)
    GLPEndPd = GetPeriod(GLCompany.NumberPds, GLPrint.EndDate Mod 100, GLCompany.FirstPeriod)

    ' )))))))))))))))))))))))))))
    GLPrint.LowerCaseDate = True
    ' )))))))))))))))))))))))))))
    
    GetDates GLCompany.ID
    DateSetup

    frmProgress.Show
    frmProgress.MousePointer = vbHourglass
    frmProgress.lblMsg1 = "Free Format: " & GLCompany.Name & vbCr & PRGlobal.Description
    frmProgress.Refresh
    
    ' )))))))))))))))))))))))))))))))
    GLPrint.PrtZeroBal = False
    GLPrint.PrtAcctNum = False
    
    cmdPrint_Click
    ' )))))))))))))))))))))))))))))))

End Sub
Private Sub CmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()

    ' loop thru columns
    ' determine how many years of data is needed
    '
    ' also - store data from GLColumn in dim array
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
    
    ' output format
    GLPrint.WidePrint = True
    If PrintCols <= 5 Then
        PrtInit "Port"
        SetFont 8, Equate.Portrait
    Else
        PrtInit "Land"
        SetFont 9, Equate.LandScape
    End If
    
    ' loop thru the accounts in account # order
    ' for the range in GLPrint
    If SchedGlobID = 0 Then
        LoopAccts
    Else        ' use GLFFSched
        LoopFFSched
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

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
    
    If InStr(1, "0MN", GLAccount.AcctType, vbTextCompare) Then Type0MN
        
    If GLAccount.AcctType = "T" Then TypeT
        
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
            Y = CStr(GLFFSched.Account)
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
    If Ln >= MaxLines - 3 Then
        FormFeed
        Ln = 0
    End If
    LineFeeds = GLFFSched.LineFeeds
    If LineFeeds >= 40 And Ln < LineFeeds Then
        LineFeeds = LineFeeds - Ln
    End If
    For I = 1 To LineFeeds
        Ln = Ln + 1
        If Ln >= MaxLines - 5 Then
            FormFeed
            Ln = 0
        End If
    Next I

End Sub

Private Sub PrintAmtLine()

Dim AmtFlag As Boolean
Dim DollarSign As String

    If GLPrint.PrtAcctNum = True Then
        X = GLAccount.Account & " " & FullDesc
        ' x = GLAccount.Account & " " & GLAccount.AcctType & " " & FullDesc
    Else
        X = FullDesc
    End If
    
    PrintValue(1) = X:         FormatString(1) = "a40"  ' 41?
    
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
                    ' percent format
                    PrintValue(J) = "":                     FormatString(J) = "a1"
                    PrintValue(J + 1) = ColVal(I):          FormatString(J + 1) = "p6"
                ElseIf GLPrint.RoundDollars = True Then
                    ' round to nearest dollar
                    PrintValue(J) = DollarSign:             FormatString(J) = "a1"
                    PrintValue(J + 1) = Round(ColVal(I), 0): FormatString(J + 1) = "i14"
                Else
                    ' normat amount print
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
            ColVal(K) = Total(J + K)  ' * GLSign(ColBudget(k))
        Else
            If ColEndNum(K) = 99 Then
                ' pct base
                ColVal(K) = Div0(Total(J + ColStartNum(K)), PctBaseAmt(ColStartNum(K)))
            Else
                ' pct between columns
                ColVal(K) = Div0(Total(J + ColStartNum(K)), Total(J + ColEndNum(K)))
            End If
        End If

NxtT:
    
    Next K

    PrintAmtLine
    ClearTotals

End Sub

Private Sub Type0MN()

Dim FFCol As Byte

    If GLAccount.AcctType = "M" Then ClearTotals
    
    ' get the lookup values for the line
    For FFCol = 1 To 20
        
        If ColColType(FFCol) <> 0 Then
        
            If ColColCat(FFCol) = "Mth" Then
                ColVal(FFCol) = GetMthVal(GLFFSched.Account, FFCol)
            End If
        
        End If
    
    Next FFCol
    
    ' loop for sign rev
    For FFCol = 1 To 20
        ColVal(FFCol) = ColVal(FFCol) * GLSign(ColBudget(FFCol))
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
                            ' percent base calc
                            If PctBaseAmt(ColStartNum(FFCol)) <> 0 Then
                                ColVal(FFCol) = Div0(ColVal(ColStartNum(FFCol)), PctBaseAmt(ColStartNum(FFCol)))
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

            End If
            
            ' update totals
            For K = 0 To 4
                Total(FFCol + 20 * K) = Total(FFCol + 20 * K) + ColVal(FFCol)
            Next K
        
        End If
        
    Next FFCol

    If GLAccount.AcctType = "M" Then Exit Sub       ' 2670
                
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

    For Clr = 1 To 20 * TotalLvl
        
        ' reverse the sign?
        If GLAccount.AcctType = "T" Then
            
            ' total level 5
            If Clr >= 81 Then
                
                ' budget column or using FFSched and sign revers set
                If ColBudget(Clr Mod 20) = 1 Or (SchedGlobID <> 0 And GLFFSched.SignReverse = 1) Then
                    
                    Total(Clr) = -Total(Clr)
                
                End If
            
            End If
        
        End If

        If Clr < 81 And GLAccount.AcctType = "C" Then
            Total(Clr) = 0
        End If
        
    Next Clr

End Sub

Private Function GLSign(ByVal BudgFlag As Byte) As Integer

Dim LineFlag As Byte
    
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

