VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmaMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmnOpen 
      Left            =   2280
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmaMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DBName As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsCompany As ADODB.Recordset
Dim x, Prog As String
Dim Flg As Boolean
Dim Results As frmResults
Public xDB As New XArrayDB
Public CompanyID As Long
Dim BatchNum As Long

Private Sub Form_Load()
   
Dim R1 As Long
Dim r2 As Long
Dim i As Long
Dim j As Long
Dim booX As Boolean
Dim FileChannel As Long

Dim frmResults As New frmResults
Dim frmMultDiv As New frmMultDiv
Dim frmDeleteAccts As New frmDeleteAccts
Dim frmCopy As New frmCopy
Dim frmCopyBB As New frmCopyBB
Dim frmGLUrange As New frmGLUrange
Dim frmYearEnd As New frmYearEnd
   
Dim MResp As Integer
Dim x As String
Dim y As String
Dim vv As Variant
   
'   Set GLCompany = New cGLCompany
'   Set GLAccount = New cGLAccount
'   Set GLAmount = New cGLAmount
'   Set GLHistory = New cGLHistory
'   Set GLDescription = New cGLDescription
'   Set Equate = New cEquate
'   Set GLPrint = New cGLPrint
'   Set GLBatch = New cGLBatch
'
'   SetEquates
'
'   x = Command
'
'   If x = "" Then       ' nothing on the command line
'      x = "19/DeleteAccts/C:\Balint\Data\GLSystem.mdb/JimBo"
'      x = "55/GLFileCopy/C:\Balint\Data\GLSystem.mdb/JimBo"
'      x = "55/GLFileCopy/C:\Balint\Data\GLSystem.mdb/JimBo"
''      x = "19/ClearGLBudget/C:\Balint\Data\GLSystem.mdb/JimBo"
'      x = "63/YearEnd/C:\Balint\Data\GLSystem.mdb/jim"
'      x = "69/UpdateBatch/C:\Balint\Data\GLSystem.mdb/jim/1"
'
'
'      x = "59//GLMultDiv/C:\Balint\Data\GLSystem.mdb/JimBo"
'      x = "59//CopyBB/C:\Balint\Data\GLSystem.mdb/JimBo"
'      x = "59//ClearGLAmount/C:\Balint\Data\GLSystem.mdb/jim"
'
'   End If
'
''   If cmdline(x, ID, Password, Prog, SysFile, User, BatchNum) = False Then
''      MsgBox "Bad command line !!!"
''   End If
'
'   CNDesOpen (SysFile)
'   CompanyID = ID
'
'   If Not GLCompany.GetData(CompanyID) Then
'      MsgBox "Company record not found ID# " & CompanyID
'      End
'   End If
'
'   DBName = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
'   CNOpen DBName, Password
'
'   Prog = StrConv(Prog, vbUpperCase)
'
'   GLPrint.GetData User, Flg
'
'   Select Case Prog
'
'      Case "GLMULTDIV"
'
'         Set frmMultDiv = New frmMultDiv
'
'         frmMultDiv.lblCompanyName = GLCompany.Name
'         frmMultDiv.tdbLoAccount = 1
'         frmMultDiv.tdbHiAccount = 999999999
'
'         Do
'
'            frmMultDiv.Show vbModal
'
'            If Not Response Then Exit Do
'
'            If frmMultDiv.optMultiply Then
'               x = "Mult"
'            Else
'               x = "Div"
'            End If
'
'            If frmMultDiv.optShow Then
'               y = "Show"
'            Else
'               y = "Go"
'            End If
'
'            Set uDB = GLMultDiv(frmMultDiv.tdbLoAccount.Value, _
'                                frmMultDiv.tdbHiAccount.Value, _
'                                x, _
'                                frmMultDiv.tdbValue, _
'                                y)
'
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Multiply/Divide Account Numbers"
'            frmResults.lblMsg2 = ""
'            frmResults.lblMsg3 = ""
'            For i = 1 To uDB.UpperBound(1)
'                frmResults.List1.AddItem uDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'            If y = "Show" Then
'               MResp = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
'               If MResp = vbNo Then Exit Do
'            Else
'               Exit Do   ' done
'            End If
'         Loop
'
'      Case "DELETEACCTS"
'
'         Set frmDeleteAccts = New frmDeleteAccts
'         frmDeleteAccts.lblCompanyName = GLCompany.Name
'         frmDeleteAccts.tdbLoValue = 1
'         frmDeleteAccts.tdbHiValue = 999999999
'
'         Do
'
'            frmDeleteAccts.Show vbModal
'            If Not Response Then Exit Do
'
'            If frmDeleteAccts.optRegular Then
'               x = "Acct"
'            Else
'               x = "Sub"
'            End If
'
'            If frmDeleteAccts.optShow Then
'               y = "Show"
'            Else
'               y = "Go"
'            End If
'
'            Set uDB = DeleteAccts(x, _
'                                  frmDeleteAccts.tdbLoValue, _
'                                  frmDeleteAccts.tdbHiValue, _
'                                  y, _
'                                  frmDeleteAccts.chkDelHist)
'
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Delete Accounts"
'            frmResults.lblMsg2 = ""
'            frmResults.lblMsg3 = ""
'            For i = 1 To uDB.UpperBound(1)
'                frmResults.List1.AddItem uDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'            If y = "Show" Then
'               MResp = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
'               If MResp = vbNo Then Exit Do
'            Else
'               Exit Do   ' done
'            End If
'
'         Loop
'
'      Case "GLFILECOPY"
'
'         Set frmCopy = New frmCopy
'         frmCopy.lblCompanyName = GLCompany.Name
'         frmCopy.Show vbModal
'         If Response Then
'            x = Left(GLCompany.FileName, 1) & ":\Balint\Data\" & frmCopy.txtFileName & ".mdb"
'
'            Set uDB = GLFileCopy(CompanyID, _
'                                 x, _
'                                 Password, _
'                                 frmCopy.txtCompName, _
'                                 frmCopy.chkClear, _
'                                 SysFile)
'
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Copy file to:"
'            frmResults.lblMsg2 = x
'            frmResults.lblMsg3 = ""
'            For i = 1 To uDB.UpperBound(1)
'                frmResults.List1.AddItem uDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'
'         End If
'
'      Case "COPYBB"
'
'         Set frmCopyBB = New frmCopyBB
'         frmCopyBB.lblCompanyName = GLCompany.Name
'         frmCopyBB.tdbLoAccount = 1
'         frmCopyBB.tdbHiAccount = 999999999
'
'         Do
'
'            frmCopyBB.Show vbModal
'            If Not Response Then Exit Do
'
'            If frmCopyBB.optMain Then
'               x = "Main"
'            Else
'               x = "Sub"
'            End If
'
'            If frmCopyBB.optShow Then
'               y = "Show"
'            Else
'               y = "Go"
'            End If
'
'            Set uDB = CopyBB(frmCopyBB.tdbLoAccount, _
'                             frmCopyBB.tdbHiAccount, _
'                             frmCopyBB.tdbLoValue, _
'                             frmCopyBB.tdbHiValue, _
'                             x, _
'                             GLCompany.SubDigits, _
'                             y)
'
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Copy Branch / Budget"
'            frmResults.lblMsg2 = ""
'            frmResults.lblMsg3 = ""
'            For i = 1 To uDB.UpperBound(1)
'                frmResults.List1.AddItem uDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'            If y = "Show" Then
'               MResp = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
'               If MResp = vbNo Then Exit Do
'            Else
'               Exit Do   ' done
'            End If
'
'         Loop
'
'      Case "CLEARGLAMOUNT"
'
'         Set frmGLUrange = New frmGLUrange
'         frmGLUrange.lblCompanyName = GLCompany.Name
'
''         frmGLUrange.lblSuspenseAcct.Visible = False
''         frmGLUrange.tdbSuspenseAcct.Visible = False
'
'         frmGLUrange.Show vbModal
'
'         If Response Then
'
'            xDB.ReDim 0, 5, 0, 0
'            xDB(1, 0) = GLCompany.Name & " Clear and Update"
'            xDB(2, 0) = "Fiscal Year: " & frmGLUrange.cmbFiscalYear & " " & _
'                        "Start Period: " & frmGLUrange.cmbStartPeriod & " " & _
'                        "End Period: " & frmGLUrange.cmbEndPeriod
'            xDB(3, 0) = " "
'            xDB(4, 0) = String(40, "=")
'            xDB(5, 0) = " "
'
'            If frmGLUrange.optClear Then
'               booX = True
'            Else
'               booX = False
'            End If
'
'            Set uDB = ClearGLAmount(frmGLUrange.cmbFiscalYear, _
'                                    frmGLUrange.cmbFiscalYear, _
'                                    frmGLUrange.StartPd, _
'                                    frmGLUrange.EndPD, _
'                                    booX)
'
'            xDBAssign
'
'
'
'
'            Set uDB = UpdateGLAmount(frmGLUrange.cmbFiscalYear, _
'                                     frmGLUrange.cmbFiscalYear, _
'                                     frmGLUrange.StartPd, _
'                                     frmGLUrange.EndPD, _
'                                     frmGLUrange.tdbSuspenseAcct, _
'                                     ID)
'
'            xDBAssign
'
'            Set uDB = MathUpdate(frmGLUrange.cmbFiscalYear, _
'                                 frmGLUrange.cmbFiscalYear, _
'                                 frmGLUrange.StartPd, _
'                                 frmGLUrange.EndPD)
'
'            xDBAssign
'
'
'TMP:
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Clear and Update GL Amounts"
'            frmResults.lblMsg2 = ""
'            frmResults.lblMsg3 = ""
'            For i = 1 To xDB.UpperBound(1)
'                frmResults.List1.AddItem xDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'
'         End If
'
'      Case "CLEARGLBUDGET"
'
'         Set frmGLUrange = New frmGLUrange
'         frmGLUrange.lblCompanyName = GLCompany.Name
'         frmGLUrange.lblSuspenseAcct.Visible = False
'         frmGLUrange.tdbSuspenseAcct.Visible = False
'         frmGLUrange.Caption = "Clear GL Budget Amounts"
'         frmGLUrange.Frame1.Visible = False
'         frmGLUrange.Show vbModal
'
'         If Response Then
'
'            Set uDB = ClearGLBudget(frmGLUrange.cmbFiscalYear, _
'                                    frmGLUrange.cmbFiscalYear, _
'                                    frmGLUrange.StartPd, _
'                                    frmGLUrange.EndPD)
'
'            Set frmResults = New frmResults
'            frmResults.lblCompanyName = GLCompany.Name
'            frmResults.lblMsg1 = "Clear GL Budget Amounts"
'            frmResults.lblMsg2 = ""
'            frmResults.lblMsg3 = ""
'            For i = 1 To uDB.UpperBound(1)
'                frmResults.List1.AddItem uDB(i, 0)
'            Next i
'            frmResults.Show vbModal
'
'         End If
'
'      Case "UPDATEBATCH"
'
''         Set frmGLUrange = New frmGLUrange
''         frmGLUrange.lblCompanyName = GLCompany.Name
'
'         ' get the batch number
''         GLBatch.OpenRS
'         If Not GLBatch.GetBatch(BatchNum) Then
'            MsgBox "Update failed! Batch #: " & BatchNum & " does not exist!", vbCritical + vbOKOnly, "Update Amounts"
'            End
'         End If
'
'         xDB.ReDim 0, 5, 0, 0
'         xDB(1, 0) = GLCompany.Name & " Clear and Update"
'         xDB(2, 0) = "Fiscal Year: " & GLBatch.FiscalYear & " " & _
'                     "Start Period: " & GLBatch.Period & " " & _
'                     "End Period: " & GLBatch.Period
'         xDB(3, 0) = " "
'         xDB(4, 0) = String(40, "=")
'         xDB(5, 0) = " "
'
'         booX = False  ' don't delete history - clear and reupdate it
'
'         Set uDB = ClearGLAmount(GLBatch.FiscalYear, _
'                                 GLBatch.FiscalYear, _
'                                 GLBatch.Period, _
'                                 GLBatch.Period, _
'                                 booX)
'
'         xDBAssign
'
'         ' create suspense account
'         Set uDB = UpdateGLAmount(GLBatch.FiscalYear, _
'                                  GLBatch.FiscalYear, _
'                                  GLBatch.Period, _
'                                  GLBatch.Period, _
'                                  0, _
'                                  ID)
'
'         xDBAssign
'
'
'         Set uDB = MathUpdate(GLBatch.FiscalYear, _
'                              GLBatch.FiscalYear, _
'                              GLBatch.Period, _
'                              GLBatch.Period)
'
'         xDBAssign
'
'         Set frmResults = New frmResults
'         frmResults.lblCompanyName = GLCompany.Name
'         frmResults.lblMsg1 = "Clear and Update GL Amounts"
'         frmResults.lblMsg2 = ""
'         frmResults.lblMsg3 = ""
'         For i = 1 To xDB.UpperBound(1)
'             frmResults.List1.AddItem xDB(i, 0)
'         Next i
'         frmResults.Show vbModal
'
'      Case "YEAREND"
'
'         ' get Account and Amount (prev year) record sets
'         GLAccount.OpenRS
'
'         ' test first p rec
'         If Not (GLAccount.GetAccount(GLCompany.FirstPAcct)) Then
'            MsgBox "First P record NOT FOUND !! " & GLCompany.FirstPAcct, vbCritical + vbOKOnly, "Year End"
'            Exit Sub
'         End If
'
'         If GLAccount.AcctType <> "P" Then
'            MsgBox "First P record wrong type: " & _
'                   GLCompany.FirstPAcct & " " & GLAccount.AcctType, vbCritical + vbOKOnly, "Year End"
'            Exit Sub
'         End If
'
'         ' test N record
'         If Not (GLAccount.GetAccount(GLCompany.NetProfitAcct)) Then
'            MsgBox "N record NOT FOUND !! " & GLCompany.NetProfitAcct, vbCritical + vbOKOnly, "Year End"
'            Exit Sub
'         End If
'
'         If GLAccount.AcctType <> "N" Then
'            MsgBox "N record wrong type: " & _
'                   GLCompany.NetProfitAcct & " " & GLAccount.AcctType, vbCritical + vbOKOnly, "Year End"
'            Exit Sub
'         End If
'
'         If GLCompany.FirstPAcct <= GLCompany.NetProfitAcct Then
'            MsgBox "First P account: " & GLCompany.FirstPAcct & _
'                   "must be greater then N account#: " & GLCompany.NetProfitAcct, _
'                   vbCritical + vbOKOnly, "Year End"
'
'            Exit Sub
'         End If
'
'         Set frmYearEnd = New frmYearEnd
'         frmYearEnd.Show vbModal
'
'         If Not Response Then End  ' exit button was selected
'
'         Set uDB = YearEnd(frmYearEnd.tdbFiscalYear, _
'                           frmYearEnd.tdbRetEarn, _
'                           frmYearEnd.tdbAcct01, _
'                           frmYearEnd.tdbAcct02, _
'                           frmYearEnd.tdbAcct03, _
'                           frmYearEnd.tdbAcct04, _
'                           frmYearEnd.tdbAcct05, _
'                           frmYearEnd.tdbAcct06, _
'                           frmYearEnd.tdbAcct07, _
'                           frmYearEnd.tdbAcct08, _
'                           frmYearEnd.tdbAcct09, _
'                           frmYearEnd.tdbAcct10)
'
'
'         frmResults.lblCompanyName = GLCompany.Name
'         frmResults.lblMsg1 = "Year End Process"
'         frmResults.lblMsg2 = ""
'         frmResults.lblMsg3 = ""
'
'         For i = 1 To uDB.UpperBound(1)
'             frmResults.List1.AddItem uDB(i, 0)
'         Next i
'         frmResults.Show vbModal
'
'   End Select
'
'   End
'
'   ' ---------------
'
'
'   For i = 1 To uDB.UpperBound(1)
'       Results.List1.AddItem uDB(i, 0)
'   Next i
'
'   Results.Show vbModal
'
'   End
'
'
''   ' start fy, end fy, start pd, end pd, susp acct, company ID#
''   Set xdb = UpdateGLAmount(2004, 2004, 1, 1, 111, 2)
''   Set xdb = DeleteAccts("Acct", 1310, 1310, "Go", True)
''   Set xdb = GLFileCopy(2, "c:\Balint\Data\Copy07.mdb", False)
''   Set xdb = ClearGLAmount(2003, 2003, 12, 12, True)
''   Set xdb = GLMultDiv(207000, 208050, "Mult", 10, "Go")
''   Set xdb = CopyBB(3000, 3999, 4, 5, "Sub", 1, "Go")
''
''   Set xDB = MathUpdate(2003, 2003, 1, 6)
''
''
''   End
''
''
''   If Not GLAccount.GetAcctRecSet(0, 0) Then
''      MsgBox "No Records"
''      End
''   End If
''
''   Set xdb = UpdateGLAmount(2003, 2003, 2, 2, 10100, _
''                        cCompany.NetProfitAcct, cCompany.FirstPAcct)
''
''
''   MsgBox xdb(1, 0) & vbCrLf & xdb(2, 0)
''
'''   Set xdb = ClearGLAmount(2003, 2003, 1, 1)
''
''
''   End
'
'
'End Sub
'
'
''Private Sub MTest()
''
''Dim mx As New XArrayDB
''
''Dim Desc As String
''
''Dim nVal As Double
''Dim LastVal As Double
''
''Dim Op As String
''Dim Amt As Currency
''
''Dim Mo As Byte
''Dim Yr As Long
''Dim StartFY As Long
''Dim EndFY As Long
''Dim StartPd As Byte
''Dim EndPd As Byte
''
''Dim i As Long
''Dim j As Long
''
''Dim SStart As Integer
''Dim SEnd As Integer
''Dim ItemCount As Integer
''
''Dim SLen As Integer
''
''   StartFY = 2003
''   EndFY = 2003
''   StartPd = 1
''   EndPd = 1
''
''   Mo = 1
''   Yr = 2003
''
''   Desc = "2180T2620"
'''   Desc = "2200DN.01"
'''  Desc = "2200"
''
''
''   x = "SELECT GLAccount.*, GLAmount.FiscalYear, " & _
''       "GLAmount.Amount01, GLAmount.Amount02, GLAmount.Amount03, " & _
''       "GLAmount.Amount04, GLAmount.Amount05, GLAmount.Amount06, " & _
''       "GLAmount.Amount07, GLAmount.Amount08, GLAmount.Amount09, " & _
''       "GLAmount.Amount10, GLAmount.Amount11, GLAmount.Amount12, GLAmount.Amount13 " & _
''       "FROM GLAccount LEFT JOIN GLAmount on " & _
''       "(GLAccount.Account = GLAmount.Account AND " & _
''       "GLAmount.FiscalYear >= " & StartFY & " AND " & _
''       "GLAmount.FiscalYear <= " & EndFY & ") ORDER BY GLAccount.Account"
''
''   rsInit x, cn, rs
''
''   If rs.BOF And rs.EOF Then
''      MsgBox "Join Failed !"
''      End
''   End If
''
''   mx.ReDim 0, 0, 1, 3
''
''   SLen = Len(Desc)
''
''   ' check start of string
''   If InStr(1, "N-.0123456789", Mid(Desc, 1, 1), vbTextCompare) = 0 Then
''      MsgBox "Bad string start: " & Mid(Desc, 1, 1)
''      End
''   End If
''
''   ' first value is a number not an account
''   If Mid(Desc, 1, 1) = "N" Then
''      mx(0, 2) = 1
''      i = 1
''   Else
''      i = 0
''   End If
''
''   x = ""
''
''   ' loop for the numbers
''   Do
''
''      i = i + 1
''
''      If i > Len(Desc) Then Exit Do
''
''      If InStr(1, "-.0123456789", Mid(Desc, i, 1), vbTextCompare) = 0 Then
''         If ItemCount <> 0 Then mx.AppendRows (1)
''         mx(ItemCount, 1) = x
''         x = ""
''         ItemCount = ItemCount + 1
''         If Mid(Desc, i + 1, 1) = "N" Then i = i + 1
''         If i >= Len(Desc) Then Exit Do
''      Else
''         x = x & Mid(Desc, i, 1)
''      End If
''
''   Loop
''
''   If x <> "" Then
''      mx.AppendRows (1)
''      mx(ItemCount, 1) = x
''      ItemCount = ItemCount + 1
''   End If
''
''   ItemCount = 0
''   i = 1
''
''   ' loop for the operators
''   Do Until i > Len(Desc)
''
''      x = Mid(Desc, i, 1)
''      If InStr(1, "ASMDT", x) <> 0 Or x = " " Then
''
''         ItemCount = ItemCount + 1
''
''         If x = " " Then
''            mx(ItemCount, 3) = "A"
''         Else
''            mx(ItemCount, 3) = x
''         End If
''
''         If Mid(Desc, i + 1, 1) = "N" Then
''            i = i + 1
''            mx(ItemCount, 2) = 1
''         Else
''            mx(ItemCount, 2) = 0
''         End If
''      End If
''
''      i = i + 1
''
''   Loop
''
''   ' first argument must be an acct # - find it
''   Amt = GetAmount(mx(0, 1), mx(0, 1), False, Mo)
''
''   ' zero out if totaling a range of accts
''   If mx.UpperBound(1) >= 1 And mx(1, 3) = "T" Then
''      Amt = 0
''   End If
''
''   LastVal = mx(0, 1)
''
''   For i = 1 To mx.UpperBound(1)
''
''       nVal = mx(i, 1)
''       Op = mx(i, 3)
''
''       If Op = "A" Then
''          If mx(i, 2) = 0 Then
''             Amt = Amt + GetAmount(nVal, nVal, False, Mo)
''          Else
''             Amt = Amt + nVal
''          End If
''       ElseIf Op = "S" Then
''          If mx(i, 2) = 0 Then
''             Amt = Amt - GetAmount(nVal, nVal, False, Mo)
''          Else
''             Amt = Amt - nVal
''          End If
''       ElseIf Op = "D" Then
''          If mx(i, 2) = 0 Then
''             Amt = Amt / GetAmount(nVal, nVal, False, Mo)
''          Else
''             Amt = Amt / nVal
''          End If
''       ElseIf Op = "M" Then
''          If mx(i, 2) = 0 Then
''             Amt = Amt * GetAmount(nVal, nVal, False, Mo)
''          Else
''             Amt = Amt * nVal
''          End If
''       ElseIf Op = "T" Then
''          If mx(i, 2) = 0 Then
''             Amt = Amt + GetAmount(LastVal, nVal, True, Mo)
''          Else
''             Amt = Amt + nVal
''          End If
''       End If
''
''       LastVal = nVal
''
''   Next i
''
''   MsgBox Format(Amt, "Currency")
''
''   End
''
''

End Sub

''
''
''Private Function GetAmount(ByVal LoAcct As Long, _
''                           ByVal HiAcct As Long, _
''                           ByVal ZeroOnly As Boolean, _
''                           ByVal Pd As Byte) As Currency
''
''Dim Acct As Long
''
''   GetAmount = 0
''
''   x = "Account = " & LoAcct
''   rs.Find x, 0, adSearchForward, 1
''
''   ' the LoAcct must be found or a zero is returned
''   If rs.EOF Then Exit Function
''
''   Do Until rs!Account > HiAcct
''
''      If ZeroOnly And rs!AcctType <> "0" Then GoTo NxtAcct
''
''      If Pd = 1 Then GetAmount = GetAmount + rs!Amount01
''      If Pd = 2 Then GetAmount = GetAmount + rs!Amount02
''      If Pd = 3 Then GetAmount = GetAmount + rs!Amount03
''      If Pd = 4 Then GetAmount = GetAmount + rs!Amount04
''      If Pd = 5 Then GetAmount = GetAmount + rs!Amount05
''      If Pd = 6 Then GetAmount = GetAmount + rs!Amount06
''      If Pd = 7 Then GetAmount = GetAmount + rs!Amount07
''      If Pd = 8 Then GetAmount = GetAmount + rs!Amount08
''      If Pd = 9 Then GetAmount = GetAmount + rs!Amount09
''      If Pd = 10 Then GetAmount = GetAmount + rs!Amount10
''      If Pd = 11 Then GetAmount = GetAmount + rs!Amount11
''      If Pd = 12 Then GetAmount = GetAmount + rs!Amount12
''      If Pd = 13 Then GetAmount = GetAmount + rs!Amount13
''
''NxtAcct:
''
''      rs.MoveNext
''      If rs.EOF Then Exit Function
''
''   Loop
''
''End Function


Private Sub xDBAssign()

Dim i As Integer
Dim j As Integer
    
    For i = 1 To uDB.UpperBound(1)
        xDB.AppendRows
        j = xDB.UpperBound(1)
        xDB(j, 0) = uDB(i, 0)
    Next i

End Sub

