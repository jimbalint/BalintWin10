VERSION 5.00
Begin VB.Form aMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "aMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As String
Dim y As String
   
Dim DBName As String
Dim CompanyID As Integer
Dim xDB As XArrayDB
   
Dim GLPrintAll As frmGLPrint
   
Dim FY, StartPd, EndPD As Integer
   
Dim flg As Boolean
   
Private Sub Form_Load()

   Set GLCompany = New cGLCompany
   Set GLAccount = New cGLAccount
   Set GLAmount = New cGLAmount
   Set GLHistory = New cGLHistory
'   Set GLDescription = New cGLDescription
   Set GLJournal = New cGLJournal
   Set GLPrint = New cGLPrint
   Set Equate = New cEquate
   Set GLBatch = New cGLBatch
   Set GLUser = New cGLUser
   
   
'   SetEquates
'
'   x = Command
'
'   If x = "" Then       ' nothing on the command line
''      x = "16/GLHistJnl/C:\Balint\Data\GLSystem.mdb/JimBo"
''      x = "16/ChartOfAccts/C:\Balint\Data\GLSystem.mdb/JimBo"
''      x = "16/PrintDesc/C:\Balint\Data\GLSystem.mdb/JimBo"
''      x = "16/PrintGLAccount/C:\Balint\Data\GLSystem.mdb/JimBo"
'
'' ESCOTT
''      x = "65/GLHistJnl/S:\Balint\Data\GLSystem.mdb/jim"
'
'      x = "65/GLHistJnl/C:\Balint\Data\GLSystem.mdb/jim"
'
'
'
'      x = "59/golf/GLHistJnl/C:\Balint\Data\GLSystem.mdb/jim"
'      x = "59/golf/ChartOfAccts/C:\Balint\Data\GLSystem.mdb/jim"
'      x = "59/golf/DetailGL/C:\Balint\Data\GLSystem.mdb/jim"
'      x = "59/golf/PrintGLAccount/C:\Balint\Data\GLSystem.mdb/jim"
'      x = "59/golf/PrintDesc/C:\Balint\Data\GLSystem.mdb/jim"
'
'   End If
'
'   If cmdline(x, ID, Password, Prog, SysFile, User, BatchNum) = False Then
'      MsgBox "Bad command line !!!"
'   End If
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
'
'   CNOpen DBName, Password
'
'   Prog = StrConv(Prog, vbUpperCase)
'
'   GLPrint.GetData User, Flg
'
'   ' new GLPrint record was created for the user
'   ' load defaults from the GLCompany file
'   If Flg = True Then
'      GLPrint.LowAccount = 1
'      GLPrint.HiAccount = 999999999
'      GLPrint.LowBranchAcct = GLCompany.LowBranch
'      GLPrint.HiBranchAcct = GLCompany.HiBranch
'      GLPrint.LowConsAcct = GLCompany.LowConsolidated
'      GLPrint.HiConsAcct = GLCompany.HiConsolidated
'      GLPrint.Save (Equate.RecPut)
'   End If
'
'' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'' GLHistJnl 2004, 8, 10, 0, 0
'' Prvw.vsp.EndDoc
'' Prvw.Show vbModal
'' End
'' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'   ' initialize the GLPrint screen
'   Set GLPrintAll = New frmGLPrint
'   GLPrintAll.lblCompName = GLCompany.Name
'   GLPrintAll.Caption = "Print Set Up  " & User
'
'   Select Case Prog
'
'      Case "CHARTOFACCTS"
'         GLPrintAll.lblProgName = "Chart Of Accounts"
'         GLPrintAll.optBranch.Enabled = False
'         GLPrintAll.optBudget.Enabled = False
'         GLPrintAll.cmdOptions.Enabled = False
'         DisableFrame2
'         DisableFrame3
'         DisableFrame4
'         DisableJS
'         DisableBranchRange
'         DisableConsRange
'         DisableFY
'         DisablePdRange
'         GLPrintAll.Show vbModal
'         If Response Then
'            If GLPrintAll.optRegular Then
'               x = "Reg"
'            Else
'               x = "Cons"
'            End If
'            ChartOfAccts GLPrintAll.txtLoAccount, _
'                         GLPrintAll.txtHiAccount, _
'                         x, _
'                         GLCompany.SubDigits
'         End If
'
'      Case "PRINTDESC"
'         GLPrintAll.lblProgName = "Print Description File"
'         GLPrintAll.optBranch.Enabled = False
'         GLPrintAll.optBudget.Enabled = False
'         GLPrintAll.cmdOptions.Enabled = False
'         DisableFrame1
'         DisableFrame2
'         DisableFrame3
'         DisableFrame4
'         DisableJS
'         DisableBranchRange
'         DisableConsRange
'         DisableFY
'         DisablePdRange
'
'         GLPrintAll.lblLoAccount = "Low Desc Number:"
'         GLPrintAll.lblHiAccount = "Hi Desc Number:"
'
'         GLPrintAll.Show vbModal
'
'         If Response Then
'
'            PrintDesc GLPrintAll.txtLoAccount, GLPrintAll.txtHiAccount
'
'         End If
'
'      Case "PRINTGLACCOUNT"
'
'         GLPrintAll.lblProgName = "Print GL Account File"
'         DisableFrame1
'         DisableFrame2
'         DisableFrame3
'         DisableFrame4
'         DisableJS
'         DisableFY
'         DisablePdRange
'         DisableConsRange
'
'         GLPrintAll.Show vbModal
'
'         If Response Then
'
'            PrintGLAccount GLPrintAll.txtLoAccount, _
'                           GLPrintAll.txtHiAccount, _
'                           GLPrintAll.txtLoCons, _
'                           GLPrintAll.txtHiCons, _
'                           GLPrintAll.txtLoBranch, _
'                           GLPrintAll.txtHiBranch, _
'                           GLCompany.SubDigits
'
'         End If
'
'      Case "GLHISTJNL"
'
'         If BatchNum = 0 Then
'
'            GLPrintAll.AllJnlOption = True
'            GLPrintAll.lblProgName = "Data Entry Journal"
'            GLPrintAll.cmdOptions.Enabled = False
'            DisableFrame1
'            DisableFrame2
'            DisableFrame3
'            DisableFrame4
'            DisableAccountRange
'            DisableBranchRange
'            DisableConsRange
'            GLPrintAll.Show vbModal
'
'            If Response Then
'               GLHistJnl GLPrintAll.FiscalYear, _
'                         GLPrintAll.StartPd, _
'                         GLPrintAll.EndPD, _
'                         GLPrintAll.JournalSource, _
'                         0
'            End If
'
'         Else
'
'            If Not GLBatch.GetBatch(BatchNum) Then
'               MsgBox "Batch Not Found !!! " & GLBatch.BatchNumber
'               End
'            End If
'
'            Response = True
'
'            GLHistJnl GLBatch.FiscalYear, _
'                      GLBatch.Period, _
'                      GLBatch.Period, _
'                      GLBatch.JournalSource, _
'                      BatchNum
'
'         End If
'
'      Case "DETAILGL"
'         GLPrintAll.lblProgName = "Detail General Ledger"
'
'         DisableFrame2
'         DisableFrame3
'         DisableFrame4
'
'         GLPrintAll.optBudget.Enabled = False
'
'         GLPrintAll.cmbJournalSource.Visible = False
'         GLPrintAll.lblJournalSource.Visible = False
'
'         GLPrintAll.Show vbModal
'
'         If GLPrint.RegBraCon = Equate.Regular Then x = "Reg"
'         If GLPrint.RegBraCon = Equate.Branch Then x = "Bra"
'         If GLPrint.RegBraCon = Equate.Consol Then x = "Con"
'
'         If Response Then
'
'            DetailGL x, _
'                     GLPrintAll.FiscalYear, _
'                     GLPrintAll.StartPd, _
'                     GLPrintAll.EndPD, _
'                     GLPrintAll.txtLoAccount, _
'                     GLPrintAll.txtHiAccount, _
'                     GLPrintAll.txtLoCons, _
'                     GLPrintAll.txtHiCons, _
'                     GLPrintAll.txtLoBranch, _
'                     GLPrintAll.txtHiBranch, _
'                     GLCompany.SubDigits, _
'                     GLPrint.SepPage, _
'                     GLPrint.PrtZeroBal, _
'                     CompanyID
'
'         End If
'   End Select
'
'   ' show the preview screen
'   If Response Then
'      Prvw.vsp.EndDoc
'      Prvw.Show vbModal
'   End If
'
'   End
'
'
''   PrintGLAccount 0, 0, 0, 0, 0, 0, 0
''   PrintGLAccount 0, 0, 0, 0, 3, 3, 1
''   DetailGL "Reg", 2005, 1, 3, 0, 0, 0, 0, 0, 0, 0, False
''   GLHistJnl 2004, 1, 1, 0, 0
'
'
'   End
'
End Sub


Private Sub DisableFrame1()

    GLPrintAll.fraType1.Enabled = False
    GLPrintAll.optNormal.Enabled = False
    GLPrintAll.optBranch.Enabled = False
    GLPrintAll.optConsolidated.Enabled = False
    GLPrintAll.optBudget.Enabled = False

End Sub

Private Sub DisableFrame2()

    GLPrintAll.fraType2.Enabled = False
    GLPrintAll.optStatements.Enabled = False
    GLPrintAll.optSchedules.Enabled = False

End Sub

Private Sub DisableFrame3()

    GLPrintAll.fraType3.Enabled = False
    GLPrintAll.optRegular.Enabled = False
    GLPrintAll.optComparative.Enabled = False

End Sub

Private Sub DisableFrame4()

    GLPrintAll.fraType4.Enabled = False
    GLPrintAll.optBoth.Enabled = False
    GLPrintAll.optBalanceSheet.Enabled = False
    GLPrintAll.optIncomeStatement.Enabled = False

End Sub

Private Sub DisableAccountRange()

    GLPrintAll.txtLoAccount.Enabled = False
    GLPrintAll.txtHiAccount.Enabled = False
    GLPrintAll.lblLoAccount.Enabled = False
    GLPrintAll.lblHiAccount.Enabled = False
    
End Sub

Private Sub DisableBranchRange()

    GLPrintAll.txtLoBranch.Enabled = False
    GLPrintAll.txtHiBranch.Enabled = False
    GLPrintAll.lblLoBranch.Enabled = False
    GLPrintAll.lblHiBranch.Enabled = False
    
End Sub

Private Sub DisableConsRange()

    GLPrintAll.txtLoCons.Enabled = False
    GLPrintAll.txtHiCons.Enabled = False
    GLPrintAll.lblLoCons.Enabled = False
    GLPrintAll.lblHiCons.Enabled = False
    
End Sub

Private Sub DisableJS()

    GLPrintAll.cmbJournalSource.Visible = False
    GLPrintAll.lblJournalSource.Enabled = False

End Sub

Private Sub DisablePdRange()

    GLPrintAll.cmbStartPeriod.Visible = False
    GLPrintAll.cmbEndPeriod.Visible = False
    GLPrintAll.lblStartPd.Enabled = False
    GLPrintAll.lblEndPeriod.Enabled = False

End Sub

Private Sub DisableFY()

    GLPrintAll.lblFY.Enabled = False
    GLPrintAll.cmbFiscalYear.Visible = False

End Sub
