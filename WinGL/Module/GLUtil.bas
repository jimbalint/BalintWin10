Attribute VB_Name = "modGLUtil"
Option Explicit

Dim Count1 As Long
Dim Count2 As Long
Dim SQLString As String
Dim xDB As New XArrayDB

Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset

Const RecAdd As Boolean = True
Const RecPut As Boolean = False

Dim x As String
Dim Y As String

Dim i As Long
Dim j As Long

Dim Acct As Long
Dim nPlaces As Integer
Dim Amt As Currency

Dim rsFlg As Boolean
Dim rsflg2 As Boolean

' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'  variables for MathUpdate
Dim FirstM As Long
Dim FirstN As Long
Dim FirstP As Long
Dim FirstL As Long
Dim CurFormat As String

Dim yr As Long
Dim Mo As Long

Dim Amount(13) As Currency
Dim BudAmount(13) As Currency

Dim SignMode As Integer
Dim GAcct As Long
Dim LAcct As Long
Dim Desc As String

Dim G(10) As Currency
Dim BG(10) As Currency

Dim Ct As Long

' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Public Function ClearGLAmount(ByVal StartFY As Long, _
                              ByVal EndFY As Long, _
                              ByVal StartPd As Byte, _
                              ByVal EndPD As Byte, _
                              ByVal DeleteHist As Boolean) As XArrayDB
                              
   
   frmProgress.lblMsg1 = "Clearing Totals for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Now Gathering Information ... "
   frmProgress.Show
   Ct = 0
   
   rsInit "SELECT * FROM GLAmount WHERE GLAmount.FiscalYear >= " & StartFY & _
          " AND GLAmount.FiscalYear <= " & EndFY, _
          cn, rs
               
   xDB.ReDim 0, 2, 0, 0
   
   ' no records found
   If rs.BOF And rs.EOF Then
      xDB(1, 0) = Stamp("Error!!! ")
      xDB(2, 0) = "No amount records found for the FY range: " & _
                  CStr(StartFY) & " to: " & CStr(EndFY)
      Set ClearGLAmount = xDB
      Exit Function
   End If
   
   rs.MoveFirst
   Do Until rs.EOF
      
      Ct = Ct + 1
      If Ct = 1 Or Ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & rs!Account
         frmProgress.lblMsg2.Refresh
      End If
      
      If 1 >= StartPd And 1 <= EndPD Then rs!Amount01 = 0
      If 2 >= StartPd And 2 <= EndPD Then rs!Amount02 = 0
      If 3 >= StartPd And 3 <= EndPD Then rs!Amount03 = 0
      If 4 >= StartPd And 4 <= EndPD Then rs!Amount04 = 0
      If 5 >= StartPd And 5 <= EndPD Then rs!Amount05 = 0
      If 6 >= StartPd And 6 <= EndPD Then rs!Amount06 = 0
      If 7 >= StartPd And 7 <= EndPD Then rs!Amount07 = 0
      If 8 >= StartPd And 8 <= EndPD Then rs!Amount08 = 0
      If 9 >= StartPd And 9 <= EndPD Then rs!Amount09 = 0
      If 10 >= StartPd And 10 <= EndPD Then rs!Amount10 = 0
      If 11 >= StartPd And 11 <= EndPD Then rs!Amount11 = 0
      If 12 >= StartPd And 12 <= EndPD Then rs!Amount12 = 0
      If 13 >= StartPd And 13 <= EndPD Then rs!Amount13 = 0
      Count1 = Count1 + 1
      rs.Update
      rs.MoveNext
   Loop
   
   xDB(1, 0) = Stamp("Clear GLAmount started: ")
   xDB(2, 0) = Count1 & " AMOUNT records cleared for FY " & StartFY & " to: " & EndFY & _
            PeriodName(StartFY, StartPd, GLCompany.FirstPeriod, GLCompany.NumberPds) _
            & " to: " & _
            PeriodName(EndFY, EndPD, GLCompany.FirstPeriod, GLCompany.NumberPds)

   If DeleteHist Then
      
      frmProgress.lblMsg2 = "Now DELETING History Records ..."
      frmProgress.lblMsg2.Refresh
      
      x = "DELETE * FROM GLHistory WHERE FiscalYear >= " & StartFY & _
          " AND FiscalYear <= " & EndFY & " AND Period >= " & StartPd & _
          " AND Period <= " & EndPD
      rsInit x, cn, rs
       
      xAddRow xDB, "GLHistory records from " & StartFY & " Pd " & StartPd & _
                  " To " & EndFY & " Pd " & EndPD & " have been deleted."
   
      x = "DELETE * FROM GLBatch WHERE FiscalYear >= " & StartFY & _
          " AND FiscalYear <= " & EndFY & " AND Period >= " & StartPd & _
          " AND Period <= " & EndPD
      rsInit x, cn, rs
   
      ' scan for GLAmount records where there are
      ' no history for the FY
      ' 2015-12-12
      Dim rsFY As ADODB.Recordset
      x = " SELECT DISTINCT(FiscalYear) as FY " & _
          " FROM GLAmount " & _
          " WHERE FiscalYear BETWEEN " & StartFY & " AND " & EndFY
      rsInit x, cn, rsFY
      rsFY.MoveFirst
      Do Until rsFY.EOF
            x = " SELECT COUNT(1) as GLHCount " & _
                " FROM GLHistory " & _
                " WHERE FiscalYear = " & rsFY!FY
            Dim rsGLH As ADODB.Recordset
            rsInit x, cn, rsGLH
            Do Until rsGLH.EOF
                If rsGLH!GLHCount = 0 Then
                    Dim gResp As Integer
                    gResp = MsgBox("OK to delete ALL GL data for FY: " & rsFY!FY & "?", vbExclamation + vbOKCancel)
                    If gResp = vbOK Then
                        x = " DELETE * FROM GLAmount WHERE FiscalYear = " & rsFY!FY
                        cn.Execute x
                    End If
                End If
                rsGLH.MoveNext
            Loop
            rsFY.MoveNext
      Loop
   
   Else
      rs.Close
   End If
   
   frmProgress.Hide
   
   Set ClearGLAmount = xDB
   Set rs = Nothing

End Function

Public Function ClearGLBudget(ByVal StartFY As Long, _
                              ByVal EndFY As Long, _
                              ByVal StartPd As Byte, _
                              ByVal EndPD As Byte) As XArrayDB
                              
   x = "SELECT * FROM GLAmount WHERE GLAmount.FiscalYear >= " & StartFY & _
          " AND GLAmount.FiscalYear <= " & EndFY

   rsInit x, cn, rs
               
   xDB.ReDim 0, 2, 0, 0
   
   ' no records found
   If rs.BOF And rs.EOF Then
      xDB(1, 0) = Stamp("Error!!! ")
      xDB(2, 0) = "No amount records found for the FY range: " & _
                  CStr(StartFY) & " to: " & CStr(EndFY)
      Set ClearGLBudget = xDB
      Exit Function
   End If
   
   rs.MoveFirst
   Do Until rs.EOF
      
      ' count the actual number of accounts cleared
      If 1 >= StartPd And 1 <= EndPD And rs!Budget01 <> 0 Then Count1 = Count1 + 1
      If 2 >= StartPd And 2 <= EndPD And rs!Budget02 <> 0 Then Count1 = Count1 + 1
      If 3 >= StartPd And 3 <= EndPD And rs!Budget03 <> 0 Then Count1 = Count1 + 1
      If 4 >= StartPd And 4 <= EndPD And rs!Budget04 <> 0 Then Count1 = Count1 + 1
      If 5 >= StartPd And 5 <= EndPD And rs!Budget05 <> 0 Then Count1 = Count1 + 1
      If 6 >= StartPd And 6 <= EndPD And rs!Budget06 <> 0 Then Count1 = Count1 + 1
      If 7 >= StartPd And 7 <= EndPD And rs!Budget07 <> 0 Then Count1 = Count1 + 1
      If 8 >= StartPd And 8 <= EndPD And rs!Budget08 <> 0 Then Count1 = Count1 + 1
      If 9 >= StartPd And 9 <= EndPD And rs!Budget09 <> 0 Then Count1 = Count1 + 1
      If 10 >= StartPd And 10 <= EndPD And rs!Budget10 <> 0 Then Count1 = Count1 + 1
      If 11 >= StartPd And 11 <= EndPD And rs!Budget11 <> 0 Then Count1 = Count1 + 1
      If 12 >= StartPd And 12 <= EndPD And rs!Budget12 <> 0 Then Count1 = Count1 + 1
      If 13 >= StartPd And 13 <= EndPD And rs!Budget13 <> 0 Then Count1 = Count1 + 1
      
      If 1 >= StartPd And 1 <= EndPD Then rs!Budget01 = 0
      If 2 >= StartPd And 2 <= EndPD Then rs!Budget02 = 0
      If 3 >= StartPd And 3 <= EndPD Then rs!Budget03 = 0
      If 4 >= StartPd And 4 <= EndPD Then rs!Budget04 = 0
      If 5 >= StartPd And 5 <= EndPD Then rs!Budget05 = 0
      If 6 >= StartPd And 6 <= EndPD Then rs!Budget06 = 0
      If 7 >= StartPd And 7 <= EndPD Then rs!Budget07 = 0
      If 8 >= StartPd And 8 <= EndPD Then rs!Budget08 = 0
      If 9 >= StartPd And 9 <= EndPD Then rs!Budget09 = 0
      If 10 >= StartPd And 10 <= EndPD Then rs!Budget10 = 0
      If 11 >= StartPd And 11 <= EndPD Then rs!Budget11 = 0
      If 12 >= StartPd And 12 <= EndPD Then rs!Budget12 = 0
      If 13 >= StartPd And 13 <= EndPD Then rs!Budget13 = 0
      
      rs.Update
      rs.MoveNext
   Loop
   
   xDB(1, 0) = Stamp("OK")
   xDB(2, 0) = Count1 & " BUDGET records cleared for FY " & StartFY & " to: " & EndFY & _
               PeriodName(StartFY, StartPd, GLCompany.FirstPeriod, GLCompany.NumberPds) & _
               " To: " & _
               PeriodName(EndFY, EndPD, GLCompany.FirstPeriod, GLCompany.NumberPds)

   Set ClearGLBudget = xDB

   rs.Close
   Set rs = Nothing

End Function

Public Function UpdateGLAmount(ByVal StartFY As Long, _
                               ByVal EndFY As Long, _
                               ByVal StartPd As Byte, _
                               ByVal EndPD As Byte, _
                               ByVal SuspenseAcct As Long, _
                               ByVal ID As Long) As XArrayDB
                               
                               
Dim GLAccount As New cGLAccount
Dim GLHistory As New cGLHistory
Dim GLAmount As New cGLAmount
Dim GLCompany As New cGLCompany

Dim Equate As New cEquate

Dim xdb2 As XArrayDB
Dim TotalCredits As Currency
Dim TotalDebits As Currency
Dim NetPL As Currency
Dim TotalSusp As Currency

Dim RecCount As Long
Dim SuspCount As Long
Dim RSPos As Long
                               
Dim yr As Long
Dim Pd As Byte
Dim gFlg As Boolean
                               
Dim AmtFormat As String
                               
AmtFormat = "###,###,###,##0.00"
   
   
   Ct = 0
   frmProgress.lblMsg1 = "Now Updating Amounts for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Gathering Information .... "
   frmProgress.Show
   
   ' init xDB
   xDB.ReDim 0, 2, 0, 0
   xDB(1, 0) = " "
   xDB(2, 0) = Stamp("Amount update started")
   
   ' get the company info
   If Not GLCompany.GetData(ID) Then
      xAddRow xDB, Stamp("Company record NF: " & ID)
      xAddRow xDB, "Update not completed !!!"
      Set UpdateGLAmount = xDB
      Exit Function
   End If
                               
   ' verify GLCompany info
   If GLCompany.NetProfitAcct <> 0 And GLCompany.FirstPAcct <= GLCompany.NetProfitAcct Then
      xAddRow xDB, Stamp("Error: First P Rec: " & GLCompany.FirstPAcct & _
                        " First N Rec: " & GLCompany.NetProfitAcct)
      xAddRow xDB, "Update not completed !!!"
      Set UpdateGLAmount = xDB
      Exit Function
   End If

'   ' clear first
'   Set xDB = ClearGLAmount(StartFY, EndFY, StartPd, EndPD, False)
   
   ' get the GLAmount record set
   If Not GLAmount.GetRecordSet(0, 0, StartFY, EndFY) Then
'      MsgBox "Amount recs will be added " & StartFY & " " & EndFY
   End If
   
   ' get the GLAccount record set
   GLAccount.GetAllAccounts
   
   ' verify the suspense acct
   If SuspenseAcct <> 0 Then
      If Not GLAccount.Find(SuspenseAcct) Then
         xAddRow xDB, " "
         xAddRow xDB, "Suspense acct# not found: " & SuspenseAcct
         Set UpdateGLAmount = xDB
         GLAccount.CloseRS
         Exit Function
      ElseIf GLAccount.AcctType <> "0" Then
         xAddRow xDB, " "
         xAddRow xDB, "Suspense acct# " & SuspenseAcct & " Wrong Type: " & GLAccount.AcctType
         Set UpdateGLAmount = xDB
         GLAccount.CloseRS
         Exit Function
      End If
   End If
   
   ' get history record set
   For yr = StartFY To EndFY
       For Pd = StartPd To EndPD
       
           ' get history record set
           If Not GLHistory.QueryByFiscalYearByPeriod(yr, Pd) Then GoTo NextPd
                         
           ' comment lines
           xAddRow xDB, " "
           xAddRow xDB, Stamp("History Fiscal Year: " & yr & " Pd: " & Pd)
                         
           Ct = 0
                         
           ' loop thru the record set
           Do
              
              Ct = Ct + 1
              If Ct = 1 Or Ct Mod 10 = 0 Then
                 frmProgress.lblMsg2 = "History Period: " & yr & "/" & Pd & " Record: " & Format(Ct, "##,###,##0")
                 frmProgress.lblMsg2.Refresh
              End If
              
              ' 1440
              If GLHistory.Account < 0 Then GoTo NxtHist
              
' >>>>>>>>>> !!! fix GLImport !!!
              ' 1460
'              If Not InStr(1, "NORTW", GLHistory.HisType, vbTextCompare) Then
'                 GoTo NxtHist
'              End If
' >>>>>>>>>>
           
              ' suspend ?
              If GLHistory.Account = 0 Then
                 
                 If SuspenseAcct = 0 Then GoTo NxtHist ' skip it
                 ' If Not GLAccount.QueryByAccount(SuspenseAcct) Then GoTo NxtHist
                 If Not GLAccount.GetAccount(SuspenseAcct) Then GoTo NxtHist
                 TotalSusp = TotalSusp + GLHistory.Amount
                 xAddRow xDB, "Placed in suspense: " & GLHistory.Account & " " & Format(GLHistory.Amount)
              
              ElseIf Not GLAccount.Find(GLHistory.Account) Then
                 
                 If SuspenseAcct = 0 Then
                    GLAccount.Clear
                    GLAccount.Account = GLHistory.Account
                    GLAccount.Description = "New Account"
                    GLAccount.AcctType = "0"
                    GLAccount.Save (RecAdd)
                    xAddRow xDB, "New Account ---> " & GLHistory.Account
                 Else
                    If Not GLAccount.Find(SuspenseAcct) Then GoTo NxtHist
                    TotalSusp = TotalSusp + GLHistory.Amount
                    xAddRow xDB, "Placed in suspense: " & GLHistory.Account & " " & Format(GLHistory.Amount)
                 End If
              Else
        
'                  Debug.Print "GLAccount found: " & GLAccount.Account
              
              End If
           
              ' update it - create glamount if dne
              RSPos = GLAmount.sFind(GLAccount.Account, yr, RSPos)
              If RSPos = 0 Then
                 GLAmount.Clear
                 GLAmount.Account = GLAccount.Account
                 GLAmount.FiscalYear = yr
                 GLAmount.Save (RecAdd)
              Else
'                 MsgBox "GLAmount found: " & GLAccount.Account & " " & Yr
              End If
              
              ' add the history amount to the GLAmount Bucket
              If GLHistory.HisType <> "B" Then
                 GLAmount.AddAmount Pd, GLHistory.Amount
              Else
                 GLAmount.AddBudgAmount Pd, GLHistory.Amount
              End If
              
              GLAmount.Save (RecPut)
           
              ' update totals  2370  2380
              If GLHistory.Amount > 0 And GLHistory.HisType <> "B" Then
                 TotalDebits = TotalDebits + GLHistory.Amount
              Else
                 TotalCredits = TotalCredits + GLHistory.Amount
              End If
           
              ' 2390
              If GLHistory.Account > GLCompany.FirstPAcct And GLHistory.HisType <> "B" Then
                 NetPL = NetPL + GLHistory.Amount
              End If
              
              RecCount = RecCount + 1
              
NxtHist:
              If Not GLHistory.GetNext Then Exit Do
              
           Loop
           
           ' update N record   2410
           If GLAmount.sFind(GLCompany.NetProfitAcct, yr, 0) = 0 Then
              GLAmount.Clear
              GLAmount.Account = GLCompany.NetProfitAcct
              GLAmount.FiscalYear = yr
              GLAmount.Save (RecAdd)
           End If
           
           GLAmount.AddAmount Pd, NetPL
           GLAmount.Save (RecPut)
   
           GLHistory.CloseRS
   
           ' update the comment log
           xAddRow xDB, "Debit Amount: " & vbTab & vbTab & Format(TotalDebits, AmtFormat)
           xAddRow xDB, "Credit Amount: " & vbTab & vbTab & Format(TotalCredits, AmtFormat)
           xAddRow xDB, "Check Balance: " & vbTab & vbTab & Format(TotalDebits + TotalCredits, AmtFormat)
           xAddRow xDB, "Net Profit- Loss" & vbTab & vbTab & Format(NetPL, AmtFormat)
           
           If TotalSusp <> 0 Then
              xAddRow xDB, "Placed in Suspense: " & Format(TotalSusp, AmtFormat)
           End If
           
           xAddRow xDB, "Number of GLHistory records updated: " & Format(RecCount, "###,###,##0")
   
           NetPL = 0
   
NextPd:
       Next Pd
   
       ' clear the totals
       TotalDebits = 0
       TotalCredits = 0
       NetPL = 0
       TotalSusp = 0
       RecCount = 0
   
   Next yr
      
   frmProgress.Hide
      
   Set UpdateGLAmount = xDB

End Function

Public Function DeleteAccts(ByVal AcctSub As String, _
                            ByVal LoValue As Long, _
                            ByVal HiValue As Long, _
                            ByVal ShowGo As String, _
                            ByVal DelHistAmt As Boolean) As XArrayDB
                            
Dim GLAccount As New cGLAccount
Dim i As Integer
Dim x As String
                               
                               
   If ShowGo = "Show" Then
      frmProgress.lblMsg1 = "Displaying Account Deletes for: " & GLCompany.Name
   Else
      frmProgress.lblMsg1 = "Performing Account Deletes for: " & GLCompany.Name
   End If
   frmProgress.lblMsg2 = "Gathering Information ... "
   frmProgress.Show
   Ct = 0
                               
   xDB.ReDim 0, 0, 0, 0
                               
   If AcctSub <> "Acct" And AcctSub <> "Sub" Then
      xAddRow xDB, "Bad Acct / Sub parameter !!! " & AcctSub
      Set DeleteAccts = xDB
      Exit Function
   End If
   
   If ShowGo <> "Show" And ShowGo <> "Go" Then
      xAddRow xDB, "Bad Show / Go parameter !!! " & ShowGo
      Set DeleteAccts = xDB
      Exit Function
   End If
                               
   ' calculate number of places for sub
   If AcctSub = "Sub" Then
      nPlaces = DecPlaces(LoValue)
   End If
                               
   GLAccount.OpenRS
   
   Do
      
      Ct = Ct + 1
      If Ct = 1 Or Ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
         frmProgress.lblMsg2.Refresh
      End If
      
      If AcctSub = "Acct" Then
         If GLAccount.Account < LoValue Or GLAccount.Account > HiValue Then
            GoTo NextAcct
         End If
      Else             ' by sub acct
         If GLAccount.Account Mod 10 ^ nPlaces < LoValue Then GoTo NextAcct
         If GLAccount.Account Mod 10 ^ nPlaces > HiValue Then GoTo NextAcct
      End If
      
'      x = GLAccount.Account & " " & AcctDes(GLAccount.Account)
      
      x = GLAccount.Account & " " & GLAccount.Description
      
      Acct = GLAccount.Account
      
      ' delete it
      If ShowGo = "Go" Then
         If GLAccount.DeleteCurrentRecord Then
            x = x & " Deleted"
         
            If DelHistAmt Then
               
               Y = "DELETE * FROM GLHistory WHERE Account = " & Acct
               rsInit Y, cn, rs2
               Set rs2 = Nothing
               
               Y = "DELETE * FROM GLAmount WHERE Account = " & Acct
               rsInit Y, cn, rs2
               Set rs2 = Nothing
               
            End If
         
         Else
            x = x & " Error!"
         End If
      End If
   
      xAddRow xDB, x
   
NextAcct:
      If Not GLAccount.GetNextAcct Then Exit Do
   
   Loop
                            
   frmProgress.Hide
                            
   Set DeleteAccts = xDB
   GLAccount.CloseRS
                            
End Function

Public Sub GLFileCopy(ByVal CopyFrom As String, ByVal CopyTo As String)

Dim x As String
Dim OldName As String
Dim FileName As String
Dim PRCompanyID As Long
Dim ChkFileName As String
Dim Pos As Integer
Dim xCopyTo As String
        
    ' check to see if the target file already exists
    On Error Resume Next
    GetAttr (CopyTo)
    If Err.Number = 0 Then
        MsgBox FileName & vbCr & "Already exists!", vbExclamation, "GL File Copy"
        Exit Sub
    End If
   
    On Error GoTo 0
   
    ' check to see if the file name already exists in GLCompany
'    If GLCompany.GetByName(FileName) Then
'        MsgBox FileName & vbCr & "Already exists in the GLCompany file!", vbExclamation, "GL File Copy"
'        Exit Sub
'    End If
   
    ' copy the .mdb file
    On Error Resume Next
   
    ' store the copy from name
'    x = Left(App.Path, 1) & Mid(GLCompany.FileName, 2, Len(GLCompany.FileName) - 1)
   
    ' close the connection
    cn.Close
    cnDes.Close
   
    FileCopy CopyFrom, CopyTo
   
    If Err.Number <> 0 Then
        MsgBox "File copy FAILED !!! " & vbCr & Err.Description & " " & Trim(x) & " " & Trim(FileName), vbExclamation
        Exit Sub
    Else
        On Error GoTo 0
    End If
   
    ' re-open the connection to the NEW file
    CNDesOpen (SysFile)

    ' re-get the original company record
    If GLCompany.GetData(GLUser.LastCompany) = False Then
        MsgBox "Re-get error?", vbExclamation, "GL File Copy"
        Exit Sub
    End If

    ' get the original PRCompany record
    SQLString = "SELECT * FROM PRCompany WHERE GLCompanyID = " & GLCompany.ID
    If PRCompany.GetBySQL(SQLString) = True Then
        PRCompanyID = PRCompany.CompanyID
    Else
        PRCompanyID = 0
    End If
    
'    If PRCompany.GetByFileName(GLCompany.FileName) Then
'        PRCompanyID = PRCompany.CompanyID
'    Else
'        PRCompanyID = 0
'    End If

    ' if BalintFolder used - make drive letter "X"
    If BalintFolder <> "" Then
        Pos = InStrRev(CopyTo, "\", Len(CopyTo), vbTextCompare)
        xCopyTo = "X:\Balint\Data\" & Mid(CopyTo, Pos + 1, Len(CopyTo) - Pos)
    Else
        xCopyTo = CopyTo
    End If

    ' make new GLCompany record
    
    GLCompany.FileName = xCopyTo
    
    GLCompany.Name = frmCopy.txtCompName
    If Not GLCompany.Save(Equate.RecAdd) Then
        MsgBox "New GLCompany save failed!", vbExclamation, "GL File Copy"
        Exit Sub
    End If
   
    ' make new PRCompany record
    ' If PRCompanyID <> 0 Then
    PRCompany.FileName = xCopyTo
    PRCompany.Name = frmCopy.txtCompName
    PRCompany.GLCompanyID = GLCompany.ID
    PRCompany.Save (Equate.RecAdd)
    ' End If
   
    ' delete the GL Info
    If Not CNOpen(CopyTo, dbPwd) Then
        MsgBox "Copy complete - clear failed !!!", vbExclamation, "GL File Copy"
        Exit Sub
    End If
    
    ' remove the GL Info ?
    If frmCopy.chkCopyGL = 0 Then
        
        SQLString = "DELETE * FROM GLAccount"
        cn.Execute SQLString
        
        SQLString = "DELETE * FROM GLAmount"
        cn.Execute SQLString
        
        SQLString = "DELETE * FROM GLBatch"
        cn.Execute SQLString
        
        SQLString = "DELETE * FROM GLBranch"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM GLColumn"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM GLHistory"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM GLJournal"
        cn.Execute SQLString
   
    Else        ' clear history and amounts?
    
        If frmCopy.chkClearGL Then
            GLAmount.DeleteAll
            GLHistory.DeleteAll
            GLBatch.DeleteAll
        End If
    
    End If
   
    If frmCopy.chkCopyPR = 0 Then
        
        SQLString = "DELETE * FROM PRAdjust"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRBatch"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRDepartment"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRDist"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PREELists"
        cn.Execute SQLString
        
        SQLString = "DELETE * FROM PREmployee"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRGLUpd"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRHist"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRItem"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRItemHist"
        cn.Execute SQLString
   
        SQLString = "DELETE * FROM PRW2"
        cn.Execute SQLString
    
        SQLString = "DELETE * FROM PRW2City"
        cn.Execute SQLString
    
        SQLString = "DELETE * FROM PRW2State"
        cn.Execute SQLString
    
    Else
   
        ' clear amounts?
        If frmCopy.chkClearPR Then
            
            SQLString = "DELETE * FROM PRBatch"
            cn.Execute SQLString
            
            SQLString = "DELETE * FROM PRDist"
            cn.Execute SQLString
            
            SQLString = "DELETE * FROM PRHist"
            cn.Execute SQLString
            
            SQLString = "DELETE * FROM PRItemHist"
            cn.Execute SQLString
            
        End If
    
    End If
   
    ' copy the laser check setup if it exists
    ' *** for Eaglowski only ***
    On Error Resume Next
    If BalintFolder = "" Then
        ChkFileName = Left(App.Path, 1) & ":\Balint\Data\PRCKeag" & Format(PRCompanyID, "000000") & ".mdb"
    Else
        ChkFileName = BalintFolder & "\Data\PRCKeag" & Format(PRCompanyID, "000000") & ".mdb"
    End If
    
    GetAttr (ChkFileName)
    If Err.Number = 0 Then      ' go ahead and copy it
        If BalintFolder = "" Then
            x = Left(App.Path, 1) & ":\Balint\Data\PRCKeag" & Format(PRCompany.CompanyID, "000000") & ".mdb"
        Else
            x = BalintFolder & "\Data\PRCKeag" & Format(PRCompany.CompanyID, "000000") & ".mdb"
        End If
        FileCopy ChkFileName, x
    End If
    
    ' save the new company id to the user file
    If Not GLCompany.GetByName(xCopyTo) Then
       MsgBox "Copy Error? ", vbCritical, "GL File Copy"
       GoBack
    End If
    
    If GLUser.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       GLUser.LastCompany = GLCompany.ID
       GLUser.Save (Equate.RecPut)
    End If
    
    MsgBox CopyFrom & " Copy complete to: " & CopyTo, vbInformation, "GL File Copy"
   
End Sub

Public Function GLMultDiv(ByVal LoAcct As Long, _
                          ByVal HiAcct As Long, _
                          ByVal MultDiv As String, _
                          ByVal MDBy As Integer, _
                          ByVal AcctBase As Boolean, _
                          ByVal ShowGo As String) As XArrayDB
   
Dim i As Long
Dim j As Long
Dim k As Long

Dim SubDig As Integer
   
   SubDig = GLCompany.SubDigits
   
   If ShowGo = "Go" Then
      If MultDiv = "Mult" Then
         frmProgress.lblMsg1 = "Now Performing Account Multiply by: " & MDBy & " for: " & GLCompany.Name
      Else
         frmProgress.lblMsg1 = "Now Performing Account Divide by: " & MDBy & " for: " & GLCompany.Name
      End If
   Else
      If MultDiv = "Mult" Then
         frmProgress.lblMsg1 = "Now Displaying Account Multiply by: " & MDBy & " for: " & GLCompany.Name
      Else
         frmProgress.lblMsg1 = "Now Displaying Account Divide by: " & MDBy & " for: " & GLCompany.Name
      End If
   End If
   
   frmProgress.lblMsg2 = "Now Gathering Information .... "
   frmProgress.Show
   Ct = 0
   
   xDB.ReDim 0, 0, 0, 0
   
   xAddRow xDB, Stamp("Account Multiply/Divide")
   If MultDiv = "Mult" Then
      xAddRow xDB, "Multiply accounts " & LoAcct & " to: " & HiAcct
   Else
      xAddRow xDB, "Divide accounts " & LoAcct & " to: " & HiAcct
   End If
   xAddRow xDB, "By: " & MDBy
   xAddRow xDB, " "
   
   If MultDiv <> "Mult" And MultDiv <> "Div" Then
      x = "Bad parameter Mult/Div"
      GoTo MDErr
   End If

   If ShowGo <> "Show" And ShowGo <> "Go" Then
      x = "Bad parameter Show/Go"
      GoTo MDErr
   End If

   If LoAcct = 0 And HiAcct = 0 Then
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLAccount ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLAccount ORDER BY Account"
      End If
   Else
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLAccount WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLAccount WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account"
      End If
   End If
   
   rsInit x, cn, rs
   
   rsFlg = True
   
   If rs.BOF And rs.EOF Then
      x = "No accounts found in the range: " & LoAcct & " To: " & HiAcct
      GoTo MDErr
   End If
   
   ' if just show - open up second record set for lookup to check for errors
   If ShowGo = "Show" Then
      x = "SELECT Account FROM GLAccount"
      rsInit x, cn, rs2
      rsflg2 = True
   End If
   
   rs.MoveFirst
   
   Do Until rs.EOF
      
      Ct = Ct + 1
      If Ct = 1 Or Ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "Updating Account File: " & rs!Account
         frmProgress.lblMsg2.Refresh
      End If
      
      If MultDiv = "Mult" Then
         If AcctBase = False Then
            i = rs!Account * MDBy
         Else
            k = rs!Account Mod 10 ^ SubDig ' branch
            j = Int(rs!Account / 10 ^ SubDig) * MDBy * 10 ^ SubDig ' base
            i = j + k
         End If
      Else
         i = Int(rs!Account / MDBy)
      End If
      
      x = "Account #: " & rs!Account & " Move To #: " & i
            
      If ShowGo = "Go" Then
         rs.Fields("Account") = i
         On Error Resume Next
         rs.Update
         If Err.Number <> 0 Then
            x = x & " ERROR !!! " & Err.Number
            GoTo MDErr
         Else
            x = x & " Complete !"
         End If
         On Error GoTo 0
      Else
         Y = "Account = " & i
         rs2.Find Y, 0, adSearchForward, 1
         If Not rs2.EOF Then     ' was found - can't do !!!
            x = x & " ERROR !!! Already exists !!!"
            GoTo MDErr
         End If
         x = x & " OK"
      End If
            
      xAddRow xDB, x
      
      rs.MoveNext
   
   Loop
      
   rs.Close
   Set rs = Nothing
      
   If ShowGo = "Show" Then
      Set GLMultDiv = xDB
      Exit Function
   End If
      
   ' update GLAmount
   If LoAcct = 0 And HiAcct = 0 Then
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLAmount ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLAmount ORDER BY Account"
      End If
   Else
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLAmount WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLAmount WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account"
      End If
   End If
   
   rsInit x, cn, rs
      
   Ct = 0
      
   If Not (rs.BOF = True And rs.EOF = True) Then
      
       Ct = Ct + 1
       If Ct = 1 Or Ct Mod 100 = 0 Then
          frmProgress.lblMsg2 = "Updating Amount File: " & rs!Account
          frmProgress.lblMsg2.Refresh
       End If
      
       rs.MoveFirst
    
       Count1 = 0
    
       Do Until rs.EOF
    
          If MultDiv = "Mult" Then
             If AcctBase = False Then
                i = rs!Account * MDBy
             Else
                k = rs!Account Mod 10 ^ SubDig ' branch
                j = Int(rs!Account / 10 ^ SubDig) * MDBy * 10 ^ SubDig ' base
                i = j + k
             End If
          Else
             i = rs!Account / MDBy
          End If
    
          rs.Fields("Account") = i
    
          On Error Resume Next
          rs.Update
          If Err.Number <> 0 Then
             x = "GLAmount Error: " & rs!Account
             GoTo MDErr
          End If
          On Error GoTo 0
    
          Count1 = Count1 + 1
    
          rs.MoveNext
    
       Loop
    
       rs.Close
       Set rs = Nothing
    
       x = Count1 & " Amount records updated"
       xAddRow xDB, x
   
   End If
   
   Count1 = 0
   
   Ct = 0
   
   ' update GLHistory
   If LoAcct = 0 And HiAcct = 0 Then
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLHistory ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLHistory ORDER BY Account"
      End If
   Else
      If MultDiv = "Mult" Then
         x = "SELECT * FROM GLHistory WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account DESC"
      Else
         x = "SELECT * FROM GLHistory WHERE Account >= " & LoAcct & _
             " AND Account <= " & HiAcct & _
             " ORDER BY Account"
      End If
   End If
   
   rsInit x, cn, rs
      
   If Not (rs.BOF = True And rs.EOF = True) Then
      
        Ct = Ct + 1
        If Ct = 1 Or Ct Mod 100 = 0 Then
           frmProgress.lblMsg2 = "Updating History File: " & rs!Account
           frmProgress.lblMsg2.Refresh
        End If
      
        rs.MoveFirst
        
        Count1 = 0
        
        Do Until rs.EOF
           
           If MultDiv = "Mult" Then
              If AcctBase = False Then
                 i = rs!Account * MDBy
              Else
                 k = rs!Account Mod 10 ^ SubDig ' branch
                 j = Int(rs!Account / 10 ^ SubDig) * MDBy * 10 ^ SubDig ' base
                 i = j + k
              End If
           Else
              i = rs!Account / MDBy
           End If
           
           rs.Fields("Account") = i
           
           On Error Resume Next
           rs.Update
           If Err.Number <> 0 Then
              x = "GLHistory Error: " & rs!Account
              GoTo MDErr
           End If
           On Error GoTo 0
           
           Count1 = Count1 + 1
           
           rs.MoveNext
        
        Loop
           
        x = Count1 & " History records updated"
        xAddRow xDB, x
      
   End If
   
   Set GLMultDiv = xDB
   
   frmProgress.Hide
   
   Exit Function
   
MDErr:
   
   xAddRow xDB, x
   Set GLMultDiv = xDB
   
   If rsFlg Then
      rs.Close
      Set rs = Nothing
   End If
   
   If rsflg2 Then
      rs2.Close
      Set rs2 = Nothing
   End If

End Function

Public Function CopyBB(ByVal LoAcct As Long, _
                       ByVal HiAcct As Long, _
                       ByVal ValFrom As Long, _
                       ByVal ValTo As Long, _
                       ByVal MainSub As String, _
                       ByVal SubDigits As Integer, _
                       ByVal ShowGo As String) As XArrayDB
                       
Dim bFLg As Boolean
                                              
   xDB.ReDim 0, 0, 0, 0
                                              
   frmProgress.lblMsg1 = "Copying Branch/Budget for: " & GLCompany.Name
   frmProgress.lblMsg2 = "Now gathering records .... "
   frmProgress.Show
   Ct = 0
                                              
   ' range echo back
   xAddRow xDB, Stamp("Copy Branches ")
   xAddRow xDB, "Account # from: " & LoAcct & " to: " & HiAcct
   If MainSub = "Main" Then
      xAddRow xDB, "Copy Main from: " & ValFrom & " to: " & ValTo
   Else
      xAddRow xDB, "Copy Sub from: " & ValFrom & " to: " & ValTo
   End If
   xAddRow xDB, "Number of Sub Digits: " & SubDigits
   xAddRow xDB, " "
   
   bFLg = False
                       
   If MainSub <> "Main" And MainSub <> "Sub" Then
      x = "Parameter error: Main / Sub"
      xAddRow xDB, x
      Exit Function
   End If
   
   x = "SELECT * FROM GLAccount " & _
       "WHERE Account >= " & LoAcct & " AND " & _
       "Account <= " & HiAcct & " ORDER BY Account"
   
   rsInit x, cn, rs
   
   rsFlg = True
   
   If rs.BOF And rs.EOF Then
      x = "No Accounts Founds in " & LoAcct & " to " & HiAcct
      GoTo BBExit
   End If
   
   x = "SELECT * FROM GLAccount ORDER BY Account"
   rsInit x, cn, rs2
   
   rsflg2 = True
   
   rs.MoveFirst
   
   Do Until rs.EOF
            
      Ct = Ct + 1
      If Ct = 1 Or Ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On Account: " & rs!Account
         frmProgress.lblMsg2.Refresh
      End If
            
      If MainSub = "Main" Then
         If Int(rs!Account / 10 ^ SubDigits) <> ValFrom Then GoTo BBCycle
         Acct = (ValTo * 10 ^ SubDigits) + (rs!Account Mod 10 ^ SubDigits)
      Else
         If rs!Account Mod 10 ^ SubDigits <> ValFrom Then GoTo BBCycle
         Acct = Int(rs!Account / 10 ^ SubDigits) * 10 ^ SubDigits + ValTo
      End If
   
      ' see if the account already exists
      x = "Account = " & Acct
      rs2.Find x, 0, adSearchForward, 1
        
      If Not rs2.EOF Then
         x = rs!Account & " copy to: " & Acct & " FAILED - already exists !"
         xAddRow xDB, x
         GoTo BBCycle
      End If
      
      bFLg = True
      
      If ShowGo = "Show" Then
         xAddRow xDB, "Account #: " & rs!Account & " will be copied to: " & Acct
         GoTo BBCycle
      End If
         
      ' assign the fields and add to the second record set
      With rs2
         .AddNew
         .Fields("Account") = Acct
         .Fields("AcctType") = rs!AcctType
         .Fields("Description") = rs!Description & ""
         .Fields("DescNumber") = rs!DescNumber
         .Fields("TotalLevel") = rs!TotalLevel
         .Fields("PrintTab") = rs!PrintTab
         .Fields("LineFeeds") = rs!LineFeeds
         .Fields("BSColumn") = rs!BSColumn
         .Fields("AllStatements") = rs!AllStatements
         .Fields("AllSchedules") = rs!AllSchedules
         .Fields("BranchAcct") = rs!BranchAcct
         .Fields("ConsAcct") = rs!ConsAcct
         .Fields("TotalOnLedger") = rs!TotalOnLedger
         .Fields("DollarSign") = rs!DollarSign
         .Fields("SignRevStmt") = rs!SignRevStmt
         .Fields("SignRevSched") = rs!SignRevSched
         .Update
      End With
   
      xAddRow xDB, "Account #: " & rs!Account & " copied to: " & Acct
      
BBCycle:
   
      rs.MoveNext
   
   Loop

BBExit:
   
   frmProgress.Hide
   
   If Not bFLg Then
      x = "No matching accounts found !"
      xAddRow xDB, x
   End If
   
   Set CopyBB = xDB
   
   If rsFlg Then
      rs.Close
      Set rs = Nothing
   End If
   
   If rsflg2 Then
      rs2.Close
      Set rs2 = Nothing
   End If
   
End Function
Public Function MathUpdate(ByVal StartFY As Long, _
                           ByVal EndFY As Long, _
                           ByVal StartPd As Byte, _
                           ByVal EndPD As Byte) As XArrayDB

Dim GLDescription As New cGLDescription
Dim GLAmount As New cGLAmount

Dim mx As New XArrayDB
Dim MAcct As Long
Dim SLen As Integer
Dim Desc As String
Dim ItemCount As Integer
Dim BudAmt, Amt As Currency
Dim LastVal As Long
Dim BkMark As Variant

Dim Op As String
Dim nVal As Double

    frmProgress.lblMsg1 = "Now Performing Math Update for: " & GLCompany.Name
    frmProgress.lblMsg2 = "Now Gathering Information ... "
    frmProgress.Show
    Ct = 0
    rsFlg = False

    CurFormat = "###,###,##0.00"

    xDB.ReDim 0, 0, 0, 0

    xAddRow xDB, " "
    xAddRow xDB, Stamp("Math Update:")
    xAddRow xDB, "Start FY: " & StartFY & " End FY: " & EndFY
    xAddRow xDB, "Start Pd: " & StartPd & " End Pd: " & EndPD
    xAddRow xDB, " "

    ' check to see if any amounts exist ?
    x = "SELECT Account FROM GLAmount WHERE " & _
       "FiscalYear >= " & StartFY & " AND " & _
       "FiscalYear <= " & EndFY

    rsInit x, cn, rs

    rsFlg = True

    If rs.BOF And rs.EOF Then
        xAddRow xDB, "No Amounts found for: " & StartFY & " to: " & EndFY
        GoTo ExitUp
    End If

    rs.Close

    ' create GLAmount records for GLAccount T records if necessary
    x = "SELECT Account, AcctType FROM GLAccount WHERE AcctType = 'T' OR AcctType = '.' OR AcctType = 'M'"
   
    rsInit x, cn, rs
    rs.MoveFirst
   
    GLAmount.OpenRS
       
    Ct = 0
   
    Do Until rs.EOF
      
        Ct = Ct + 1
        If Ct = 1 Or Ct Mod 10 = 0 Then
            frmProgress.lblMsg2 = "Checking for Amount Records: " & Ct
            frmProgress.lblMsg2.Refresh
        End If
      
        For yr = StartFY To EndFY
                              
            If Not GLAmount.Find(rs!Account, yr) Then
                GLAmount.Clear
                GLAmount.Account = rs!Account
                GLAmount.FiscalYear = yr
                GLAmount.Save (Equate.RecAdd)
                           
                xAddRow xDB, "GLAmount record added: " & rs!Account & " " & yr
            End If
          
        Next yr
          
        rs.MoveNext
       
    Loop
    
    rs.Close
    
    frmProgress.lblMsg2 = "Gathering Account and Amount information .... "
    frmProgress.lblMsg2.Refresh
    
    x = "SELECT GLAccount.*, GLAmount.FiscalYear, " & _
        "GLAmount.Amount01, GLAmount.Amount02, GLAmount.Amount03, " & _
        "GLAmount.Amount04, GLAmount.Amount05, GLAmount.Amount06, " & _
        "GLAmount.Amount07, GLAmount.Amount08, GLAmount.Amount09, " & _
        "GLAmount.Amount10, GLAmount.Amount11, GLAmount.Amount12, GLAmount.Amount13, " & _
        "GLAmount.Budget01, GLAmount.Budget02, GLAmount.Budget03, " & _
        "GLAmount.Budget04, GLAmount.Budget05, GLAmount.Budget06, " & _
        "GLAmount.Budget07, GLAmount.Budget08, GLAmount.Budget09, " & _
        "GLAmount.Budget10, GLAmount.Budget11, GLAmount.Budget12, GLAmount.Budget13 " & _
        "FROM GLAccount LEFT JOIN GLAmount on " & _
        "(GLAccount.Account = GLAmount.Account AND " & _
        "GLAmount.FiscalYear >= " & StartFY & " AND " & _
        "GLAmount.FiscalYear <= " & EndFY & ") ORDER BY GLAccount.Account"
    
    rsInit x, cn, rs
    
    If rs.BOF And rs.EOF Then
        xAddRow xDB, "No accounts/amounts found: "
        GoTo ExitUp
    End If
    
    For yr = StartFY To EndFY
        For Mo = StartPd To EndPD
    
            xAddRow xDB, " "
            xAddRow xDB, Stamp("TOTALS Update for FY = " & yr & _
                               PeriodName(yr, Mo, GLCompany.FirstPeriod, GLCompany.NumberPds))
            xAddRow xDB, " "
    
            ' clear variables
            FirstM = 0
            FirstN = 0
            FirstP = 0
            FirstL = 0
            For i = 1 To 10
                G(i) = 0
                BG(i) = 0
            Next i
    
            ' loop thru the balance sheet accounts
            rs.MoveFirst
    
            Ct = 0
    
            Do Until rs.EOF
    
                Ct = Ct + 1
                If Ct = 1 Or Ct Mod 100 = 0 Then
                    frmProgress.lblMsg2 = "Gathering Totals " & rs!Account & " " & yr & "/" & Mo
                    frmProgress.lblMsg2.Refresh
                End If
    
                If rs!FiscalYear <> yr Then GoTo MCycle
                  
                ' assign the amounts if available
                Amount(1) = nNull(rs!Amount01)
                Amount(2) = nNull(rs!Amount02)
                Amount(3) = nNull(rs!Amount03)
                Amount(4) = nNull(rs!Amount04)
                Amount(5) = nNull(rs!Amount05)
                Amount(6) = nNull(rs!Amount06)
                Amount(7) = nNull(rs!Amount07)
                Amount(8) = nNull(rs!Amount08)
                Amount(9) = nNull(rs!Amount09)
                Amount(10) = nNull(rs!Amount10)
                Amount(11) = nNull(rs!Amount11)
                Amount(12) = nNull(rs!Amount12)
                Amount(13) = nNull(rs!Amount13)
                                
                ' =====================================================
                BudAmount(1) = nNull(rs!Budget01)
                BudAmount(2) = nNull(rs!Budget02)
                BudAmount(3) = nNull(rs!Budget03)
                BudAmount(4) = nNull(rs!Budget04)
                BudAmount(5) = nNull(rs!Budget05)
                BudAmount(6) = nNull(rs!Budget06)
                BudAmount(7) = nNull(rs!Budget07)
                BudAmount(8) = nNull(rs!Budget08)
                BudAmount(9) = nNull(rs!Budget09)
                BudAmount(10) = nNull(rs!Budget10)
                BudAmount(11) = nNull(rs!Budget11)
                BudAmount(12) = nNull(rs!Budget12)
                BudAmount(13) = nNull(rs!Budget13)
                  
                ' 1000
                If rs!AcctType = "C" And Not rs!BranchAcct And Not rs!ConsAcct Then
                    ClearTotals
                End If
    
                ' 1010 check types
                If rs!AcctType = "I" Then SignMode = 1
                If rs!AcctType = "E" Then SignMode = -1
                If rs!AcctType = "M" And FirstM = 0 Then FirstM = rs!Account
                If rs!AcctType = "N" And FirstN = 0 Then FirstN = rs!Account
                If rs!AcctType = "P" And FirstP = 0 Then GrandTotals
                If rs!AcctType = "L" And FirstL = 0 Then GrandTotals
                If rs!AcctType = "T" And Not rs!BranchAcct And Not rs!ConsAcct Then
                    TotalsX
                End If
                If rs!AcctType <> "N" And rs!AcctType <> "0" Then GoTo MCycle
                  
                ' type 0 and N   !!!!!!!!!!!
                For i = 1 To 5
                    G(i) = G(i) + Amount(Mo)
                    BG(i) = BG(i) + BudAmount(Mo)
                Next i
                  
                If rs!AcctType = "N" Then
                    G(8) = Amount(Mo)
                    BG(8) = BudAmount(Mo)
                End If
                  
                If SignMode = 1 Then
                    G(10) = G(10) + Amount(Mo)
                    BG(10) = BG(10) + Amount(Mo)
                Else
                    G(9) = G(9) + Amount(Mo)
                    BG(9) = BG(9) + Amount(Mo)
                End If
    
MCycle:
                rs.MoveNext
    
            Loop
        Next Mo
    Next yr
                
    ' regular totals 1380
    xAddRow xDB, "Grand Totals FY = " & EndFY & " " & _
                 PeriodName(yr, StartPd, GLCompany.FirstPeriod, GLCompany.NumberPds) & " To: " & _
                 PeriodName(yr, EndPD, GLCompany.FirstPeriod, GLCompany.NumberPds)
    
       
    xAddRow xDB, "Total Assets" & vbTab & vbTab & Format(G(6), CurFormat)
    xAddRow xDB, "Total Liabilities" & vbTab & vbTab & Format(G(7), CurFormat)
    xAddRow xDB, "Check Balance" & vbTab & vbTab & Format(G(6) + G(7), CurFormat)
    xAddRow xDB, " "
    xAddRow xDB, "Total Net Profit -Loss" & vbTab & vbTab & Format(G(8), CurFormat)
    xAddRow xDB, "Total Expense" & vbTab & vbTab & Format(G(9), CurFormat)
    xAddRow xDB, "Total Income" & vbTab & vbTab & Format(G(10), CurFormat)
    xAddRow xDB, " "
    
    ' ====================================================================================
    ' update math records
    GLDescription.OpenRS
    
    ' On Error GoTo ErrMsg
    
    For yr = StartFY To EndFY
        For Mo = StartPd To EndPD
                       
            Ct = 0
           
            xAddRow xDB, Stamp("Update General Ledger Math Records For: " & yr & "/" & Mo)
           
            x = "Account = " & FirstM
            rs.Find x, 0, adSearchForward, 1
            If rs.EOF Then GoTo EndMath
          
            Do Until rs.EOF
          
                Ct = Ct + 1
                If Ct = 1 Or Ct Mod 100 = 0 Then
                    frmProgress.lblMsg2 = "Updating Math Records: " & rs!Account & " " & yr & "/" & Mo
                    frmProgress.lblMsg2.Refresh
                End If
          
                If rs!AcctType <> "M" Then GoTo NextM
                If rs!FiscalYear <> yr Then GoTo NextM
          
                ' save the spot
                BkMark = rs.Bookmark
          
                mx.Clear
                mx.ReDim 0, 0, 1, 3
             
                Desc = rs!Description
             
                If rs!DescNumber = 0 Then
                    SLen = Len(Desc)
                Else
                    If GLDescription.Find(rs!DescNumber) Then
                        Desc = GLDescription.Description & rs!Description
                    Else
                        Desc = rs!Description
                    End If
                End If
             
                If Desc = "" Then GoTo NextM
                If IsNull(Desc) Then GoTo NextM
             
                ' check start of string
                If InStr(1, "N-.0123456789", Mid(Desc, 1, 1), vbTextCompare) = 0 Then
                    xAddRow xDB, "ERROR Acct#: " & rs!Account & " Bad M Description: " & Desc
                    GoTo NextM
                End If
             
                ' first value is a number not an account
                If Mid(Desc, 1, 1) = "N" Then
                    mx(0, 2) = 1
                    i = 1
                Else
                    i = 0
                End If
             
                x = ""
                ItemCount = 0
                   
                ' loop for the numbers
                Do
                    i = i + 1
                    If i > Len(Desc) Then Exit Do
                
                    If InStr(1, "-.0123456789", Mid(Desc, i, 1), vbTextCompare) = 0 Then
                        If ItemCount <> 0 Then mx.AppendRows (1)
                        mx(ItemCount, 1) = x
                        x = ""
                        ItemCount = ItemCount + 1
                        If Mid(Desc, i + 1, 1) = "N" Then i = i + 1
                        If i >= Len(Desc) Then Exit Do
                    Else
                        x = x & Mid(Desc, i, 1)
                    End If
                
                Loop
                
                If x <> "" Then
                    mx.AppendRows (1)
                    mx(ItemCount, 1) = x
                    ItemCount = ItemCount + 1
                End If
                
                ItemCount = 0
                i = 1
             
                ' loop for the operators
                Do Until i > Len(Trim(Desc))
                    x = Mid(Desc, i, 1)
                    If InStr(1, "ASMDT", x) <> 0 Or x = " " Then
                        ItemCount = ItemCount + 1
                        If x = " " Then
                            mx(ItemCount, 3) = "A"
                        Else
                            mx(ItemCount, 3) = x
                        End If
                                                        
                        If Mid(Desc, i + 1, 1) = "N" Then
                            i = i + 1
                            mx(ItemCount, 2) = 1
                        Else
                            mx(ItemCount, 2) = 0
                        End If
                   
                    End If
                    i = i + 1
                Loop
             
                ' first argument must be an acct # - find it
                Amt = GetAmount(mx(0, 1), mx(0, 1), False, Mo)
                BudAmt = GetAmount(mx(0, 1), mx(0, 1), False, Mo, True)
          
                If mx.UpperBound(1) >= 1 And mx(1, 3) = "T" Then
                    Amt = 0
                End If
             
                LastVal = mx(0, 1)
             
                For i = 1 To mx.UpperBound(1)
             
                    On Error Resume Next
             
                    nVal = mx(i, 1)
                       
                    If Err.Number <> 0 Then
                    End If
                        
                    On Error GoTo 0
                       
                    Op = mx(i, 3)
                 
                    If Op = "A" Then
                        If mx(i, 2) = 0 Then
                            Amt = Amt + GetAmount(nVal, nVal, False, Mo)
                            BudAmt = BudAmt + GetAmount(nVal, nVal, False, Mo, True)
                        Else
                            Amt = Amt + nVal
                            BudAmt = BudAmt + nVal
                        End If
                    ElseIf Op = "S" Then
                        If mx(i, 2) = 0 Then
                            Amt = Amt - GetAmount(nVal, nVal, False, Mo)
                            BudAmt = BudAmt - GetAmount(nVal, nVal, False, Mo, True)
                        Else
                            Amt = Amt - nVal
                            BudAmt = BudAmt - nVal
                        End If
                    ElseIf Op = "D" Then
                        If mx(i, 2) = 0 Then
                            Amt = Div0(Amt, GetAmount(nVal, nVal, False, Mo))
                            BudAmt = Div0(BudAmt, GetAmount(nVal, nVal, False, Mo, True))
                        Else
                            Amt = Div0(Amt, CLng(nVal))
                            BudAmt = Div0(BudAmt, CLng(nVal))
                        End If
                    ElseIf Op = "M" Then
                        If mx(i, 2) = 0 Then
                            Amt = Amt * GetAmount(nVal, nVal, False, Mo)
                            BudAmt = BudAmt * GetAmount(nVal, nVal, False, Mo, True)
                        Else
                            Amt = Amt * nVal
                            BudAmt = BudAmt + nVal
                        End If
                    ElseIf Op = "T" Then
                        If mx(i, 2) = 0 Then
                            Amt = Amt + GetAmount(LastVal, nVal, True, Mo)
                            BudAmt = BudAmt + GetAmount(LastVal, nVal, True, Mo, True)
                        Else
                            Amt = Amt + nVal
                            BudAmt = BudAmt + nVal
                        End If
                    End If
                              
                    LastVal = nVal
                              
                Next i
          
                ' go back to the BookMark and assign the amount
                rs.Bookmark = BkMark
                  
                If Mo = 1 Then rs!Amount01 = Amt
                If Mo = 2 Then rs!Amount02 = Amt
                If Mo = 3 Then rs!Amount03 = Amt
                If Mo = 4 Then rs!Amount04 = Amt
                If Mo = 5 Then rs!Amount05 = Amt
                If Mo = 6 Then rs!Amount06 = Amt
                If Mo = 7 Then rs!Amount07 = Amt
                If Mo = 8 Then rs!Amount08 = Amt
                If Mo = 9 Then rs!Amount09 = Amt
                If Mo = 10 Then rs!Amount10 = Amt
                If Mo = 11 Then rs!Amount11 = Amt
                If Mo = 12 Then rs!Amount12 = Amt
                If Mo = 13 Then rs!Amount13 = Amt
                 
                If Mo = 1 Then rs!Budget01 = BudAmt
                If Mo = 2 Then rs!Budget02 = BudAmt
                If Mo = 3 Then rs!Budget03 = BudAmt
                If Mo = 4 Then rs!Budget04 = BudAmt
                If Mo = 5 Then rs!Budget05 = BudAmt
                If Mo = 6 Then rs!Budget06 = BudAmt
                If Mo = 7 Then rs!Budget07 = BudAmt
                If Mo = 8 Then rs!Budget08 = BudAmt
                If Mo = 9 Then rs!Budget09 = BudAmt
                If Mo = 10 Then rs!Budget10 = BudAmt
                If Mo = 11 Then rs!Budget11 = BudAmt
                If Mo = 12 Then rs!Budget12 = BudAmt
                If Mo = 13 Then rs!Budget13 = BudAmt
                
                rs.Update
    
NextM:
                rs.MoveNext
               
            Loop           ' next rs
           
            xAddRow xDB, Stamp("Completed")
            xAddRow xDB, " "
           
        Next Mo
       
    Next yr
    
EndMath:
    
    
       ' =============================================================================
       ' update percentage records
       ' =============================================================================
       
       For yr = StartFY To EndFY
           For Mo = StartPd To EndPD
           
               xAddRow xDB, Stamp("Percentage record update for: " & yr & "/" & Mo)
           
               x = "Account = " & FirstN
               rs.Find x, 0, adSearchForward, 1
               If rs.EOF Then
                  xAddRow xDB, "First N find error: " & FirstN
                  GoTo ExitUp
               End If
               
               Ct = 0
               
               Do Until rs.EOF
               
                  Ct = Ct + 1
                  If Ct = 1 Or Ct Mod 100 = 0 Then
                     frmProgress.lblMsg2 = "Updating Percentage Records: " & rs!Account & " " & yr & "/" & Mo
                     frmProgress.lblMsg2.Refresh
                  End If
               
                  If rs!AcctType <> "." Then GoTo NextP
                  
                  BkMark = rs.Bookmark
                  
                  If IsNumeric(rs!Description) = False Then
                       GoTo NextP
                  End If
                  x = "Account = " & rs!Description
                  Acct = rs!Account
                  
                  rs.Find x, 0, adSearchForward, 1
                  
                  If rs.EOF Then
                     xAddRow xDB, "ERROR: Pct Acct#: " & Acct & _
                                 " points to non-existent Account: " & x
                     rs.Bookmark = BkMark
                     GoTo NextP
                  End If
               
                  If IsNull(rs!Amount01) Then
                     Amt = 0
                  Else
                     If Mo = 1 Then Amt = rs!Amount01
                     If Mo = 2 Then Amt = rs!Amount02
                     If Mo = 3 Then Amt = rs!Amount03
                     If Mo = 4 Then Amt = rs!Amount04
                     If Mo = 5 Then Amt = rs!Amount05
                     If Mo = 6 Then Amt = rs!Amount06
                     If Mo = 7 Then Amt = rs!Amount07
                     If Mo = 8 Then Amt = rs!Amount08
                     If Mo = 9 Then Amt = rs!Amount09
                     If Mo = 10 Then Amt = rs!Amount10
                     If Mo = 11 Then Amt = rs!Amount11
                     If Mo = 12 Then Amt = rs!Amount12
                     If Mo = 13 Then Amt = rs!Amount13
                  End If
                  
                  rs.Bookmark = BkMark
                  
                  If Mo = 1 Then rs!Amount01 = Amt
                  If Mo = 2 Then rs!Amount02 = Amt
                  If Mo = 3 Then rs!Amount03 = Amt
                  If Mo = 4 Then rs!Amount04 = Amt
                  If Mo = 5 Then rs!Amount05 = Amt
                  If Mo = 6 Then rs!Amount06 = Amt
                  If Mo = 7 Then rs!Amount07 = Amt
                  If Mo = 8 Then rs!Amount08 = Amt
                  If Mo = 9 Then rs!Amount09 = Amt
                  If Mo = 10 Then rs!Amount10 = Amt
                  If Mo = 11 Then rs!Amount11 = Amt
                  If Mo = 12 Then rs!Amount12 = Amt
                  If Mo = 13 Then rs!Amount13 = Amt
                  
                  rs.Update
                   
NextP:
                  rs.MoveNext
          
               Loop
           Next Mo
       
           xAddRow xDB, Stamp("Complete")
           xAddRow xDB, " "
       
       Next yr
    
ExitUp:
    
       If rsFlg Then
          rs.Close
          Set rs = Nothing
       End If
    
       Set MathUpdate = xDB
    
       frmProgress.Hide

End Function

Private Sub TotalsX()             ' 1180
   
   GAcct = rs!Account
   
   If Mo = 1 Then rs!Amount01 = G(rs!TotalLevel)
   If Mo = 2 Then rs!Amount02 = G(rs!TotalLevel)
   If Mo = 3 Then rs!Amount03 = G(rs!TotalLevel)
   If Mo = 4 Then rs!Amount04 = G(rs!TotalLevel)
   If Mo = 5 Then rs!Amount05 = G(rs!TotalLevel)
   If Mo = 6 Then rs!Amount06 = G(rs!TotalLevel)
   If Mo = 7 Then rs!Amount07 = G(rs!TotalLevel)
   If Mo = 8 Then rs!Amount08 = G(rs!TotalLevel)
   If Mo = 9 Then rs!Amount09 = G(rs!TotalLevel)
   If Mo = 10 Then rs!Amount10 = G(rs!TotalLevel)
   If Mo = 11 Then rs!Amount11 = G(rs!TotalLevel)
   If Mo = 12 Then rs!Amount12 = G(rs!TotalLevel)
   If Mo = 13 Then rs!Amount13 = G(rs!TotalLevel)
   
   If Mo = 1 Then rs!Budget01 = BG(rs!TotalLevel)
   If Mo = 2 Then rs!Budget02 = BG(rs!TotalLevel)
   If Mo = 3 Then rs!Budget03 = BG(rs!TotalLevel)
   If Mo = 4 Then rs!Budget04 = BG(rs!TotalLevel)
   If Mo = 5 Then rs!Budget05 = BG(rs!TotalLevel)
   If Mo = 6 Then rs!Budget06 = BG(rs!TotalLevel)
   If Mo = 7 Then rs!Budget07 = BG(rs!TotalLevel)
   If Mo = 8 Then rs!Budget08 = BG(rs!TotalLevel)
   If Mo = 9 Then rs!Budget09 = BG(rs!TotalLevel)
   If Mo = 10 Then rs!Budget10 = BG(rs!TotalLevel)
   If Mo = 11 Then rs!Budget11 = BG(rs!TotalLevel)
   If Mo = 12 Then rs!Budget12 = BG(rs!TotalLevel)
   If Mo = 13 Then rs!Budget13 = BG(rs!TotalLevel)
   
   rs.Update
   
   ClearTotals

End Sub

Private Sub ClearTotals()        ' 1220
   For i = 1 To rs!TotalLevel
       If rs!AcctType = "C" Then
          G(i) = 0
          BG(i) = 0
       End If
       If i <> 5 Then
          G(i) = 0
          BG(i) = 0
       End If
   Next i
End Sub

Private Sub GrandTotals()        ' 1280

   If rs!AcctType = "L" Then
      FirstL = rs!Account
      G(6) = G(5)
      BG(6) = BG(5)
   End If
   
   If rs!AcctType = "P" Then
      FirstP = rs!Account
      G(7) = G(5)
      BG(7) = BG(5)
   End If
   
   If GAcct <> 0 Then


'  ????????????????????????????????????????????  continue if not assigned
'             gacct is assigned when the first T record not branch or cons flagged is read
'       If GAcct = 0 Then
'          MsgBox "GAcct = 0 ???"
'          End
'       End If
       
       LAcct = rs!Account
       x = "Account = " & GAcct
       rs.Find x, 0, adSearchForward, 1
       
       If rs.EOF Then
          MsgBox "GAcct find " & GAcct & " ???"
          End
       End If
       
       If Mo = 1 Then rs!Amount01 = rs!Amount01 + G(2)
       If Mo = 2 Then rs!Amount02 = rs!Amount02 + G(2)
       If Mo = 3 Then rs!Amount03 = rs!Amount03 + G(2)
       If Mo = 4 Then rs!Amount04 = rs!Amount04 + G(2)
       If Mo = 5 Then rs!Amount05 = rs!Amount05 + G(2)
       If Mo = 6 Then rs!Amount06 = rs!Amount06 + G(2)
       If Mo = 7 Then rs!Amount07 = rs!Amount07 + G(2)
       If Mo = 8 Then rs!Amount08 = rs!Amount08 + G(2)
       If Mo = 9 Then rs!Amount09 = rs!Amount09 + G(2)
       If Mo = 10 Then rs!Amount10 = rs!Amount10 + G(2)
       If Mo = 11 Then rs!Amount11 = rs!Amount11 + G(2)
       If Mo = 12 Then rs!Amount12 = rs!Amount12 + G(2)
       If Mo = 13 Then rs!Amount13 = rs!Amount13 + G(2)
     
       If Mo = 1 Then rs!Budget01 = rs!Budget01 + BG(2)
       If Mo = 2 Then rs!Budget02 = rs!Budget02 + BG(2)
       If Mo = 3 Then rs!Budget03 = rs!Budget03 + BG(2)
       If Mo = 4 Then rs!Budget04 = rs!Budget04 + BG(2)
       If Mo = 5 Then rs!Budget05 = rs!Budget05 + BG(2)
       If Mo = 6 Then rs!Budget06 = rs!Budget06 + BG(2)
       If Mo = 7 Then rs!Budget07 = rs!Budget07 + BG(2)
       If Mo = 8 Then rs!Budget08 = rs!Budget08 + BG(2)
       If Mo = 9 Then rs!Budget09 = rs!Budget09 + BG(2)
       If Mo = 10 Then rs!Budget10 = rs!Budget10 + BG(2)
       If Mo = 11 Then rs!Budget11 = rs!Budget11 + BG(2)
       If Mo = 12 Then rs!Budget12 = rs!Budget12 + BG(2)
       If Mo = 13 Then rs!Budget13 = rs!Budget13 + BG(2)
     
       rs.Update
       
       ' find original account
       x = "Account = " & LAcct
       rs.Find x, 0, adSearchForward, 1
    
       If rs.EOF Then
          MsgBox "Refind error: " & LAcct
          End
       End If

   End If

   ' 1360
   For i = 1 To 5
       G(i) = 0
   Next i

End Sub
Private Function GetAmount(ByVal LoAcct As Long, _
                           ByVal HiAcct As Long, _
                           ByVal ZeroOnly As Boolean, _
                           ByVal Pd As Byte, _
                           Optional Budg As Boolean) As Currency
   
Dim Acct As Long
   
    GetAmount = 0
   
    x = "Account = " & LoAcct
    rs.Find x, 0, adSearchForward, 1
   
    ' the LoAcct must be found or a zero is returned
    If rs.EOF Then Exit Function
      
    Do Until rs!Account > HiAcct
   
        If ZeroOnly And rs!AcctType <> "0" Then GoTo NxtAcct
      
        If Budg = False Then
      
            If Not IsNull(rs!Amount01) Then
                If Pd = 1 Then GetAmount = GetAmount + rs!Amount01
                If Pd = 2 Then GetAmount = GetAmount + rs!Amount02
                If Pd = 3 Then GetAmount = GetAmount + rs!Amount03
                If Pd = 4 Then GetAmount = GetAmount + rs!Amount04
                If Pd = 5 Then GetAmount = GetAmount + rs!Amount05
                If Pd = 6 Then GetAmount = GetAmount + rs!Amount06
                If Pd = 7 Then GetAmount = GetAmount + rs!Amount07
                If Pd = 8 Then GetAmount = GetAmount + rs!Amount08
                If Pd = 9 Then GetAmount = GetAmount + rs!Amount09
                If Pd = 10 Then GetAmount = GetAmount + rs!Amount10
                If Pd = 11 Then GetAmount = GetAmount + rs!Amount11
                If Pd = 12 Then GetAmount = GetAmount + rs!Amount12
                If Pd = 13 Then GetAmount = GetAmount + rs!Amount13
            End If

        Else

            If Not IsNull(rs!Budget01) Then
                If Pd = 1 Then GetAmount = GetAmount + rs!Budget01
                If Pd = 2 Then GetAmount = GetAmount + rs!Budget02
                If Pd = 3 Then GetAmount = GetAmount + rs!Budget03
                If Pd = 4 Then GetAmount = GetAmount + rs!Budget04
                If Pd = 5 Then GetAmount = GetAmount + rs!Budget05
                If Pd = 6 Then GetAmount = GetAmount + rs!Budget06
                If Pd = 7 Then GetAmount = GetAmount + rs!Budget07
                If Pd = 8 Then GetAmount = GetAmount + rs!Budget08
                If Pd = 9 Then GetAmount = GetAmount + rs!Budget09
                If Pd = 10 Then GetAmount = GetAmount + rs!Budget10
                If Pd = 11 Then GetAmount = GetAmount + rs!Budget11
                If Pd = 12 Then GetAmount = GetAmount + rs!Budget12
                If Pd = 13 Then GetAmount = GetAmount + rs!Budget13
            End If

        End If

NxtAcct:
   
        rs.MoveNext
        If rs.EOF Then Exit Function
   
    Loop

End Function


Function Stamp(ByVal Msg As String) As String
   Stamp = Msg & " " & Format(Now, "Long Date") & " @ " & Format(Now, "Long Time")
End Function

Sub xAddRow(ByRef xDB As XArrayDB, ByVal Msg As String)
   xDB.AppendRows 1
   xDB(xDB.UpperBound(1), 0) = Msg
End Sub

Function DecPlaces(ByVal i As Integer) As Integer
   Do Until i = 0
      i = Int(i / 10)
      DecPlaces = DecPlaces + 1
   Loop
End Function

Public Function YearEnd(ByVal FiscalYear As Integer, _
                        ByVal RetEarn As Long, _
                        ByVal rs As ADODB.Recordset) As XArrayDB
                        
Dim Amount As Currency
Dim NetProfit As Currency
Dim RecCount As Long
Dim TlDebits As Currency
Dim TlCredits As Currency
        
    frmProgress.lblMsg1 = "Processing Year End for: " & GLCompany.Name
    frmProgress.lblMsg2 = "Now Gathering Information .... "
    frmProgress.Show
    Ct = 0
    
    xDB.ReDim 0, 2, 0, 0
    xDB(1, 0) = Stamp("Year End process started: ")
    xDB(2, 0) = "Close Fiscal Year: " & FiscalYear - 1 & " to: " & FiscalYear
   
    xAddRow xDB, " "
    xAddRow xDB, " "
   
    ' open the GLHistory file
    GLHistory.OpenRS
   
    ' create batch record
    GLBatch.AddBatch FiscalYear, 1
    BatchNum = GLBatch.BatchNumber
   
    ' open record set for year closing to
    If Not (GLAmount.GetRecordSet(1, 999999999, FiscalYear, FiscalYear)) Then
    End If
   
    ' loop thru GLAccount
    '  GLAccount record set with LEFT JOIN for GLAmount for year closing from
    If Not GLAccount.GetRecordSetsNoBudget(FiscalYear - 1, FiscalYear - 1) Then
        MsgBox "No Accounts Found !!!", vbOKOnly + vbCritical, "Year End"
        Exit Function
    End If
   
    Do
   
        Ct = Ct + 1
        If Ct = 1 Or Ct Mod 100 = 0 Then
            frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
            frmProgress.lblMsg2.Refresh
        End If
   
        If InStr(1, "MT0.", GLAccount.AcctType) = 0 Then GoTo NextAcct
   
        ' skip specified accounts
        SQLString = "GLDesc = " & GLAccount.Account
        rs.Find SQLString, 0, adSearchForward, 1
        If Not rs.EOF Then GoTo NextAcct
   
        If GLAccount.Account = GLCompany.RetEarnAcct Then GoTo NextAcct
      
        ' get amount from fiscal year being closed
        Amount = GLAccount.GetCurrAmount(1, 13)
      
        ' create amount record
        CreateAmount GLAccount.Account, FiscalYear
      
        ' only write balance sheet accounts with an amount
        If GLAccount.Account > GLCompany.FirstPAcct Or Amount = 0 Then GoTo NextAcct
   
        ' write to GLAmount
        GLAmount.Amount01 = GLAmount.Amount01 + Amount
        GLAmount.Save (Equate.RecPut)
      
        ' write to GLHistory if type 0
        If GLAccount.AcctType = "0" Then
         
            GLHistory.Clear
            GLHistory.Account = GLAccount.Account
            GLHistory.Period = 1
            GLHistory.FiscalYear = FiscalYear
            GLHistory.Amount = Amount
            GLHistory.Description = "OPENING ENTRY"
            GLHistory.Reference = "FY TOTALS"
            GLHistory.JournalSource = 1
            GLHistory.HisType = "A"
            GLHistory.BatchNumber = BatchNum
            GLHistory.Save (Equate.RecAdd)
      
            xAddRow xDB, "Account#: " & Format(GLAccount.Account, "########0") & _
                      "  Amount: " & Format(Amount, "###,###,##0.00")
      
      
            ' update info for the batch record
            RecCount = RecCount + 1
            If Amount >= 0 Then
                TlDebits = TlDebits + Amount
            Else
                TlCredits = TlCredits + Amount
            End If
      
        End If
   
NextAcct:
        If Not GLAccount.GetNext Then Exit Do
       
    Loop

    ' get other account amounts
    Amount = 0
    NetProfit = 0
    
    rs.MoveFirst
    If Not rs.EOF Then
                    
        Do
            
            If Not IsNull(rs!GLDesc) Then
                Acct = rs!GLDesc
                If GLAccount.Find(Acct) Then
                    NetProfit = NetProfit + GLAccount.GetCurrAmount(1, 13)
                    CreateAmount GLAccount.Account, FiscalYear
                End If
            End If
            rs.MoveNext
            If rs.EOF Then Exit Do
        
        Loop

    End If

    ' ======================================================================
    ' process the "N" record
    If Not GLAccount.Find(GLCompany.NetProfitAcct) Then
        MsgBox "Net Profit Acct NF: " & GLCompany.NetProfitAcct, vbCritical + vbOKOnly
        Exit Function
    End If
   
    NetProfit = NetProfit + GLAccount.GetCurrAmount(1, 13)
    CreateAmount GLAccount.Account, FiscalYear
   
    ' ======================================================================
   
    ' get the amount from Retained Earnings
    ' If Not GLAccount.Find(GLCompany.RetEarnAcct) Then
   
    ' 03/11/08 - take from param list
    If Not GLAccount.Find(RetEarn) Then
        MsgBox "Retained Earnings Acct NF: " & GLCompany.RetEarnAcct, vbCritical + vbOKOnly
        Exit Function
    End If
   
    NetProfit = NetProfit + GLAccount.GetCurrAmount(1, 13)
   
    CreateAmount GLAccount.Account, FiscalYear
   
    GLAmount.Amount01 = GLAmount.Amount01 + NetProfit
    GLAmount.Save (Equate.RecPut)
   
    ' write to GLHistory if type 0
    GLHistory.Clear
    GLHistory.Account = GLAccount.Account
    GLHistory.Period = 1
    GLHistory.FiscalYear = FiscalYear
    GLHistory.Amount = NetProfit
    GLHistory.Description = "RETAIN EARNINGS"
    GLHistory.Reference = "FY TOTALS"
    GLHistory.JournalSource = 1
    GLHistory.HisType = "A"
    GLHistory.BatchNumber = BatchNum
    GLHistory.Save (Equate.RecAdd)
   
    ' update info for the batch record
    RecCount = RecCount + 1
    If NetProfit >= 0 Then
        TlDebits = TlDebits + NetProfit
    Else
        TlCredits = TlCredits + NetProfit
    End If
   
    ' write last close to GLCompany
    GLCompany.LastClose = FiscalYear * 10 ^ 4 + GLCompany.FirstPeriod * 100 + 1
   
    ' update the last batch number in GLCompany
    GLCompany.LastBatch = BatchNum
   
    GLCompany.Save (Equate.RecPut)
   
    xAddRow xDB, "RETAINED EARNINGS - Account#: " & Format(GLAccount.Account, "########0") & _
                "  Amount: " & Format(NetProfit, "###,###,##0.00")
   
    xAddRow xDB, " "
    xAddRow xDB, " "
    xAddRow xDB, Stamp("Fiscal Year close complete to: " & FiscalYear)
   
    ' update the batch record
    If GLBatch.GetBatch(BatchNum) Then
        GLBatch.Credits = TlCredits
        GLBatch.Debits = TlDebits
        GLBatch.JournalSource = 1
        GLBatch.Records = RecCount
        GLBatch.CreateUser = GLUser.ID
        GLBatch.UpdateUser = GLUser.ID
        GLBatch.Created = Now
        GLBatch.Updated = Now
        GLBatch.Save (Equate.RecPut)
    End If
   
    Set YearEnd = xDB
   
    frmProgress.Hide
   
End Function
                        

Private Sub CreateAmount(ByVal Account As Long, ByVal FY As Long)
      
   ' see if GLAmount to close to exists - create if it does not
   If Not GLAmount.Find(Account, FY) Then
      GLAmount.Clear
      GLAmount.Account = Account
      GLAmount.FiscalYear = FY
      GLAmount.Save (Equate.RecAdd)
   End If

End Sub

