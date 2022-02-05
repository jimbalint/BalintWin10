Attribute VB_Name = "modStart"
Private Sub Main()

Dim X As String
Dim FileExt As String
    
    frmSplash.Show

    Set Equate = New cEquate
    Set User = New cGLUser
    Set PRCompany = New cPRCompany
    Set PRDepartment = New cPRDepartment
    Set PREmployee = New cPREmployee
    Set PREquate = New cPREquate
    Set PRBatch = New cPRBatch
    Set PRHist = New cPRHist
    Set PRItem = New cPRItem
    Set PRTotal = New cPRTotal
    Set PRItemHist = New cPRItemHist
    Set PRDist = New cPRDist
    Set PRFWTTable = New cPRFWTTable
    Set PRGlobal = New cPRGlobal
    Set PRGLUpd = New cPRGLUpd

    Set GLAccount = New cGLAccount
    Set GLBatch = New cGLBatch
    Set GLCompany = New cGLCompany
    Set GLHistory = New cGLHistory
    Set GLJournal = New cGLJournal
    Set GLPrint = New cGLPrint

    Set PRCity = New cPRCity
    Set PRState = New cPRState
    Set JCJob = New cJCJob
    Set PRTimeSheet = New cPRTimeSheet
    
    ' Set PRW2Box = New cPRW2Box

    SetEquates
    
    OpenTab = 2
    
    X = Command()
    
    If X = "" Then         ' set for testing
       BalintFolder = "c:\Balint"
       BalintFolder = "\\vboxsrv\vm-share\Balint"
       dbPwd = ""
       PRBatchID = 0
       BatchNum = PRBatchID
       BatchNumber = PRBatchID
       ProgName = UCase("OHW2")                  ''''''  Select from cases below
       SysFile = "s:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       MenuName = GetCmd(X, "MenuName", "Str")
       PRBatchID = GetCmd(X, "Batch", "Num")
       Period = GetCmd(X, "Period", "Num")
       BalintFolder = GetCmd(X, "BalintFolder", "Str")
    End If
    
    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
    ' 2012-11-04 - replace ^ with space for folder name
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If
    
    ' new ADO?
    Dim NewFile As String
    NewFile = Replace(SysFile, ".mdb", ".accdb")
    If Len(Dir(NewFile, vbNormal)) Then
        SysFile = NewFile
        FileExt = ".accdb"
        modPRGlobal.NewADO = True
    Else
        FileExt = ".mdb"
        modPRGlobal.NewADO = False
    End If
    
    ' =========================================================================================
    ' check for required info
    If ProgName = "" Then
       MsgBox "Error - Program Name not given", vbExclamation, "PR Reports"
       End
    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbExclamation, "PR Reports"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbExclamation, "PR Reports"
       End
    End If
    ' =========================================================================================

    ' connect to the system data base
    If Not SysOpen(SysFile) Then
       MsgBox "Error connecting to: " & SysFile, vbExclamation, "PR Maintenance"
       End
    End If
    
    ' get the user record
    If Not User.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbExclamation, "PR Maintenance"
       End
    End If
    
    ' find the last GL company file id in PRCompany
    If (IsNull(User.LastCompany) Or User.LastCompany = 0) Then
       MsgBox "GLCompany ID not assigned ! ", vbExclamation, "PR Maintenance"
       End
    End If

    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.GLCompanyID = " & User.LastCompany
    If Not PRCompany.GetBySQL(SQLString) Then
        MsgBox "PRCompany.GLCompanyID record NF: " & User.LastCompany, vbExclamation
        GoBack
    End If

    ' open the company database
    If BalintFolder = "" Then
        X = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
        ' 2016-04-23
        X = "\Balint\Data\" & FNameOnly(PRCompany.FileName)
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(PRCompany.FileName)
    End If
    
    If FileExt = ".accdb" Then X = Replace(LCase(X), ".mdb", ".accdb")
    
    CNOpen X, dbPwd
    CompanyID = PRCompany.CompanyID
    
    ' open the GL Company
    If Not GLCompany.GetData(PRCompany.GLCompanyID) Then
        MsgBox "GLCompany ID record NF: " & PRCompany.GLCompanyID, vbCritical
        End
    End If
    
    ' 2018-01-05 - fix patch for Eaglowski - Altitude Mansfield
    ' -- fix PRItem.ItemType -- synch to EER
    ' -- EER item type was changed after it was added for the employee
    If PRCompany.CompanyID = 238 And InStr(LCase(PRCompany.Name), "mansfield") > 0 Then
        Dim rsEag As New ADODB.Recordset
        
        SQLString = "select EmployeeID from PRItem where EmployeeID <> 0 and EmployerItemID = 1 and ItemType = 4"
        rsInit SQLString, cn, rsEag
        If rsEag.RecordCount <> 0 Then
            SQLString = "update PRItem set ItemType = 5 where EmployeeID <> 0 and EmployerItemID = 1"
            cn.Execute SQLString
        End If
        
        SQLString = "select EmployeeID from PRItem where EmployeeID <> 0 and EmployerItemID = 2 and ItemType = 4"
        rsInit SQLString, cn, rsEag
        If rsEag.RecordCount <> 0 Then
            SQLString = "update PRItem set ItemType = 5 where EmployeeID <> 0 and EmployerItemID = 2"
            cn.Execute SQLString
        End If
        
    End If
    
    ' perform field sweeps - in NewField module
    FieldSweep
    
    Unload frmSplash
    
    Select Case ProgName
        Case "CHECKRECON"
            frmCheckRecon.Show
        Case "CHECKREG"
            frmCheckReg.Show
        Case "CITYLIST"
            frmCityList.Show
        Case "CITYTAX"
            frmCityTaxRpt.Show
        Case "DEPOSIT"
            frmDeposit.Show
        Case "DIRDEP"
            frmDirectDep.Show
        Case "EELIST"
            frmLists.Show
        Case "ENTRYFORM"
            frmEntry.Show
        Case "GLUPDATE"
            frmGLUpdate.Show
        Case "NEWHIRE"
            frmNewHire.Show
        Case "OHBUC"
            frmOHBUC.Show
        Case "QTRRPTS"
            frmPRQtrlyRpts.Show
        Case "TEST"
            Form1.Show
        Case "YECITYTAX"
            frmYECityTaxRpt.Show
        Case "WAGEBYJOB"
            OpenTab = 3
            frmWageByJob.Show
        Case "ERNDED"
            frmErnDed.Show
        Case "ITEMLISTING"
            frmItemListing.Show
        Case "CERTREG"
            frmCertReg.Show
        Case "DPTDIST"
            frmDptDist.Show
        Case "1099"
            frm1099.Show
        Case "FUTA940"
            frmFUTA940.Show
        Case "OHW2"
            frmOHW2.FileExt = FileExt
            frmOHW2.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            End
    End Select

End Sub

