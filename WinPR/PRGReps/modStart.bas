Attribute VB_Name = "modStart"
Private Sub Main()

Dim x As String
Dim I As Long
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
    Set GLPrint = New cGLPrint
    Set GLCompany = New cGLCompany

    Set PRCity = New cPRCity
    Set PRState = New cPRState
    
    Set PRW2 = New cPRW2
    Set PRW2City = New cPRW2City
    Set PRW2State = New cPRW2State
    
    Set PRTimeSheet = New cPRTimeSheet
    Set JCJob = New cJCJob
    
    SetEquates
    
    ' ////////////////////////
    x = Command()
    
    If x = "" Then         ' set for testing
       BalintFolder = "\\vboxsrv\vm-share\balint"
       BalintFolder = "c:\Balint"
       dbPwd = ""
       ' ProgName = UCase("ITEMDETAIL")
       ProgName = UCase("test")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       ' UserID = 15
       BackName = ""
       MenuName = ""
       BatchNum = 155
       BatchNumber = BatchNum
       PRBatchID = BatchNum
       BatchNumbr = BatchNum
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(x, "dbPwd", "Str")
       ProgName = UCase(GetCmd(x, "ProgName", "Str"))
       SysFile = GetCmd(x, "SysFile", "Str")
       UserID = GetCmd(x, "UserID", "Num")
       BackName = GetCmd(x, "BackName", "Str")
       MenuName = GetCmd(x, "MenuName", "Str")
       BatchNum = GetCmd(x, "Batch", "Num")
       BatchNumber = BatchNum
       PRBatchID = BatchNum
       Period = GetCmd(x, "Period", "Num")
       BalintFolder = GetCmd(x, "BalintFolder", "Str")
    End If
    
    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
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
       MsgBox "Error - Program Name not given", vbExclamation, "PR Utilities"
       End
    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbExclamation, "PR Utilities"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbExclamation, "PR Utilities"
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
       MsgBox "PRCompany ID not assigned ! ", vbExclamation, "PR Maintenance"
       End
    End If

    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.GLCompanyID = " & User.LastCompany
    If Not PRCompany.GetBySQL(SQLString) Then
        MsgBox "PRCompany record NF: " & User.LastCompany, vbCritical
        End
    End If

'    ' get the company record from the system data base
'    If User.LastPRCompany <> 0 Then
'        If Not PRCompany.GetBySQL("SELECT * FROM PRCompany WHERE PRCompany.CompanyID = " & CStr(User.LastPRCompany)) Then
'           MsgBox "Company ID not found ! " & User.LastPRCompany, vbExclamation, "PR Maintenance"
'           End
'        End If
'    End If

    ' open the company database
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
        ' 2016-04-23
        x = "\Balint\Data\" & FNameOnly(PRCompany.FileName)
    Else
        x = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(PRCompany.FileName)
    End If
    
    If FileExt = ".accdb" Then x = Replace(LCase(x), ".mdb", ".accdb")
    
    CNOpen x, dbPwd
    CompanyID = PRCompany.CompanyID
    
    
    ' =========================================
    ' get the employee record
'    If PREmployee.GetByID(1) = False Then
'        MsgBox "---EmployeeID not found: " & PRW2.EmployeeID, vbExclamation
'        GoBack
'    End If
    
    ' =========================================
    
    ' perform field sweeps - in NewField module
    FieldSweep
    OpenTab = 2
    
    Unload frmSplash

    Select Case ProgName

        Case "TEST"
            frm941_2022_March.Show
        Case "CHECKPRINT"
            frmCheckPrint.Show
        Case "FORM941"
            With frm941_Select
                .Show vbModal
                I = .Form941
                Unload frm941_Select
                If I = 1 Then
                    frm941Entry.Show
                ElseIf I = 2 Then
                    frm941_2010A.Show
                ElseIf I = 3 Then
                    frm941_2011A.Show
                ElseIf I = 4 Then
                    frm941_2012A.Show
                ElseIf I = 5 Then
                    frm941_2013A.Show
                ElseIf I = 6 Then
                    ' added on 2014-07-18
                    frm941_2013A2.Show
                ElseIf I = 7 Then
                    ' added on 2014-08-26
                    frm941_2014.Show
                ElseIf I = 8 Then
                    ' added on 2017-04-08
                    frm941_2017.Show
                ElseIf I = 9 Then
                    ' added on 2020-07-15
                    frm941_2020_June.Show
                ElseIf I = 10 Then
                    ' added on 2021-08-04
                    frm941_2021_June.Show
                ElseIf I = 11 Then
                    ' added on 2022-04-09
                    frm941_2022_March.Show
                Else
                    GoBack
                End If
            End With
        Case "FORM941_2010A"
            frm941_2010A.Show
        Case "FORM941_2011A"
            frm941_2011A.Show
        Case "TAXWAGE"
            frmTaxWage.Show
        Case "TEST"
            form1.Show
        Case "BATCHLIST"
            frmBatchList.Show
        Case "ITEMDETAIL"
            frmItemDetail.Show
        Case "CHECKDETAIL"
            frmCheckDetail.Show
        Case "W2"
            frmW2.Show
        Case "EARNSUMMARY"
            frmEarnSumm.Show
        Case "TEST"
            form1.Show
        Case "LISTS"
            frmLists.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            End

    End Select
    
End Sub

