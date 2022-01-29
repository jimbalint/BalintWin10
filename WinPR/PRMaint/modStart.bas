Attribute VB_Name = "modStart"
Private Sub Main()

Dim x As String
Dim NewFlag As Boolean
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
    Set PRW2Box = New cPRW2Box
    
    Set JCCustomer = New cJCCustomer
    Set JCJob = New cJCJob
    Set PRTimeSheet = New cPRTimeSheet
    Set QBAccount = New cQBAccount
    Set QBUpdate = New cQBUpdate

    Set PRCounty = New cPRCounty

    Set Notes = New cNotes

    SetEquates
    
    OpenTab = 2
    
    x = Command()
    
    If x = "" Then         ' set for testing
       BalintFolder = "\\vboxsrv\vm-share\Balint"
       BalintFolder = "c:\Balint"
       dbPwd = ""
       ProgName = UCase("EMPLOYER")
       ' ProgName = UCase("test2")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 1
       BatchNumber = 1
       PRBatchID = 1
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(x, "dbPwd", "Str")
       ProgName = UCase(GetCmd(x, "ProgName", "Str"))
       SysFile = GetCmd(x, "SysFile", "Str")
       UserID = GetCmd(x, "UserID", "Num")
       BackName = GetCmd(x, "BackName", "Str")
       BatchNum = GetCmd(x, "Batch", "Num")
       MenuName = GetCmd(x, "MenuName", "Str")
       Period = GetCmd(x, "Period", "Num")
       BalintFolder = GetCmd(x, "BalintFolder", "Str")
    End If
    
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

    ' no PRCompany info - check for other files also ....
    NewFlag = False
    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.GLCompanyID = " & User.LastCompany
    If Not PRCompany.GetBySQL(SQLString) Then
        If GLCompany.GetData(User.LastCompany) = False Then
            MsgBox "GL Company NF: ", vbExclamation
            GoBack
        End If
        
        PRCompany.Clear
        PRCompany.Name = GLCompany.Name
        PRCompany.FileName = GLCompany.FileName
        PRCompany.GLCompanyID = GLCompany.ID
        PRCompany.Save (Equate.RecAdd)
                
        NewFlag = True
    
        ' update GLCompany
    
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
    
    If TableExists("PRHist", cn) = False Then HistCreate
    If TableExists("PRItem", cn) = False Then ItemCreate
    If TableExists("PRBatch", cn) = False Then PRBatchCreate
    If TableExists("PRDepartment", cn) = False Then DepartmentCreate
    If TableExists("PRDist", cn) = False Then DistCreate
    If TableExists("PREmployee", cn) = False Then EmployeeCreate
    If TableExists("PRGLUpd", cn) = False Then GLUpdCreate
    If TableExists("PRItemHist", cn) = False Then ItemHistCreate
    
    ' perform field sweeps - in NewField module
    FieldSweep
 
'''' ****************************************
'''EmpID = 1
'''frmTaxWage.Show vbModal
'''End
'''' ****************************************
    
    Unload frmSplash
    
 ' frmJCGetQBData.Show vbModal
 ' End
    
    
 'frmQBAccts.Show vbModal
 'End
 
' frmDeductBasis.EmployeeID = 0
' frmDeductBasis.ItemID = 13
' frmDeductBasis.Show vbModal
' End
    
    Select Case ProgName

        Case "EMPLOYEE"
            frmEmpList.Show
        Case "EMPLOYER"
            frmCompany.Show
        Case "CITY"
            frmPRCity.cmdSelect.Visible = False
            frmPRCity.Show
        Case "STATE"
            frmPRState.Show
        Case "DEPARTMENT"
            frmDepartment.Show
        Case "TEST"
            Form1.Show
        Case "GLUPD"
            frmGLUpd.Show
        Case "EESELECT"
            frmEmployeeSelect.Show
        Case "TAXSWEEP"
            frmTaxSweep.Show
        Case "ASSIGNCITY"
            frmAssignCity.Show
        Case "JCLIST"
            MsgBox "This form is not available", vbInformation
            GoBack
            frmJCList.Show
        Case "JCJOBMAINT"
            OpenTab = 3
            frmJCJobMaint.Show
        Case "JCGETQBDATA"
            OpenTab = 3
            frmJCGetQBData.Show
        Case "TIMESHEET"
            OpenTab = 3
            frmPRTimeSheet.Show
        Case "JOBUPDATE"
            OpenTab = 3
            frmQBJobUpdate.Show
        Case "TSPRINT"
            OpenTab = 3
            frmPRTSPrint.Show
        Case "TEST2"
            Form2.Show
        Case "COUNTY"
            frmPRCounty.Show
        Case "HISTIMPORT"
            frmPRHistImport.Show
        Case "QBREGISTER"
            frmQBRegister.Show
        Case "PWMAINT"
            frmPWMaint.Show
        Case "PURGE"
            frmPRPurge.Show
        Case "TEST3"
            Form3.Show
        Case "TEST4"
            Form4.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            End

    End Select
    
    ' frmEmpList.Show
    ' frmCompany.Show
    ' frmEmpForm.Show
    ' Form1.Show

End Sub

