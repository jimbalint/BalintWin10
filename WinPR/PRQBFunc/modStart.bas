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

    Set PRCity = New cPRCity
    Set PRState = New cPRState
    Set GLPrint = New cGLPrint
    
    Set JCCustomer = New cJCCustomer
    Set JCJob = New cJCJob
    Set PRTimeSheet = New cPRTimeSheet
    Set QBAccount = New cQBAccount
    Set QBUpdate = New cQBUpdate

    Set PRCounty = New cPRCounty

    Set Notes = New cNotes

    SetEquates
    
    OpenTab = 2
    
    X = Command()
    
    If X = "" Then         ' set for testing
       BalintFolder = "g:"
       dbPwd = ""
       ProgName = UCase("taxpay")
       ' ProgName = UCase("test2")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 202
       BatchNumber = BatchNum
       PRBatchID = BatchNum
       BatchNumbr = BatchNum
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       BatchNum = GetCmd(X, "Batch", "Num")
       BatchNumber = BatchNum
       PRBatchID = BatchNum
       BatchNumbr = BatchNum
       Period = GetCmd(X, "Period", "Num")
       BalintFolder = GetCmd(X, "BalintFolder", "Str")
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
    Else
        FileExt = ".mdb"
    End If
    
    ' *** force date range screen for TaxPay ***
    If ProgName = "TAXPAY" Then
        BatchNum = 0
        BatchNumber = 0
        BatchNumbr = 0
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
        MsgBox "PRCompany record NF: " & User.LastCompany, vbExclamation
        GoBack
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
        X = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
        ' 2016-04-23
        X = "\Balint\Data\" & FNameOnly(PRCompany.FileName)
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(PRCompany.FileName)
    End If
    
    If FileExt = ".accdb" Then X = Replace(LCase(X), ".mdb", ".accdb")
    
    CNOpen X, dbPwd
    CompanyID = PRCompany.CompanyID
    
    If TableExists("QBAccount", cn) = False Then
        QBAccountCreate
    End If
    
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

        Case "TEST"
            Form1.Show
        Case "TAXPAY"
            frmTaxPay.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            End

    End Select
    
    ' frmEmpList.Show
    ' frmCompany.Show
    ' frmEmpForm.Show
    ' Form1.Show

End Sub


