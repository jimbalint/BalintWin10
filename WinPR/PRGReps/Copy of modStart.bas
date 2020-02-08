Attribute VB_Name = "modStart"
Private Sub Main()

Dim X As String

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

    Set PRCity = New cPRCity
    Set PRState = New cPRState
    Set PRW2Box = New cPRW2Box

    SetEquates
    
    X = Command()
    
    If X = "" Then         ' set for testing
       dbPwd = ""
       ProgName = UCase("form941")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 14
       BatchNumber = 14
       PRBatchID = 14
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       BatchNum = GetCmd(X, "Batch", "Num")
       PRBatchID = BatchNum
       Period = GetCmd(X, "Period", "Num")
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
    If Not User.GetSQL("SELECT * FROM Users WHERE ID = " & UserID) Then
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
    X = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
    CNOpen X, dbPwd
    CompanyID = PRCompany.CompanyID
    
    ' perform field sweeps - in NewField module
    FieldSweep
    
 
'''' ****************************************
'''EmpID = 1
'''frmTaxWage.Show vbModal
'''End
'''' ****************************************
    
    Select Case ProgName

        Case "CHECKPRINT"
            frmCheckPrint.Show
        Case "FORM941"
            frm941Entry.Show
        Case "TEST"
            Form1.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            End

    End Select
    
    ' frmEmpList.Show
    ' frmCompany.Show
    ' frmEmpForm.Show
    ' Form1.Show

End Sub

