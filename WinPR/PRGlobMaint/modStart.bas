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

    Set PRCity = New cPRCity
    Set PRState = New cPRState
    ' Set PRW2Box = New cPRW2Box

    SetEquates
    OpenTab = 2
    X = Command()
    
    If X = "" Then         ' set for testing
       BalintFolder = "g:"
       dbPwd = ""
       ProgName = UCase("test")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 0
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       ProgName = "GlobMaint"       ' only one choice !!!
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       BatchNum = GetCmd(X, "Batch", "Num")
       MenuName = GetCmd(X, "MenuName", "Str")
       Period = GetCmd(X, "Period", "Num")
       BalintFolder = GetCmd(X, "BalintFolder", "Str")
    End If
    
    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
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
    CNOpen X, dbPwd
    CompanyID = PRCompany.CompanyID
    
'    ' open the GL Company
'    If Not GLCompany.GetData(PRCompany.GLCompanyID) Then
'        MsgBox "GLCompany ID record NF: " & PRCompany.GLCompanyID, vbCritical
'        End
'    End If
    
    ' perform field sweeps - in NewField module
    FieldSweep
    
    ' frmTest.Show
    frmGlobalMaint.Show

End Sub

