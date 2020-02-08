Attribute VB_Name = "modStart"
Private Sub Main()

Dim X As String
Dim RC As Long

    frmSplash.Show
    frmSplash.MousePointer = vbHourglass

    Set Equate = New cEquate
    Set User = New cGLUser
    
    Set PRCity = New cPRCity
    Set PRCompany = New cPRCompany
    Set PRDepartment = New cPRDepartment
    Set PREmployee = New cPREmployee
    Set PREquate = New cPREquate
    Set PRItem = New cPRItem
    Set PRState = New cPRState
    Set PRW2Box = New cPRW2Box
    Set PRBatch = New cPRBatch
    Set PRHist = New cPRHist
    Set PRDist = New cPRDist
    Set PRGlobal = New cPRGlobal
    Set PRFWTTable = New cPRFWTTable
    Set PRItemHist = New cPRItemHist

    Set PRTotal = New cPRTotal
    Set PRGLUpd = New cPRGLUpd

    Set JCCustomer = New cJCCustomer
    Set JCJob = New cJCJob
    Set PRTimeSheet = New cPRTimeSheet
    Set QBAccount = New cQBAccount
    Set GLCompany = New cGLCompany

    SetEquates
    
    OpenTab = 2
    
    X = Command()
    FilterSw = 0
    LandSw = 0
    QtrEnding = ""

    If X = "" Then         ' SET FOR TESTING
        BalintFolder = "\\vboxsrv\vm-share\balint"
        dbPwd = ""
        ProgName = UCase("Form941")
        SysFile = "e:\Balint\Data\GLSystem.mdb"
        UserID = 2
        BackName = ""
        MenuName = ""
        jbFlag = True
    Else
        dbPwd = GetCmd(X, "dbPwd", "Str")
        ProgName = UCase(GetCmd(X, "ProgName", "Str"))
        SysFile = GetCmd(X, "SysFile", "Str")
        UserID = GetCmd(X, "UserID", "Num")
        BackName = GetCmd(X, "BackName", "Str")
        BalintFolder = GetCmd(X, "BalintFolder", "Str")
        MenuName = GetCmd(X, "MenuName", "Str")
        jbFlag = False
    
'        If MenuName <> "" Then BackName = MenuName
    End If
    
    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If
    
    ' =========================================================================================
    ' check for required info
'    If ProgName = "" Then
'       MsgBox "Error - Program Name not given", vbCritical, "PR Utilities"
'       End
'    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbCritical, "PR Utilities"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbCritical, "PR Utilities"
       End
    End If
    ' =========================================================================================

    ' connect to the system data base
    If Not SysOpen(SysFile) Then
       MsgBox "Error connecting to: " & SysFile, vbCritical, "PR Print"
       End
    End If
    
    ' get the user record
    If Not User.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbCritical, "PR Print"
       End
    End If
    
    ' find the last GL company file id in PRCompany
    If (IsNull(User.LastCompany) Or User.LastCompany = 0) Then
       MsgBox "GLCompany ID not assigned ! ", vbExclamation, "PR Maintenance"
       End
    End If

'    ' ***** R&C fix - 07/07/2010 *****
'    If PRCompany.GetByID(64) Then
'        If InStr(1, LCase(PRCompany.Name), "tobacco") Then
'            PRCompany.GLCompanyID = 259
'            PRCompany.Save (Equate.RecPut)
'            For RC = 65 To 72
'                SQLString = "DELETE * FROM PRCompany WHERE CompanyID = " & RC
'                cnDes.Execute SQLString
'            Next RC
'        End If
'    End If
'    ' ***** R&C fix - 07/07/2010 *****
    
'    ' ***** R&C fix - 07/15/2010 *****
'    If PRCompany.GetByID(73) Then
'        If InStr(1, LCase(PRCompany.Name), "rolling") Then
'            PRCompany.GLCompanyID = 260
'            PRCompany.Save (Equate.RecPut)
'            For RC = 74 To 77
'                SQLString = "DELETE * FROM PRCompany WHERE CompanyID = " & RC
'                cnDes.Execute SQLString
'            Next RC
'        End If
'    End If
'    ' ***** R&C fix - 07/07/2010 *****

'    ' ***** R&C fix - 01/20/2011 *****
'    Dim rcc As Integer
'    If PRCompany.GetByID(83) Then
'        If InStr(1, LCase(PRCompany.Name), "jfs") Then
'            PRCompany.GLCompanyID = 188
'            PRCompany.Save (Equate.RecPut)
'            For RC = 82 To 88
'                If RC <> 83 Then
'                    rcc = rcc + 1
'                    SQLString = "DELETE * FROM PRCompany WHERE CompanyID = " & RC
'                    cnDes.Execute SQLString
'                End If
'            Next RC
'            MsgBox "JFS sweep complete - # removed: " & rcc, vbInformation
'        End If
'    End If
'    ' ***** R&C fix - 07/07/2010 *****
    
'    ' ***** R&C fix - 01/24/2011 *****
'    If PRCompany.GetByID(90) Then
'        If InStr(1, LCase(PRCompany.Name), "lazar") Then
'            rcc = 0
'            PRCompany.GLCompanyID = 124
'            PRCompany.Save (Equate.RecPut)
'            For RC = 91 To 96
'                rcc = rcc + 1
'                SQLString = "DELETE * FROM PRCompany WHERE CompanyID = " & RC
'                cnDes.Execute SQLString
'            Next RC
'            MsgBox "Lazar sweep complete - # removed: " & rcc, vbInformation
'        End If
'    End If
'    ' ***** R&C fix - 07/07/2010 *****

    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.GLCompanyID = " & User.LastCompany
    If Not PRCompany.GetBySQL(SQLString) Then
        MsgBox "PRCompany record NF: " & User.LastCompany, vbExclamation
        GoBack
    End If
    
    SQLString = "SELECT * FROM GLCompany WHERE ID = " & User.LastCompany
    If Not GLCompany.GetBySQL(SQLString) Then
        MsgBox "GLCompany record NF: " & User.LastCompany, vbExclamation
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

    ' HC city switch sweep
    If LCase(Mid(PRCompany.Name, 1, 9)) = "hernandez" And PRCompany.FederalID = "20-3413203" Then
        SQLString = "SELECT * FROM PRDist WHERE CityID = 46 and YearMonth >= 201001 and YearMonth <= 201012"
        If PRDist.GetBySQL(SQLString) = True Then
            Do
                PRDist.CityID = 82
                PRDist.Save (Equate.RecPut)
                If PRDist.GetNext = False Then Exit Do
            Loop
            MsgBox "Hernandez: Cleveland moved to Clv Hts", vbInformation
        End If
    End If

    ' perform field sweeps - in NewField module
    FieldSweep

    ' *** frmEntry.Show ***
    frmBatchList.Show

End Sub


