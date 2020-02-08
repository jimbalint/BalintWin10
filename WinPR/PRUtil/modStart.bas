Attribute VB_Name = "modStart"
Private Sub Main()

Dim X As String
Dim FName, DriveLetter As String

    ' save "C:"
    DriveLetter = Left(App.Path, 2)

    Set Equate = New cEquate
    Set User = New cGLUser
    Set PRCompany = New cPRCompany
    Set PRDepartment = New cPRDepartment
    Set PREmployee = New cPREmployee
    Set PREquate = New cPREquate

    Set PRCity = New cPRCity
    Set PRState = New cPRState

    Set PRHist = New cPRHist
    Set PRDist = New cPRDist
    Set PRItem = New cPRItem
    Set PRItemHist = New cPRItemHist

    Set PRAdjust = New cPRAdjust
    Set PRBatch = New cPRBatch
    Set PREELists = New cPREELists
    Set PRGlobal = New cPRGlobal
    Set PRFWTTable = New cPRFWTTable
    Set PRGLUpd = New cPRGLUpd

    Set GLCompany = New cGLCompany
    Set PRCounty = New cPRCounty
    Set JCCustomer = New cJCCustomer
    Set JCJob = New cJCJob
    
    ' Set PRHistTotal = New cPRHistTotal

    SetEquates
    
    OpenTab = 2
    
    X = Command()
    
    If X = "" Or X = "prcompany" Then         ' set for testing
        dbPwd = ""
        ProgName = UCase("prhist")
        ' ProgName = "x"
        SysFile = "c:\Balint\Data\GLSystem.mdb"
        SysFile = "\Balint\Data\GLSystem.mdb"
        UserID = 2
        BackName = ""
        BatchNum = 0
        Period = 0         ' yyyypp
        TextFileName = "c:\balint\data\PRH11901.txt"
        dbName = "c:\balint\data\ZBARCOSECURITY.mdb"
    Else
        dbPwd = GetCmd(X, "dbPwd", "Str")
        ProgName = UCase(GetCmd(X, "ProgName", "Str"))
        SysFile = GetCmd(X, "SysFile", "Str")
        UserID = GetCmd(X, "UserID", "Num")
        BackName = GetCmd(X, "BackName", "Str")
        BatchNum = GetCmd(X, "Batch", "Num")
        MenuName = GetCmd(X, "MenuName", "Str")
        Period = GetCmd(X, "Period", "Num")
        TextFileName = GetCmd(X, "txtName", "Str")
        dbName = GetCmd(X, "dbName", "Str")
    End If
    
    ' non-standard folders
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If
    
    ' =========================================================================================
    ' check for required info
    If ProgName = "" Then
       MsgBox "Error - Program Name not given", vbCritical, "PR Utilities"
       End
    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbCritical, "PR Utilities"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbCritical, "PR Utilities"
       End
    End If
    
    If TextFileName = "" Then
       MsgBox "Error - TextFileName not given", vbCritical, "PR Utilities"
       End
    End If
    
    If dbName = "" Then
       MsgBox "Error - dbName not given", vbCritical, "PR Utilities"
       End
    End If
    
    ' =========================================================================================

    ' connect to the system data base
    If Not SysOpen(SysFile) Then
       MsgBox "Error connecting to: " & SysFile, vbCritical, "PR Utilities"
       End
    End If
        
    ' create the PRCompany table in GLSystem.MDB and exit
    If X = "prcompany" Then
        GlobalCreate
        CompanyCreate
        MsgBox "PRCompany Created", vbInformation
        End
    End If
        
    ' get the user record
    If Not User.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbCritical, "PR Utilities"
       End
    End If
    
'    ' use the last company id
'    If (IsNull(User.LastCompany) Or User.LastCompany = 0) Then
'       MsgBox "Company ID not assigned ! ", vbCritical, "PR Utilities"
'       End
'    End If
'
'    ' get the company record from the system data base
'    If User.LastCompany <> 0 Then
'        If Not PRCompany.GetByID(User.LastPRCompany) Then
'           MsgBox "Company ID not found ! " & GLUser.LastCompany, vbCritical, "GL Utilities"
'           End
'        End If
'    End If

    ' open the company database
    ' use from the MDB file instead of the command line
    ' problem with MDB file names with spaces
    If Not GLCompany.GetData(User.LastCompany) Then
        MsgBox "GL CompanyID not found! " & User.LastCompany, vbExclamation
        End
    End If
    dbName = GLCompany.FileName
    FName = DriveLetter & Mid(dbName, 3, Len(dbName) - 2)
    
    CNOpen FName, dbPwd
    
    ' frmTest.Show
    If ProgName = "JOB" Then
        cn.Execute "DROP TABLE JCCustomer"
        If TableExists("JCCustomer", cn) = False Then
            CustomerCreate
            MsgBox "JCCustomer table created", vbInformation
        End If
        cn.Execute "DROP TABLE JCJob"
        If TableExists("JCJob", cn) = False Then
            JobCreate
            MsgBox "JCJob table created", vbInformation
        End If
        GoBack
    Else
        frmStart.Show
    End If

End Sub
