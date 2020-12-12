Attribute VB_Name = "modStart"
Private Sub Main()   ' *** project execution starts here ***

Dim x As String
Dim AcctDesc As Byte

    Set GLCompany = New cGLCompany
    Set GLAccount = New cGLAccount
    Set GLAmount = New cGLAmount
    Set GLHistory = New cGLHistory
    Set GLDescription = New cGLDescription
    Set Equate = New cEquate
    Set GLPrint = New cGLPrint
    Set GLBatch = New cGLBatch
    Set GLUser = New cGLUser
    Set GLJournal = New cGLJournal
    Set GLFFSched = New cGLFFSched
    Set GLFFColumn = New cGLFFColumn
    Set PREquate = New cPREquate
    Set PRGlobal = New cPRGlobal
    
    SetEquates

    OpenTab = 1

    x = Command()
    
    ' ---------------------------------------
    ' - Program List
    ' GLHistJnl
    ' ChartOfAccounts
    ' PrintDesc
    ' PrintGLAccount
    ' GLHistJnl
    ' DetailGL
    ' Statement
    ' TrialBal
    ' ---------------------------------------
    
    If x = "" Then         ' set for testing
        BalintFolder = "c:\Balint"
        BalintFolder = "\\vboxsrv\vm-share\balint"
        dbPwd = "golf"
        ProgName = UCase("statement")
        SysFile = "c:\Balint\Data\GLSystem.mdb"
        UserID = 2
        BackName = ""
        BatchNum = 0
        TestMode = True
    Else
        dbPwd = GetCmd(x, "dbPwd", "Str")
        ProgName = UCase(GetCmd(x, "ProgName", "Str"))
        SysFile = GetCmd(x, "SysFile", "Str")
        UserID = GetCmd(x, "UserID", "Num")
        BackName = GetCmd(x, "BackName", "Str")
        BatchNum = GetCmd(x, "Batch", "Num")
        MenuName = GetCmd(x, "MenuName", "Str")
        AcctDesc = GetCmd(x, "AcctDesc", "Num")
        BalintFolder = GetCmd(x, "BalintFolder", "Str")
        TestMode = False
    End If

    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If

    ' =========================================================================================
    ' check for required info
'    If ProgName = "" Then
'       MsgBox "Error - Program Name not given", vbCritical, "GL Utilities"
'       End
'    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbCritical, "GL Utilities"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbCritical, "GL Utilities"
       End
    End If
    ' =========================================================================================
    
    ' new ADO?
    Dim NewFile As String
    NewFile = Replace(SysFile, ".mdb", ".accdb")
    If Len(Dir(NewFile, vbNormal)) Then
        SysFile = NewFile
        FileExt = ".accdb"
        NewADO = True
    Else
        FileExt = ".mdb"
        NewADO = False
    End If
    
    ' connect to the system data base
    If Not CNDesOpen(SysFile) Then
       MsgBox "Error connecting to: " & SysFile, vbCritical, "GL Utilities"
       End
    End If
    
    ' get the user record
    If Not GLUser.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbCritical, "GL Utilities"
       End
    End If
    
    ' use the last company id
    If IsNull(GLUser.LastCompany) Or GLUser.LastCompany = 0 Then
       MsgBox "Company ID not assigned ! ", vbCritical, "GL Utilities"
       End
    End If
    
    ' get the company record from the system data base
    If Not GLCompany.GetData(GLUser.LastCompany) Then
       MsgBox "Company ID not found ! " & GLUser.LastCompany, vbCritical, "GL Utilities"
       End
    End If
       
    ' open the company database
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    Else
        x = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(GLCompany.FileName)
    End If
    
    If NewADO Then
        x = Replace(x, ".mdb", ".accdb")
    Else
        x = Replace(x, ".accdb", ".mdb")
    End If
    
    CNOpen x, dbPwd

    CompanyID = GLUser.LastCompany

    MainMenu.Show

End Sub
