Attribute VB_Name = "modStart"
' Public DBName As String


Private Sub Main()   ' *** project execution starts here ***

Dim x As String
Dim b As Long
Dim FileExt As String

    Set GLCompany = New cGLCompany
    Set GLAccount = New cGLAccount
    Set GLAmount = New cGLAmount
    Set GLHistory = New cGLHistory
    Set GLDescription = New cGLDescription
    Set Equate = New cEquate
    Set GLPrint = New cGLPrint
    Set GLBatch = New cGLBatch
    Set GLUser = New cGLUser
    Set GLBranch = New cGLBranch
    Set GLJournal = New cGLJournal
    Set PREquate = New cPREquate
    Set PRGlobal = New cPRGlobal
    Set GLFFSched = New cGLFFSched
    Set GLFFColumn = New cGLFFColumn

    SetEquates
    
    x = Command()
    
    ' go back to GL tab
    OpenTab = 1
    
    If x = "" Then         ' set for testing
       BalintFolder = ""
       dbPwd = ""
       ProgName = UCase("account")
       ' ProgName = UCase("ffschedule")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 3
    Else
       dbPwd = GetCmd(x, "dbPwd", "Str")
       ProgName = UCase(GetCmd(x, "ProgName", "Str"))
       SysFile = GetCmd(x, "SysFile", "Str")
       UserID = GetCmd(x, "UserID", "Num")
       BackName = GetCmd(x, "BackName", "Str")
       BatchNum = GetCmd(x, "Batch", "Num")
       MenuName = GetCmd(x, "MenuName", "Str")
       BalintFolder = GetCmd(x, "BalintFolder", "Str")
    End If

    If SysFile = "" Then SysFile = "\Balint\Data\GLSystem.mdb"
    
    ' non-standard folders
    If BalintFolder <> "" Then
        SysFile = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
    End If

    ' =========================================================================================
    ' check for required info
    If ProgName = "" Then
       MsgBox "Error - Program Name not given", vbCritical, "GL Utilities"
       End
    End If
    
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
    ' not needed if using user maint
    If ProgName <> "USER" Then
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
            ' 2016-04-23
            x = "\Balint\Data\" & mdbName(GLCompany.FileName)
        Else
            x = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(GLCompany.FileName)
        End If
        
        If NewADO Then
            x = Replace(x, ".mdb", ".accdb")
        Else
            x = Replace(x, ".accdb", ".mdb")
        End If
        
        dbName = x
        CNOpen x, dbPwd
        CompanyID = GLUser.LastCompany
    
    End If
        
    ' execute the call
    Select Case ProgName
       
        Case "ACCOUNT"
            frmAccount.Show
        Case "JOURNAL"
            frmJournal.Show
        Case "USER"
            frmUsers.Show
        Case "DESCRIPTIONS"
            frmDescriptions.Show
        Case "COMPANY"
            CompanyForm.Show
        Case "FFSCHEDULE"
            frmFFSchedule.Show
        Case "FFCOLUMN"
            frmFFColumn.Show
        Case "ACCOUNTCHANGE"
            frmAccountChange.Show
    
    End Select

End Sub
