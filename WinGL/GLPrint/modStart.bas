Attribute VB_Name = "modStart"
Private Sub Main()   ' *** project execution starts here ***

Dim X As String
Dim AcctDesc As Byte
Dim bProgName As String

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

    X = Command()
    
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
    
    If X = "" Then         ' set for testing
       BalintFolder = "g:"
       dbPwd = "golf"
       ProgName = UCase("statement")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       UserID = 10      ' mary - richlak
       BackName = ""
       BatchNum = 0
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       BatchNum = GetCmd(X, "Batch", "Num")
       MenuName = GetCmd(X, "MenuName", "Str")
       AcctDesc = GetCmd(X, "AcctDesc", "Num")
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
        X = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(GLCompany.FileName)
    End If
    CNOpen X, dbPwd

    CompanyID = GLUser.LastCompany

' frmAcctLookup.Show vbModal
' End

    ' handle call to print jnl for a batch
    If ProgName = "GLHISTJNL" And BatchNum <> 0 Then
                 
        If BalintFolder = "" Then
            If InStr(1, BackName, "GLEntryADO.exe", vbTextCompare) Then
                BackName = "\Balint\GLEntryADO.exe"
            Else
                BackName = "\Balint\GLEntry.exe"
            End If
        Else
            If InStr(1, BackName, "GLEntryADO.exe", vbTextCompare) Then
                BackName = "c:\Balint\GLEntryADO.exe"
            Else
                BackName = "c:\Balint\GLEntry.exe"
            End If
        End If
        
         ' get the horz nudge
         TabValue = 0
         SQLString = " SELECT * FROM PRGlobal WHERE Description = 'GLTab' " & _
                     " AND UserID = " & GLUser.ID
         If PRGlobal.GetBySQL(SQLString) = True Then
            TabValue = PRGlobal.Var1
         End If
        
        GLHistJnl 0, 0, 0, 0, BatchNum, AcctDesc
        Prvw.vsp.EndDoc
        Prvw.Show vbModal
        GoBack
    ElseIf ProgName = "FREEFORMAT" Then
        ' Form1.Show
        frmFreeFormat.Show
    ElseIf ProgName = "TEST2" Then
        Form2.Show
    Else
       frmGLPrint.Show
    End If

End Sub
