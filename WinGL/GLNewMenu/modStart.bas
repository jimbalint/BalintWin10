Attribute VB_Name = "modStart"
' Public DBName As String


Private Sub Main()   ' *** project execution starts here ***

Dim x As String
Dim b As Long
Dim I, J, K As Long

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
    Set PRCompany = New cPRCompany
    Set PREmployee = New cPREmployee
    Set PRState = New cPRState
    Set PREquate = New cPREquate
    Set PRGlobal = New cPRGlobal
    Set PRHist = New cPRHist
    Set PRDist = New cPRDist
    Set PRCity = New cPRCity
    Set PRFWTTable = New cPRFWTTable

    SetEquates

    x = Command()
        
    ' for non-standard paths
    ' location for the Balint folder
    On Error Resume Next
    Open "C:\Balint\Init.txt" For Input As #1
    If Err.Number = 0 Then
        Line Input #1, BalintFolder
        Close #1
    Else
        BalintFolder = ""
    End If
    
    If BalintFolder = "" Then
        If CNDesOpen("\Balint\Data\GLSystem.mdb") = False Then
            MsgBox "error opening \Balint\Data\GLSystem.mdb", vbCritical
            End
        End If
    Else
        Dim SysFl As String
        SysFl = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
        If CNDesOpen(Trim(SysFl)) = False Then
            ' eag test - 20190620
            MsgBox "*** Error opening: " & SysFl, vbCritical
            ' >>>> End
        End If
    End If
    
    If x = "" Then         ' open sys file and login
        SysFile = "\Balint\Data\GLSystem.mdb"
        OpenTab = 5
        frmLogin.Show vbModal
        If Response = False Then End
    
'        UserID = 2
'        SQLString = "SELECT * FROM Users WHERE Logon = 'jim'"
'        If Not GLUser.GetBySQL(SQLString) Then End
    
    Else                    ' back from menu selection
        dbPwd = GetCmd(x, "dbPwd", "Str")
        ProgName = UCase(GetCmd(x, "ProgName", "Str"))
        SysFile = GetCmd(x, "SysFile", "Str")
        UserID = GetCmd(x, "UserID", "Num")
        BackName = GetCmd(x, "BackName", "Str")
        BatchNum = GetCmd(x, "Batch", "Num")
        OpenTab = GetCmd(x, "OpenTab", "Num")
        BalintFolder = GetCmd(x, "BalintFolder", "Str")
        SysFile = "\Balint\Data\GLSystem.mdb"
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

'    ' connect to the system data base
'    If Not CNDesOpen(SysFile) Then
'       MsgBox "Error connecting to: " & SysFile, vbCritical, "GL Utilities"
'       End
'    End If
    
    ' get the user record
    If Not GLUser.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbCritical, "GL Utilities"
       End
    End If
    
    ' use the last company id
    ' not needed if using user maint
            
'    If IsNull(GLUser.LastCompany) Or GLUser.LastCompany = 0 Then
'        MsgBox "Company ID not assigned ! ", vbCritical, "GL Utilities"
'        End
'    End If
    
    If IsNull(GLUser.LastCompany) = False And GLUser.LastCompany <> 0 Then
    
        ' get the company record from the system data base
        If Not GLCompany.GetData(GLUser.LastCompany) Then
            GLUser.LastCompany = 0
            GLUser.LastPRCompany = 0
            GLUser.Save (Equate.RecPut)
            frmMainMenu.lblCompanyName = "No Company Loaded"
           
        Else
            
            ' open the company database
            If BalintFolder = "" Then
                x = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
                DBName = x
                
            Else
                
                ' get the .mdb file name
                ' from the right until the first "\" in found
                K = Len(GLCompany.FileName)
                For I = K To 1 Step -1
                    If Mid(GLCompany.FileName, I, 1) = "\" Then
                        Exit For
                    End If
                Next I
                If I = 0 Then
                    MsgBox "Error in company database name: " & GLCompany.FileName, vbExclamation
                    End
                End If
                x = Replace(BalintFolder, "^", " ") & "\Data\" & Mid(GLCompany.FileName, I + 1, K)
                DBName = x
            
            End If
                
            CNOpen x, dbPwd
            CompanyID = GLUser.LastCompany
            
            ' ??? needed for menu lblCompany after file copy ???
            GLCompany.GetData (GLUser.LastCompany)
            frmMainMenu.lblCompanyName = GLCompany.Name
    
        End If
    
    Else
        
        frmMainMenu.lblCompanyName = "No Company Loaded"
        
    End If
        
    frmMainMenu.Show
    
    
'    ' execute the call
'    Select Case ProgName
'
'       Case "ACCOUNT"
'          frmAccount.Show
'       Case "JOURNAL"
'          frmJournal.Show
'       Case "USER"
'          frmUsers.Show
'       Case "DESCRIPTIONS"
'          frmDescriptions.Show
'
'    End Select

End Sub

'Public Function TableExists(ByVal TableName As String, _
'                            ByRef adoConn As ADODB.Connection) _
'                            As Boolean
'
'Dim cm As ADODB.Command
'Dim frs As ADODB.Recordset
'Dim FldFlag As Boolean
'Dim fString As String
'
'    ' see if the field is already in the Table
'    Set frs = New ADODB.Recordset
'    frs.CursorLocation = adUseClient
'    frs.CursorType = adOpenStatic
'    frs.LockType = adLockBatchOptimistic
'    Set frs = adoConn.OpenSchema(adSchemaColumns)
'
'    TableExists = False
'
'    Do Until frs.EOF = True
'
'        If frs!Table_Name = TableName Then
'            TableExists = True
'            Exit Do
'        End If
'
'       frs.MoveNext
'
'   Loop
'
'End Function



