Attribute VB_Name = "modStart"
Private Sub Main()

Dim x As String
Dim NewFlag As Boolean

    frmSplash.Show

    Set Equate = New cEquate
    Set User = New cGLUser
    Set PRCompany = New cPRCompany
    Set PREquate = New cPREquate
    Set PRGlobal = New cPRGlobal
    Set GLCompany = New cGLCompany
    Set PRFWTTable = New cPRFWTTable

    Set JCCustomer = New cJCCustomer
    Set JCJob = New cJCJob
    Set QBAccount = New cQBAccount
    Set QBUpdate = New cQBUpdate

    Set InvBody = New cInvBody
    Set InvEquate = New cInvEquate
    Set InvHeader = New cInvHeader
    Set InvStock = New cInvStock
    Set InvGlobal = New cInvGlobal
    
    Set PRCity = New cPRCity
    Set PRState = New cPRState
    
    Set Notes = New cNotes
    
    SetEquates
    InvSetEquates
    
    OpenTab = 4
    
    x = Command()
    
    If x = "" Then         ' set for testing
       dbPwd = ""
       
        Dim Sel As Integer
'        X = "1=Process / 2=Stock / 3=Global / 4=GlobalQB"
'        Sel = InputBox(X)
        
        Sel = 1
        
        Select Case Sel
            Case 0:     End
            Case 1:     ProgName = UCase("process")
            Case 2:     ProgName = UCase("stockmaint")
            Case 3:     ProgName = UCase("global")
            Case 4:     ProgName = UCase("globalqb")
            Case 5:     ProgName = UCase("qbjob")
        End Select
       'ProgName = UCase("test2")
       'ProgName = UCase("custmsg")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       SysFile = "\\vboxsrv\vm-share\balint\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 1
       BatchNumber = 1
       PRBatchID = 1
       Period = 0         ' yyyypp
       BalintFolder = "\\vboxsrv\vm-share\balint\"
    Else
       dbPwd = GetCmd(x, "dbPwd", "Str")
       ProgName = UCase(GetCmd(x, "ProgName", "Str"))
       SysFile = GetCmd(x, "SysFile", "Str")
       UserID = GetCmd(x, "UserID", "Num")
       BackName = GetCmd(x, "BackName", "Str")
       BatchNum = GetCmd(x, "Batch", "Num")
       Period = GetCmd(x, "Period", "Num")
       BalintFolder = GetCmd(x, "BalintFolder", "Str")
    End If
    
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
        PRCompany.Save (Equate.RecAdd)
                
        NewFlag = True
    
    End If

'    ' get the company record from the system data base
'    If User.LastPRCompany <> 0 Then
'        If Not PRCompany.GetBySQL("SELECT * FROM PRCompany WHERE PRCompany.CompanyID = " & CStr(User.LastPRCompany)) Then
'           MsgBox "Company ID not found ! " & User.LastPRCompany, vbExclamation, "PR Maintenance"
'           End
'        End If
'    End If

    ' open the company database
    ' open the company database
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
    Else
        x = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(PRCompany.FileName)
    End If
    
    If NewADO Then
        x = Replace(x, ".mdb", ".accdb")
    Else
        x = Replace(x, ".accdb", ".mdb")
    End If
    
    
    CNOpen x, dbPwd
    CompanyID = PRCompany.CompanyID
    
    ' ***********************************
'    InvGlobalCreate True
'    StockCreate True
'    HeaderCreate True
'    BodyCreate True
    ' ***********************************
    
    If TableExists("JCCustomer", cn) = False Then CustomerCreate
    If TableExists("JCJob", cn) = False Then JobCreate
    If TableExists("InvStock", cn) = False Then StockCreate
    If TableExists("InvHeader", cn) = False Then HeaderCreate
    If TableExists("InvBody", cn) = False Then BodyCreate
    If TableExists("InvGlobal", cnDes) = False Then InvGlobalCreate
    
    ' ========================================================
    ' clear invoices not update to QB
'    SQLString = "SELECT * FROM InvHeader WHERE InvoiceDate <> 0"
'    boo = InvHeader.GetBySQL(SQLString)
'    If boo = True Then
'        Do
'            If InvHeader.QBInvoiceID = "" Then
'                InvHeader.InvoiceDate = 0
'                InvHeader.rsPut
'            End If
'            If InvHeader.GetNext = False Then Exit Do
'        Loop
'    End If
    ' ========================================================
    
    Unload frmSplash
    
    ' ==================================
'    Dim InvStk0 As New cInvStock
'    Dim InvStkJ As New cInvStock
'    Dim ch As Integer
'
'    ch = FreeFile
'    Open "f:\asend\kp.txt" For Output As #ch
'
'    SQLString = "SELECT * FROM InvStock WHERE JobID = 0"
'    If InvStk0.GetBySQL(SQLString) = True Then
'        Do
'            SQLString = "SELECT * FROM InvStock WHERE QBID = '" & InvStk0.QBID & "'" & _
'                        " AND JobID <> 0 ORDER BY JobID"
'            If InvStkJ.GetBySQL(SQLString) = True Then
'                Do
'                    Print #ch, InvStk0.QBID & " " & InvStkJ.QBID & " " & InvStkJ.JobID
'                    If InvStkJ.GetNext = False Then Exit Do
'                Loop
'            End If
'            If InvStk0.GetNext = False Then Exit Do
'        Loop
'    End If
'    End
'
    ' ==================================
    
    ' stock file fix
    If PRCompany.CompanyID = 4 Then
        SQLString = "SELECT * FROM InvGlobal WHERE Description = 'NB Stock Fix'"
        If InvGlobal.GetBySQL(SQLString) = False Then
            
            SQLString = "DELETE * FROM InvStock WHERE JobID >= 479"
            cn.Execute SQLString
            
            SQLString = "INSERT INTO InvGlobal (Description) VALUES ('NB Stock Fix')"
            cnDes.Execute SQLString
        
        End If
    End If
    
'SQLString = "select * from InvStock where StockID = 16769"
'If InvStock.GetBySQL(SQLString) Then
'    MsgBox (InvStock.JobID)
'End If
'
'SQLString = "select * from JCJob where JobID = 33"
'If JCJob.GetBySQL(SQLString) Then
'    MsgBox (JCJob.CompanyName)
'End If
'
'SQLString = "select * from InvHeader where InvoiceNumber = 34530"
'If InvHeader.GetBySQL(SQLString) Then
'    MsgBox (InvHeader.SoldJobID)
'End If
'
'End

    Select Case ProgName
        Case "PROCESS":         frmInvProcess.Show
        Case "STOCKMAINT":      frmInvStockMaint.Show
        Case "GLOBAL":          frmInvGlobalMaint.Show
        Case "CUSTMSG":         frmInvMessage.Show
        Case "GLOBALQB":        frmInvGlobalQB.Show
        Case "QBJOB"
            frmJCGetQBData.Show vbModal
            GoBack
    End Select
    
    'frmInvProcess.Show
    'frmInvStockMaint.Show
    'frmTest.Show
    'frmInvGlobalMaint.Show
    'frmInvFind.Show

'    JCJob.GetByID (6)
'    frmInvPriceLookup.JobID = 6
'    frmInvPriceLookup.Init
'    frmInvPriceLookup.Show vbModal
'    End

End Sub

