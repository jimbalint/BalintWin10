Attribute VB_Name = "modStart"
Private Sub Main()

Dim X As String
Dim NewFlag As Boolean
Dim FileExt As String

    frmSplash.Show

    Set Equate = New clsEquate
    Set User = New cGLUser
    Set GLCompany = New cGLCompany
    Set PRCompany = New cPRCompany
    Set PRGlobal = New cPRGlobal
    Set Form99 = New clsForm99
    Set Field99 = New clsField99
    Set Payee99 = New clsPayee99
    Set Detail99 = New clsDetail99
    Set PRState = New cPRState

    SetEquates
    
    OpenTab = 5
    
    Dim CmdLine As String
    X = Command()
    CmdLine = X
    
    If CmdLine = "" Then         ' set for testing
       BalintFolder = "c:\Balint"
        ' BalintFolder = ""
       dbPwd = ""
       ProgName = UCase("payee")
       ' ProgName = UCase("test2")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 1
       BatchNumber = 1
       PRBatchID = 1
       Period = 0         ' yyyypp
    Else
       dbPwd = GetCmd(X, "dbPwd", "Str")
       ProgName = UCase(GetCmd(X, "ProgName", "Str"))
       SysFile = GetCmd(X, "SysFile", "Str")
       UserID = GetCmd(X, "UserID", "Num")
       BackName = GetCmd(X, "BackName", "Str")
       BatchNum = GetCmd(X, "Batch", "Num")
       Period = GetCmd(X, "Period", "Num")
       BalintFolder = GetCmd(X, "BalintFolder", "Str")
    End If
    
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
        mod99Global.NewADO = True
    Else
        FileExt = ".mdb"
        mod99Global.NewADO = False
    End If
    
    ' =========================================================================================
    ' check for required info
    If ProgName = "" Then
       MsgBox "Error - Program Name not given", vbExclamation, "Win 1099"
       End
    End If
    
    If SysFile = "" Then
       MsgBox "Error - System File Name not given", vbExclamation, "Win 1099"
       End
    End If

    If UserID = 0 Then
       MsgBox "Error - User ID not given", vbExclamation, "Win 1099"
       End
    End If
    ' =========================================================================================

    ' connect to the system data base
    If Not SysOpen(SysFile) Then
       MsgBox "Error connecting to: " & SysFile, vbExclamation, "Win 1099"
       End
    End If

    ' get the user record
    If Not User.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) Then
       MsgBox "User ID not found: " & UserID, vbExclamation, "Win 1099"
       End
    End If
    
    ' find the last GL company file id in glcompany
    If (IsNull(User.LastCompany) Or User.LastCompany = 0) Then
       MsgBox "User.LastCompany not assigned ! ", vbExclamation, "Win 1099"
       End
    End If

    ' get the company record from the system data base
    If Not GLCompany.GetData(User.LastCompany) Then
       MsgBox "Company ID not found ! " & User.LastCompany, vbCritical, "GL Utilities"
       End
    End If
       
    ' open the Win1099 DB - always MDB
    If BalintFolder = "" Then
        X = Mid(App.Path, 1, 2) & "\Balint\Data\Win1099.mdb"
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\Win1099.mdb"
    End If
    ' If FileExt = ".accdb" Then X = Replace(LCase(X), ".mdb", ".accdb")
    CN99Open X

    ' open the company database
    If BalintFolder = "" Then
        X = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(GLCompany.FileName)
    End If
    
    If FileExt = ".accdb" Then X = Replace(LCase(X), ".mdb", ".accdb")
    
    CNOpen X, dbPwd
    CompanyID = GLCompany.ID
    
    
' =========================================================
'    SQLString = "DROP TABLE Payee99"
'    cn.Execute SQLString
'    SQLString = "DROP TABLE Detail99"
'    cn.Execute SQLString
    
    ' *******************
    If TableExists("Payee99", cn) = False Then Payee99Create
    If TableExists("Detail99", cn) = False Then Detail99Create
    ' *******************
    
    ' =========
    ' create forms for the new year
    ' copies from the previous year
    ' parameter is the copy TO year
    ' creates in \balint\data
    ' OLD: copy to \Balint\Data_1099 for the install
    ' NEW: Payee Maint - Init New Year button has logic for new forms
    ' CopyForms 2021
    ' End
    
    
'    SQLString = "select * from Form99 order by TaxYear desc"
'    If Form99.GetBySQL(SQLString) Then
'        Do
'            MsgBox Form99.FormType & " " & Form99.TaxYear
'            If Not Form99.GetNext Then Exit Do
'        Loop
'    End If
'    End
    
'    SQLString = "select * from Field99 where FormType = 'NEC'"
'    If Field99.GetBySQL(SQLString) Then
'        Do
'            MsgBox Field99.FormType & " " & Field99.TaxYear & " " & Field99.FieldTitle
'            If Not Field99.GetNext Then Exit Do
'        Loop
'    End If
'    End
    
    
'    NewForm 2019, "MISC", "NEC"
'    End
  
    
'    End
' =========================================================
    
    ' perform field sweeps - in NewField module
    ' FieldSweep
 
    Unload frmSplash
        
'     frmPayeeList.Show vbModal
'     PrintForm99 "MISC", 2011, True
'      End
    ' ============================================
    

' **********************************************************
' *** use for test runs ***
'If CmdLine = "" Then
'    HorzNudge = 4
'    VertNudge = 4
'    Create2021Forms "NEC"
'    PrintForm99 "NEC", 2021, True
'    End
'End If

' **********************************************************
    Select Case ProgName

        Case "CREATE"
''''            ' **** Win1099.mdb distributed w/ install ****
''''            SQLString = "DROP TABLE Form99"
''''            cn99.Execute SQLString
''''            SQLString = "DROP TABLE Field99"
''''            cn99.Execute SQLString
''''            If TableExists("Form99", cn99) = False Then Form99Create
''''            If TableExists("Field99", cn99) = False Then
''''                Field99Create
''''                Create2011Forms
''''            End If
''''            MsgBox "1099 Forms for 2011 have been created!", vbInformation
''''            End
        Case "PAYEE"
            frmPayeeList.Show
        Case "PAYER"
            frmPayer.Show
        Case "SDIMPORT"
            SDImport
        Case "PRINT"
            frmPrint99.Show
        Case "REPORT"
            frmReport.Show
        Case Else
            MsgBox "Selection NF: " & ProgName, vbExclamation
            GoBack

    End Select
    
    ' frmEmpList.Show
    ' frmCompany.Show
    ' frmEmpForm.Show
    ' Form1.Show

End Sub

