Attribute VB_Name = "modStart"

Private Sub Main()   ' *** project execution starts here ******

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
    Set PRCompany = New cPRCompany
    Set GLFFSched = New cGLFFSched
    Set GLFFColumn = New cGLFFColumn
    Set PREquate = New cPREquate
    Set PRGlobal = New cPRGlobal

    SetEquates

'       Case "CLEARGLAMOUNT"
'       Case "GLFILECOPY"
'       Case "GLMULTDIV"
'       Case "DELETEACCTS"
'       Case "COPYBB"
'       Case "YEAREND"
'       Case "IMPORT"
'       Case "UPDATEB"
'       Case "NEWFILE"
'       Case "AcctImport"

    x = Command()
    
    OpenTab = 1
    
    If x = "" Then         ' set for testing
       BalintFolder = "c:\Balint"
       dbPwd = ""
       ProgName = UCase("CLEARGLAMOUNT")
       SysFile = "c:\Balint\Data\GLSystem.mdb"
       UserID = 2
       BackName = ""
       BatchNum = 234
       Period = 201904       ' yyyypp
       ' Period = 0
       DBName = "c:\Balint\Data\KirtlandHills.mdb"
    Else
       dbPwd = GetCmd(x, "dbPwd", "Str")
       ProgName = UCase(GetCmd(x, "ProgName", "Str"))
       SysFile = GetCmd(x, "SysFile", "Str")
       UserID = GetCmd(x, "UserID", "Num")
       BackName = GetCmd(x, "BackName", "Str")
       BatchNum = GetCmd(x, "Batch", "Num")
       MenuName = GetCmd(x, "MenuName", "Str")
       Period = GetCmd(x, "Period", "Num")
       DBName = GetCmd(x, "dbName", "Str")
       BalintFolder = GetCmd(x, "BalintFolder", "Str")
    End If

    If SysFile = "" Then
        SysFile = "\Balint\Data\GLSystem.mdb"
    End If
    
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
    If (IsNull(GLUser.LastCompany) Or GLUser.LastCompany = 0) And ProgName <> "IMPORT" Then
       MsgBox "Company ID not assigned ! ", vbCritical, "GL Utilities"
       End
    End If
    
    ' get the company record from the system data base
    If GLUser.LastCompany <> 0 Then
        If Not GLCompany.GetData(GLUser.LastCompany) Then
               MsgBox "Company ID not found ! " & GLUser.LastCompany, vbCritical, "GL Utilities"
           End
        End If
    End If
       
    ' write the LastPRCompany if available to the user record
'    SQLString = "SELECT * FROM PRCompany WHERE FileName = '" & GLCompany.FileName & "'"
'    If PRCompany.GetBySQL(SQLString) Then
'        GLUser.LastPRCompany = PRCompany.CompanyID
'        GLUser.Save (Equate.RecPut)
'    End If
       
    ' open the company database
    ' If ProgName <> "PRIMPORT" And ProgName <> "IMPORT" And ProgName <> "NEWFILE" And ProgName <> "HISTIMPORT" Then
    If ProgName <> "PRIMPORT" And ProgName <> "IMPORT" And ProgName <> "NEWFILE" Then
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
    End If
    
    ' execute the call
    Select Case ProgName
       
       Case "CLEARGLAMOUNT"
          If Period = 0 Then
             frmGLUrange.Show    ' update using the screen
          Else                   ' update for a period w/out entry screen
             GLBatch.FiscalYear = Int(Period / 100)
             GLBatch.Period = Period Mod 100
             frmUpdBatch.Show
          End If
       Case "GLFILECOPY"
          frmCopy.Show
       Case "GLMULTDIV"
          frmMultDiv.Show
       Case "DELETEACCTS"
          frmDeleteAccts.Show
       Case "COPYBB"
          frmCopyBB.Show
       Case "YEAREND"
          frmYearEnd.Show
       Case "IMPORT"
          SDImport "GL"
       Case "NEWFILE"
          MakeNewFile
       Case "HISTIMPORT"
          SDImport "Hst"
       Case "FFIMPORT"
          SDImport "GLFF"
       Case "PRIMPORT"
          SDImport "PR"
       Case "UPDATEB"
          frmUpdBatch.Show
       Case "SWEEP"
          frmSweep.Show
       Case "ACCTIMPORT"
          frmAcctImport.Show
    
    End Select

End Sub
