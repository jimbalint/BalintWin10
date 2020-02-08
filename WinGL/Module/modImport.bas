Attribute VB_Name = "modImport"
Option Explicit

Dim RCT As Long
Dim SQLQuery As String
Dim BByte As Byte
Dim c As Currency
Dim i As Integer
Dim j As Integer
Dim w As String
Dim x As String
Dim Y As Long
Dim z As Long
Dim dType As String
Dim AsciiChannel As Integer
Dim Ct As Long
Dim GLFName As String
Dim jbName As String
Dim TextName As String
Dim TaskID As Long
Dim DriveLetter As String
Dim Response
Dim BatchMode As Boolean
Dim FileCount As Integer
Dim ClientNum As Integer
Dim ClientName As String
Dim FirstFY As Long
Dim HistCount As Long

Dim LowYear As Integer
Dim HiYear As Integer
Dim LowPeriod As Byte
Dim HiPeriod As Byte

Dim Idb As XArrayDB

Dim LastFY As Integer
Dim LastPd As Byte
Dim LastJS As Byte
Dim HFlag As Byte
Dim TlCredits As Currency
Dim TlDebits As Currency
Dim HRecCount As Long
Dim IType As Byte
Dim SString As String

Dim ColRec, LastColRec, LastFormat As Byte
    
Public Sub SDImport(ByVal ImportType As String)

Dim NewFlag As Boolean
Dim PRCFlag As Boolean
Dim ShellString As String
Dim txtTitle As String
Dim GlCoID As Long
Dim OvwFlag As Boolean
Dim FName, PRC, GLC As String
    
Dim PRCompany As New cPRCompany
    
    ' store drive letter and colon
    DriveLetter = Left(App.Path, 2)
    CNDesOpen (DriveLetter & "\balint\data\GLSystem.mdb")
    
    ' ImportType
    '    = "New"    - new GL client from blank          IType = 1
    '    = "GL"     - new GL client from SD import      IType = 2
    '    = "Hst"    - GL history import only            IType = 3
    '    = "PR"     - new PR client from SD Import      IType = 4
    '    = "GLFF"   - GL Free Format only               IType = 5
           
    NoFieldCheck = True     ' don't check for field updates on connection
    txtTitle = ""
    NewFlag = False
    If ImportType = "New" Then NewFlag = True
    
    SQLString = "SELECT * FROM Users WHERE ID = " & UserID
    If Not GLUser.GetBySQL(SQLString) Then
        MsgBox "User Not Found: " & UserID, vbCritical
        End
    End If
    ' get the GLCompany record
    If Not GLCompany.GetData(GLUser.LastCompany) Then
        MsgBox "GLCompany NF: " & GLUser.LastCompany, vbCritical
        End
    End If
    DBName = GLCompany.FileName
    
    ' make a new DB ???
    If ImportType = "GL" Or ImportType = "PR" Then
        If ImportType = "GL" Then
            x = "Place GL data in existing Database: " & DBName & " ?"
        Else
            x = "Place PR data in existing Database: " & DBName & " ?"
        End If
        Select Case MsgBox(x, vbQuestion + vbYesNoCancel, "Windows Accounting Import")
            Case vbCancel
                GoBack
            Case vbYes
                NewFlag = False
            Case vbNo
                NewFlag = True
        End Select
        If NewFlag = False Then
            If ImportType = "GL" Then
                x = "OK to overwrite all existing GL data in: " & DBName & " ?"
            Else
                x = "OK to overwrite all existing PR data in: " & DBName & " ?"
            End If
            If MsgBox(x, vbYesNo + vbApplicationModal + vbCritical, "File Import") = vbNo Then
                GoBack
            End If
        End If
    End If
    
    If Not TableExists("PRCompany", cnDes) Then CompanyCreate
    
    ' *** create new MDB file ***
    If NewFlag Then        ' enter data base/client name if necessary
        Do
            x = InputBox("Enter NEW Data Base Name (Computer file name):", _
                         "File will be saved in " & DriveLetter & _
                         "\Balint\Data")
            If x = "" Then GoBack
            GLFName = DriveLetter & "\Balint\Data\" & x & ".mdb"
            If GLCopy(GLFName) Then Exit Do    ' successful - move on
        Loop
        
        ' no password yet
        CNOpen GLFName, ""
                
        ' !!! the new mdb file has been created - set it up !!!
                
        ' >>>>> create GLCompany and PRCompany records
        ' is the company record already there (mdb overwrite) ?
        If GLCompany.GetByName(GLFName) Then
            GlCoID = GLCompany.ID
        Else
            GLCompany.Clear
            GLCompany.Name = "New Company " & GLFName
            GLCompany.FileName = GLFName
            If Not GLCompany.Save(Equate.RecAdd) Then
                MsgBox "GLCompany Save Error!", vbCritical
                End
            End If
                            
            ' find it again - wtf ???
            If Not GLCompany.GetByName(GLFName) Then
                MsgBox "GLCompany add error: " & GLFName, vbCritical
                End
            End If
            GlCoID = GLCompany.ID
        End If
        
        SQLString = "SELECT * FROM PRCompany WHERE FileName = " & "'" & GLFName & "'"
        If PRCompany.GetBySQL(SQLString) Then
        Else
            PRCompany.OpenRS
            PRCompany.Clear
            PRCompany.Name = "New Company " & DBName
            PRCompany.FileName = GLFName
            PRCompany.GLCompanyID = GLCompany.ID
            PRCompany.Save (Equate.RecAdd)
        End If
        
        ' >>>>> update user record
        If GLUser.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) = False Then
            MsgBox "GLUser error! " & UserID
            End
        End If
        GLUser.LastCompany = GlCoID
        GLUser.LastPRCompany = PRCompany.CompanyID
        GLUser.Save (Equate.RecPut)
    
    Else    ' open the existing db
        
        ' *************************************************
        ' fix for menu problem
        ' GLUser.LastPRCompany is not properly assigned
        If IsNull(GLUser.LastCompany) = True Or GLUser.LastCompany = 0 Then
            MsgBox "User Last GL Company Not Assigned!", vbExclamation
            End
        End If
        SQLString = "SELECT * FROM PRCompany WHERE GLCompanyID = " & GLUser.LastCompany
        If Not PRCompany.GetBySQL(SQLString) Then   ' make a NEW PRCompany record
            PRCompany.OpenRS
            PRCompany.Clear
            PRCompany.Name = "New Company " & DBName
            PRCompany.FileName = GLFName
            PRCompany.GLCompanyID = GLCompany.ID
            PRCompany.Save (Equate.RecAdd)
        End If
        GLUser.LastPRCompany = PRCompany.CompanyID
        GLUser.Save (Equate.RecPut)
        ' *************************************************
        
        FName = DriveLetter & Mid(DBName, 3, Len(DBName) - 2)
        CNOpen FName, dbPwd
        GLFName = FName
    
    End If
        
    ' create any tables that dont exist
    If Not TableExists("GLAccount", cn) Then AccountCreate
    If Not TableExists("GLAmount", cn) Then AmountCreate
    If Not TableExists("GLBatch", cn) Then BatchCreate
    If Not TableExists("GLBranch", cn) Then BranchCreate
    ' If Not TableExists("GLColumn", cn) Then ColumnCreate
    If Not TableExists("GLFFSched", cn) Then GLFFSchedCreate
    If Not TableExists("GLHistory", cn) Then HistoryCreate
    If Not TableExists("GLJournal", cn) Then JournalCreate
    If Not TableExists("GLPrint", cn) Then PrintCreate
    
    If Not TableExists("PRAdjust", cn) Then AdjustCreate
    If Not TableExists("PRBatch", cn) Then PRBatchCreate
    If Not TableExists("PRDepartment", cn) Then DepartmentCreate
    If Not TableExists("PRDist", cn) Then DistCreate
    If Not TableExists("PREELists", cn) Then EEListsCreate
    If Not TableExists("PREmployee", cn) Then EmployeeCreate
    If Not TableExists("PRGLUpd", cn) Then GLUpdCreate
    If Not TableExists("PRHist", cn) Then HistCreate
    If Not TableExists("PRItem", cn) Then ItemCreate
    If Not TableExists("PRItemHist", cn) Then ItemHistCreate
    
    If ImportType <> "Hst" And ImportType <> "GLFF" Then
        ClientName = InputBox("Enter Client Name (Company name for report titles): ", "SuperDOS Import")
        If ClientName = "" Then GoBack
                
        ' get the company records in case not already found
        If Not GLCompany.GetData(GLUser.LastCompany) Then
            MsgBox "GLCompany Error! " & GLUser.LastCompany, vbCritical, "SuperDOS Import"
            End
        End If
        GLCompany.Name = ClientName
        GlCoID = GLCompany.ID
        If Not GLCompany.Save(Equate.RecPut) Then
            MsgBox "GL Company Save Error", vbCritical
            End
        End If
        
        If Not PRCompany.GetByID(GLUser.LastPRCompany) Then
            MsgBox "PRCompany Error! " & GLUser.LastPRCompany, vbCritical, "SuperDOS Import"
            End
        End If
        PRCompany.Name = ClientName
        PRCompany.Save (Equate.RecPut)
    End If
                
    txtTitle = ""
    Select Case ImportType
        Case "New"
            IType = 1
        Case "GL"
            IType = 2
            txtTitle = "Select GL Client file to import"
        Case "Hst"
            IType = 3
            txtTitle = "Select GL History file to import"
        Case "PR"
            IType = 4
            txtTitle = "Select PR Client file to import"
        Case "GLFF"
            IType = 5
            txtTitle = "SELECT GL Free Format file to import"
        Case Else
            MsgBox "Bad command argument: " & ImportType, vbCritical
            End
   
    End Select
                   
    ' select import file name
    If txtTitle <> "" Then
                   
        If MsgBox(txtTitle, vbOKCancel + vbInformation, "Windows Accounting") = vbCancel Then GoBack
                   
        ' default text file names
        If ImportType = "GL" Then
            TextName = GetTxtName("GLX*.txt")
            If TextName = "" Then GoBack
        ElseIf ImportType = "Hst" Then
            TextName = GetTxtName("GLH*.txt")
            If TextName = "" Then GoBack
        ElseIf ImportType = "PR" Then
            TextName = GetTxtName("PRX*.txt")
            If TextName = "" Then GoBack
        ElseIf ImportType = "GLFF" Then
            TextName = GetTxtName("GLF*.txt")
            If TextName = "" Then GoBack
        End If
                   
    End If
                   
    FirstFY = 0
    LastFY = 0
    LastPd = 0
    LastJS = 0
    BatchNum = 0
   
    BatchMode = False
                
    If ImportType = "GL" Or ImportType = "Hst" Or ImportType = "GLFF" Then
        ' GL Client Import or GL History Import
        Import
    ElseIf ImportType = "PR" Then
    
        ShellString = DriveLetter & "\Balint\PRUtil.exe" & _
            " ProgName=Import " & _
            " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
            " UserID=" & UserID & _
            " BackName=" & DriveLetter & "\Balint\GLMenu.exe" & _
            " txtName=" & TextName & _
            " dbName=" & GLFName
            
        ' database password if required
        If dbPwd <> "" Then
            ShellString = ShellString & " dbPWd=" & dbPwd
        End If
        
        TaskID = Shell(ShellString, vbMaximizedFocus)

        End
    
    End If
    
    GoBack
    
End Sub

Public Sub MakeNewFile()

Dim NewFlag As Boolean
Dim PRCFlag As Boolean
Dim ShellString As String
Dim txtTitle As String
Dim GlCoID As Long
Dim OvwFlag As Boolean
Dim FName, PRC, GLC As String
    
Dim PRCompany As New cPRCompany
    
    ' 2015-05-27
    ' copied from SDImport
    ' trim out extra stuff for SuperDOS import
    ' make compatible w/ Init.txt
    
    ' store drive letter and colon
    DriveLetter = Left(App.Path, 2)
    
    If BalintFolder = "" Then
        CNDesOpen (DriveLetter & "\balint\data\GLSystem.mdb")
    Else
        x = Replace(BalintFolder, "^", " ") & "\Data\GLSystem.mdb"
        CNDesOpen (x)
    End If
    
    ' ImportType
    '    = "New"    - new GL client from blank          IType = 1
    '    = "GL"     - new GL client from SD import      IType = 2
    '    = "Hst"    - GL history import only            IType = 3
    '    = "PR"     - new PR client from SD Import      IType = 4
    '    = "GLFF"   - GL Free Format only               IType = 5
           
    NoFieldCheck = True     ' don't check for field updates on connection
    txtTitle = ""
    NewFlag = True
    
    SQLString = "SELECT * FROM Users WHERE ID = " & UserID
    If Not GLUser.GetBySQL(SQLString) Then
        MsgBox "User Not Found: " & UserID, vbCritical
        End
    End If
    ' get the GLCompany record
    If Not GLCompany.GetData(GLUser.LastCompany) Then
        MsgBox "GLCompany NF: " & GLUser.LastCompany, vbCritical
        End
    End If
    DBName = GLCompany.FileName
    
    If Not TableExists("PRCompany", cnDes) Then CompanyCreate
    
    ' *** create new MDB file ***
    Dim InputFName As String
    Do
        
        x = InputBox("Enter NEW Data Base Name (Computer file name):", _
                     "File will be saved in " & "\Balint\Data")
        If x = "" Then GoBack
        
        If BalintFolder = "" Then
            GLFName = DriveLetter & "\Balint\Data\" & x & ".mdb"
        Else
            GLFName = Replace(BalintFolder, "^", " ") & "\Data\" & x & ".mdb"
        End If
        
        InputFName = x
        
        If GLCopy(GLFName) Then Exit Do    ' successful - move on
    
    Loop
    
    ' no password yet
    CNOpen GLFName, ""
            
    ' !!! the new mdb file has been created - set it up !!!
            
    ' >>>>> create GLCompany and PRCompany records
    ' is the company record already there (mdb overwrite) ?
    If GLCompany.GetByName(GLFName) Then
        GlCoID = GLCompany.ID
    Else
        GLCompany.Clear
        GLCompany.Name = "New Company " & GLFName
        
        Dim FindName As String
        If BalintFolder = "" Then
            GLCompany.FileName = GLFName
            FindName = GLFName
        Else
            GLCompany.FileName = "X:\Balint\Data\" & InputFName & ".mdb"
            FindName = Replace(BalintFolder, "^", " ") & "\Data\" & InputFName & ".mdb"
        End If
        
        If Not GLCompany.Save(Equate.RecAdd) Then
            MsgBox "GLCompany Save Error!", vbCritical
            End
        End If
                        
        ' find it again - wtf ???
'        If Not GLCompany.GetByName(GLFName) Then
'            MsgBox "GLCompany add error: " & GLFName, vbCritical
'            End
'        End If
        GlCoID = GLCompany.ID
    End If
    
    SQLString = "SELECT * FROM PRCompany WHERE FileName = " & "'" & GLFName & "'"
    If PRCompany.GetBySQL(SQLString) Then
    Else
        PRCompany.OpenRS
        PRCompany.Clear
        PRCompany.Name = "New Company " & DBName
            
        If BalintFolder = "" Then
            PRCompany.FileName = GLFName
        Else
            PRCompany.FileName = "X:\Balint\Data\" & InputFName & ".mdb"
        End If
        
        PRCompany.GLCompanyID = GLCompany.ID
        PRCompany.Save (Equate.RecAdd)
    End If
    
    ' >>>>> update user record
    If GLUser.GetBySQL("SELECT * FROM Users WHERE ID = " & UserID) = False Then
        MsgBox "GLUser error! " & UserID
        End
    End If
    GLUser.LastCompany = GlCoID
    GLUser.LastPRCompany = PRCompany.CompanyID
    GLUser.Save (Equate.RecPut)
    
        
    ' create any tables that dont exist
    If Not TableExists("GLAccount", cn) Then AccountCreate
    If Not TableExists("GLAmount", cn) Then AmountCreate
    If Not TableExists("GLBatch", cn) Then BatchCreate
    If Not TableExists("GLBranch", cn) Then BranchCreate
    ' If Not TableExists("GLColumn", cn) Then ColumnCreate
    If Not TableExists("GLFFSched", cn) Then GLFFSchedCreate
    If Not TableExists("GLHistory", cn) Then HistoryCreate
    If Not TableExists("GLJournal", cn) Then JournalCreate
    If Not TableExists("GLPrint", cn) Then PrintCreate
    
    If Not TableExists("PRAdjust", cn) Then AdjustCreate
    If Not TableExists("PRBatch", cn) Then PRBatchCreate
    If Not TableExists("PRDepartment", cn) Then DepartmentCreate
    If Not TableExists("PRDist", cn) Then DistCreate
    If Not TableExists("PREELists", cn) Then EEListsCreate
    If Not TableExists("PREmployee", cn) Then EmployeeCreate
    If Not TableExists("PRGLUpd", cn) Then GLUpdCreate
    If Not TableExists("PRHist", cn) Then HistCreate
    If Not TableExists("PRItem", cn) Then ItemCreate
    If Not TableExists("PRItemHist", cn) Then ItemHistCreate
    
    ClientName = InputBox("Enter Client Name (Company name for report titles): ", "SuperDOS Import")
    If ClientName = "" Then GoBack
            
    ' get the company records in case not already found
    If Not GLCompany.GetData(GLUser.LastCompany) Then
        MsgBox "GLCompany Error! " & GLUser.LastCompany, vbCritical, "SuperDOS Import"
        End
    End If
    GLCompany.Name = ClientName
    GlCoID = GLCompany.ID
    If Not GLCompany.Save(Equate.RecPut) Then
        MsgBox "GL Company Save Error", vbCritical
        End
    End If
    
    If Not PRCompany.GetByID(GLUser.LastPRCompany) Then
        MsgBox "PRCompany Error! " & GLUser.LastPRCompany, vbCritical, "SuperDOS Import"
        End
    End If
    PRCompany.Name = ClientName
    PRCompany.Save (Equate.RecPut)
                
    GoBack

End Sub

Private Sub Import()
   
Dim HistYear1 As Integer
Dim HistPd1 As Byte
Dim HistYear2 As Integer
Dim HistPd2 As Byte

Dim yr As Integer
Dim Pd As Byte

Dim FF1, FF2 As Long

Dim hrs As ADODB.Recordset
   
    ' get record count
    frmProgress.Caption = "GL SuperDOS Import"
    frmProgress.lblMsg2 = "Clearing Files ..."
    frmProgress.lblMsg1 = "Performing import for: " & ClientName
    frmProgress.Show
    frmProgress.MousePointer = vbHourglass
   
    ' clear the tables for new imports
    If IType = 2 Then
        DropTable "GLAccount", cn
        AccountCreate
        
        DropTable "GLAmount", cn
        AmountCreate
        
        DropTable "GLBatch", cn
        BatchCreate
        
        DropTable "GLBranch", cn
        BranchCreate
        
        DropTable "GLHistory", cn
        HistoryCreate
        
        DropTable "GLJournal", cn
        JournalCreate
        
        DropTable "GLPrint", cn
        PrintCreate
        
        DropTable "GLFFSched", cn
        GLFFSchedCreate
        
    End If
   
    GLDescription.OpenRS
    GLAccount.OpenRS
    GLAmount.OpenRS
    GLBatch.OpenRS
    GLBranch.OpenRS
    GLHistory.OpenRS
    GLJournal.OpenRS
'     GLColumn.OpenRS
    GLFFSched.OpenRS
    GLFFColumn.OpenRS
    PRGlobal.OpenRS
   
    FileCount = 1
   
NextFile:
   frmProgress.lblMsg2 = "Counting records to import ..." & TextName
   
   AsciiChannel = FreeFile
   Open TextName For Input As AsciiChannel
   
   RCT = 0
   
   Do
   
      Line Input #AsciiChannel, x
      
      RCT = RCT + 1
      
      If Mid(x, 2, 3) = "END" Then Exit Do
      
      If RCT Mod 100 = 0 Then
         frmProgress.lblMsg2 = "Counting Records: " & CStr(RCT) & " " & TextName
         frmProgress.lblMsg2.Refresh
      End If
   
   Loop
   
   Ct = 0
   
   Close #AsciiChannel
   Open TextName For Input As AsciiChannel
   
   LastFormat = 0
   LastColRec = 0
   ColRec = 0
   
   Do
      
      Input #AsciiChannel, dType
      
      ' first line for history only import
      If Ct = 0 And IType = 3 Then
         
         If Mid(dType, 1, 8) <> "HISTONLY" Then
            MsgBox "Bad import file for history only !!!" & vbCr & dType, vbCritical
            End
         End If
         
         HistYear1 = CInt(Mid(dType, 9, 4))
         HistPd1 = CByte(Mid(dType, 13, 2))
      
         HistYear2 = CInt(Mid(dType, 15, 4))
         HistPd2 = CByte(Mid(dType, 19, 2))
      
         If HistYear1 = 0 Or HistPd1 = 0 Or HistYear2 = 0 Or HistPd2 = 0 Then
            MsgBox "Bad Year/Period to import: " & x, vbCritical
            End
         End If
         
         ' confirm to proceed
         MsgResponse = MsgBox("OK to import GL History only" & vbCr & _
                              "From " & HistYear1 & " Pd " & HistPd1 & " To: " & HistYear2 & " Pd " & HistPd2, _
                              vbQuestion + vbYesNo + vbDefaultButton1, GLCompany.Name)
                
         If MsgResponse = vbNo Then Exit Sub
         
         frmProgress.lblMsg1 = "Deleting History and Batch records ..."
         frmProgress.lblMsg1.Refresh
         
         ' clear the glamount file
         ClearGLAmount HistYear1, HistYear2, HistPd1, HistPd2, False
         
         ' loop thru ranges to clear
         For yr = HistYear1 To HistYear2
             
             For Pd = HistPd1 To HistPd2
                 
                 If Pd = 14 Then
                    yr = yr + 1
                    Pd = 1
                 End If
         
                 ' delete history records for year/period
                 SString = "DELETE * FROM GLHistory WHERE FiscalYear = " & yr & _
                           " AND Period = " & Pd
                 rsInit SString, cn, hrs
         
                 ' delete batch records
                 SString = "DELETE * FROM GLBatch WHERE FiscalYear = " & yr & _
                           " AND Period = " & Pd
                 rsInit SString, cn, hrs
            
            Next Pd
         
         Next yr
         
         GLHistory.OpenRS
      
         ' get first data line of import
         Input #AsciiChannel, dType
      
      End If
      
      Ct = Ct + 1
      If Ct = 1 Or Ct Mod 100 = 0 Then
         frmProgress.lblMsg2 = "On record: " & CStr(Ct) & " of: " & CStr(RCT) & " " & TextName
         frmProgress.Refresh
      End If
      
      Select Case dType
         Case "END"           ' End Of File
              Exit Do
         Case "CMP"           ' Company Info
              Company
         Case "ACT"           ' Account info
              Account
         
'         Case "AMT"           ' Amount Info
'              Amount
         
'         Case "BUD"           ' Budget Amount
'              Budget
         
         Case "COL"           ' Column Info
              FFColumn
         Case "FOR"           ' free format schedule
              FFormat
         Case "BRA"           ' Branch Info
              Branch
         Case "JRN"           ' Journal Info
              Journal
         Case "DES"           ' Description
              Description
         Case "HIS"           ' History
              History
         Case "DELHIS"        ' delete history - used if exporting only history from SuperDOS
              DelHistory      '     will delete all history records for the FY/Period from Access
      
      End Select
   
   Loop
   
   ' check for next file
   FileCount = FileCount + 1
   x = Mid(TextName, 1, 21) & Format(FileCount, "00") & ".txt"
   On Error Resume Next
   GetAttr (x)
   If Err.Number = 0 Then    ' exists - process it
      On Error GoTo 0
      TextName = x
      GoTo NextFile
   End If
   
   ' save the last batch if necessary
   If HistCount <> 0 Then
      GLBatch.GetBatch (BatchNum)
      GLBatch.Credits = GLBatch.Credits + TlCredits
      GLBatch.Debits = GLBatch.Debits + TlDebits
      GLBatch.Records = GLBatch.Records + HistCount
      GLBatch.Created = Now
      GLBatch.CreateUser = GLUser.ID
      GLBatch.Updated = Now
      GLBatch.UpdateUser = GLUser.ID
      GLBatch.FiscalYear = LastFY
      GLBatch.Period = LastPd
      GLBatch.JournalSource = LastJS
      GLBatch.Save (Equate.RecPut)
   End If
   
   GLHistory.CloseRS
   
   ' assign the last batch number to the company file
   GLCompany.LastBatch = BatchNum
   GLCompany.Save (Equate.RecPut)
       
   
   If LowYear <> 0 Then
   
      For i = LowYear To HiYear
   
         ' amount update
         Set Idb = UpdateGLAmount(i, i, LowPeriod, HiPeriod, 0, GLCompany.ID)
     
         ' math update
         Set Idb = MathUpdate(i, i, LowPeriod, HiPeriod)
      
      Next i
   
   End If
   
    ' fill in journal sources if necessary
    For Y = 1 To 10
        If Not GLJournal.GetData(Y) Then
            GLJournal.JournalSource = Y
            GLJournal.JournalName = "Jnl " & Y
            GLJournal.Save (Equate.RecAdd)
        End If
    Next Y
    
    If IType = 2 Then
        ' write the first fiscal year to the company record
        If FirstFY < 1990 Or FirstFY > Year(Now() + 2) Then FirstFY = Year(Now())
        GLCompany.FirstFiscalYear = FirstFY
        GLCompany.Save (Equate.RecPut)
    End If
    
    MsgBox "All Done ...", vbOKOnly + vbInformation, "SuperDOS GL Import"
    frmProgress.lblMsg2 = "On record: " & CStr(Ct) & " of: " & CStr(RCT)
    frmProgress.Refresh
    frmProgress.Hide
    Unload frmProgress
   
End Sub

Private Sub Company()
   
   GLCompany.Clear
   
   ' add GLCompany record if dne
   If Not GLCompany.GetByName(GLFName) Then
      GLCompany.Clear
      GLCompany.FileName = GLFName
      GLCompany.Save Equate.RecAdd
        
      ' get it again ...
      GLCompany.GetByName (GLFName)
      
   End If
   
   GLCompany.FileName = GLFName
   
   For i = 1 To 15
       
       Input #AsciiChannel, x
       
       If x <> "" Then
   
          If i = 1 Then GLCompany.Name = x
          
          If i = 2 Then
             
             If Mid(x, 3, 1) = "/" Then
                GLCompany.LastUpdate = CLng(Mid(x, 7, 4)) * 10000
                GLCompany.LastUpdate = GLCompany.LastUpdate + CLng(Mid(x, 1, 2)) * 100
                GLCompany.LastUpdate = GLCompany.LastUpdate + CLng(Mid(x, 4, 2))
             Else
                GLCompany.LastUpdate = CLng(x)
             End If
             
          End If
          
          If i = 3 Then
             
             If Mid(x, 3, 1) = "/" Then
                GLCompany.LastClose = CLng(Mid(x, 7, 4)) * 10000
                GLCompany.LastClose = GLCompany.LastClose + CLng(Mid(x, 1, 2)) * 100
                GLCompany.LastClose = GLCompany.LastClose + CLng(Mid(x, 4, 2))
             Else
                GLCompany.LastClose = CLng(x)
             End If
             
          End If
          
          If i = 4 Then GLCompany.RetEarnAcct = CLng(x)
          If i = 5 Then GLCompany.SuspAcct = CLng(x)
          If i = 6 Then GLCompany.NetProfitAcct = CLng(x)
          If i = 7 Then GLCompany.FirstPAcct = CLng(x)
          If i = 8 Then GLCompany.PctBaseAcct = CLng(x)
          If i = 9 Then GLCompany.SubDigits = CByte(x)
          If i = 10 Then GLCompany.NumberPds = CByte(x)
          If i = 11 Then GLCompany.FirstPeriod = CByte(x)
       
          If i = 12 Then GLCompany.LowBranch = CLng(x)
          If i = 13 Then GLCompany.HiBranch = CLng(x)
          If i = 14 Then GLCompany.LowConsolidated = CLng(x)
          If i = 15 Then GLCompany.HiConsolidated = CLng(x)
       
       End If
       
   Next i
 
   GLCompany.Name = ClientName
 
   ' save it
   GLCompany.Save (Equate.RecPut)
       
End Sub

Private Sub Account()
   
   GLAccount.Clear
   
   For i = 1 To 10
      
      Input #AsciiChannel, x
      
      If i = 3 And x = "" Then
         GLAccount.DescNumber = 0
         GLAccount.Description = ""
      End If
      
      If x <> "" Then
         
         If i = 1 Then GLAccount.Account = CLng(x)
         If i = 2 Then GLAccount.AcctType = x
         
         ' description
         If i = 3 Then
            
            If Mid(x, 1, 1) = "," Then
               
               w = ""
               
               For Y = 2 To Len(x)
                   If Mid(x, Y, 1) = " " Then Exit For
                   w = w & Mid(x, Y, 1)
               Next Y
               
               GLAccount.DescNumber = CLng(w)
               GLAccount.Description = Mid(x, Y + 1)
            
'        MsgBox x & " " & GLAccount.DescNumber & " " & GLAccount.Description
            
            Else
               
               GLAccount.DescNumber = 0
               GLAccount.Description = x
            
            End If
         
         End If
         
         If i = 4 Then GLAccount.TotalLevel = CByte(x)
         If i = 5 Then GLAccount.PrintTab = CByte(x)
         If i = 6 Then GLAccount.LineFeeds = CByte(x)
         If i = 7 Then GLAccount.BSColumn = CByte(x)
         If i = 8 Then
            BByte = CByte(x)
            If BByte And 2 ^ 7 Then GLAccount.AllStatements = True
            If BByte And 2 ^ 6 Then GLAccount.AllSchedules = True
            If BByte And 2 ^ 5 Then GLAccount.BranchAcct = True
            If BByte And 2 ^ 4 Then GLAccount.ConsAcct = True
            If BByte And 2 ^ 3 Then GLAccount.TotalOnLedger = True
            If BByte And 2 ^ 2 Then GLAccount.DollarSign = True
            If BByte And 2 ^ 1 Then GLAccount.SignRevStmt = True
            If BByte And 2 ^ 0 Then GLAccount.SignRevSched = True
         End If
'             If i = 9 Then glaccount.Date1 = CLng(x)
'             If i = 10 Then glaccount.Date2 = CLng(x)
             
      End If
      
   Next i
   
   GLAccount.Save (Equate.RecAdd)
   
End Sub
Private Sub Amount()
   
   GLAmount.Clear
   
   For i = 1 To 15
      
      Input #AsciiChannel, x
      
      If x <> "" Then

         If i = 1 Then GLAmount.Account = CLng(x)
         If i = 2 Then GLAmount.FiscalYear = CLng(x)
         If i = 3 Then GLAmount.Amount01 = CDec(x)
         If i = 4 Then GLAmount.Amount02 = CDec(x)
         If i = 5 Then GLAmount.Amount03 = CDec(x)
         If i = 6 Then GLAmount.Amount04 = CDec(x)
         If i = 7 Then GLAmount.Amount05 = CDec(x)
         If i = 8 Then GLAmount.Amount06 = CDec(x)
         If i = 9 Then GLAmount.Amount07 = CDec(x)
         If i = 10 Then GLAmount.Amount08 = CDec(x)
         If i = 11 Then GLAmount.Amount09 = CDec(x)
         If i = 12 Then GLAmount.Amount10 = CDec(x)
         If i = 13 Then GLAmount.Amount11 = CDec(x)
         If i = 14 Then GLAmount.Amount12 = CDec(x)
         If i = 15 Then GLAmount.Amount13 = CDec(x)
         
      End If
   
   Next i
   
   GLAmount.Save (Equate.RecAdd)
   
   If FirstFY = 0 Or GLAmount.FiscalYear < FirstFY Then FirstFY = GLAmount.FiscalYear
   
End Sub

Private Sub FFormat()

    GLFFSched.Clear
    
    For i = 1 To 7
    
        Input #AsciiChannel, x

        If x <> "" Then
            
            If i = 1 Then
                GLFFSched.ReportID = CByte(x)
            
                If LastFormat = 0 Or GLFFSched.ReportID <> LastFormat Then
                    PRGlobal.Clear
                    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSched
                    PRGlobal.UserID = GLCompany.ID
                    PRGlobal.Description = GLCompany.Name & " " & GLFFSched.ReportID
                    PRGlobal.Byte10 = GLFFSched.ReportID
                    PRGlobal.Save (Equate.RecAdd)
                End If
                LastFormat = GLFFSched.ReportID
            
            End If
            
            If i = 2 Then GLFFSched.SortOrder = CLng(x)
            
            If i = 3 Then
                GLFFSched.Account = CLng(x)
                If GLFFSched.Account < 0 Then
                    GLFFSched.SignReverse = 1
                    GLFFSched.Account = -GLFFSched.Account
                End If
            End If
            
            If i = 4 Then GLFFSched.PercentBase = CLng(x)
            If i = 5 Then GLFFSched.PrintTab = CByte(x)
            If i = 6 Then GLFFSched.LineFeeds = CByte(x)
            If i = 7 Then GLFFSched.AltDesc = x
            
        End If
        
    Next i
        
    GLFFSched.GlobalID = PRGlobal.GlobalID
    GLFFSched.Save (Equate.RecAdd)

End Sub

Private Sub FFColumn()
    
    GLFFColumn.Clear

    For i = 1 To 7

        Input #AsciiChannel, x

        If x <> "" Then

            ' assigned later when PRGlobal records are added
            ' If i = 1 Then GLColumn.ReportID = CLng(x)
            If i = 1 Then
                ColRec = CByte(x)
                PRGlobal.Byte10 = CByte(x)
            
                If LastColRec = 0 Or ColRec <> LastColRec Then
                    PRGlobal.Clear
                    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFColumn
                    PRGlobal.Save (Equate.RecAdd)
                End If
                LastColRec = ColRec
            
            End If
            
            If i = 2 Then GLFFColumn.ColNum = CByte(x)
            
            ' If i = 3 Then GLColumn.Description = x
            If i = 3 Then
                PRGlobal.Description = GLCompany.ID & " " & _
                                        GLCompany.Name & " " & _
                                        x
            End If
            
            If i = 4 Then GLFFColumn.StartNum = CByte(x)
            If i = 5 Then GLFFColumn.EndNum = CByte(x)
            If i = 6 Then GLFFColumn.PrintTab = CByte(x)

            If i = 7 Then
                BByte = CByte(x)
                
                If BByte And 2 ^ 4 Then GLFFColumn.ColType = Equate.ColAdd
                ' divide type overrides column add
                If BByte And 2 ^ 3 Then GLFFColumn.ColType = Equate.ColDivide
                
                If BByte And 2 ^ 2 Then GLFFColumn.NonPrint = True
                If BByte And 2 ^ 1 Then GLFFColumn.FiscalYear = 1
                If BByte And 2 ^ 0 Then GLFFColumn.Budget = 1
                
            End If

        End If

    Next i

    If GLFFColumn.ColType = 0 Then
        If GLFFColumn.StartNum = 1 And GLFFColumn.EndNum >= 12 Then
            GLFFColumn.ColType = Equate.ColAllPd
        ElseIf GLFFColumn.StartNum = GLFFColumn.EndNum Then
            GLFFColumn.ColType = Equate.ColCurrPd
        Else
            GLFFColumn.ColType = Equate.ColCustom
        End If
    End If

    GLFFColumn.GlobalID = PRGlobal.GlobalID
    
    GLFFColumn.Save (Equate.RecAdd)
    PRGlobal.Save (Equate.RecPut)

End Sub
Private Sub Branch()

End Sub
Private Sub Journal()
   
   GLJournal.Clear
   
   For i = 1 To 2
   
      Input #AsciiChannel, x
      
      If x <> "" Then
         If i = 1 Then GLJournal.JournalSource = CInt(x)
         If i = 2 Then GLJournal.JournalName = x
      End If
   Next i
   
   GLJournal.Save (Equate.RecAdd)
   
End Sub
Private Sub Description()
   
Dim DNumber As Long
Dim Desc As String
   
   GLDescription.Clear
   
   For i = 1 To 2
      Input #AsciiChannel, x
      If x <> "" Then
         
         If i = 1 Then GLDescription.Number = CLng(x)
         If i = 1 Then DNumber = CLng(x)
         
         If i = 2 Then
            Desc = x & ""
         End If
      
      End If
   Next i
   
   ' overwrite the existing description if the record exists
   If GLDescription.Find(DNumber) Then
      GLDescription.Description = Desc
      GLDescription.Save (Equate.RecPut)
   Else
      GLDescription.Description = Desc
      GLDescription.Save (Equate.RecAdd)
   End If

End Sub
Private Sub History()

   GLHistory.Clear
   
   For i = 1 To 10
      
      Input #AsciiChannel, x
   
      If x <> "" Then
         If i = 1 Then GLHistory.Account = CLng(x)
         If i = 2 Then GLHistory.FiscalYear = CLng(x)
         If i = 3 Then GLHistory.Period = CByte(x)
         If i = 4 Then GLHistory.Amount = CDec(x)
         If i = 5 Then GLHistory.Reference = x
         If i = 6 Then GLHistory.Description = x
         If i = 7 Then GLHistory.SourceCode = CByte(x)
         If i = 8 Then GLHistory.JournalSource = CByte(x)
         If i = 9 Then GLHistory.HisType = x
         If i = 10 Then
            BByte = CByte(x)
            If BByte And 2 ^ 0 Then GLHistory.UpdateFlag = True
         End If
      
      End If
   
   Next i
   
   ' budget history record ??? - add 100 to the journal source number   js 2 --> 102
   If GLHistory.HisType = "B" Then GLHistory.JournalSource = GLHistory.JournalSource + 100
   
   ' assign the batch number - see if Yr or Pd or JS has changed
   HFlag = 0
   
   If LastFY = 0 Or GLHistory.FiscalYear <> LastFY Then HFlag = 1
   If LastPd = 0 Or GLHistory.Period <> LastPd Then HFlag = 1
   If LastJS = 0 Or GLHistory.JournalSource <> LastJS Then HFlag = 1
   
   If HFlag Then      ' create new batch record
      
      ' save the prior batch if not first time through
      If LastFY <> 0 Then
         
         GLBatch.GetBatch (BatchNum)
         GLBatch.Credits = GLBatch.Credits + TlCredits
         GLBatch.Debits = GLBatch.Debits + TlDebits
         GLBatch.Records = GLBatch.Records + HistCount
         GLBatch.Created = Now
         GLBatch.CreateUser = GLUser.ID
         GLBatch.Updated = Now
         GLBatch.UpdateUser = GLUser.ID
         GLBatch.FiscalYear = LastFY
         GLBatch.Period = LastPd
         GLBatch.JournalSource = LastJS
         GLBatch.Save (Equate.RecPut)
      
         ' see if batch already exists
         x = "SELECT * FROM GLBatch WHERE FiscalYear = " & GLHistory.FiscalYear & _
             " AND Period = " & GLHistory.Period & _
             " AND JournalSource = " & GLHistory.JournalSource
         If Not GLBatch.GetByString(x) Then
            GLBatch.Clear
            GLBatch.AddBatch GLHistory.FiscalYear, GLHistory.Period
            BatchNum = GLBatch.BatchNumber
         Else
            BatchNum = GLBatch.BatchNumber
         End If
         
      Else        ' first batch
         
         GLBatch.Clear
         GLBatch.AddBatch GLHistory.FiscalYear, GLHistory.Period
         BatchNum = GLBatch.BatchNumber
      
      End If
      
      HistCount = 0
      TlCredits = 0
      TlDebits = 0
   
   End If
   
   LastFY = GLHistory.FiscalYear
   LastPd = GLHistory.Period
   LastJS = GLHistory.JournalSource
   
   If GLHistory.Amount >= 0 Then
      TlDebits = TlDebits + GLHistory.Amount
   Else
      TlCredits = TlCredits + GLHistory.Amount
   End If
   
   HistCount = HistCount + 1
   
   GLHistory.BatchNumber = BatchNum
   
   GLHistory.Save (Equate.RecAdd)
   
   ' assign the post date
   HRecCount = HRecCount + 1
   GLHistory.PostDate = DateSerial(Year(Now()), Month(Now()), Day(Now())) - 1 + _
                        TimeSerial(0, 0, HRecCount)
                        
   GLHistory.Save (Equate.RecPut)
   
   ' save for company record
   If FirstFY = 0 Or GLHistory.FiscalYear < FirstFY Then FirstFY = GLHistory.FiscalYear
   
   ' save lo/hi fy and period for amount/math update
   If LowYear = 0 Or GLHistory.FiscalYear < LowYear Then LowYear = GLHistory.FiscalYear
   If GLHistory.FiscalYear > HiYear Then HiYear = GLHistory.FiscalYear
   If LowPeriod = 0 Or GLHistory.Period < LowPeriod Then LowPeriod = GLHistory.Period
   If GLHistory.Period > HiPeriod Then HiPeriod = GLHistory.Period
   
End Sub

Private Sub DelHistory()

Dim DelYear As String
Dim DelPeriod As String

      Input #AsciiChannel, x
      DelYear = x
      
      Input #AsciiChannel, x
      DelPeriod = x
      
      x = "DELETE * FROM GLHistory WHERE FiscalYear = " & DelYear & " AND Period = " & DelPeriod
      
      GLHistory.GetByString x
      
End Sub

Private Sub Budget()
   
Dim Acct As Long
Dim FY As Integer
   
   On Error GoTo Berr
   
   GLAmount.Clear
   
   Input #AsciiChannel, x
   Acct = CLng(x)
   
   Input #AsciiChannel, x
   FY = CInt(x)
   
   If Not GLAmount.Find(Acct, FY) Then
      GLAmount.Clear
      GLAmount.Account = Acct
      GLAmount.FiscalYear = FY
      GLAmount.Save (Equate.RecAdd)
   End If
   
   For i = 1 To 13
      
      Input #AsciiChannel, x
      
      If x <> "" Then

         If i = 1 Then GLAmount.Budget01 = CDec(x)
         If i = 2 Then GLAmount.Budget02 = CDec(x)
         If i = 3 Then GLAmount.Budget03 = CDec(x)
         If i = 4 Then GLAmount.Budget04 = CDec(x)
         If i = 5 Then GLAmount.Budget05 = CDec(x)
         If i = 6 Then GLAmount.Budget06 = CDec(x)
         If i = 7 Then GLAmount.Budget07 = CDec(x)
         If i = 8 Then GLAmount.Budget08 = CDec(x)
         If i = 9 Then GLAmount.Budget09 = CDec(x)
         If i = 10 Then GLAmount.Budget10 = CDec(x)
         If i = 11 Then GLAmount.Budget11 = CDec(x)
         If i = 12 Then GLAmount.Budget12 = CDec(x)
         If i = 13 Then GLAmount.Budget13 = CDec(x)
         
      End If
   
   Next i
   
   GLAmount.Save (Equate.RecPut)
   
   Exit Sub
   
Berr:
   MsgBox Ct & " " & FY
   
End Sub


Function GetTxtName(ByVal WildCard As String) As String
      
Dim OPath As String

   ' store original path
   OPath = App.Path
      
   frmaMain.cmnOpen.Filter = "Export Files|" & WildCard
   frmaMain.cmnOpen.DefaultExt = ".txt"
   frmaMain.cmnOpen.DialogTitle = "Select File to Import"
   jbName = Left(App.Path, 2) & "\Balint\Data"
   frmaMain.cmnOpen.InitDir = jbName
   frmaMain.cmnOpen.ShowOpen
   GetTxtName = frmaMain.cmnOpen.FileName
   If GetTxtName = "" Then Exit Function

   ' restore original drive and path
   ChDrive (Left(OPath, 2))
   ChDir (OPath)

End Function

Function GLCopy(ByVal FName As String) As Boolean
    
Dim DestName As String
Dim BlankName As String
Dim Opt As Long

    On Error Resume Next
    GetAttr (FName)
    If Err.Number = 0 Then
       Opt = vbYesNo + vbCritical + vbDefaultButton2
       Response = MsgBox(FName & " already exists! " & vbCrLf & _
                  "Would you like to overwrite it ?", Opt, _
                  "SuperDOS GL Import")
       If Response = vbNo Then
          On Error GoTo 0
          GLCopy = False
          Exit Function
       End If
    End If
 
    If BalintFolder = "" Then
        BlankName = DriveLetter & "\Balint\Blank\Blank.mdb"
    Else
        BlankName = BalintFolder & "\Blank\Blank.mdb"
    End If
    
    On Error Resume Next
    FileCopy BlankName, FName
    If Err.Number <> 0 Then
        MsgBox Err.Description & " File copy FAILED !" & vbCr & Trim(BlankName) & vbCr & Trim(FName), vbCritical
        On Error GoTo 0
        GLCopy = False
        Exit Function
    Else
        On Error GoTo 0
        GLCopy = True
    End If
   
End Function

Public Sub DropTable(ByVal TableName As String, _
                      ByVal adoCn As ADODB.Connection)

' *** Drop a table if it exists ***

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim TableFlag As Boolean
Dim FString As String
                         
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoCn.OpenSchema(adSchemaColumns)
       
    TableFlag = False
       
    Do Until frs.EOF = True
              
        If frs!Table_Name = TableName Then
            TableFlag = True
            Exit Do
        End If
      
        frs.MoveNext
   
    Loop

    frs.Close
    
    ' table does not exist
    If TableFlag = False Then Exit Sub

    FString = "DROP TABLE " & TableName
    adoCn.Execute FString

End Sub

Public Function TableExistsI(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String

    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)

    TableExistsI = False

    Do Until frs.EOF = True

        If frs!Table_Name = TableName Then
            TableExistsI = True
            Exit Do
        End If

       frs.MoveNext

   Loop

End Function

Private Sub AccountCreate()
    
    SQLString = "CREATE TABLE GLAccount ( " & _
                        "[ID] Counter, CONSTRAINT glacctIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "GLAccount", "Account", "Long", cn
    AddField "GLAccount", "AcctType", "Char (1)", cn
    AddField "GLAccount", "Description", "Char (200)", cn
    AddField "GLAccount", "DescNumber", "Long", cn
    AddField "GLAccount", "TotalLevel", "Byte", cn
    AddField "GLAccount", "PrintTab", "Byte", cn
    AddField "GLAccount", "LineFeeds", "Byte", cn
    AddField "GLAccount", "BSColumn", "Byte", cn
    AddField "GLAccount", "AllStatements", "Logical", cn
    AddField "GLAccount", "AllSchedules", "Logical", cn
    AddField "GLAccount", "BranchAcct", "Logical", cn
    AddField "GLAccount", "ConsAcct", "Logical", cn
    AddField "GLAccount", "TotalOnLedger", "Logical", cn
    AddField "GLAccount", "DollarSign", "Logical", cn
    AddField "GLAccount", "SignRevStmt", "Logical", cn
    AddField "GLAccount", "SignRevSched", "Logical", cn
    AddField "GLAccount", "Date1", "Long", cn
    AddField "GLAccount", "Date2", "Long", cn

End Sub

Private Sub AmountCreate()

    SQLString = "CREATE TABLE GLAmount ( " & _
                        "[ID] Counter, CONSTRAINT glamtIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString
    
    AddField "GLAmount", "Account", "Long", cn
    AddField "GLAmount", "FiscalYear", "Integer", cn
    AddField "GLAmount", "Amount01", "Currency", cn
    AddField "GLAmount", "Amount02", "Currency", cn
    AddField "GLAmount", "Amount03", "Currency", cn
    AddField "GLAmount", "Amount04", "Currency", cn
    AddField "GLAmount", "Amount05", "Currency", cn
    AddField "GLAmount", "Amount06", "Currency", cn
    AddField "GLAmount", "Amount07", "Currency", cn
    AddField "GLAmount", "Amount08", "Currency", cn
    AddField "GLAmount", "Amount09", "Currency", cn
    AddField "GLAmount", "Amount10", "Currency", cn
    AddField "GLAmount", "Amount11", "Currency", cn
    AddField "GLAmount", "Amount12", "Currency", cn
    AddField "GLAmount", "Amount13", "Currency", cn
    AddField "GLAmount", "Budget01", "Currency", cn
    AddField "GLAmount", "Budget02", "Currency", cn
    AddField "GLAmount", "Budget03", "Currency", cn
    AddField "GLAmount", "Budget04", "Currency", cn
    AddField "GLAmount", "Budget05", "Currency", cn
    AddField "GLAmount", "Budget06", "Currency", cn
    AddField "GLAmount", "Budget07", "Currency", cn
    AddField "GLAmount", "Budget08", "Currency", cn
    AddField "GLAmount", "Budget09", "Currency", cn
    AddField "GLAmount", "Budget10", "Currency", cn
    AddField "GLAmount", "Budget11", "Currency", cn
    AddField "GLAmount", "Budget12", "Currency", cn
    AddField "GLAmount", "Budget13", "Currency", cn

End Sub

Private Sub BatchCreate()

    SQLString = "CREATE TABLE GLBatch ( " & _
                        "[ID] Counter, CONSTRAINT glbatIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString
    
    AddField "GLBatch", "FiscalYear", "Integer", cn
    AddField "GLBatch", "Period", "Byte", cn
    AddField "GLBatch", "BatchNumber", "Long", cn
    AddField "GLBatch", "Debits", "Currency", cn
    AddField "GLBatch", "Created", "DateTime", cn
    AddField "GLBatch", "Updated", "DateTime", cn
    AddField "GLBatch", "CreateUser", "Long", cn
    AddField "GLBatch", "UpdateUser", "Long", cn
    AddField "GLBatch", "Records", "Long", cn
    AddField "GLBatch", "Credits", "Currency", cn
    AddField "GLBatch", "JournalSource", "Integer", cn

End Sub

Private Sub BranchCreate()

    SQLString = "CREATE TABLE GLBranch ( " & _
                        "[ID] Counter, CONSTRAINT glbraIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString
    
    AddField "GLBranch", "BranchNumber", "Integer", cn
    AddField "GLBranch", "Name", "Char (60)", cn

End Sub

Private Sub ColumnCreate()

    SQLString = "CREATE TABLE GLColumn ( " & _
                        "[ID] Counter, CONSTRAINT glcolIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString
    
    AddField "GLColumn", "ReportID", "Long", cn
    AddField "GLColumn", "ColumnNum", "Byte", cn
    AddField "GLColumn", "Description", "Char (30)", cn
    AddField "GLColumn", "Value1", "Byte", cn
    AddField "GLColumn", "Value2", "Byte", cn
    AddField "GLColumn", "PrintTab", "Byte", cn
    AddField "GLColumn", "bColumn", "Byte", cn
    AddField "GLColumn", "bMonth", "Byte", cn
    AddField "GLColumn", "bBudget", "Byte", cn
    AddField "GLColumn", "bPercent", "Byte", cn
    AddField "GLColumn", "bNonPrint", "Byte", cn
    AddField "GLColumn", "StartFY", "Long", cn
    AddField "GLColumn", "EndFY", "Long", cn
    AddField "GLColumn", "Operation", "Byte", cn

End Sub

Private Sub HistoryCreate()

    SQLString = "CREATE TABLE GLHistory ( " & _
                        "[ID] Counter, CONSTRAINT glhisIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString

    AddField "GLHistory", "Account", "Long", cn
    AddField "GLHistory", "FiscalYear", "Integer", cn
    AddField "GLHistory", "Period", "Byte", cn
    AddField "GLHistory", "BatchNumber", "Long", cn
    AddField "GLHistory", "Amount", "Currency", cn
    AddField "GLHistory", "Reference", "Char (20)", cn
    AddField "GLHistory", "Description", "Char (20)", cn
    AddField "GLHistory", "SourceCode", "Byte", cn
    AddField "GLHistory", "JournalSource", "Byte", cn
    AddField "GLHistory", "HisType", "Char (1)", cn
    AddField "GLHistory", "UpdateFlag", "Logical", cn
    AddField "GLHistory", "PostDate", "DateTime", cn

End Sub

Private Sub JournalCreate()

    SQLString = "CREATE TABLE GLJournal ( " & _
                        "[ID] Counter, CONSTRAINT gljnlIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString

    AddField "GLJournal", "JournalSource", "Integer", cn
    AddField "GLJournal", "JournalName", "Char (60)", cn

End Sub

Private Sub PrintCreate()

    SQLString = "CREATE TABLE GLPrint ( " & _
                        "[ID] Counter, CONSTRAINT glprtIDKey PRIMARY KEY ([ID]) ) "
                        
    cn.Execute SQLString

    ' *************************************************
    AddField "GLPrint", "gUser", "Char (8)", cn
    ' *************************************************
    
    AddField "GLPrint", "ReportName", "Char (30)", cn
    AddField "GLPrint", "FiscalYear", "Long", cn
    AddField "GLPrint", "BeginDate", "Long", cn
    AddField "GLPrint", "EndDate", "Long", cn
    AddField "GLPrint", "ReportDate", "Long", cn
    AddField "GLPrint", "LowAccount", "Long", cn
    AddField "GLPrint", "HiAccount", "Long", cn
    AddField "GLPrint", "SepPage", "Logical", cn
    AddField "GLPrint", "SupprCP", "Logical", cn
    AddField "GLPrint", "UseMathRec", "Logical", cn
    AddField "GLPrint", "PrtAcctNum", "Logical", cn
    AddField "GLPrint", "PrtZeroBal", "Logical", cn
    AddField "GLPrint", "RoundDollars", "Logical", cn
    AddField "GLPrint", "WidePrint", "Logical", cn
    AddField "GLPrint", "LowerCaseDate", "Logical", cn
    AddField "GLPrint", "LowBranchAcct", "Long", cn
    AddField "GLPrint", "HiBranchAcct", "Long", cn
    AddField "GLPrint", "LowConsAcct", "Long", cn
    AddField "GLPrint", "HiConsAcct", "Long", cn
    
    ' *************************************************
    AddField "GLPrint", "gOutput", "Char (60)", cn
    ' *************************************************
    
    AddField "GLPrint", "Copies", "Integer", cn
    AddField "GLPrint", "RegBraCon", "Byte", cn
    AddField "GLPrint", "StaSch", "Byte", cn
    AddField "GLPrint", "RegCmp", "Byte", cn
    AddField "GLPrint", "PrintBIB", "Byte", cn

End Sub
