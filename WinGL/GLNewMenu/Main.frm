VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   " BALINT GENERAL LEDGER SYSTEM"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9630
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStmt 
      Caption         =   "&Statements"
      Height          =   855
      Left            =   6960
      Picture         =   "Main.frx":0CFA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdTB 
      Caption         =   "&Trial Bal."
      Enabled         =   0   'False
      Height          =   855
      Left            =   5340
      Picture         =   "Main.frx":1004
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdMaint 
      Caption         =   "&Maint."
      Height          =   855
      Left            =   3720
      Picture         =   "Main.frx":1446
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      Caption         =   "&Open"
      Height          =   855
      Left            =   2100
      Picture         =   "Main.frx":1888
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "&New"
      Height          =   855
      Left            =   480
      Picture         =   "Main.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog msDialog 
      Left            =   6720
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCName 
      Caption         =   "Current Client:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblFname 
      Caption         =   "File Location:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblUserName 
      Caption         =   "glUserName"
      Height          =   255
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Current User:"
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOn 
         Caption         =   "Log User On"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu mru1 
         Caption         =   ""
      End
      Begin VB.Menu mru2 
         Caption         =   ""
      End
      Begin VB.Menu mru3 
         Caption         =   ""
      End
      Begin VB.Menu mru4 
         Caption         =   ""
      End
      Begin VB.Menu m4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuAccount 
         Caption         =   "Account"
      End
      Begin VB.Menu mnuBranch 
         Caption         =   "Branch"
      End
      Begin VB.Menu mnuCompany 
         Caption         =   "Company Info"
      End
      Begin VB.Menu mnuJournal 
         Caption         =   "Journal Sources"
      End
      Begin VB.Menu m3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDescriptions 
         Caption         =   "Descriptions"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Users"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "&Data Entry"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu FinStmt 
         Caption         =   "&Financial Statements"
      End
      Begin VB.Menu mnuTrialBalance 
         Caption         =   "&Trial Balance"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuSDImport 
         Caption         =   "&SuperDOS Import"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLFName As String
Dim x As String
Dim CompanyID As Long
Dim Flg As Boolean
Dim BatchNum As Long

'Private Sub cmdMaint_Click()
'   mnuAccount_Click
'End Sub
'
'Private Sub cmdOpen_Click()
'   mnuOpen_Click
'End Sub
'
'Private Sub cmdStmt_Click()
'    Dim TaskNumber As Variant
'
''    RetValue = ExecCmd("\balint\GLSTMT.exe " & glFileName(0))
'
'    TaskNumber = Shell("\balint\GLSTMT.exe " & glFileName(0), vbMaximizedFocus)
'
'End Sub
'
'Private Sub cmdTB_Click()
'   mnuTrialBalance_Click
'End Sub
'
'Private Sub Command1_Click()
'   mnuNew_Click
'End Sub
'
'
'Private Sub FinStmt_Click()
'    Dim TaskNumber As Variant
'
''    RetValue = ExecCmd("\balint\GLSTMT.exe " & glFileName(0))
'
'    TaskNumber = Shell("\balint\GLSTMT.exe " & glFileName(0), vbMaximizedFocus)
'End Sub
'
'Private Sub Form_Load()
'
'    mnuTrialBalance.Enabled = False
'
'    Set GLCompany = New cGLCompany
'    Set GLAccount = New cGLAccount
'    Set GLAmount = New cGLAmount
'    Set GLBranch = New cGLBranch
''    Set GLColumn = New cGLColumn
'    Set GLCompany = New cGLCompany
'    Set GLDescription = New cGLDescription
'    Set GLHistory = New cGLHistory
'    Set GLPrint = New cGLPrint
'    Set GLUser = New cGLUser
'    Set Equate = New cEquate
'
'    SetEquates
'
'    x = Command
'
'    If x = "" Then       ' nothing on the command line
'       x = "53//User/C:\Balint\Data\GLSystem.mdb/jim"
'       x = "58/golf/Account/C:\Balint\Data\GLSystem.mdb/jim"
'    End If
'
'    If cmdline(x, ID, Password, Prog, SysFile, User, BatchNum) = False Then
'       MsgBox "Bad command line !!!"
'    End If
'
'    CNDesOpen (SysFile)
'    CompanyID = ID
'
'    If Not GLCompany.GetData(CompanyID) Then
'       MsgBox "Company record not found ID# " & CompanyID
'       End
'    End If
'
'    ' get the user record
'    If Not GLUser.GetSQL("SELECT * FROM Users WHERE Logon = '" & User & "'") Then
'       MsgBox "User not found! " & vbCrLf & User, vbCritical, "GL Main"
'       Unload Me
'       End
'    End If
'
'    UserID = GLUser.ID
'
'    DBName = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
'
'    CNOpen DBName, Password
'
'    Prog = StrConv(Prog, vbUpperCase)
'
'    GLDescription.OpenRS
'    GLPrint.GetData User, Flg
'
'    Select Case Prog
'
'       Case "ACCOUNT"
'          frmAccount.DBName = DBName
'          frmAccount.Show vbModal
'
'       Case "JOURNAL"
'          frmJournal.Show vbModal
'
'       Case "BRANCH"
'          frmBranches.Show vbModal
'
'       Case "USER"
'          frmUsers.Show vbModal
'
'    End Select
'
'    End
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    WriteParams
'End Sub
'
'Private Sub mnuAccount_Click()
'    If CNOpen(glFileName(0), Password) Then frmAccount.Show vbModal, Me
'End Sub
'
'Private Sub mnuBranch_Click()
'    If CNOpen(glFileName(0), Password) Then frmBranches.Show vbModal, Me
'End Sub
'
'Private Sub mnuCompany_Click()
'    If CNOpen(glFileName(0), Password) Then frmCompany.Show vbModal, Me
'End Sub
'
'Private Sub mnuConfig_Click()
'    frmConfig.Show vbModal, Me
'End Sub
'
'Private Sub mnuDescriptions_Click()
'    frmDescriptions.Show vbModal, Me
'End Sub
'
'Private Sub mnuExit_Click()
'    Unload Me
'End Sub
'
'Private Sub mnuJournal_Click()
'    If CNOpen(glFileName(0), Password) Then frmJournal.Show vbModal, Me
'End Sub
'
'Private Sub mnuLogOn_Click()
'    frmLogin.Show vbModal, Me
'    lblUserName = glUserName
'End Sub
'
'Private Sub mnuNew_Click()
'    Dim temp As String
'    Dim FF As Integer
'
'    Dim fso, mdbFile
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    FF = MsgBox("First select the NEW DataBase name to save under", vbInformation + vbOKCancel, "New GL DataBase")
'    If FF = vbCancel Then Exit Sub
'
'    temp = glFileName(0)
'
''    On Error Resume Next
''    msDialog.Filter = "Client Files|*.mdb"
''    msDialog.DefaultExt = ".mdb"
''    msDialog.DialogTitle = "Enter GL file name to CREATE"
''    msDialog.InitDir = "\Balint\Data"
''    msDialog.ShowOpen
''    If Not Err.Number = 0 Then Exit Sub
''    glFileName(0) = msDialog.FileName
'
'    glFileName(0) = InputBox("Enter a name for the NEW DataBase File", _
'                            "Will be saved in \Balint\Data")
'    If glFileName(0) = "" Then Exit Sub
''    If Mid(glFileName(0), Len(glFileName(0)) - 4, 4) <> ".mdb" Then
'
'
'     glFileName(0) = Left(App.Path, 2) & "\Balint\Data\" & glFileName(0) & ".mdb"
'
''    End If
'
'    ' warn if exists
'    On Error Resume Next
'    FF = FreeFile
'    Open glFileName(0) For Input As #FF
'
'    If Err.Number = 0 Then
'       MsgBox "Permission denied: " & glFileName(0) & " already exists!", vbExclamation
'       Close #FF
'       Exit Sub
'    End If
'
'    Close #FF
'
'    On Error GoTo 0
'
'    glFileName(4) = glFileName(3)
'    glFileName(3) = glFileName(2)
'    glFileName(2) = glFileName(1)
'    glFileName(1) = glFileName(0)
'    LabelMRU
'
'    ' copy the blank file to the selected file
'    GLFName = Left(App.Path, 2) & "\balint\data\glblank.mdb"
'    Set mdbFile = fso.getfile(GLFName)
'    mdbFile.Copy (glFileName(0))
'
'    Response = CNOpen(glFileName(0), Password)
'    If Response = False Then
'       MsgBox "File Create Error !", vbCritical
'       End
'    End If
'
'    ' create the company record
'    GLCompany.Clear
'    GLCompany.Name = "New Client"
'    GLCompany.FileName = glFileName(0)
'    GLCompany.Save 0, Equate.RecAdd
'
'    glCompanyID(0) = GLCompany.ID
'
'    If Response = False Then
'       MsgBox "Company create error!", vbCritical
'       End
'    End If
'
'    ' update the company record
'    ' frmCompany.Show vbModal
'
'    ' display to screen
'    OpenCompany (glFileName(0))
'
'    FF = MsgBox("Next you will specify where the IMPORT text file is", vbInformation + vbOKCancel)
'    If FF = vbCancel Then Exit Sub
'
'    Dim TaskNumber As Variant
'
''    RetValue = ExecCmd("\balint\GLImport.exe " & glFileName(0))
'
'    TaskNumber = Shell("\balint\GLImport.exe " & glFileName(0), vbMaximizedFocus)
'
'End Sub
'
'Private Sub mnuOpen_Click()
'    Dim temp As String
'    Dim temp2 As Long
'
'    temp = glFileName(0)
'    temp2 = glCompanyID(0)
'
'    On Error Resume Next
'    msDialog.InitDir = "\Balint\Data"
'    msDialog.Filter = "Client Files|*.mdb"
'    msDialog.ShowOpen
'    If Not Err.Number = 0 Then Exit Sub
'
'    glFileName(0) = msDialog.FileName
'
'    If CNOpen(glFileName(0), Password) Then
'
'        glFileName(4) = glFileName(3)
'        glCompanyID(4) = glCompanyID(3)
'
'        glFileName(3) = glFileName(2)
'        glCompanyID(3) = glCompanyID(2)
'
'        glFileName(2) = glFileName(1)
'        glCompanyID(2) = glCompanyID(1)
'
'        glFileName(1) = glFileName(0)
'        glCompanyID(1) = glCompanyID(0)
'
'        LabelMRU
'    Else
'        glFileName(0) = temp
'        glCompanyID(0) = temp2
'    End If
'
'    OpenCompany (glFileName(0))
'    glCompanyID(0) = GLCompany.ID
'
'End Sub
'
'Private Sub mnuSDImport_Click()
'   Dim i As Integer
'   i = MsgBox("ALL records will be overwritten", vbExclamation + vbOKCancel, "SuperDOS Import")
'   If i = vbCancel Then Exit Sub
'   Dim TaskNumber As Variant
'
''   RetValue = ExecCmd("\balint\GLImport.exe " & glFileName(0))
'
'   TaskNumber = Shell("\balint\GLImport.exe " & glFileName(0), vbMaximizedFocus)
'
'End Sub
'
'Private Sub mnuStatements_Click()
'
'End Sub
'
'Private Sub mnuTrialBalance_Click()
'    Dim TaskNumber As Variant
'
''    RetValue = ExecCmd("\balint\GLTrialBal.exe " & glFileName(0))
'
'    TaskNumber = Shell("\balint\GLTrialBal.exe " & glFileName(0), vbMaximizedFocus)
'
'End Sub
'
'Private Sub mnuUsers_Click()
'    frmUsers.Show vbModal, Me
'End Sub
'
'Private Sub ReadParams()
'
'    glFileName(1) = "..."
'    glFileName(2) = "..."
'    glFileName(3) = "..."
'    glFileName(4) = "..."
'
'    glCompanyID(1) = 0
'    glCompanyID(2) = 0
'    glCompanyID(3) = 0
'    glCompanyID(4) = 0
'
'    glLoadLast = True
'
'    Dim fid As Integer
'    fid = FreeFile()
'    On Error GoTo glErr
'
'    GLFName = Left(App.Path, 2) & "\Balint\Data\defaults.gl"
'    Open GLFName For Input As #fid
'    Dim strLine As String
'
'    Input #fid, strLine
'
'    If Not strLine = "..." And strLine <> "0" Then
'        glCompanyID(1) = strLine
'        GLCompany.GetData (glCompanyID(1))
'        glFileName(1) = GLCompany.FileName
'    End If
'
'    Input #fid, strLine
'    If Not strLine = "..." And strLine <> "0" Then
'        glCompanyID(2) = strLine
'        GLCompany.GetData (glCompanyID(2))
'        glFileName(2) = GLCompany.FileName
'    End If
'
'    Input #fid, strLine
'    If Not strLine = "..." And strLine <> "0" Then
'        glCompanyID(3) = strLine
'        GLCompany.GetData (glCompanyID(3))
'        glFileName(3) = GLCompany.FileName
'    End If
'
'    Input #fid, strLine
'    If Not strLine = "..." And strLine <> "0" Then
'        glCompanyID(4) = strLine
'        GLCompany.GetData (glCompanyID(4))
'        glFileName(4) = GLCompany.FileName
'    End If
'
'    Input #fid, strLine
'    glLoadLast = CBool(strLine)
'    If glLoadLast = True Then
'        If Not glFileName(1) = "..." Then
'
'            glFileName(0) = glFileName(1)
'            glCompanyID(0) = glCompanyID(1)
'
'            GLCompany.GetData (glCompanyID(0))
'
'            GLCompany.ID = glCompanyID(0)
'            lblCName.Caption = GLCompany.Name
'            lblFname.Caption = GLCompany.FileName
'            Me.Refresh
'
'        End If
'    End If
'
'glErr:
'    LabelMRU
'    Close #fid
'End Sub
'
'Private Sub WriteParams()
'    Dim fid As Integer
'    On Error GoTo glErr
'    fid = FreeFile()
'
'    GLFName = Left(App.Path, 2) & "\Balint\Data\defaults.gl"
'    Open GLFName For Output As #fid
'    If Err.Number Then Exit Sub
'
'    Print #fid, glCompanyID(0)
'    Print #fid, glCompanyID(1)
'    Print #fid, glCompanyID(2)
'    Print #fid, glCompanyID(3)
'    Print #fid, CStr(glLoadLast)
'glErr:
'    Close #fid
'End Sub
'
'Private Sub LabelMRU()
'    mru1.Caption = "&1-" & glFileName(1)
'    mru2.Caption = "&2-" & glFileName(2)
'    mru3.Caption = "&3-" & glFileName(3)
'    mru4.Caption = "&4-" & glFileName(4)
'End Sub
'
'Private Sub mru1_Click()
'    Dim temp As String
'    temp = glFileName(0)
'    glFileName(0) = glFileName(1)
'    If CNOpen(glFileName(0), Password) Then
'        glFileName(1) = glFileName(0)
'        LabelMRU
'    Else
'        glFileName(0) = temp
'    End If
'    OpenCompany (glFileName(0))
'    glCompanyID(0) = GLCompany.ID
'
'End Sub
'
'Private Sub mru2_Click()
'    Dim temp As String
'    temp = glFileName(0)
'    glFileName(0) = glFileName(2)
'    If CNOpen(glFileName(0), Password) Then
'        glFileName(2) = glFileName(1)
'        glFileName(1) = glFileName(0)
'        LabelMRU
'    Else
'        glFileName(0) = temp
'    End If
'    OpenCompany (glFileName(0))
'    glCompanyID(0) = GLCompany.ID
'
'End Sub
'
'Private Sub mru3_Click()
'    Dim temp As String
'    temp = glFileName(0)
'    glFileName(0) = glFileName(3)
'    If CNOpen(glFileName(0), Password) Then
'        glFileName(3) = glFileName(2)
'        glFileName(2) = glFileName(1)
'        glFileName(1) = glFileName(0)
'        LabelMRU
'    Else
'        glFileName(0) = temp
'    End If
'    OpenCompany (glFileName(0))
'    glCompanyID(0) = GLCompany.ID
'End Sub
'
'Private Sub mru4_Click()
'    Dim temp As String
'    temp = glFileName(0)
'    glFileName(0) = glFileName(4)
'    If CNOpen(glFileName(0), Password) Then
'        glFileName(4) = glFileName(3)
'        glFileName(3) = glFileName(2)
'        glFileName(2) = glFileName(1)
'        glFileName(1) = glFileName(0)
'        LabelMRU
'    Else
'        glFileName(0) = temp
'    End If
'    OpenCompany (glFileName(0))
'    glCompanyID(0) = GLCompany.ID
'
'End Sub
'Private Sub OpenCompany(ByVal FileName As String)
'    GLCompany.GetByName (FileName)
'    lblFname.Caption = "File Location: " & GLCompany.FileName
'    lblCName.Caption = "Current Client: " & GLCompany.Name
'    Me.Refresh
'End Sub
