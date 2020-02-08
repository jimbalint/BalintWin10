VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainMenu 
   Caption         =   " GENERAL LEDGER MENU"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11145
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   0  'User
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   " Payroll Demo "
      Height          =   3135
      Left            =   7920
      TabIndex        =   12
      Top             =   1680
      Width           =   2895
      Begin VB.CommandButton cmdPRMaintEE 
         Caption         =   "Employee Maint"
         Height          =   495
         Left            =   480
         TabIndex        =   16
         Top             =   1120
         Width           =   2055
      End
      Begin VB.CommandButton cmdPRReport 
         Caption         =   "Reports"
         Enabled         =   0   'False
         Height          =   495
         Left            =   480
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdPREntry 
         Caption         =   "Data Entry"
         Height          =   495
         Left            =   480
         TabIndex        =   14
         Top             =   1760
         Width           =   2055
      End
      Begin VB.CommandButton cmdPRMaintER 
         Caption         =   "Employer Maint"
         Height          =   495
         Left            =   480
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":0EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":11E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":14FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1816
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":1B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":25C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open File"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AcctMaint"
            Description     =   "AcctMain"
            Object.ToolTipText     =   "Account Maintenance"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DataEntry"
            Object.ToolTipText     =   "Data Entry"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DetailGL"
            Object.ToolTipText     =   "Detail General Ledger"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stmt"
            Object.ToolTipText     =   "Financial Statements"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Windows GL"
            ImageIndex      =   9
         EndProperty
      EndProperty
      MouseIcon       =   "MainMenu.frx":28DC
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version: 2.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Balint and Associates - Windows GL"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1245
      TabIndex        =   10
      Top             =   840
      Width           =   8655
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblCityStateZip 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "No Company Loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "USER:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblUserName 
      Caption         =   "No user is Currently Loged In"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label lblUserLogon 
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSDImport 
         Caption         =   "&SuperDOS CLIENT Import"
      End
      Begin VB.Menu mnuHistImport 
         Caption         =   "SuperDOS &HISTORY Import"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New File"
      End
      Begin VB.Menu mnuFileSwitchUser 
         Caption         =   "S&witch User"
      End
      Begin VB.Menu mnuSetPassword 
         Caption         =   "Set &Password"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMaint 
      Caption         =   "Maintenance"
      Begin VB.Menu mnuFMCompany 
         Caption         =   "&Company"
      End
      Begin VB.Menu menuFMAccounts 
         Caption         =   "&Accounts / Amounts"
      End
      Begin VB.Menu mnuFMHistory 
         Caption         =   "&History"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFMReportSetup 
         Caption         =   "&Report Setup"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFMBranch 
         Caption         =   "&Branch"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFMJournalSource 
         Caption         =   "&Journal Source"
      End
      Begin VB.Menu mnuFMDescription 
         Caption         =   "&Descriptions"
      End
      Begin VB.Menu mnuFMUsers 
         Caption         =   "&Users"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "Data Entry"
   End
   Begin VB.Menu mnuStatements 
      Caption         =   "Statements"
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuDEJournal 
         Caption         =   "Data Entry &Journal"
      End
      Begin VB.Menu mnuDetailGL 
         Caption         =   "Detail &General Ledger"
      End
      Begin VB.Menu mnuChartOfAccounts 
         Caption         =   "&Chart of Accounts"
      End
      Begin VB.Menu mnuPrintGLAccount 
         Caption         =   "&Print GLAccount"
      End
      Begin VB.Menu mnuTrialBal 
         Caption         =   "&Trial Balance"
      End
      Begin VB.Menu mnuPrintDesc 
         Caption         =   "Print &Description File"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "Utility"
      Begin VB.Menu mnuClearAmounts 
         Caption         =   "Clear &Amounts And Update"
      End
      Begin VB.Menu mnuClearBudget 
         Caption         =   "Clear &Budget Amounts"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUpdateAmounts 
         Caption         =   "&Update Amounts"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFiscalClose 
         Caption         =   "Fiscal Year &Closing"
      End
      Begin VB.Menu mnuDeleteAccounts 
         Caption         =   "&Delete Accounts"
      End
      Begin VB.Menu mnuMultDivide 
         Caption         =   "&Multiply/Divide Accounts"
      End
      Begin VB.Menu mnuCopyBB 
         Caption         =   "C&opy Branch / Budget"
      End
      Begin VB.Menu mnuFileCopy 
         Caption         =   "&File Copy"
      End
      Begin VB.Menu mnuBatchUtilities 
         Caption         =   "Batch &Utilities"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TaskID As Double
Dim x As String
Dim com As New rCompany
Dim DriveLetter As String
Dim ID2 As Long
Dim I As Integer
Dim UID As Long
Public Password As String
Public dbPassWord As String
Dim Cmd As String

Public Sub SetCompany(ByVal ID As Long)
    
Dim db As DAO.Database
Dim FName As String
        
    Response = False
    
    If ID > 0 And com.GetRecord(ID) Then
        
        ' see if a password is required
        On Error Resume Next
        
        ' database expected to be on same drive as .exe
        '   drv letter case ???
        FName = com.FileName
        If Mid(App.Path, 1, 2) <> Mid(com.FileName, 1, 2) Then
           FName = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
        End If
        
        If dbPassWord = "" Then
           
           Set db = OpenDatabase(FName)
    
           If Err.Number = 0 Then ' no password required - OK to continue
              On Error GoTo 0
              db.Close
              Set db = Nothing
              Response = True
              dbPassWord = ""
           Else                ' get the password
              If Err.Description = "Not a valid password." Then
                 On Error GoTo 0
                 frmEnterDBPassword.FileName = FName
                 frmEnterDBPassword.lblCompanyName = com.Name
                 frmEnterDBPassword.lblFileName = FName
                 frmEnterDBPassword.Show vbModal
              Else            ' other error
                 MsgBox "Error opening " & FName & vbCrLf & _
                        Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "Windows GL"
                 Unload Me
                 End
              End If
           End If
        
        Else     ' dbPassword from the command line
        
           Set db = OpenDatabase(Name:=FName, _
                                 Options:=False, _
                                 ReadOnly:=False, _
                                 Connect:=";pwd=" & dbPassWord)
           If Err.Number <> 0 Then
              MsgBox "Database password error !", vbCritical
              Unload Me
              End
           End If
           On Error GoTo 0
        
'           dbPassWord = ""
           db.Close
           Set db = Nothing
           
           Response = True
        
        End If
    End If
    
    If Not Response Then      ' password failed
        
        MenuEnable False
        lblFileName = ""
        lblCompanyName = ""
        curCompany = 0
        CompanyID = 0
    
    Else
        
        MenuEnable True
        lblFileName = Mid(com.FileName, 3, Len(com.FileName) - 2)
        lblCompanyName = com.Name
        lblAddress = com.Address1
        lblCityStateZip = com.City
        If Not com.City = "" Then lblCityStateZip = lblCityStateZip & " " & com.State
        If Not com.ZipCode = "" Then lblCityStateZip = lblCityStateZip & " " & com.ZipCode
        curCompany = ID
        CompanyID = ID
        Me.Caption = "Windows GL - " & com.Name
    
        ' assign to user file
        If User.GetSQL("SELECT * FROM Users WHERE Users.ID = " & CStr(UserID)) = 1 Then
           User(1).LoadLastCompany = True
           User(1).LastCompany = CompanyID
           User(1).PutRecord UserID
        End If
        
        Response = True
    
    End If

End Sub

Private Sub cmdChangeCompany_Click()
    CompanyList.Show vbModal
End Sub

Private Sub cmdChangeUser_Click()
    OnLogOn
End Sub

Private Sub cmdDataEntry_Click()
    On Error GoTo glErr
    If lblFileName = "" Then
        MsgBox "No File Name Selected"
'        cmdChangeCompany.SetFocus
        Exit Sub
    End If

    RetValue = ExecCmd("\balint\glentry.exe " & curUser & " " & CStr(curCompany))

'    Shell "\balint\glentry.exe " & curUser & " " & CStr(curCompany), vbNormalFocus
    
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdEditUser_Click()
    frmUser.ID = curUser
    frmUser.Init
    frmUser.Show 'vbModal
    lblUserLogon = Logon
    lblUserName = UserName
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdJournalStm_Click()
    MsgBox "Journal Statements Call"
End Sub

Private Sub cmdMaint_Click()
    On Error GoTo glErr
    If lblFileName = "" Then
        MsgBox "No File Name Selected"
'        cmdChangeCompany.SetFocus
        Exit Sub
    End If
    
    RetValue = ExecCmd("\balint\glmaint.exe " & curUser)
    
'    Shell "\balint\glmaint.exe " & curUser, vbNormalFocus
    
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdNewCompany_Click()
    On Error GoTo glErr
    CompanyForm.ID = 0
    CompanyForm.Init
    CompanyForm.Show vbModal
    Unload CompanyForm
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdNewUser_Click()
    frmUser.ID = 0
    frmUser.Init
    frmUser.Show vbModal
End Sub

Private Sub cmdSystem_Click()
    frmSystem.Show vbModal
End Sub

Private Sub cmdTrialBalance_Click()
    MsgBox "Trial Balance Call"
End Sub



Private Sub cmdPREntry_Click()
    NewCall "PREntry", ""
End Sub

Private Sub cmdPRMaintEE_Click()
    NewCall "PRMaint", "Employee"
End Sub

Private Sub cmdPRMaintER_Click()
    NewCall "PRMaint", "Employer"
End Sub

Private Sub Form_Load()
    
    DriveLetter = Left(App.Path, 2)
    OnLogOn

End Sub

Private Sub OnLogOn()
    
    ' see if logon and password passed on the command line
    '   set to Balint after first time through - don't process - switch user
    
    If Cmd <> "Balint" Then
        Cmd = Command()
        UserID = GetCmd(Cmd, "UserID", "Num")
        dbPassWord = GetCmd(Cmd, "dbPwd", "Str")
    End If
                
    ' connect to the GLSystem database
    Set cnDes = New ADODB.Connection
    cnDes.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnDes.ConnectionString = "\Balint\Data\GLSystem.mdb"
    cnDes.Open
    
    ' --- error checking ???
    
    ' *****
    ' UpdateCheck True, cnDes
        
    If UserID = 0 Then
       
       frmLogin.Show vbModal        ' show the screen
       curUser = frmLogin.ID
       UID = frmLogin.ID
    
    Else

        SQLString = "SELECT * FROM Users WHERE ID = " & UserID
        rsInit SQLString, cnDes, adoRS
       
        If adoRS.RecordCount = 0 Then
            MsgBox "Logon: " & Logon & " not found !!!"
            Unload Me
            End
        End If
    
        Response = True   ' login from command line success !!
    
        ' set variables for assignment below
        Logon = adoRS!Logon
        UserName = adoRS!Name
        CompanyID = adoRS!LastCompany
    
        curUser = adoRS!ID
        UID = adoRS!ID
    
        If IsNull(adoRS!LastCompany) Then adoRS.Fields("LastCompany") = 0
        
        ' *****
        ' Me.SetCompany (adoRS!LastCompany)
       
        Response = True
       
    End If
    
    If Not Response Then
        
        Unload Me
        End
        
        curUser = 0
        lblUserLogon = ""
        lblUserName = "No user logged on!"
    
    Else
        
        lblUserLogon = Logon
        lblUserName = UserName
    
    End If

    Cmd = "Balint"

End Sub

Private Sub menuFMAccounts_Click()

    NewCall "GLMaint", "Account"
    
End Sub

Private Sub mnuBatchUtilities_Click()

    MenuCall "GLUtil", "x"
    
End Sub

Private Sub mnuChartOfAccounts_Click()

    NewCall "GLPrint", "ChartOfAccounts"
    
End Sub

Private Sub mnuClearAmounts_Click()
    
    NewCall "GLUtil", "ClearGLAmount"

End Sub

Private Sub mnuClearBudget_Click()

    NewCall "GLUtil", "ClearGLBudget"

End Sub

Private Sub mnuCopyBB_Click()
    
    NewCall "GLUtil", "CopyBB"
    
End Sub

Private Sub mnuDataEntry_Click()
    
    NewCall "GLEntry", "GLEntry"
    
End Sub

Private Sub mnuDEJournal_Click()
    
    NewCall "GLPrint", "GLHistJnl"
    
End Sub

Private Sub mnuDeleteAccounts_Click()

    NewCall "GLUtil", "DeleteAccts"
    
End Sub

Private Sub mnuDetailGL_Click()
    
    NewCall "GLPrint", "DetailGL"
    
End Sub

Private Sub mnuExit_Click()
    mnuFileExit_Click
End Sub

Private Sub mnuFileCopy_Click()

    NewCall "GLUtil", "GLFileCopy"
    
End Sub

Private Sub mnuFileExit_Click()
    
    If User.GetSQL("SELECT * FROM Users WHERE Users.ID = " & CStr(UserID)) = 1 Then
       
'       User.Name = frmLogin.txtName
'       User.Login = LogOn
'       User.Password = frmLogin.txtPassword
'       User.ID = UID
       
       User(1).LoadLastCompany = True
       User(1).LastCompany = CompanyID
       User(1).PutRecord UserID
    
    End If
    
    End

End Sub

Private Sub mnuFileNew_Click()
    
    NewCall "GLUtil", "NewFile"

End Sub

Private Sub mnuFileOpen_Click()
    
    dbPassWord = ""
    CompanyList.Show vbModal
    
End Sub

Private Sub mnuFileSDImport_Click()
    
    ID2 = CompanyID
    
    NewCall "GLUtil", "Import"
    
    End   ' !!!
    
    ' open the company just imported - get from the user record
    If User.GetSQL("SELECT * FROM Users WHERE LogOn = '" & lblUserLogon & "'") Then
       If ID2 <> User.LastCompany Then
          Me.SetCompany (User.LastCompany)
       End If
    End If
    
End Sub

Private Sub mnuFileSwitchUser_Click()
    dbPassWord = ""
    UserID = 0
    OnLogOn
End Sub

Private Sub mnuFiscalClose_Click()

    NewCall "GLUtil", "YearEnd"
    
End Sub

Private Sub mnuFMBranch_Click()

    MenuCall "GLMaint", "Branch"
    
End Sub

Private Sub mnuFMCompany_Click()
    CompanyForm.ID = CompanyID
    CompanyForm.Show vbModal
    Me.SetCompany CompanyID
End Sub

Private Sub mnuFMDescription_Click()
    
    NewCall "GLMaint", "Descriptions"

End Sub

Private Sub mnuFMJournalSource_Click()

    NewCall "GLMaint", "Journal"
    
End Sub

Private Sub mnuFMUsers_Click()

    NewCall "GLMaint", "User"

End Sub

Private Sub mnuHistImport_Click()
    
    NewCall "GLUtil", "HistImport"

End Sub

Private Sub mnuMultDivide_Click()
    
    NewCall "GLUtil", "GLMultDiv"

End Sub

Private Sub MenuCall(ByVal ModuleName As String, ByVal ProgName As String)
    
    If ModuleName = "GLEntry" Then
       
       x = DriveLetter & "\Balint\GLEntry.exe " & _
           "UserID=" & UserID & " " & _
           "Password=" & Password & " " & _
           "CompanyID=" & CompanyID
       
'       RetValue = ExecCmd(x)
       
       TaskID = Shell(x, vbMaximizedFocus)
       
       Exit Sub
    
    End If
    
    x = DriveLetter & "\Balint\" & ModuleName & ".exe " & _
        CompanyID & "/" & _
        dbPassWord & "/" & _
        ProgName & "/" & _
        DriveLetter & "\Balint\Data\GLSystem.mdb" & "/" & _
        Logon
    
'    RetValue = ExecCmd(x)
    
    TaskID = Shell(x, vbNormalFocus)
    AppActivate TaskID

End Sub

Private Sub NewCall(ByVal ModuleName As String, ByVal ProgName As String)

    x = DriveLetter & "\Balint\" & ModuleName & ".exe" & _
        " ProgName=" & ProgName & _
        " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
        " UserID=" & UserID & _
        " BackName=" & DriveLetter & "\Balint\GLMenu.exe"

    ' database password if required
    If dbPassWord <> "" Then
       x = x & " dbPWd=" & dbPassWord
    End If
        
    TaskID = Shell(x, vbMaximizedFocus)
'    AppActivate TaskID

    Unload Me
    End

End Sub


Private Sub mnuPrintDesc_Click()

    NewCall "GLPrint", "PrintDesc"
    
End Sub

Private Sub mnuPrintGLAccount_Click()

    NewCall "GLPrint", "PrintGLAccount"
    
End Sub

Private Sub mnuSetPassword_Click()
    
    frmSetDBPassword.lblCompanyName = Me.lblCompanyName
    frmSetDBPassword.lblFileName = Me.lblFileName
    frmSetDBPassword.Show vbModal
    
End Sub

Private Sub mnuStatements_Click()
    
'    Me.WindowState = vbMinimized
    NewCall "GLPrint", "Statement"
'    Me.WindowState = vbMaximized
'    Me.SetFocus

End Sub

Private Sub mnuTrialBal_Click()
    
    NewCall "GLPrint", "TrialBal"

End Sub

Private Sub mnuUpdateAmounts_Click()

    MenuCall "GLUtil", "UpdateGLAmount"
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
       Case "Open"
            CompanyList.Show vbModal
       Case "AcctMaint"
            NewCall "GLMaint", "Account"
'            MenuCall "GLMaint", "Account"
       Case "DetailGL"
            NewCall "GLPrint", "DetailGL"
'            MenuCall "GLPrint", "DetailGL"
       Case "DataEntry"
            NewCall "GLEntry", "Entry"
       Case "Stmt"
            NewCall "GLPrint", "Statement"
       Case "Exit"
            End
    End Select

End Sub

Private Sub MenuEnable(ByVal TF As Boolean)

    ' allow user maint
    mnuMaint.Enabled = True
    Me.mnuFMUsers.Enabled = True
    
    Me.mnuFMCompany.Enabled = TF
    Me.menuFMAccounts.Enabled = TF
    Me.mnuFMHistory.Enabled = TF
    Me.mnuFMReportSetup.Enabled = TF
    Me.mnuFMBranch.Enabled = TF
    Me.mnuFMJournalSource.Enabled = TF
    Me.mnuFMDescription.Enabled = TF
    
    mnuDataEntry.Enabled = TF
    mnuStatements.Enabled = TF
    mnuReports.Enabled = TF
    mnuUtility.Enabled = TF
    Me.mnuSetPassword.Enabled = TF

End Sub
