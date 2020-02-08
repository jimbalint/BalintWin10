VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login for General Ledger"
   ClientHeight    =   1425
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   841.937
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   390
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblUsername 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bUserOk As Boolean
Private bPassOk As Boolean
Private SuperLogon As String
Private mcn As ADODB.Connection
Private mrs As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    CheckUser
    If bUserOk = False Then
        MsgBox "User Not Found for Login"
        txtUserName = ""
        txtUserName.SetFocus
        Exit Sub
    End If
    If bPassOk = False Then
        MsgBox "Password is not correct"
        txtPassword = ""
        txtPassword.SetFocus
        Exit Sub
    End If
    txtPassword_Change
End Sub

Private Sub Form_Load()
    bPassOk = False
    bUserOk = False
End Sub

Private Sub txtPassword_Change()
    If bUserOk = False Then Exit Sub
    If UCase(mrs!Password) = UCase(txtPassword) Then
        glUserName = mrs!Name
        glUserID = mrs!ID
        mrs.MoveFirst
        If mrs!logon = SuperLogon Then
            glSuperUser = True
        Else
            glSuperUser = False
        End If
        Unload Me
    End If
End Sub

Private Sub txtUserName_LostFocus()
    CheckUser
End Sub

Private Sub CheckUser()
    Set mcn = New ADODB.Connection
    mcn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\Balint\Users.mdb"
    mcn.Open
    SetAdo mcn, mrs, "select [ID],[Name],[Logon],[Password] from Users"
    mrs.MoveFirst
    SuperLogon = mrs!logon
    mrs.Close
    SetAdo mcn, mrs, "select [ID],[Name],[Logon],[Password] from Users where Logon = '" & txtUserName & "'"
    If mrs.RecordCount Then bUserOk = True
End Sub

