VERSION 5.00
Begin VB.Form UserForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Record"
   ClientHeight    =   4545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685.336
   ScaleMode       =   0  'User
   ScaleWidth      =   4323.846
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFullName 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1200
      TabIndex        =   6
      Top             =   3720
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   390
      Left            =   2400
      TabIndex        =   7
      Top             =   3720
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "&Full Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Integer
Public cc As New ccUsers

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo glerr
    Dim cc As New ccUsers
    cc.login = txtUserName
    cc.name = txtFullName
    cc.password = txtPassword
    cc.PutRecord ID
    Unload Me
    Exit Sub
glerr:
    MsgBox Error(Err.Number)
End Sub

Private Sub Form_Load()
    OnInit
End Sub

Public Sub OnInit()
    If ID = 0 Then
        txtUserName = ""
        txtFullName = ""
        txtPassword = ""
    Else
        Dim cc As New ccUsers
        If 1 = cc.SetSQL("select * from Users where ID=" & ID) Then
            txtUserName = cc(1).login
            txtFullName = cc(1).name
            txtPassword = cc(1).password
        End If
    End If
End Sub
