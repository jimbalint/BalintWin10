VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Record"
   ClientHeight    =   5010
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5370
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2960.073
   ScaleMode       =   0  'User
   ScaleWidth      =   5042.14
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFullName 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1440
      TabIndex        =   6
      Top             =   4320
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   390
      Left            =   2640
      TabIndex        =   7
      Top             =   4320
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "&Full Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   2160
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
    On Error GoTo glErr
    Dim cc As New ccUsers
    cc.login = txtUserName
    cc.name = txtFullName
    cc.password = txtPassword
    cc.PutRecord ID
    Unload Me
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Private Sub Form_Load()
    Init
End Sub

Public Sub Init()
    If ID = 0 Then
        txtUserName = ""
        txtFullName = ""
        txtPassword = ""
    Else
        Dim cc As New ccUsers
        If 1 = cc.GetSQL("select * from Users where ID=" & ID) Then
            txtUserName = cc(1).login
            txtFullName = cc(1).name
            txtPassword = cc(1).password
        End If
    End If
End Sub
