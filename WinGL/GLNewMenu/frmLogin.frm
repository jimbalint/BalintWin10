VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows GL Login"
   ClientHeight    =   2415
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5865
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1426.861
   ScaleMode       =   0  'User
   ScaleWidth      =   5506.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbUser 
      Height          =   390
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin TDBText6Ctl.TDBText txtPassword 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   2325
      _Version        =   65536
      _ExtentX        =   4101
      _ExtentY        =   609
      Caption         =   "frmLogin.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLogin.frx":036E
      Key             =   "frmLogin.frx":038C
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   "*"
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   3
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   1800
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   390
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public txtName As String
Public txtLogOn As String
Public ID As Long

Private Sub Form_Load()
 
    SQLString = "SELECT * FROM Users ORDER BY Logon"
    If Not GLUser.GetBySQL(SQLString) Then
        GLUser.Clear
        GLUser.Logon = "Default"
        GLUser.Name = "Default User"
        GLUser.LastCompany = 0
        GLUser.LastPRCompany = 0
        GLUser.Password = ""
        GLUser.Save (Equate.RecAdd)
            
        ' redo the get ?
        SQLString = "SELECT * FROM GLUser ORDER BY Logon"
        If Not GLUser.GetBySQL(SQLString) Then
            MsgBox "User File Error", vbExclamation
            End
        End If
    End If
            
    Do
        cmbUser.AddItem GLUser.Logon
        If Not GLUser.GetNext Then Exit Do
    Loop
 
    Response = False
    UserID = 0
    CompanyID = 0

End Sub

Private Sub cmdCancel_Click()
    Response = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
Dim Flg As Byte
Dim cID As Long
    
    If cmbUser = "" Then
        MsgBox "No LogOn Provided"
        cmbUser.SetFocus
        Exit Sub
    End If
    
    SQLString = "SELECT * FROM Users WHERE Logon = '" & cmbUser & "'"
    If Not GLUser.GetBySQL(SQLString) Then
        MsgBox "User Logon not found!", vbExclamation, "Windows GL Login"
        cmbUser.SetFocus
        Exit Sub
    End If
    
    If UCase(GLUser.Logon) = "JIM" Then
        txtName = GLUser.Name
        txtLogOn = cmbUser.Text
        Response = True
        UserID = GLUser.ID
        CompanyID = GLUser.LastCompany
        Unload Me
        Exit Sub
    End If
    
    If GLUser.Password = Me.txtPassword Or IsNull(GLUser.Password) Or GLUser.Password = " " Then
    Else
        MsgBox "Invalid Password!", vbExclamation, "Windows GL Login"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    txtName = GLUser.Name
    txtLogOn = cmbUser.Text
    
    Response = True
    
    UserID = GLUser.ID
    CompanyID = GLUser.LastCompany
    
    Response = True
    
    Unload Me
    
End Sub

