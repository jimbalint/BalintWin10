VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmSetDBPassword 
   Caption         =   "Set DataBase Password"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   615
      Left            =   3278
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin TDBText6Ctl.TDBText tdbOldPassword 
      Height          =   375
      Left            =   5198
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "SetDBPassword.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "SetDBPassword.frx":0064
      Key             =   "SetDBPassword.frx":0082
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   5678
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin TDBText6Ctl.TDBText tdbNewPassword 
      Height          =   375
      Left            =   5198
      TabIndex        =   1
      Top             =   3120
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "SetDBPassword.frx":00C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "SetDBPassword.frx":012A
      Key             =   "SetDBPassword.frx":0148
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
   Begin TDBText6Ctl.TDBText tdbConfirmPassword 
      Height          =   375
      Left            =   5198
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   661
      Caption         =   "SetDBPassword.frx":018C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "SetDBPassword.frx":01F0
      Key             =   "SetDBPassword.frx":020E
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
   Begin VB.Label Label3 
      Caption         =   "Confirm New Password:"
      Height          =   255
      Left            =   1478
      TabIndex        =   9
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "New Password:"
      Height          =   255
      Left            =   1478
      TabIndex        =   8
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password (blank for none): "
      Height          =   255
      Left            =   1478
      TabIndex        =   7
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   8055
   End
End
Attribute VB_Name = "frmSetDBPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rs As Recordset
Dim x As String

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Me.tdbConfirmPassword <> Me.tdbNewPassword Then
       MsgBox "Password confirm failed !", vbExclamation + vbOKOnly, "Set DataBase Password"
       Exit Sub
    End If
    
    ' confirm old password
    On Error Resume Next
    
    SysName = "\balint\data\glsystem.mdb"
    SysName = Me.lblFileName
    
    If Me.tdbOldPassword = "" Then
       Set db = OpenDatabase(Name:=SysName, _
                             Options:=True, _
                             ReadOnly:=False)
    Else
       
       Set db = OpenDatabase(Name:=SysName, _
                             Options:=True, _
                             ReadOnly:=False, _
                             Connect:=";pwd=" & Me.tdbOldPassword)
    End If
    
'    On Error GoTo 0
    
    If Err Then
       If Err.Description = "Not a valid password." Then
          MsgBox "Invalid Old Password!", vbExclamation + vbOKOnly, "Set DataBase Password"
       Else
          MsgBox "DataBase Open Error: " & Err.Description & " " & _
                 Err.Number, vbExclamation + vbOKOnly, "Set DataBase Password"
       End If
       Exit Sub
    Else     ' OK
    
       ' change it
       db.NewPassWord Me.tdbOldPassword, Me.tdbNewPassword
       db.Close
       
       MsgBox "The Password has been changed.", vbInformation + vbOKOnly, "Set DataBase Password"
       
    End If
    
    MainMenu.Password = Me.tdbNewPassword
    
    Me.Hide

End Sub
