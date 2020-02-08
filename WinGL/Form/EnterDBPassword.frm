VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmEnterDBPassword 
   Caption         =   "Windows GL"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin TDBText6Ctl.TDBText tdbPassword 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   661
      Caption         =   "EnterDBPassword.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "EnterDBPassword.frx":006C
      Key             =   "EnterDBPassword.frx":008A
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4508
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2228
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      Caption         =   "FileName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "CompanyName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter DataBase Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   735
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
   End
End
Attribute VB_Name = "frmEnterDBPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As String
Public FileName As String
Private Sub Form_Load()
    Me.tdbPassword = ""
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Me.Hide

'    On Error Resume Next
'
'    Set db = OpenDatabase(Name:=FileName, _
'                          Options:=False, _
'                          ReadOnly:=False, _
'                          Connect:=";pwd=" & Me.tdbPassword)
'
'    If Err.Description <> "" Then
'       If Err.Description = "Not a valid password." Then
'          MsgBox "Invalid Password!", vbExclamation + vbOKOnly, "Windows GL"
'       Else
'          MsgBox "Database Error: " & Err.Description & " " & Err.Number, vbCritical + vbOKOnly, "Windows GL"
'          Unload Me
'       End If
'    Else
'       db.Close
'       Set db = Nothing
'       Response = True
'       MainMenu.dbPassword = Me.tdbPassword
'       Unload Me
'    End If

End Sub

