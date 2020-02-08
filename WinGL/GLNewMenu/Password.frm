VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmPassword 
   Caption         =   "Set Password"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin TDBText6Ctl.TDBText txtPass2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Password.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Password.frx":006C
      Key             =   "Password.frx":008A
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
   Begin TDBText6Ctl.TDBText txtPass1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Password.frx":00CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Password.frx":013A
      Key             =   "Password.frx":0158
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
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Password..."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Retype to Confirm..."
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NewPassWord As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtPass1 = txtPass2 Then
       NewPassWord = txtPass1
       Unload Me
    Else
       MsgBox "Password confirmation failed !", vbExclamation + vbOKOnly, "User Password Set"
       txtPass1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    NewPassWord = "ERROR"
End Sub

