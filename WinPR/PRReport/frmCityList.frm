VERSION 5.00
Begin VB.Form frmCityList 
   Caption         =   "City Rate List"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4703
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1583
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.OptionButton optName 
      Caption         =   "Sort By City Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2543
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton optNumber 
      Caption         =   "Sort By City Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2543
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "COMPANY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   23
      TabIndex        =   4
      Top             =   360
      Width           =   7695
   End
End
Attribute VB_Name = "frmCityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Me.KeyPreview = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
   CityList
End Sub

