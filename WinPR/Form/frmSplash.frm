VERSION 5.00
Begin VB.Form frmSplash 
   Caption         =   "Balint Accounting"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4695
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   2160
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Balint and Associates Windows Accounting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   6015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Show
    Me.Refresh
End Sub
