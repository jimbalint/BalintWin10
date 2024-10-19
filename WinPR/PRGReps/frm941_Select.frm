VERSION 5.00
Begin VB.Form frm941_Select 
   Caption         =   "Select Form 941 Version"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9555
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMar24 
      Caption         =   "March 2024"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   5400
      Width           =   3735
   End
   Begin VB.CommandButton cmdApr23 
      Caption         =   "April 2023"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   4560
      Width           =   3735
   End
   Begin VB.CommandButton cmdMar22 
      Caption         =   "March 2022"
      Height          =   615
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   3720
      Width           =   3735
   End
   Begin VB.CommandButton cmdJun21 
      Caption         =   "June 2021"
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton cmdJun20 
      Caption         =   "June 2020"
      Height          =   615
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   2355
      TabIndex        =   2
      Top             =   8280
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "10/19/2024"
      Height          =   495
      Left            =   6720
      TabIndex        =   5
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select which 941 Form you have"
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   1260
      Width           =   5715
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   7875
   End
End
Attribute VB_Name = "frm941_Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Form941 As Byte



Private Sub Form_Load()
    Me.KeyPreview = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub


Private Sub cmdJun20_Click(Index As Integer)
    ' added 2020-07-15
    Form941 = 9
    Me.Hide
End Sub

Private Sub cmdJun21_Click(Index As Integer)
    Form941 = 10
    Me.Hide
End Sub

Private Sub cmdMar22_Click(Index As Integer)
    Form941 = 11
    Me.Hide
End Sub

Private Sub cmdApr23_Click()
    Form941 = 12
    Me.Hide
End Sub

Private Sub cmdMar24_Click()
    Form941 = 13
    Me.Hide
End Sub

