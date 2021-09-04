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
   Begin VB.CommandButton cmdJun21 
      Caption         =   "June 2021"
      Height          =   615
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdJun20 
      Caption         =   "June 2020"
      Height          =   615
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton cmdJan17 
      Caption         =   "January 2017"
      Height          =   615
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   3120
      Width           =   3735
   End
   Begin VB.CommandButton cmdJan14 
      Caption         =   "January 2014"
      Height          =   615
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   2160
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

Private Sub cmdFeb2010_Click()
    Form941 = 1
    Me.Hide
End Sub

Private Sub cmdApril2010_Click()
    Form941 = 2
    Me.Hide
End Sub
Private Sub cmdJan2011_Click()
    Form941 = 3
    Me.Hide
End Sub

Private Sub cmdJan2012_Click()
    Form941 = 4
    Me.Hide
End Sub

Private Sub cmdJan2013_Click(Index As Integer)
    Form941 = 5
    Me.Hide
End Sub

Private Sub cmdJan13v2_Click(Index As Integer)
    ' added 2014-07-18
    Form941 = 6
    Me.Hide
End Sub

Private Sub cmdJan14_Click(Index As Integer)
    
    ' added 2014-08-26
    Form941 = 7
    Me.Hide

End Sub
Private Sub cmdJan17_Click(Index As Integer)
    
    ' added 2017-04-08
    Form941 = 8
    Me.Hide

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

