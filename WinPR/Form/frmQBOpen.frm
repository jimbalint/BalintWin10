VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQBOpen 
   Caption         =   "Open QuickBooks Data File"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLookUp 
      Height          =   495
      Left            =   10440
      Picture         =   "frmQBOpen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtFileName 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   10095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6480
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Select the company QuickBooks File:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   1590
      Left            =   120
      Picture         =   "frmQBOpen.frx":030A
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmQBOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileName As String

Private Sub Form_Load()

    FileName = ""
    
    Me.lblCompanyName = PRCompany.Name
    
    Me.KeyPreview = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    FileName = ""
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    FileName = Me.txtFileName
    Me.Hide
End Sub

Private Sub cmdLookUp_Click()
                        
    ' ask for the file name
    Me.CommonDialog1.DialogTitle = "QB File to open"
    Me.CommonDialog1.Filter = "QB Data Files (*.qbw)|*.qbw"
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        Me.txtFileName = Me.CommonDialog1.FileName
    Else
        Me.txtFileName = ""
    End If

End Sub

