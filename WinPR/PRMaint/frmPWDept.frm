VERSION 5.00
Begin VB.Form frmPWDept 
   Caption         =   "Prevailing Wage by Department"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10305
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbWorkCat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10440
      TabIndex        =   0
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Classification:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmPWDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsWorkCat As New ADODB.Recordset
Dim rsCounty As New ADODB.Recordset

Private Sub Form_Load()

    ' use PRGlobal for storage
    
    


    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

