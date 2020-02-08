VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   " General Ledger Configuration"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox chkLoadLast 
      Caption         =   " Load Last File on Startup"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If glLoadLast = False Then
        chkLoadLast.Value = 0
    Else
        chkLoadLast.Value = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    glLoadLast = chkLoadLast.Value
End Sub
