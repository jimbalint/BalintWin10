VERSION 5.00
Begin VB.Form frmSweep 
   Caption         =   "GL Sweeps"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
End
Attribute VB_Name = "frmSweep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    GLPrint.OpenRS
    GLPrint.GetData "JimBo", False
    MsgBox GLPrint.Output
    
    End

End Sub
