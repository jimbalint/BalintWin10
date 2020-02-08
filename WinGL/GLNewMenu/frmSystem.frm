VERSION 5.00
Begin VB.Form frmSystem 
   Caption         =   " System Maintenance"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmSystem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDescGrid 
      Caption         =   "GENERAL LEDGER Descriptions"
      Height          =   975
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdVersion 
      Caption         =   "DataBase Version"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Converts all Data Files to Most Recent Version"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Entry and Editing of the General Ledger Description List"
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDescGrid_Click()
    DescGrid.Show vbModal
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdVersion_Click()
    On Error GoTo glErr
    
    ExecCmd ("\Balint\Versions.exe")

'    Shell "\balint\version.exe", vbNormalFocus
    
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub
