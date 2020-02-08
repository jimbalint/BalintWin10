VERSION 5.00
Begin VB.Form frmGLPrint2 
   Caption         =   "GLPrint - Other Options"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   Icon            =   "GLPrint2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold Print"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame frmWide 
      Caption         =   " Wide Statements "
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   2400
      Width           =   3135
      Begin VB.OptionButton optLandscape 
         Caption         =   "Landscape"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCompressed 
         Caption         =   "Compressed"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkLowerCaseDate 
      Caption         =   "Lower case date"
      Height          =   495
      Left            =   938
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CheckBox chkRoundDollars 
      Caption         =   "Round to nearest dollar"
      Height          =   495
      Left            =   4058
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkPrintZeroBal 
      Caption         =   "Print Zero Balance Accounts"
      Height          =   495
      Left            =   938
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkPrtAcctNum 
      Caption         =   "Print Account Numbers"
      Height          =   495
      Left            =   4058
      TabIndex        =   3
      Top             =   1020
      Width           =   2415
   End
   Begin VB.CheckBox chkUseMathRecs 
      Caption         =   "Use Math Records"
      Height          =   495
      Left            =   938
      TabIndex        =   2
      Top             =   1020
      Width           =   2415
   End
   Begin VB.CheckBox chkSupprCP 
      Caption         =   "Suppress Current Period"
      Height          =   495
      Left            =   4058
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CheckBox chkSepPage 
      Caption         =   "GL Separate Pages"
      Height          =   495
      Left            =   938
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   615
      Left            =   4178
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   615
      Left            =   2018
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   3600
      Width           =   3255
   End
End
Attribute VB_Name = "frmGLPrint2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
 Me.Label1 = GLPrint.ReportName
    
    CheckSet GLPrint.SepPage, Me.chkSepPage
    CheckSet GLPrint.SupprCP, Me.chkSupprCP
    CheckSet GLPrint.UseMathRec, Me.chkUseMathRecs
    CheckSet GLPrint.PrtAcctNum, Me.chkPrtAcctNum
    CheckSet GLPrint.PrtZeroBal, Me.chkPrintZeroBal
    CheckSet GLPrint.RoundDollars, Me.chkRoundDollars
    CheckSet GLPrint.LowerCaseDate, Me.chkLowerCaseDate
   
    If GLPrint.WidePrint = True Then
        frmGLPrint2.optCompressed = True
        frmGLPrint2.optLandscape = False
    Else
        frmGLPrint2.optCompressed = False
        frmGLPrint2.optLandscape = True
    End If

    If GLPrint.Output = "Bold" Then
        Me.chkBold = 1
    Else
        Me.chkBold = 0
    End If
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   
   If frmGLPrint2.chkSepPage = 1 Then
      GLPrint.SepPage = True
   Else
      GLPrint.SepPage = False
   End If
   
   If frmGLPrint2.chkSupprCP = 1 Then
      GLPrint.SupprCP = True
   Else
      GLPrint.SupprCP = False
   End If
   
   If frmGLPrint2.chkUseMathRecs = 1 Then
      GLPrint.UseMathRec = True
   Else
      GLPrint.UseMathRec = False
   End If
      
   If frmGLPrint2.chkPrtAcctNum = 1 Then
      GLPrint.PrtAcctNum = True
   Else
      GLPrint.PrtAcctNum = False
   End If
   
   If frmGLPrint2.chkPrintZeroBal = 1 Then
      GLPrint.PrtZeroBal = True
   Else
      GLPrint.PrtZeroBal = False
   End If
   
   If frmGLPrint2.chkRoundDollars = 1 Then
      GLPrint.RoundDollars = True
   Else
      GLPrint.RoundDollars = False
   End If
   
   If frmGLPrint2.chkLowerCaseDate = 1 Then
      GLPrint.LowerCaseDate = True
   Else
      GLPrint.LowerCaseDate = False
   End If
   
   If frmGLPrint2.optCompressed = True Then
      GLPrint.WidePrint = True
   Else
      GLPrint.WidePrint = False
   End If
   
   If Me.chkBold = 1 Then
      GLPrint.Output = "Bold"
   Else
      GLPrint.Output = ""
   End If
   
   Unload Me
   
End Sub

Private Sub CheckSet(ByVal tf As Boolean, ByRef cb As CheckBox)

    If tf = True Then
        cb = 1
    Else
        cb = 0
    End If

End Sub

