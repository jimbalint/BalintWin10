VERSION 5.00
Begin VB.Form frmCopy 
   Caption         =   "Copy Company"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkClearPR 
      Caption         =   "Clear PR Amounts and History?"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   3840
      Width           =   4095
   End
   Begin VB.CheckBox chkCopyPR 
      Caption         =   "Copy PR Information?"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3240
      Width           =   4095
   End
   Begin VB.CheckBox chkCopyGL 
      Caption         =   "Copy GL Information?"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox txtCompName 
      Height          =   420
      Left            =   4245
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CheckBox chkClearGL 
      Caption         =   "Clear GL Amounts and History ?"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   4245
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
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
      Height          =   495
      Left            =   5625
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "File will be stored in the \Balint\Data Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   11
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name:"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Data Base File Name to copy to:"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   585
      TabIndex        =   8
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As String
Dim I As Integer
Public FName As String
Dim FileExt As String
    
Private Sub Form_Load()

Dim d As Variant
Dim s As Long

    Me.lblCompanyName = GLCompany.Name

    Me.chkCopyGL = 1
    Me.chkCopyPR = 1

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub cmdOK_Click()
    
Dim GName, FromName As String
Dim Pos As Integer
    
    If NewADO Then
        FileExt = ".accdb"
    Else
        FileExt = ".mdb"
    End If
    
    If Me.chkCopyGL = 0 And Me.chkCopyPR = 0 Then
        MsgBox "No Action Selected", vbInformation, "GL File Copy"
        GoBack
    End If
    
    If BalintFolder = "" Then
        x = Left(App.Path, 1) & ":\Balint\Data\" & frmCopy.txtFileName & FileExt
        FromName = Left(App.Path, 1) & Mid(GLCompany.FileName, 2, Len(GLCompany.FileName) - 1)
    Else
        x = BalintFolder & "\Data\" & frmCopy.txtFileName & FileExt
        ' get the string to the right of the last back slash
        GName = Trim(GLCompany.FileName)
        FromName = ""
        Pos = InStrRev(GName, "\", Len(GName), vbTextCompare)
        If Pos <= 0 Then
            MsgBox "Invalid GL Company file name: " & GName, vbCritical
            GoBack
        End If
        FromName = BalintFolder & "\Data\" & Mid(GName, Pos + 1, Len(GName) - Pos)
    End If
            
    FName = x
            
    ' confirm before copy
    I = MsgBox("Copy " & FromName & vbCr & _
               " to: " & x & " ?", _
                vbQuestion + vbYesNo + vbDefaultButton2, "GL File Copy")
            
    If I = vbNo Then GoBack
            
    GLFileCopy FromName, x
            
    GoBack

End Sub


Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

