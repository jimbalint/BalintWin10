VERSION 5.00
Begin VB.Form DescForm 
   Caption         =   " GL Description Record"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "DescForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Number"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "DescForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public userOK As Boolean

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo glErr
    If txtNumber = "" Then
        MsgBox "Number Not Filled In"
        txtNumber.SetFocus
        Exit Sub
    End If
    If False = IsNumeric(txtNumber) Then
        MsgBox "Number is not a valid Numeric"
        txtNumber.SetFocus
        Exit Sub
    End If
    Dim cc As New ccDescriptions
    ' Check for Double Numbers Here
    cc.description = txtDescription
    cc.number = CInt(txtNumber)
    cc.PutRecord ID
    userOK = True
    Me.Hide
    Exit Sub
glErr:
    MsgBox Error(Err.number)
End Sub

Public Sub Init()
    userOK = False
    ID = 0

    If txtNumber = "" Then
        txtNumber = ""
        txtDescription = ""
    Else
        Dim cc As New ccDescriptions
        cc.GetSQL "select * from gldescriptions where Number=" & txtNumber
        If cc.Records >= 1 Then
            ID = cc(1).ID
            txtNumber = cc(1).number
            txtDescription = cc(1).description
        End If
    End If
End Sub

