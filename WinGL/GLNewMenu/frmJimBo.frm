VERSION 5.00
Begin VB.Form frmJimBo 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   615
      Left            =   7320
      TabIndex        =   5
      Top             =   5640
      Width           =   1695
   End
   Begin VB.OptionButton radClient 
      Caption         =   "Client"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin VB.OptionButton radGLSystem 
      Caption         =   "GLSystem"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox txtSQL 
      Height          =   4455
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmJimBo.frx":0000
      Top             =   600
      Width           =   11415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   600
      TabIndex        =   4
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Label lblHdr 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmJimBo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim resp As Integer
    Dim x As String
    x = "OK to execute Query on " & IIf(Me.radGLSystem.Value, "GLSystem", Me.lblHdr.Caption) & _
        vbCr & vbCr & Me.txtSQL.text
    resp = MsgBox(x, vbYesNo + vbQuestion, "JimBo Sweeps")
    If resp = vbNo Then
        Me.Hide
        Exit Sub
    End If
    If Me.radGLSystem.Value = True Then
        cnDes.Execute Me.txtSQL.text
    Else
        cn.Execute Me.txtSQL.text
    End If
End Sub

Private Sub Form_Load()
    Me.txtSQL.text = ""
    Me.radGLSystem.Value = True
End Sub
