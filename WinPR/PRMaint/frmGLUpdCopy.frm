VERSION 5.00
Begin VB.Form frmGLUpdCopy 
   Caption         =   "Copy GL Update Setup"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   4185
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox cmbCopyTo 
      Height          =   390
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
      Width           =   6255
   End
   Begin VB.ComboBox cmbCopyFrom 
      Height          =   390
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "Copy To:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Copy From:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmGLUpdCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CopyFromID, CopyToID As Long

Private Sub Form_Load()

'    trsGLT.Fields.Append "GLType", adInteger
'    trsGLT.Fields.Append "RelatedID", adDouble
'    trsGLT.Fields.Append "GLName", adVarChar, 30

    If frmGLUpd.trsGLT.RecordCount = 0 Then
        MsgBox "No Records to copy!", vbExclamation
        CopyToID = 0
        Me.Hide
    End If

    Me.lblCompanyName = PRCompany.Name
    
    PopCmb Me.cmbCopyFrom, frmGLUpd.trsGLT
    PopCmb Me.cmbCopyTo, frmGLUpd.trsGLT
    
    Me.KeyPreview = True

End Sub

Private Sub cmdExit_Click()
    CopyToID = 0
    Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub PopCmb(ByRef cmb As ComboBox, ByRef rs As ADODB.Recordset)

    With cmb
        .Clear
        rs.MoveFirst
        Do
            .AddItem rs!GLName
            .ItemData(.NewIndex) = rs!RelatedID
            rs.MoveNext
        Loop Until rs.EOF
    End With
        
End Sub
Private Sub cmdOK_Click()

    If Me.cmbCopyFrom.ListIndex = Me.cmbCopyTo.ListIndex Then
        MsgBox "Copy From/To is the same!", vbExclamation
        Exit Sub
    End If

    With Me.cmbCopyFrom
        If .ListIndex = -1 Then
            MsgBox "Please select the copy from item!", vbExclamation
            Exit Sub
        End If
        CopyFromID = .ItemData(.ListIndex)
    End With
    
    With Me.cmbCopyTo
        If .ListIndex = -1 Then
            MsgBox "Please select the copy to item!", vbExclamation
            Exit Sub
        End If
        CopyToID = .ItemData(.ListIndex)
    End With
    
    Me.Hide

End Sub


