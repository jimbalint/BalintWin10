VERSION 5.00
Begin VB.Form frmPWAdd 
   Caption         =   "Prevailing Wage Maintenance"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
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
   ScaleHeight     =   3195
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEntry 
      Height          =   390
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1245
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4005
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmPWAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Public GlobalType As Byte
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    OK = False
    Me.txtEntry.SetFocus
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    ' make sure does not already exist
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & GlobalType & _
                " AND Description = '" & Me.txtEntry & "'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        MsgBox "This entry already exists!", vbExclamation
        Exit Sub
    End If
    
    PRGlobal.Clear
    PRGlobal.TypeCode = GlobalType
    PRGlobal.Description = Me.txtEntry
    PRGlobal.Save (Equate.RecAdd)
    
    OK = True
    Me.txtEntry.SetFocus
    Me.Hide

End Sub

Public Sub Init()

    Me.txtEntry = ""
    
    Select Case GlobalType
        Case PREquate.GlobalTypePWCraft:        Me.lblMsg1 = "Enter NEW Job Craft"
        Case PREquate.GlobalTypePWCounty:       Me.lblMsg1 = "Enter NEW County"
        Case PREquate.GlobalTypePWUnion:        Me.lblMsg1 = "Enter NEW Union"
        Case Else
            MsgBox "Invalid Global Type: " & GlobalType
            GoBack
    End Select

    Me.KeyPreview = True

End Sub
