VERSION 5.00
Begin VB.Form frmInvGlobalQB 
   Caption         =   "0"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvGlobalQB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQBAccts 
      Caption         =   "&QB ACCTS"
      Height          =   615
      Left            =   7860
      TabIndex        =   12
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox chkSalesTax 
      Caption         =   "Company Charges Sales Tax?"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
   Begin VB.ComboBox cmbMiscItem 
      Height          =   375
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3000
      Width           =   4215
   End
   Begin VB.ComboBox cmbFreight 
      Height          =   375
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
   End
   Begin VB.ComboBox cmbTemplate 
      Height          =   375
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.ComboBox cmbAR 
      Height          =   375
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   2340
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
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
      Left            =   5100
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Misc Invoice Item:"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Freight Account:"
      Height          =   255
      Left            =   2700
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice Template:"
      Height          =   255
      Left            =   2580
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "A/R Account:"
      Height          =   255
      Left            =   3060
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   11055
   End
End
Attribute VB_Name = "frmInvGlobalQB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeQBSetup
    If InvGlobal.GetBySQL(SQLString) = False Then
        InvGlobal.Clear
        InvGlobal.CompanyID = PRCompany.CompanyID
        InvGlobal.TypeCode = InvEquate.GlobalTypeQBSetup
        InvGlobal.Var1 = "0"
        InvGlobal.Var2 = "0"
        InvGlobal.Var3 = "0"
        InvGlobal.Var4 = "0"
        InvGlobal.Var5 = "0"
        InvGlobal.rsAdd
    End If
    
    SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'AccountsReceivable' " & _
                " ORDER BY Name"
    ' 2020
    SQLString = "SELECT * FROM QBAccount WHERE AccountType like '%Receivable%' " & _
                " ORDER BY Name"
    If QBAccount.GetBySQL(SQLString) = True Then
        Do
            With Me.cmbAR
                .AddItem QBAccount.Name
                .ItemData(.NewIndex) = QBAccount.QBAccountID
            End With
            If QBAccount.GetNext = False Then Exit Do
        Loop
    End If
    
    SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'TEMPLATE' " & _
            " ORDER BY NAME"
    If QBAccount.GetBySQL(SQLString) = True Then
        Do
            With Me.cmbTemplate
                .AddItem QBAccount.Name
                .ItemData(.NewIndex) = QBAccount.QBAccountID
            End With
            If QBAccount.GetNext = False Then Exit Do
        Loop
    End If
    
    SQLString = "SELECT * FROM InvStock WHERE Description = 'Freight' AND JobID = 0"
    If InvStock.GetBySQL(SQLString) = True Then
        Do
            With Me.cmbFreight
                .AddItem InvStock.Description
                .ItemData(.NewIndex) = InvStock.StockID
            End With
            If InvStock.GetNext = False Then Exit Do
        Loop
    End If
    
    SQLString = "SELECT * FROM InvStock WHERE JobID = 0 ORDER BY Description"
    If InvStock.GetBySQL(SQLString) = True Then
        Do
            With Me.cmbMiscItem
                .AddItem InvStock.Description
                .ItemData(.NewIndex) = InvStock.StockID
            End With
            If InvStock.GetNext = False Then Exit Do
        Loop
    End If
    
    cmbPoint Me.cmbAR, CLng(NumValue(InvGlobal.Var1))
    cmbPoint Me.cmbTemplate, CLng(NumValue(InvGlobal.Var2))
    cmbPoint Me.cmbFreight, CLng(NumValue(InvGlobal.Var3))
    cmbPoint Me.cmbMiscItem, CLng(NumValue(InvGlobal.Var4))
    If InvGlobal.Var5 = "1" Then
        Me.chkSalesTax = 1
    Else
        Me.chkSalesTax = 0
    End If
    
    Me.KeyPreview = True

End Sub
Private Sub cmdQBAccts_Click()
    frmQBAccts.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub
Private Sub cmdOK_Click()
    If Me.cmbAR.ListIndex = -1 Then Exit Sub
    If Me.cmbTemplate.ListIndex = -1 Then Exit Sub
    If Me.cmbFreight.ListIndex = -1 Then Exit Sub
    If Me.cmbMiscItem.ListIndex = -1 Then Exit Sub
    InvGlobal.Var1 = Me.cmbAR.ItemData(Me.cmbAR.ListIndex)
    InvGlobal.Var2 = Me.cmbTemplate.ItemData(Me.cmbTemplate.ListIndex)
    InvGlobal.Var3 = Me.cmbFreight.ItemData(Me.cmbFreight.ListIndex)
    InvGlobal.Var4 = Me.cmbMiscItem.ItemData(Me.cmbMiscItem.ListIndex)
    If Me.chkSalesTax = 1 Then
        InvGlobal.Var5 = "1"
    Else
        InvGlobal.Var5 = ""
    End If
    InvGlobal.rsPut
    GoBack
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

