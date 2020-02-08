VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmEntry 
   Caption         =   "Payroll Data Entry"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHireDate 
      Caption         =   "Display Hire Date"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   4320
      Width           =   4095
   End
   Begin VB.CheckBox chkUseGLName 
      Caption         =   "Use GL Company Name"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   6600
      Width           =   4095
   End
   Begin TDBText6Ctl.TDBText tdbtxtHdrComment 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   6375
      _Version        =   65536
      _ExtentX        =   11245
      _ExtentY        =   1296
      Caption         =   "frmEntry.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntry.frx":0072
      Key             =   "frmEntry.frx":0090
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "TDBText1"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CheckBox chkDeptSepPage 
      Caption         =   "Departments on separate pages?"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   6000
      Width           =   4095
   End
   Begin VB.ComboBox cmbSortOrder 
      Height          =   390
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5400
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Print Options:"
      Height          =   1455
      Left            =   1342
      TabIndex        =   14
      Top             =   840
      Width           =   4440
      Begin VB.CheckBox chkOtherEarns 
         Caption         =   "Print Other Earnings?"
         Height          =   300
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CheckBox chkDeds 
         Caption         =   "Print Deductions?"
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Value           =   1  'Checked
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   11
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   10
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Pay Rates For:"
      Height          =   1695
      Left            =   1342
      TabIndex        =   13
      Top             =   2400
      Width           =   4440
      Begin VB.OptionButton optNone 
         Caption         =   "None"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton optHrly 
         Caption         =   "Hourly Only"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   4095
      End
      Begin VB.OptionButton optSalHrly 
         Caption         =   "Salary and Hourly"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Sort Order:"
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   135
      TabIndex        =   12
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.lblCompanyName = PRCompany.Name
    
    ' dflt sort order from PRCompany
    With Me.cmbSortOrder
        .AddItem "By EE Number"
        .AddItem "By EE Name"
        .AddItem "By Dept By EE#"
        .AddItem "By Dept By EE Name"
        .ListIndex = PRCompany.DfltSortOrder
        If .ListIndex <= 1 Then
            Me.chkDeptSepPage.Visible = False
        End If
    End With
    
    ' use GLCompany Name per user
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & User.ID & _
                " AND Description = 'EntryName'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        If PRGlobal.Var1 = "1" Then
            Me.chkUseGLName = 1
        End If
    End If
    
    ' other answers from PRGlobal by CompanyID
    Me.chkOtherEarns = 1
    Me.chkDeds = 1
    Me.optSalHrly = True
    Me.optHrly = False
    Me.optNone = False
    
    tdbTextSet Me.tdbtxtHdrComment
    Me.tdbtxtHdrComment.MaxLength = 40
    
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                " AND Var1 = 'EntryForm'"
    If PRGlobal.GetBySQL(SQLString) Then
        If PRGlobal.Var2 = "0" Then Me.chkOtherEarns = 0
        If PRGlobal.Var3 = "0" Then Me.chkDeds = 0
        If PRGlobal.Var4 = "2" Then
            Me.optSalHrly = False
            Me.optHrly = True
            Me.optNone = False
        ElseIf PRGlobal.Var4 = "3" Then
            Me.optSalHrly = False
            Me.optHrly = False
            Me.optNone = True
        End If
        If PRGlobal.Var5 = "1" Then Me.chkDeptSepPage = 1
        If PRGlobal.Var6 = "1" Then Me.chkHireDate = 1
    End If
    
    Me.KeyPreview = True
End Sub
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdOkay_Click()
    
    ' save the settings for the company
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                " AND Var1 = 'EntryForm'"
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Var1 = "EntryForm"
        PRGlobal.Save (Equate.RecAdd)
    End If
        
    If Me.chkOtherEarns = 1 Then
        PRGlobal.Var2 = "1"
    Else
        PRGlobal.Var2 = "0"
    End If
    
    If Me.chkDeds = 1 Then
        PRGlobal.Var3 = "1"
    Else
        PRGlobal.Var3 = "0"
    End If
        
    If Me.optSalHrly = True Then
        PRGlobal.Var4 = "1"
    ElseIf Me.optHrly = True Then
        PRGlobal.Var4 = "2"
    Else
        PRGlobal.Var4 = "3"
    End If
    
    If Me.chkDeptSepPage = 1 Then
        PRGlobal.Var5 = "1"
    Else
        PRGlobal.Var5 = "0"
    End If
    
    If Me.chkHireDate = 1 Then
        PRGlobal.Var6 = "1"
    Else
        PRGlobal.Var6 = "0"
    End If
    
    PRGlobal.Save (Equate.RecPut)
    
    ' use GLCompany Name per user
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & User.ID & _
                " AND Description = 'EntryName'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        PRGlobal.Var1 = Me.chkUseGLName
        PRGlobal.Save (Equate.RecPut)
    ElseIf Me.chkUseGLName = 1 Then
        PRGlobal.Clear
        PRGlobal.UserID = User.ID
        PRGlobal.Description = "EntryName"
        PRGlobal.Var1 = Me.chkUseGLName
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    EntryForm

End Sub

Private Sub cmbSortOrder_Click()
    If Me.cmbSortOrder.ListIndex <= 1 Then
        Me.chkDeptSepPage.Visible = False
    Else
        Me.chkDeptSepPage.Visible = True
    End If
End Sub


