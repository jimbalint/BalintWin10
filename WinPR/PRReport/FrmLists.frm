VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmLists 
   Caption         =   "Windows PR Lists and Labels"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   765
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Options:"
      Height          =   3375
      Left            =   3840
      TabIndex        =   11
      Top             =   1800
      Width           =   4215
      Begin VB.Frame fraOrder 
         Caption         =   " Order By: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   3855
         Begin VB.OptionButton optDept 
            Caption         =   "Dept ID / Name"
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optNumber 
            Caption         =   " &Number"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optName 
            Caption         =   " &Emp Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton optZipCode 
            Caption         =   " &Zip Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   500
         Left            =   120
         TabIndex        =   20
         Top             =   955
         Width           =   3855
         Begin VB.OptionButton optHrly 
            Caption         =   "&Hourly"
            Height          =   255
            Left            =   2640
            TabIndex        =   23
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optSal 
            Caption         =   "&Salaried"
            Height          =   255
            Left            =   1200
            TabIndex        =   22
            Top             =   180
            Width           =   1150
         End
         Begin VB.OptionButton optAllS 
            Caption         =   "A&ll"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Height          =   500
         Left            =   120
         TabIndex        =   16
         Top             =   230
         Width           =   3855
         Begin VB.OptionButton optInactive 
            Caption         =   "&Inactive"
            Height          =   255
            Left            =   2640
            TabIndex        =   19
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Acti&ve"
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optAllA 
            Caption         =   "&All"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.CheckBox chkSSN 
         Caption         =   " &Display Soc Sec Number?"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check All"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame fraLabels 
      Caption         =   "Label Selections:"
      Height          =   2895
      Left            =   8280
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ComboBox cmbNumLabels 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton optPin 
         Caption         =   "&Pin Feed"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   670
         Width           =   1215
      End
      Begin VB.OptionButton optSheet 
         Caption         =   "S&heet"
         Height          =   240
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin TDBDate6Ctl.TDBDate tdbPEDate 
         Height          =   615
         Left            =   300
         TabIndex        =   15
         Top             =   2040
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   1085
         Calendar        =   "FrmLists.frx":0000
         Caption         =   "FrmLists.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmLists.frx":017E
         Keys            =   "FrmLists.frx":019C
         Spin            =   "FrmLists.frx":01FA
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "02/10/2009"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39854
         CenturyMode     =   0
      End
      Begin VB.Label lblNumLabels 
         Caption         =   "No. of Columns"
         Height          =   255
         Left            =   280
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdUncheck 
      Caption         =   "&Uncheck All"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbReportType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   6015
   End
   Begin VB.ListBox lstDeptSelect 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4620
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   2
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Report Selection:"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblCoName 
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
      Height          =   375
      Left            =   1703
      TabIndex        =   3
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Single
    
Private Sub Form_Load()
        
    n = 0
    Me.lstDeptSelect.AddItem "Unassigned"
    Me.lstDeptSelect.ItemData(lstDeptSelect.NewIndex) = 0
    lstDeptSelect.Selected(n) = True
    n = -1
    If PRDepartment.GetBySQL("SELECT * FROM PRDepartment ORDER BY DepartmentNumber") Then
        Do
            Me.lstDeptSelect.AddItem PRDepartment.Name
            Me.lstDeptSelect.ItemData(lstDeptSelect.NewIndex) = PRDepartment.DepartmentNumber
            n = n + 1
'            lstDeptSelect.Selected(n) = True
            If Not PRDepartment.GetNext Then Exit Do
        Loop
    End If
    cmdcheck_Click
    lstDeptSelect.ListIndex = 0
    lblCoName = Trim(PRCompany.Name)
    '  Fill Report Type Selections
    cmbReportType.AddItem "Number and Name"
    cmbReportType.AddItem "Detail List"
    cmbReportType.AddItem "Employee Rate List"
    cmbReportType.AddItem "SSN Format"
    cmbReportType.AddItem "Rate Tax Listing"
    cmbReportType.AddItem "Time Card Labels"
    cmbReportType.AddItem "Mailing Labels"
    cmbReportType.ListIndex = 0             '  SET DEFAULT TO FIRST REPORT  !!!!!!!!!!
    '  Fill NumLabels (Number of label columns) Selections
    cmbNumLabels.AddItem "1"
    cmbNumLabels.AddItem "2"
    cmbNumLabels.AddItem "3"
    cmbNumLabels.Text = cmbNumLabels.List(2)
    
    Me.KeyPreview = True
    Me.optAllS = 1
    Me.optAllA = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub


Private Sub chkSalaried_Click()
    SalariedSw = 1
End Sub

Private Sub cmbReportType_Click()
    If cmbReportType.ListIndex > 4 Then      ' All Labels
        fraLabels.Visible = True
        cmbNumLabels.Visible = True
        lblNumLabels.Visible = True
    Else
        fraLabels.Visible = False
        cmbNumLabels.Visible = False
        lblNumLabels.Visible = False
        tdbPEDate.Visible = False
    End If

    If cmbReportType.ListIndex = 5 Then
        tdbPEDate.Visible = True
    End If
 
    If cmbReportType.ListIndex = 1 Then    ' Detail List
        chkSSN.Visible = True
    Else
        chkSSN.Visible = False
    End If
   
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
Dim ndx As Integer
       
    NoLabels = cmbNumLabels.ListIndex
    ' EEList ("NumberName")   ' <==== based on user selection of report
   
    ' **** department filter
    StoreDepts
   
    If cmbReportType.ListIndex = 0 Then
        EEList ("NumberName")
    ElseIf cmbReportType.ListIndex = 1 Then
        PortLand = "land"
        EEList ("DetailList")
    ElseIf cmbReportType.ListIndex = 2 Then
        EEList ("EmployeeRateList")
    ElseIf cmbReportType.ListIndex = 3 Then
        EEList ("SSNFormat")
    ElseIf cmbReportType.ListIndex = 4 Then
        EEList ("RateTaxList")
    ElseIf cmbReportType.ListIndex = 5 Then
        EEList ("TimeCardLabels")
    ElseIf cmbReportType.ListIndex = 6 Then
        EEList ("MailingLabels")
    Else
        MsgBox "Report was not selected !!!", vbCritical, "Employee Lists and Labels"
    End If
    
'    ' **** inactive filter
'    If Not Me.chkInactive And PREmployee.Inactive Then
'        Msg1 = "Includes Only Active Employees"
'    Else
'        Msg1 = "Includes Inactive Employees"
'    End If
'
'    ' **** salaried filter
'    If Not Me.chkSalaried And PREmployee.Salaried Then
'        Msg1 = "Salaried Employees Omitted"
'    Else
'        Msg1 = "Includes Salaried Employees"
'    End If

End Sub


Private Sub cmdUncheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstDeptSelect.ListCount - 1
        lstDeptSelect.Selected(ndx) = False
    Next ndx
    Me.lstDeptSelect.ItemData(lstDeptSelect.NewIndex) = 0
End Sub

Private Sub cmdcheck_Click()
        
    Dim ndx As Integer
    For ndx = 0 To lstDeptSelect.ListCount - 1
        lstDeptSelect.Selected(ndx) = True
    Next ndx
    lstDeptSelect.ListIndex = 0
End Sub

Private Sub SelectAllDepts()
    Dim ndx As Integer
    For ndx = 0 To lstDeptSelect.ListCount - 1
        lstDeptSelect.Selected(ndx) = True
    Next ndx
    
End Sub
Private Sub PEDate_LostFocus()
    NewPEDate = PEDate
End Sub
Public Sub StoreDepts()
Dim OnlyNo As String
Dim Dept As String
Dim ndx As Integer
        
    Dpts.CursorLocation = adUseClient
    Dpts.Fields.Append "Dept", adVarChar, 3, adFldIsNullable
    Dpts.Open , , adOpenDynamic, adLockOptimistic
    
    For ndx = 0 To lstDeptSelect.ListCount - 1
        If lstDeptSelect.Selected(ndx) = True Then
            If Trim(lstDeptSelect.List(ndx)) <> "Unassigned" Then
                Dpts.AddNew
                OnlyNo = Mid(lstDeptSelect.List(ndx), 6, 2)
                Dpts.Fields("Dept") = OnlyNo
                Dpts.Update
            End If
        End If
    Next ndx
End Sub

