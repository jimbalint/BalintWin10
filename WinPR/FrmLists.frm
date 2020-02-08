VERSION 5.00
Begin VB.Form frmLists 
   Caption         =   "Windows PR Lists and Labels"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10890
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
   Picture         =   "FrmLists.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check All Departments"
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox PEDate 
      Height          =   375
      Left            =   8400
      TabIndex        =   20
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame fraLabels 
      Caption         =   "Label Selection:"
      Height          =   735
      Left            =   7920
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton optPin 
         Caption         =   "&Pin Feed"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   350
         Width           =   1215
      End
      Begin VB.OptionButton optSheet 
         Caption         =   "S&heet"
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   350
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdUncheck 
      Caption         =   "Uncheck All Departments"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   5640
      Width           =   1695
   End
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
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
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
      ItemData        =   "FrmLists.frx":030A
      Left            =   4080
      List            =   "FrmLists.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CheckBox chkSSN 
      Caption         =   "&Display SS Number?"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkSalaried 
      Caption         =   "Include &Salaried?"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CheckBox chkInactive 
      Caption         =   "Include &Inactive?"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Frame fraOrder 
      Caption         =   " Order By: "
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   3960
      Width           =   3735
      Begin VB.OptionButton optZipCode 
         Caption         =   " &Zip Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optName 
         Caption         =   " N&ame"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optNumber 
         Caption         =   " &Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
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
      ItemData        =   "FrmLists.frx":030E
      Left            =   240
      List            =   "FrmLists.frx":0310
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   8040
      Picture         =   "FrmLists.frx":0312
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   975
      Left            =   4560
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label PEText 
      Caption         =   "Period Ending Date:"
      Height          =   255
      Left            =   8070
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Report Selection:"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblNumLabels 
      Caption         =   "Number of Label Columns"
      Height          =   255
      Left            =   7830
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblCoName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   3570
      TabIndex        =   7
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Single


Private Sub chkInactive_Click()
   FilterSw = 1
   InactiveSw = 1
End Sub

Private Sub chkSalaried_Click()
   FilterSw = 1
   SalariedSw = 1
End Sub

Private Sub cmbReportType_Click()
 If cmbReportType.ListIndex > 3 Then      ' All Labels
   fraLabels.Visible = True
   cmbNumLabels.Visible = True
   lblNumLabels.Visible = True
 Else
   fraLabels.Visible = False
   cmbNumLabels.Visible = False
   lblNumLabels.Visible = False
   PEDate.Visible = False
   PEText.Visible = False
 End If

 If cmbReportType.ListIndex = 4 Then
      PEDate.Visible = True
      PEText.Visible = True
 End If
 
 If cmbReportType.ListIndex = 1 Then    ' Detail List
   chkSSN.Visible = True
 Else
   chkSSN.Visible = False
 End If
   
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()
   NoLabels = cmbNumLabels.ListIndex
   ' EEList ("NumberName")   ' <==== based on user selection of report
   If cmbReportType.ListIndex = 0 Then
      EEList ("NumberName")
   ElseIf cmbReportType.ListIndex = 1 Then
      EEList ("DetailList")
   ElseIf cmbReportType.ListIndex = 2 Then
      EEList ("RateList")
   ElseIf cmbReportType.ListIndex = 3 Then
      EEList ("SSNFormat")
   ElseIf cmbReportType.ListIndex = 4 Then
      EEList ("TimeCardLabels")
   ElseIf cmbReportType.ListIndex = 5 Then
      EEList ("MailingLabels")
   Else
      MsgBox "Report was not selected !!!", vbCritical, "Employee Lists and Labels"
   End If

     
End Sub


Private Sub cmdUncheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstDeptSelect.ListCount - 1
        lstDeptSelect.Selected(ndx) = False
    Next ndx
End Sub

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
    SelectAllDepts
    lstDeptSelect.ListIndex = 0
    lblCoName = Trim(PRCompany.Name)
    '  Fill Report Type Selections
    cmbReportType.AddItem "Number and Name"
    cmbReportType.AddItem "Detail List"
    cmbReportType.AddItem "Rate List"
    cmbReportType.AddItem "SSN Format"
    cmbReportType.AddItem "Time Card Labels"
    cmbReportType.AddItem "Mailing Labels"
    cmbReportType.ListIndex = 0             '  SET DEFAULT TO FIRST REPORT  !!!!!!!!!!
    '  Fill NumLabels (Number of label columns) Selections
    cmbNumLabels.AddItem "1"
    cmbNumLabels.AddItem "2"
    cmbNumLabels.AddItem "3"
    cmbNumLabels.Text = cmbNumLabels.List(2)

End Sub


Private Sub cmdcheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstDeptSelect.ListCount - 1
        lstDeptSelect.Selected(ndx) = True
    Next ndx
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
