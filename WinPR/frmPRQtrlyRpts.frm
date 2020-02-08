VERSION 5.00
Begin VB.Form frmPRQtrlyRpts 
   Caption         =   "Payroll Quarterly Reports"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   4
      Text            =   "cmbYear"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cmbQtr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   3
      Text            =   "cmbQtr"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk2 
      Caption         =   "&Ok"
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
      Left            =   1155
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmbReportList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   0
      Text            =   "cmbReportList"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2355
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblQtr 
      Caption         =   "Quarter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2115
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblRptSel 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1155
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblCoName2 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   315
      TabIndex        =   5
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmPRQtrlyRpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private Sub cmdExit_Click()
   GoBack
End Sub

Private Sub cmdOk2_Click()
   
    If cmbQtr = 1 Then
      startmonth = 1
      Quarter = 1
      EndMonth = 3
      QtrEnding = "Quarter Ending: 03/31/" & cmbYear

    ElseIf cmbQtr = 2 Then
      startmonth = 4
      Quarter = 4
      EndMonth = 6
      QtrEnding = "Quarter Ending: 06/30/" & cmbYear
    ElseIf cmbQtr = 3 Then
      startmonth = 7
      Quarter = 7
      EndMonth = 9
      QtrEnding = "Quarter Ending: 09/30/" & cmbYear
    ElseIf cmbQtr = 4 Then
      startmonth = 10
      Quarter = 10
      EndMonth = 12
      QtrEnding = "Quarter Ending: 12/31/" & cmbYear
    End If
   
   ' ReportList ("NumberName")   ' <==== based on user selection of report
   If cmbReportList.ListIndex = 0 Then
      QtrRpts ("QtrlyFICAFWT")
   ElseIf cmbReportList.ListIndex = 1 Then
      QtrRpts ("QtrlyStateCity")
   ElseIf cmbReportList.ListIndex = 2 Then
      QtrRpts ("QtrlyFedUnemp")
   ElseIf cmbReportList.ListIndex = 3 Then
      QtrRpts ("QtrlyTipsTaxes")
   Else
      MsgBox "Report was not selected !!!", vbCritical, "Employee Lists and Labels"
   End If
   
   If cmbReportList.ListIndex = 0 Then
      PrintTotals ("QtrlyFICAFWT")
      PrintDeptTotals ("QtrlyFICAFWT")
   ElseIf cmbReportList.ListIndex = 1 Then
      PrintTotals ("QtrlyStateCity")
      PrintDeptTotals ("QtrlyStateCity")
   ElseIf cmbReportList.ListIndex = 2 Then
      PrintTotals ("QtrlyFedUnemp")
      PrintDeptTotals ("QtrlyFedUnemp")
   ElseIf cmbReportList.ListIndex = 3 Then
      PrintTotals ("QtrlyTipsTaxes")
      PrintDeptTotals ("QtrlyTipsTaxes")
   Else
      MsgBox "Report was not selected !!!", vbCritical, "Employee Lists and Labels"
   End If

End Sub

Private Sub Form_Load()
    '  Fill Report Type Selections
    lblCoName2 = Trim(PRCompany.Name)
    cmbReportList.AddItem "Payroll Quarterly FICA And FWT"
    cmbReportList.AddItem "Payroll Quarterly State and City Report"
    cmbReportList.AddItem "Payroll Quarterly Unemployment Report"
    cmbReportList.AddItem "Payroll Quarterly Tips and Taxes Report"

    cmbReportList.ListIndex = 0             '  SET DEFAULT TO FIRST REPORT  !!!!!!!!!!
    
    cmbQtr.AddItem "1"
    cmbQtr.AddItem "2"
    cmbQtr.AddItem "3"
    cmbQtr.AddItem "4"
    
    cmbQtr.ListIndex = 0
    
    cmbYear.AddItem "2008"
    cmbYear.AddItem "2007"
    cmbYear.AddItem "2006"
    cmbYear.AddItem "2005"
    cmbYear.AddItem "2004"
    cmbYear.AddItem "2003"
    cmbYear.AddItem "2002"
    cmbYear.AddItem "2001"
    cmbYear.AddItem "1999"
    cmbYear.AddItem "1998"
    
    cmbYear.ListIndex = 0
    qyear = cmbYear
End Sub

