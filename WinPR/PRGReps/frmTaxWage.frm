VERSION 5.00
Begin VB.Form frmTaxWage 
   Caption         =   "Taxable Wage Report"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6675
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   1155
      TabIndex        =   10
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   7155
      TabIndex        =   9
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sort and Subtotal By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   1095
      TabIndex        =   2
      Top             =   2040
      Width           =   8295
      Begin VB.CheckBox chkDeptTotals 
         Caption         =   "Include Department Totals?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CommandButton cmdselect 
         Caption         =   "&Selection List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optChkDate 
         Caption         =   "Check Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optEmpNo 
         Caption         =   "Employee Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Employees To Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1245
         Width           =   2295
      End
      Begin VB.Label lblCount 
         Caption         =   "count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   6
         Top             =   1245
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "&Date Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1890
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3060
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   5775
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
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmTaxWage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    frmTaxWage.lblCount = "All Employees Selected"
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
'        Me.lblCount = frmEmpSelect.rsEmp.RecordCount
        Me.txtDisplay.Text = "Batch #: " & PRBatch.BatchID & _
                             " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy") & _
                             " Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
        RangeType = PREquate.RangeTypeBatch
        BatchNumbr = PRBatchID
    End If
    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub


Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdDateRange_Click()
   frmDateRange.lblProgram = "TAXABLE WAGE REPORT"
    frmDateRange.Show vbModal
    
    If frmDateRange.optCheckDate = True Then
        OptDate = "CHECK DATE"
    ElseIf frmDateRange.optPEDate = True Then
        OptDate = "P/E DATE"
    End If
    
    If BatchNumbr > 0 Then
        txtDisplay = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        PEDate = PRBatch.PEDate
        CheckDate = PRBatch.CheckDate
        OptDate = " "
    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    
    PRBatchID = BatchNumbr

    Me.Refresh
End Sub


Private Sub cmdOK_Click()
    If frmEmpSelect.SelCount = 0 Then
        lblCount.Caption = "ALL Employees Selected"
        frmEmpSelect.AllEmployees = True
    Else
        lblCount.Caption = frmEmpSelect.rsEmp.RecordCount & " Employees Selected"
        frmEmpSelect.AllEmployees = False
    End If
    
    If StartDate = 0 And EndDate = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbCritical, "Taxable Wage Report"
    Else
        InitFlag = True
        txtDisplay = ""

        TaxableWageRpt RangeType, BatchNumbr, CLng(PEDate), CLng(StartDate), CLng(EndDate), OptDate
    End If
End Sub


Private Sub cmdselect_Click()
    frmEmpSelect.Show vbModal
End Sub

