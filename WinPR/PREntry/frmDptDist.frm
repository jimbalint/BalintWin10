VERSION 5.00
Begin VB.Form frmDptDist 
   Caption         =   "Department Distribution Report"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
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
   ScaleHeight     =   6330
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdselect 
      Caption         =   "&Select List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3135
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CheckBox chkFormFeed 
      Caption         =   "Page Feed on SubTotals"
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
      Left            =   2760
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
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
      Left            =   5055
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      HideSelection   =   0   'False
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   5775
   End
   Begin VB.Frame fraCityTaxSel 
      Caption         =   "Order and Totals By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3082
      TabIndex        =   6
      Top             =   3000
      Width           =   2520
      Begin VB.OptionButton optByEmployee 
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   260
         Width           =   1695
      End
      Begin VB.OptionButton optByDpt 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Label lblcount 
      Caption         =   "Label2"
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
      Left            =   4695
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Selection"
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
      Left            =   735
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "COMPANY NAME"
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
      Height          =   255
      Left            =   135
      TabIndex        =   7
      Top             =   360
      Width           =   8400
   End
End
Attribute VB_Name = "frmDptDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flg As Boolean

Private Sub Form_Load()
    
    ' BatchID assigned? - use it
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
        Me.txtDisplay.Text = "Batch #: " & PRBatch.BatchID & _
                             " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy") & _
                             " Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
        RangeType = PREquate.RangeTypeBatch
        BatchNumbr = PRBatchID
    End If
    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True
    lblcount.Caption = "   All Employees"
End Sub
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdDateRange_Click()

    frmDateRange.lblProgram = "DEPARTMENT DISTRIBUTION REPORT"
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

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
    If StartDate = 0 And EndDate = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbCritical, "Department Distribution Report"
    Else
        InitFlag = True
        DptDistRpt RangeType, BatchNumbr, CLng(Int(PEDate)), CLng(Int(CheckDt)), _
                   CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate
    End If
End Sub
Private Sub cmdselect_Click()
    frmEmpSelect.Show vbModal
    If frmEmpSelect.AllEmployees = False Then
        lblcount.Caption = frmEmpSelect.SelCount & " Employees Selected"
    End If
End Sub
