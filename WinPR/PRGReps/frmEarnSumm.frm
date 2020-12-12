VERSION 5.00
Begin VB.Form frmEarnSumm 
   Caption         =   "Earmings Summary Report"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkExcludeDetail 
      Caption         =   "Exclude Pay Period Detail"
      Height          =   375
      Left            =   3420
      TabIndex        =   7
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "SORT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   16
      Top             =   5040
      Width           =   7815
      Begin VB.OptionButton optEmpName 
         Caption         =   "Employee Name"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   2415
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
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Height          =   855
      Left            =   2625
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Width           =   5775
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
      Left            =   1455
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Include:"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   7815
      Begin VB.CheckBox chkPgEmp 
         Caption         =   " Separate Page for Each Employee"
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
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CheckBox chkAddr 
         Caption         =   " Print Employee Address"
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
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.CheckBox chkSSN 
         Caption         =   " Print SS Numbers"
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
         TabIndex        =   1
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdselect 
         Caption         =   "&Selection List"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
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
         TabIndex        =   12
         Top             =   1845
         Width           =   2655
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
         TabIndex        =   11
         Top             =   1845
         Width           =   2295
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
      Left            =   5400
      TabIndex        =   9
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "** NOTE:  Start and End YEAR Must Be the SAME"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   920
      Width           =   5175
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "frmEarnSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BOQDate, EOQDate, BOYDate As Date
Public mo As Long
Public yr As Long
Public Qtr1 As Date
Public Qtr2 As Date
Dim EndMth, Qtr As Byte

Private Sub Form_Load()
    frmEarnSumm.lblCount = "All Employees Selected"
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
'        Me.lblCount = frmEmpSelect.rsEmp.RecordCount
        Me.txtDisplay.text = "Batch #: " & PRBatch.BatchID & _
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
   frmDateRange.lblProgram = "EARNINGS SUMMARY REPORT"
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
        StartDate = CheckDt
        EndDate = CheckDt
        OptDate = " "
    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            MsgBox "YOU MUST USE CHECK DATE OPTION", vbCritical, "Earnings Summary Report"
        End If
    End If
    
    ' calc begin of year / start and end of quarter
    BOYDate = DateSerial(Year(StartDate), 1, 1)
    EndMth = Month(EndDate)
    If EndMth >= 10 Then
        BOQDate = DateSerial(Year(EndDate), 10, 1)
        EOQDate = DateSerial(Year(EndDate), 12, 31)
    ElseIf EndMth >= 7 Then
        BOQDate = DateSerial(Year(EndDate), 7, 1)
        EOQDate = DateSerial(Year(EndDate), 9, 30)
    ElseIf EndMth >= 4 Then
        BOQDate = DateSerial(Year(EndDate), 4, 1)
        EOQDate = DateSerial(Year(EndDate), 6, 30)
    Else
        BOQDate = DateSerial(Year(EndDate), 1, 1)
        EOQDate = DateSerial(Year(EndDate), 3, 31)
    End If
    
    PRBatchID = BatchNumbr
    Me.Refresh
End Sub

Private Sub cmdOK_Click()
Dim SYear, EYear As Long
    If frmEmpSelect.SelCount = 0 Then
        lblCount.Caption = "ALL Employees Selected"
        frmEmpSelect.AllEmployees = True
    Else
        lblCount.Caption = frmEmpSelect.rsEmp.RecordCount & " Employees Selected"
        frmEmpSelect.AllEmployees = False
    End If
       
    SYear = Format(Year(StartDate))
    EYear = Format(Year(EndDate))
  
    If SYear < EYear Then
        MsgBox "Start Year and End Year MUST BE the SAME", vbExclamation, "Earnings Summary Report"
    ElseIf CLng(StartDate) = 0 And CLng(EndDate) = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbExclamation, "Earnings Summary Report"
    ElseIf frmDateRange.OptChkPeDate And frmDateRange.optPEDate Then
        MsgBox "Please Select a BATCH or CHECK DATE RANGE", vbExclamation, "Earnings Summary Report"
    Else
        InitFlag = True
        txtDisplay = ""
        EarnSummary RangeType, BatchNumbr, CLng(Int(PEDate)), CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate
    End If
End Sub

Private Sub cmdselect_Click()
    frmEmpSelect.Show vbModal
    Me.lblCount = frmEmpSelect.SelString
End Sub


