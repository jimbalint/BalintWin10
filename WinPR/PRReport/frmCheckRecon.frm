VERSION 5.00
Begin VB.Form frmCheckRecon 
   Caption         =   "Check Reconciliation Report"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6326
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   859
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox TxtDisplay 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   2006
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmCheckRecon.frx":0000
      Top             =   1320
      Width           =   5775
   End
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "&DATE RANGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   859
      TabIndex        =   0
      Top             =   1440
      Width           =   975
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
      Left            =   413
      TabIndex        =   4
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "frmCheckRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
    frmDateRange.lblProgram = "CHECK RECONCILIATION"
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
    If StartDate = 0 And EndDate = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbCritical, "Payroll Check Reconciliation"
    Else
        InitFlag = True
        txtDisplay = ""
        CheckRecon RangeType, BatchNumbr, CLng(Int(PEDate)), CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate
    End If
End Sub


