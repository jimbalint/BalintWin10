VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCheckReg 
   Caption         =   "Check Register"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
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
   ScaleHeight     =   7380
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   22
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CheckBox chkFileOutput 
      Caption         =   "CSV Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   21
      Top             =   5760
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog cdlTextOutput 
      Left            =   240
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   615
      Left            =   3780
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3360
      TabIndex        =   13
      Top             =   4080
      Width           =   2895
      Begin VB.CheckBox chkEESubTotal 
         Caption         =   "Subtotal by Employee"
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2535
      End
      Begin VB.CheckBox chkSepTotPg 
         Caption         =   "New page for totals"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox chkIncInactiveItems 
         Caption         =   "Include Inactive Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chkTotalsOnly 
         Caption         =   "Totals ONLY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Include:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1260
      TabIndex        =   9
      Top             =   2520
      Width           =   3135
      Begin VB.CheckBox chkDed 
         Caption         =   "&Deductions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   920
         Width           =   2655
      End
      Begin VB.CheckBox chkOEAmt 
         Caption         =   "Other Earnings &Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   610
         Width           =   2775
      End
      Begin VB.CheckBox chkOEHrs 
         Caption         =   "Other &Earnings Hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   2535
      End
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
      Left            =   2490
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
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
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5220
      TabIndex        =   1
      Top             =   2520
      Width           =   3135
      Begin VB.OptionButton optCheckNo 
         Caption         =   "&Check Number"
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
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2175
      End
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
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   920
         Width           =   2415
      End
      Begin VB.OptionButton optEmpNo 
         Caption         =   "Employee Number"
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
         Left            =   360
         TabIndex        =   4
         Top             =   610
         Width           =   2415
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
      Height          =   615
      Left            =   6900
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOkay 
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
      Height          =   615
      Left            =   1260
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
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
      Left            =   5415
      TabIndex        =   20
      Top             =   1920
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
      Left            =   1425
      TabIndex        =   19
      Top             =   1965
      Width           =   2295
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "frmCheckReg"
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
    
    Me.chkDed = 1
    Me.chkOEAmt = 1
    Me.chkOEHrs = 1
    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True
    lblcount = "All Employees"
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
    
    frmDateRange.lblProgram = "DATE RANGE"
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
        PRBatchID = BatchNumbr
        CheckDate = PRBatch.CheckDate
        OptDate = " "
    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
        
    End If

    Me.Refresh
End Sub

Private Sub cmdOkay_Click()
    
    TextFileName = ""
    
    If StartDate = 0 And EndDate = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbCritical, "Payroll Check Register"
    Else
        
        ' select the output file name
        If Me.chkFileOutput Then
                        
            cdlTextOutput.CancelError = True
            
            ' set to current
            cdlTextOutput.Flags = cdlCFBoth Or cdlCFEffects
            cdlTextOutput.Filter = "Comma Separated Values|*.csv"
            Me.cdlTextOutput.FileName = User.Logon & ".csv"
            cdlTextOutput.DialogTitle = "Select a file for Text Export"
            cdlTextOutput.CancelError = True
            cdlTextOutput.InitDir = "\Balint\Data"
    
            ' call the file dialog
            On Error Resume Next
            cdlTextOutput.ShowOpen
            
            If Err.Number = 0 Then
    
                ' assign
                TextFileName = cdlTextOutput.FileName
                TextChannel2 = FreeFile
    
                Do
                    
                    On Error Resume Next
                    Open TextFileName For Output As #TextChannel2
                    
                    If Err.Number <> 0 Then
                        
                        ErrMsg = "Error Opening: " & TextFileName & vbCr & vbCr & _
                            " " & Err.Number & " " & Err.Description
                            
                        MsgResponse = MsgBox(ErrMsg, vbRetryCancel + vbExclamation, "File Open Error")
                        If MsgResponse <> vbRetry Then
                            TextChannel2 = 0
                            TextFileName = ""
                            Exit Do
                        End If
                        
                    Else
                        Exit Do
                    End If
                
                Loop
    
            End If
    
            On Error GoTo 0

        End If
        
        Me.Hide
        InitFlag = True
        CheckRegister RangeType, BatchNumbr, CLng(Int(PEDate)), CLng(Int(CheckDt)), CLng(Int(StartDate)), _
                      CLng(Int(EndDate)), OptDate, Me.chkBold
    
    End If

End Sub


Private Sub cmdselect_Click()
    frmEmpSelect.Show vbModal
    If frmEmpSelect.AllEmployees = False Then
        lblcount.Caption = frmEmpSelect.SelCount & " Employees Selected"
    End If
End Sub
