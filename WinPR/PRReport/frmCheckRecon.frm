VERSION 5.00
Begin VB.Form frmCheckRecon 
   Caption         =   "Check Reconciliation Report"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
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
   ScaleHeight     =   6465
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Report Option "
      Height          =   1935
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   7335
      Begin VB.OptionButton optNoName 
         Caption         =   "No Names Export"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   3255
      End
      Begin VB.OptionButton optCheckRecon 
         Caption         =   "Check Reconciliation"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   5640
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
Dim x, y, z As String

Private Sub Form_Load()
    
    ' BatchID assigned? - use it
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
        Me.txtDisplay.text = "Batch #: " & PRBatch.BatchID & _
                             " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy") & _
                             " Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
        RangeType = PREquate.RangeTypeBatch
        BatchNumbr = PRBatchID
    End If
    Me.lblCompanyName = PRCompany.Name
    Me.optCheckRecon = True
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
        If Me.optCheckRecon Then
            CheckRecon RangeType, BatchNumbr, CLng(Int(PEDate)), CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate
        Else
            NoNameExport
        End If
    End If
End Sub

Private Sub NoNameExport()
    SQLString = "SELECT * FROM PRHist"
 
    If RangeType = PREquate.RangeTypeBatch Then
        SQLString = Trim(SQLString) & " WHERE PRHist.BatchID = " & BatchNumbr
        Msg1 = "Batch: " & BatchNumbr
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = Trim(SQLString) & " WHERE PRHist.CheckDate >= " & CLng(StartDate) & " AND " & _
                                    " PRHist.CheckDate <= " & CLng(EndDate)
            Msg1 = "CHECK DATE RANGE: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        ElseIf OptDate = "P/E DATE" Then
             SQLString = Trim(SQLString) & " WHERE PRHist.PEDate >= " & CLng(StartDate) & " AND " & _
                                    " PRHist.PEDate <= " & CLng(EndDate)
            Msg1 = "P/E DATE RANGE: " & CDate(StartDate) & " TO: " & CDate(EndDate)
        End If
    End If

    SQLString = Trim(SQLString) & " ORDER BY PRHist.CheckNumber"

    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History Found !!!", vbExclamation, "Payroll Check Reconciliation"
        GoBack
    End If

    Const WindowsFolder = 0
    Const SystemFolder = 1
    Const TemporaryFolder = 2
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder: tempFolder = fso.GetSpecialFolder(TemporaryFolder)
    Dim TextChannel As Integer
    
    x = "CheckExport" & Right(Year(Date), 2) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
    TextFileName = tempFolder & "\" & x & ".csv"

    TextChannel = FreeFile
    Do
        On Error Resume Next
        Open TextFileName For Output As #TextChannel
        If Err.Number <> 0 Then
            ErrMsg = "Error Opening: " & TextFileName & vbCr & vbCr & _
                " " & Err.Number & " " & Err.Description
            MsgResponse = MsgBox(ErrMsg, vbRetryCancel + vbExclamation, "File Open Error")
            If MsgResponse <> vbRetry Then
                TextChannel = 0
                TextFileName = ""
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    Print #TextChannel, "Check Number, Check Date, Check Amount"
    
    Do
        Print #TextChannel, PRHist.CheckNumber & ", " & PRHist.CheckDate & ", " & PRHist.Net
        If Not PRHist.GetNext Then Exit Do
    Loop

    Close #TextChannel
    TaskID = Shell("cmd /c " & TextFileName, vbNormalFocus)
    GoBack
End Sub
