VERSION 5.00
Begin VB.Form frmCheckDetail 
   Caption         =   "Check Detail Report"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4335
   ScaleWidth      =   8145
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1530
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
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
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4845
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1485
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmCheckDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub Form_Load()
    Me.lblCompanyName.Caption = PRCompany.Name
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
            txtDisplay = "Check Date Range: " & Format(Startdate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(Startdate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    
    PRBatchID = BatchNumbr

    Me.Refresh

End Sub

Private Sub cmdOK_Click()
    
    If Startdate = 0 And EndDate = 0 And BatchNumbr = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbCritical, "Check Detail Report"
    Else
        InitFlag = True
        txtDisplay = ""

        CheckDetail RangeType, BatchNumbr, CLng(PEDate), CLng(Startdate), CLng(EndDate), OptDate
    End If

End Sub
