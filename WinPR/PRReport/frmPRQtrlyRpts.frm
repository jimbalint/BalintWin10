VERSION 5.00
Begin VB.Form frmPRQtrlyRpts 
   Caption         =   "Payroll Quarterly Reports"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   765
      TabIndex        =   3
      Top             =   1080
      Width           =   6705
      Begin VB.CheckBox chkBold 
         Caption         =   "Bold Print"
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
         Left            =   3000
         TabIndex        =   12
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Payroll Quarterly Tips and Taxes Report"
         Enabled         =   0   'False
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
         TabIndex        =   11
         Top             =   1920
         Width           =   6015
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Payroll Quarterly Unemployment Report"
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
         TabIndex        =   10
         Top             =   1440
         Width           =   6015
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Payroll Quarterly State and City Report"
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
         TabIndex        =   9
         Top             =   960
         Width           =   6015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Payroll Quarterly FICA And FWT"
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
         TabIndex        =   8
         Top             =   480
         Width           =   6015
      End
      Begin VB.ComboBox cmbQtr 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblQtr 
         Caption         =   "Quarter"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblYear 
         Caption         =   "Year"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   3240
         Width           =   615
      End
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
      Left            =   5400
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   5760
      Width           =   1695
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
      Height          =   495
      Left            =   330
      TabIndex        =   2
      Top             =   360
      Width           =   7575
   End
End
Attribute VB_Name = "frmPRQtrlyRpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim qYear, CurrYr As Long
Dim Quarter, StartMonth, EndMonth As Byte
Dim QtrEnding As String
Dim rsState As New ADODB.Recordset
Dim StateCount, StartYM, EndYM As Long

Private Sub Form_Load()
    Me.lblCompanyName = PRCompany.Name
    CurrYr = Year(Now())
        
    ' init the yr and qtr combo
    If cmbYrQtrSet(Me.cmbYear, Me.cmbQtr) = False Then GoBack
    qYear = cmbYear
    
    Me.Check1 = 1
    Me.Check2 = 1
    Me.Check3 = 1
    
    Me.KeyPreview = True
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub
    
Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
   
    If cmbQtr = 1 Then
      StartMonth = 1
      Quarter = 1
      EndMonth = 3
      QtrEnding = "QUARTER ENDING: 03/31/" & cmbYear
    ElseIf cmbQtr = 2 Then
      StartMonth = 4
      Quarter = 4
      EndMonth = 6
      QtrEnding = "QUARTER ENDING: 06/30/" & cmbYear
    ElseIf cmbQtr = 3 Then
      StartMonth = 7
      Quarter = 7
      EndMonth = 9
      QtrEnding = "QUARTER ENDING: 09/30/" & cmbYear
    ElseIf cmbQtr = 4 Then
      StartMonth = 10
      Quarter = 10
      EndMonth = 12
      QtrEnding = "QUARTER ENDING: 12/31/" & cmbYear
    End If
   
    PrtInit ("Port")
            
    If Me.chkBold Then
        Prvw.vsp.Font.Bold = True
    End If
    
    ' get the number of different states in PRHist
    If Check2 = 1 Or Check3 = 1 Then
        rsState.CursorLocation = adUseClient
        rsState.Fields.Append "StateID", adDouble
        rsState.Open , , adOpenDynamic, adLockOptimistic
        
        StartYM = Me.cmbYear.Text * 100 + StartMonth
        EndYM = Me.cmbYear.Text * 100 + EndMonth
        
        SQLString = "SELECT * FROM PRHist WHERE PRHist.YearMonth >= " & StartYM & _
                  " AND PRHist.YearMonth <= " & EndYM & _
                  " ORDER BY PRHist.StateID"
        If PRHist.GetBySQL(SQLString) Then
            Do
                SQLString = "StateID = " & PRHist.StateID
                rsState.Find SQLString, 0, adSearchForward, 1
                If rsState.EOF Then
                    rsState.AddNew
                    rsState!StateID = PRHist.StateID
                    rsState.Update
                End If
                If Not PRHist.GetNext Then Exit Do
            Loop
        End If
    End If
            
    ' ReportList ("NumberName")   ' <==== based on user selection of report
    If Check1 = 0 And Check2 = 0 And Check3 = 0 And Check4 = 0 Then
        MsgBox "A Report was not selected !!!", vbCritical, "Payroll Quarterly Reports"
        Exit Sub
    End If
    
    ' *** Federal Report ***
    If Check1 = 1 Then
        QtrRpts "QtrlyFICAFWT", QtrEnding, 0
        If Check2 Or Check3 Or Check4 Then FormFeed
    End If
    
    ' *** State and City report ***
    If Check2 = 1 Then
        ' run for each state
        rsState.MoveFirst
        Do
            QtrRpts "QtrlyStateCity", QtrEnding, rsState!StateID
            rsState.MoveNext
            If rsState.EOF = False Then FormFeed
        Loop Until rsState.EOF
        If Check3 Or Check4 Then FormFeed
    End If
    
    ' *** Unemployment report ***
    If Check3 = 1 Then
        ' only one state to report
        If rsState.RecordCount = 1 Then
            rsState.MoveFirst
            QtrRpts "QtrlyFedUnemp", QtrEnding, rsState!StateID
        Else
            ' run for the entire company first
            QtrRpts "QtrlyFedUnemp", QtrEnding, 0
            FormFeed
            rsState.MoveFirst
            Do
                QtrRpts "QtrlyFedUnemp", QtrEnding, rsState!StateID
                rsState.MoveNext
                If rsState.EOF = False Then FormFeed
            Loop Until rsState.EOF
        End If
        If Check4 Then FormFeed
    End If
    
    ' *** Tips and Taxes report ***
    If Check4 = 1 Then
        QtrRpts "QtrlyTipsTaxes", QtrEnding, 0
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    GoBack
End Sub
