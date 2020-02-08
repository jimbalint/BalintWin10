VERSION 5.00
Begin VB.Form frmGLUpdate 
   Caption         =   "Payroll to GL Update"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
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
   ScaleHeight     =   8670
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "General Ledger History Period to Update:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   908
      TabIndex        =   12
      Top             =   1920
      Width           =   4335
      Begin VB.ComboBox cmbFiscalPeriod 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   3255
      End
      Begin VB.ComboBox cmbFiscalYear 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Period:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fiscal Year:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDescription 
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
      Left            =   908
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5760
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Records to Update:  "
      Height          =   975
      Left            =   2228
      TabIndex        =   10
      Top             =   6480
      Width           =   2655
      Begin VB.OptionButton optUpdRecent 
         Caption         =   "Recent"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optUpdAll 
         Caption         =   "All"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.ComboBox cmbJournal 
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
      Left            =   908
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4800
      Width           =   3855
   End
   Begin VB.ComboBox cmbPRMonth 
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
      Left            =   908
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4808
      TabIndex        =   6
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   728
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Description:"
      Height          =   255
      Left            =   908
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblJnl 
      Caption         =   "Journal Source to use:"
      Height          =   255
      Left            =   908
      TabIndex        =   9
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Payroll History Period to Update:"
      Height          =   255
      Left            =   908
      TabIndex        =   8
      Top             =   960
      Width           =   3135
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmGLUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim trsYM As New ADODB.Recordset
Dim trsJS As New ADODB.Recordset
Dim rsFY As New ADODB.Recordset
Dim i, j As Integer
Public JS, YM As Long

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    Me.optUpdAll = True
    
    GetPRData
    GetGLData
    
    Me.txtDescription.Text = "PAYROLL " & Me.cmbPRMonth.Text

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub GetPRData()

    ' get a list of history months that exist
    trsYM.CursorLocation = adUseClient
    trsYM.Fields.Append "YearMonth", adDouble
    trsYM.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRHist ORDER BY YearMonth DESC"
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Payroll History Exists!", vbCritical
        End
    End If
    
    Do
        trsYM.Find "YearMonth = " & PRHist.YearMonth
        If trsYM.EOF Then
            trsYM.AddNew
            trsYM!YearMonth = PRHist.YearMonth
            trsYM.Update
        End If
        If Not PRHist.GetNext Then Exit Do
    Loop
    
    trsYM.MoveFirst
    Do
        x = MonthName(trsYM!YearMonth Mod 100) & " " & Int(trsYM!YearMonth / 100)
        Me.cmbPRMonth.AddItem x
        trsYM.MoveNext
    Loop Until trsYM.EOF

    Me.cmbPRMonth.ListIndex = 0

End Sub

Private Sub GetGLData()

    ' journal source list
    trsJS.CursorLocation = adUseClient
    trsJS.Fields.Append "Number", adDouble
    trsJS.Fields.Append "Name", adVarChar, 60, adFldIsNullable
    trsJS.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM GLJournal ORDER BY JournalSource"
    If Not GLJournal.GetBySQL(SQLString) Then
        MsgBox "Journal Sources not defined!", vbCritical
        End
    End If
    
    Do
        Me.cmbJournal.AddItem "# " & GLJournal.JournalSource & " " & Trim(GLJournal.JournalName)
        
        trsJS.AddNew
        trsJS!Number = GLJournal.JournalSource
        trsJS!Name = Trim(GLJournal.JournalName)
        trsJS.Update
        
        If Not GLJournal.GetNext Then Exit Do
    Loop
    Me.cmbJournal.ListIndex = 0

    ' fiscal year list
    SQLString = "SELECT DISTINCT FiscalYear FROM GLAmount ORDER BY FiscalYear DESC"
    rsInit SQLString, cn, rsFY
    If rsFY.RecordCount = 0 Then
        i = Year(Now())
        Me.cmbFiscalYear.AddItem i + 1
        Me.cmbFiscalYear.AddItem i
        Me.cmbFiscalYear.AddItem i - 1
        Me.cmbFiscalYear.ListIndex = 1
    Else
        rsFY.MoveFirst
        Do
            Me.cmbFiscalYear.AddItem rsFY!FiscalYear
            rsFY.MoveNext
        Loop Until rsFY.EOF
        Me.cmbFiscalYear.ListIndex = 0
    End If
    rsFY.Close
    
    GetGLPeriods

End Sub

Private Sub GetGLPeriods()

Dim d2 As Date
Dim YM As Long
Dim SelGLPeriod As Long

    If Me.cmbFiscalYear.Text = "" Then Exit Sub

    ' store the YM selected
    trsYM.MoveFirst
    j = Me.cmbPRMonth.ListIndex
    For i = 1 To j
        trsYM.MoveNext
    Next i
    YM = trsYM!YearMonth

    Me.cmbFiscalPeriod.Clear
    j = CInt(Me.cmbFiscalYear.Text)
    If GLCompany.FirstPeriod = 0 Then GLCompany.FirstPeriod = 1
    If GLCompany.FirstPeriod <> 1 Then
        j = j - 1
    End If
    
    SelGLPeriod = -1
    For i = 1 To 12
        d2 = DateSerial(j, i + GLCompany.FirstPeriod - 1, 1)
        If Year(d2) * 100 + Month(d2) = YM Then
            SelGLPeriod = i - 1
        End If
        Me.cmbFiscalPeriod.AddItem "Pd. # " & i & " " & Format(d2, "mmmm-yyyy")
    Next i
    
    If SelGLPeriod >= 0 Then
        Me.cmbFiscalPeriod.ListIndex = SelGLPeriod
    Else
        Me.cmbFiscalPeriod.ListIndex = 0
    End If
    
    Me.txtDescription.Text = "PAYROLL " & Me.cmbPRMonth.Text

End Sub


Private Sub cmdOK_Click()

    ' get the temprecord for what was selected
    
    trsJS.MoveFirst
    j = Me.cmbJournal.ListIndex
    For i = 1 To j
        trsJS.MoveNext
    Next i
    JS = trsJS!Number
    
    trsYM.MoveFirst
    j = Me.cmbPRMonth.ListIndex
    For i = 1 To j
        trsYM.MoveNext
    Next i
    YM = trsYM!YearMonth

    GLUpdate

    GoBack

End Sub

Private Sub cmbPRMonth_Click()
    GetGLPeriods
End Sub

