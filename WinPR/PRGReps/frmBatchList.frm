VERSION 5.00
Begin VB.Form frmBatchList 
   Caption         =   "Batch List"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4680
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraYrRange 
      Caption         =   "Year Range"
      Height          =   1215
      Left            =   1890
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
      Begin VB.ComboBox cmbEndYear 
         Height          =   390
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbStartYear 
         Height          =   390
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "End Year"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Start Year"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   5010
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   1290
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chkAllYears 
      Caption         =   "All Years"
      Height          =   270
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company"
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
      Left            =   810
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmBatchList"
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
        
    ' init the start years
    If cmbYrSet(Me.cmbStartYear) = False Then GoBack
    qYear = cmbStartYear
    
    ' init the end years
    If cmbYrSet(Me.cmbEndYear) = False Then GoBack
    qYear = cmbEndYear
    
    Me.KeyPreview = True

    ' Disable Start and End year selection if 'All Years' is checked
    If chkAllYears = 1 Then
        Me.fraYrRange.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
        cmbStartYear.Enabled = False
        cmbEndYear.Enabled = False
    Else
        Me.fraYrRange.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        cmbStartYear.Enabled = True
        cmbEndYear.Enabled = True
    End If
       
End Sub


Private Sub chkAllYears_Click()

    ' Disable Start and End year selection if 'All Years' is checked
    If chkAllYears = 1 Then
        Me.fraYrRange.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
        cmbStartYear.Enabled = False
        cmbEndYear.Enabled = False
    Else
        Me.fraYrRange.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        cmbStartYear.Enabled = True
        cmbEndYear.Enabled = True
    End If
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
    ' Set StartDate and EndDate vars
    If chkAllYears = 0 Then
        StartDate = DateSerial(cmbStartYear, 1, 1)
        EndDate = DateSerial(cmbEndYear, 12, 31)
    End If
    
    PRBatchList StartDate, EndDate
    
End Sub


Public Function cmbYrSet(ByRef cmbYr As ComboBox) As Boolean
Dim yrs As ADODB.Recordset
Dim i, j, k As Integer

    SQLString = "SELECT DISTINCT YearMonth FROM PRHist ORDER BY YearMonth DESC"
    rsInit SQLString, cn, yrs
    If yrs.RecordCount = 0 Then
        MsgBox "No Payroll History Data Found!!", vbExclamation
        cmbYrSet = False
        Exit Function
    End If
 
    cmbYrSet = True
    
    yrs.MoveFirst
    cmbYr.AddItem Int(yrs!YearMonth / 100)

    Do
        yrs.MoveNext
        If yrs.EOF Then Exit Do
        k = 0
        j = cmbYr.ListCount
        For i = 0 To j - 1
            cmbYr.ListIndex = i
            If cmbYr.Text = Int(yrs!YearMonth / 100) Then
                k = 1
                Exit For
            End If
        Next i
        If k = 0 Then
            cmbYr.AddItem (Int(yrs!YearMonth / 100))
        End If
    Loop
    cmbYr.ListIndex = 0

End Function
