VERSION 5.00
Begin VB.Form frmPRPurge 
   Caption         =   "PR History Purge"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3150
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   5580
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   2460
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   390
      Left            =   4553
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Select Year to Purge:"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmPRPurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Yr, LastYear As Long
Dim x, y, z As String
Dim i, j, k As Long
Dim SQLString As String

Private Sub Form_Load()

    ' form setups
    Me.lblCompanyName = PRCompany.Name

    Dim rs As ADODB.Recordset
    SQLString = " SELECT DISTINCT YEAR(CheckDate) AS PRYear " & _
                " FROM PRHist "
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then
        MsgBox "No PR History found!", vbInformation
        GoBack
    End If
    
    Do
        Me.cmbTaxYear.AddItem rs!PRYear
        rs.MoveNext
        If rs.EOF Then Exit Do
    Loop
    
    Me.cmbTaxYear.ListIndex = 0
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

Private Sub cmdOK_Click()
    
    z = Me.cmbTaxYear.Text
    k = Me.cmbTaxYear.Text
    
    ' >>>>
    Dim resp As Integer
    x = "******************************************" & vbCr & _
        "* Are you SURE you want to delete" & vbCr & _
        "* ALL Payroll History for: " & z & vbCr & _
        "* This data will be permanently deleted" & vbCr & _
        "* and can NOT be retrieved !!!" & vbCr & _
        "******************************************"
    resp = MsgBox(x, vbOKCancel + vbCritical, "Purge PR History")
    If resp = vbCancel Then
        GoBack
    End If
    
    frmProgress.Show
    frmProgress.lblMsg1.Caption = "Now purging ALL payroll data for: " & z
    frmProgress.lblMsg3.Caption = ""
    
    Dim dte1, dte2 As Date
    dte1 = DateSerial(k, 1, 1)
    dte2 = DateSerial(k, 12, 31)
    Dim ym1, ym2 As Long
    ym1 = k * 100 + 1
    ym2 = k * 100 + 12
    
    Dim qry(8), tbl(8) As String
    
    qry(1) = " DELETE * FROM PRAdjust " & _
             " WHERE AdjDate between '" & z & "-01-01' " & _
             " AND '" & z & "-12-31'"
    tbl(1) = "Adjustments"
    
    qry(2) = " DELETE * FROM PRBatch " & _
             " WHERE YearMonth between " & _
             ym1 & " AND " & ym2
    tbl(2) = "Batch"
    
    qry(3) = " DELETE * FROM PRDist " & _
             " WHERE YearMonth between " & _
             ym1 & " AND " & ym2
    tbl(3) = "Distribution"
    
    qry(4) = " DELETE * FROM PRItemHist " & _
             " WHERE YearMonth between " & _
             ym1 & " AND " & ym2
    tbl(4) = "Item History"
    
    qry(5) = " DELETE * FROM PRW2 " & _
             " WHERE TaxYear = " & z
    tbl(5) = "W2 History"
    
    qry(6) = " DELETE * FROM PRW2City " & _
             " WHERE TaxYear = " & z
    tbl(6) = "W2 City History"
    
    qry(7) = " DELETE * FROM PRW2State " & _
             " WHERE TaxYear = " & z
    tbl(7) = "W2 State History"
    
    qry(8) = " DELETE * FROM PRHist " & _
             " WHERE YearMonth between " & _
             ym1 & " AND " & ym2
    tbl(8) = "History"
    
    For j = 1 To 8
        frmProgress.lblMsg2.Caption = "Payroll " & tbl(8)
        frmProgress.Refresh
        cn.Execute qry(j)
    Next j
    
    frmProgress.Hide
    
    MsgBox "All Payroll History for: " & z & vbCr & _
           " has been deleted", vbInformation
    GoBack

End Sub

