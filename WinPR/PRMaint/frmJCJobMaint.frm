VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmJCJobMaint 
   Caption         =   "Job Maintenance"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10500
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE UNSED JOBS"
      Height          =   855
      Left            =   10680
      TabIndex        =   5
      Top             =   9480
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   12855
      _cx             =   22675
      _cy             =   12938
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   855
      Left            =   8160
      TabIndex        =   2
      Top             =   9480
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&REFRESH FROM QB"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Label lblJobName 
      Caption         =   "Job Name"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   12255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Works best with the QuickBooks File open to refresh!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   9480
      Width           =   2655
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   12255
   End
End
Attribute VB_Name = "frmJCJobMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim i, j As Long

Dim ActiveDrop, CityDrop As String, CustDrop As String, StatusDrop As String

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name

    If TableExists("JCCustomer", cn) = False Then CustomerCreate
    If TableExists("JCJob", cn) = False Then JobCreate

    ' init the grid dropdowns
    CityDrop = "|#0;NONE"
    CustDrop = "|#0;NONE"
    ActiveDrop = "|#0;No|#1;Yes"

    StatusDrop = ""
    For i = 0 To 5
        X = ""
        If i = PREquate.qbJobStatus_Awarded Then X = "Awarded"
        If i = PREquate.qbJobStatus_Closed Then X = "Closed"
        If i = PREquate.qbJobStatus_InProgress Then X = "In Progress"
        If i = PREquate.qbJobStatus_None Then X = "None"
        If i = PREquate.qbJobStatus_NotAwarded Then X = "Not Awarded"
        If i = PREquate.qbJobStatus_Pending Then X = "Pending"
        StatusDrop = Trim(StatusDrop) & "|#" & i & ";" & X
    Next i

    SQLString = "SELECT * FROM PRCity ORDER BY CityName"
    If PRCity.GetBySQL(SQLString) Then
        Do
            CityDrop = Trim(CityDrop) & "|#" & PRCity.CityID & ";" & Trim(PRCity.CityName)
            If PRCity.GetNext = False Then Exit Do
        Loop
    End If
    
'    SQLString = "SELECT * FROM JCCustomer ORDER BY FullName"
'    If JCCustomer.GetBySQL(SQLString) Then
'        Do
'            CustDrop = Trim(CustDrop) & "|#" & JCCustomer.CustomerID & ";" & Trim(JCCustomer.Name)
'            If JCCustomer.GetNext = False Then Exit Do
'        Loop
'    End If
 
    SQLString = "SELECT * FROM JCJob ORDER BY FullName"
    rsInit SQLString, cn, rs
    
    SetGrid rs, fg
    
    With Me.fg
        For i = 0 To .Cols - 1
            .ColWidth(i) = 0
            X = .TextMatrix(0, i)
            Select Case X
                Case "FullName"
                    .ColWidth(i) = 5000
                Case "CityID"
                    .ColWidth(i) = 3000
                    .ColComboList(i) = CityDrop
                Case "StartDate"
                    .ColWidth(i) = 1400
                    .ColFormat(i) = "mm/dd/yyyy"
                    .TextMatrix(0, i) = "Date Modified"
                Case "JobStatus"
                    .ColWidth(i) = 1200
                    .ColComboList(i) = StatusDrop
                Case "Active"
                    .ColWidth(i) = 800
                    .ColComboList(i) = ActiveDrop
            End Select
        Next i
    End With
    
'    fg.ColWidth(0) = 0          ' JobID
'    fg.ColWidth(1) = 0          ' Name
'    fg.ColWidth(2) = 5000       ' FullName
'    fg.ColWidth(3) = 0          ' CompanyName
'    fg.ColWidth(4) = 0          ' QBID
'    fg.ColWidth(5) = 0          ' QBParentID
'    fg.ColWidth(6) = 0          ' ParentID - Company Name
'    fg.ColWidth(7) = 3000       ' CityID
'    For i = 8 To 30
'        If i = 25 Then
'        ElseIf i = 26 Then
'        Else
'            fg.ColWidth(i) = 0
'        End If
'    Next i
'
'    fg.ColComboList(25) = StatusDrop
'    fg.ColWidth(25) = 1500       ' job status
'    fg.ColWidth(26) = 1500       ' date modified - JCJob.StartDate
'    fg.ColFormat(26) = "mm/dd/yy"
'
'    ' fg.ColComboList(6) = CustDrop
'    fg.ColComboList(7) = CityDrop
'
'    fg.TextMatrix(0, 1) = "Job #"
'    fg.TextMatrix(0, 2) = "Customer:Job Name"
'    fg.TextMatrix(0, 6) = "Customer Name"
'    fg.TextMatrix(0, 7) = "Assigned City"

    fg.AutoSearch = flexSearchFromTop

    Me.KeyPreview = True

    Me.Show
    fg.SetFocus
    
    If fg.Rows >= 2 Then
        fg.Row = 1
        fg.Col = 2
    End If
    
End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = fg.ColIndex("CityID") Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol <> fg.ColIndex("FullName") Then
        fg.AutoSearch = flexSearchNone
    Else
        fg.AutoSearch = flexSearchFromCursor
    End If
    
    ' display the full job name at the top of the screen
    On Error Resume Next
    If JCJob.GetByID(fg.TextMatrix(NewRow, 0)) = True Then
        Me.lblJobName = JCJob.FullName
    Else
        Me.lblJobName = ""
    End If
    Me.Refresh
    On Error GoTo 0

End Sub
Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdRefresh_Click()
    
'    If MsgBox("OK to overwrite ALL QB Customer and Job Info?", vbQuestion + vbYesNo, "QB Customer/Job Import") = vbNo Then
'        Exit Sub
'    End If
'
'    If TableExists("JCCustomer", cn) = False Then
'        CustomerCreate
'    End If
'
'    If TableExists("JCJob", cn) = False Then
'        JobCreate
'    End If
'
'    DoCustomerQueryRq "US", 5, 0
    
    frmJCGetQBData.Show vbModal
    
    rs.Requery
    fg.DataRefresh
    
    MsgBox "Import of QB Customer and Job Info Complete", vbInformation, "Balint Windows PR"

End Sub

Private Sub cmdDelete_Click()
    If rs.RecordCount = 0 Then Exit Sub
    If MsgBox("OK to remove unused jobs?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    rs.MoveFirst
    Do
        SQLString = "SELECT * FROM PRDist WHERE JobID = " & rs!JobID
        If PRDist.GetBySQL(SQLString) = False Then rs.Delete
        rs.MoveNext
    Loop Until rs.EOF
    rs.Requery
End Sub


