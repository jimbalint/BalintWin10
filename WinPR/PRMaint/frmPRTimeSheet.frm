VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPRTimeSheet 
   Caption         =   "Time Sheet Entry"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14055
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
   ScaleHeight     =   9375
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd40Hours 
      Caption         =   "40 Hrs"
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdSortHours 
      Caption         =   "SORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   23
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdSortName 
      Caption         =   "SORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   22
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdSortNum 
      Caption         =   "SORT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdDelAll 
      Caption         =   "DEL ALL"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   1680
      TabIndex        =   18
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddLine 
      Caption         =   "ADD"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4560
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   13815
      _cx             =   24368
      _cy             =   5953
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddEmp 
      Caption         =   "&ADD EMPLOYEE"
      Height          =   855
      Left            =   12360
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   2655
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   6495
      _cx             =   11456
      _cy             =   4683
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
      Height          =   495
      Left            =   12120
      TabIndex        =   3
      Top             =   8760
      Width           =   1695
   End
   Begin VB.ComboBox cmbWEDate 
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblEEName 
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   26
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label lblJobName 
      Caption         =   "Job Name"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   3960
      Width           =   12975
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Msg1"
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
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   8640
      Width           =   5655
   End
   Begin VB.Label lblTotHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   16
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblSatHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   15
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblFriHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   14
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblThuHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   13
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblWedHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblTueHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   11
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblMonHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblSunHrs 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblTotalHours 
      Caption         =   "Total Hours:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Week Of:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "frmPRTimeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WEDate As Date
Dim TotalHrsWith, TotalHrsWithout As Single
Dim LoadFlag As Boolean
Dim MaxWeeks As Integer
Dim i, j, k As Long
Dim X, Y, z As String
Dim fgRow, fgCol As Long
Dim JC As New ADODB.Recordset
Dim EMP As New ADODB.Recordset
Dim TS As New ADODB.Recordset
Dim HrFmt As String
Dim GString As String
Dim p1 As Currency

Dim CellValue As String

Dim JobDrop, DeptDrop, ItemDrop, CityDrop As String

Dim SortOrder, SortCol As Byte

Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Me.lblJobName = ""
    On Error Resume Next
    If fg.TextMatrix(NewRow, 0) = "" Then Exit Sub
    If NewRow <= 0 Then Exit Sub
    If JCJob.GetByID(fg.TextMatrix(NewRow, 0)) = True Then
        Me.lblJobName = JCJob.FullName
    End If
    On Error GoTo 0
End Sub

Private Sub Form_Load()

    Me.lblMsg1 = ""

    If TableExists("PRTimeSheet", cn) = False Then
        PRTimeSheetCreate
    End If
    
    If TableExists("JCJob", cn) = False Then
        JobCreate
    End If
    
    LoadFlag = True
    
    Me.lblCompanyName = PRCompany.Name

    ' *****
'    If TableExists("PRTimeSheet", cn) = True Then
'        cn.Execute "DROP TABLE PRTimeSheet"
'        PRTimeSheetCreate
'    End If
    ' *****

    HrFmt = "#,##0.00"
    MaxWeeks = 52
    GetWEDates
    PopDrops
    
    PopEEFG
    PopFG

    ' initial display and cursor setting
    Me.Show
    fg.SetFocus
    LoadFlag = False
    
    CalcAll
    CalcTotals 7
    ' AddLine
    
    ' ------
    LoadFlag = False
    
    EMP.MoveFirst
    PopFG
    
    Me.KeyPreview = True

End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    X = Trim(fg.TextMatrix(Row, 3))

    ' Me.lblMsg1 = ""

    On Error Resume Next
    p1 = CCur(fg.TextMatrix(Row, 3))
    If Err.Number <> 0 Then p1 = 0
    On Error GoTo 0

    If LoadFlag = True Then Exit Sub
    If IsNull(TS!BatchID) Then Exit Sub
    If TS!BatchID = 0 Then Exit Sub
    
    If fg.TextMatrix(Row, Col) <> CellValue Then
        MsgBox "Change of timesheet entry not allowed" & vbCr & vbCr & _
               "Paycheck information already entered", vbInformation
        fg.TextMatrix(Row, Col) = CellValue
    End If
    fg.Refresh

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    CellValue = fg.TextMatrix(Row, Col)

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub GetWEDates()

Dim JDate As Long

    JDate = Int(Now())
    
    ' find next Saturday
    Do
        If JDate Mod 7 = 0 Then Exit Do
        JDate = JDate + 1
    Loop
    
    With Me.cmbWEDate
        For i = 1 To MaxWeeks
            .AddItem Format(JDate - 6, "mm/dd/yy") & " To: " & Format(JDate, "mm/dd/yy")
            .ItemData(.NewIndex) = JDate
            JDate = JDate - 7
        Next i
        .ListIndex = 0
        WEDate = .ItemData(.ListIndex)
    End With

End Sub

Private Sub PopDrops()

    ' Job Cost
    ' make temp record set
    ' so can sort by name
    ' name can be from different fields
    ' depending on what is filled in
    JC.CursorLocation = adUseClient
    JC.Fields.Append "JobID", adDouble
    JC.Fields.Append "CityID", adDouble
    JC.Fields.Append "CityRate", adCurrency
    JC.Fields.Append "Name", adVarChar, 80, adFldIsNullable
    JC.Open , , adOpenDynamic, adLockOptimistic

    ' Job Drop
    ' *** only jobs w/ City Rate filled in
    SQLString = "SELECT * FROM JCJob WHERE CityID <> 0 AND Active = 1"
    If JCJob.GetBySQL(SQLString) Then
        Do
            JC.AddNew
            JC!JobID = JCJob.JobID
            JC!CityID = JCJob.CityID
                            
            ' ******************
            ' *** stuff it   ***
            ' JC!CityID = (JCJob.JobID Mod 10) + 1
            ' ******************
            
            If PRCity.GetByID(JC!CityID) Then
                JC!CityRate = PRCity.CityRate
            Else
                JC!CityRate = 0
            End If
            If Trim(JCJob.FullName) <> "" Then
                JC!Name = Mid(JCJob.FullName, 1, 80)
            ElseIf Trim(JCJob.Name) <> "" Then
                JC!Name = Mid(JCJob.Name, 1, 80)
            Else
                X = Trim(JCJob.FirstName) & " " & Trim(JCJob.MidInit) & " " & Trim(JCJob.LastName)
                If X = "" Then
                    JC!Name = "Job ID: " & JCJob.JobID
                Else
                    JC!Name = Mid(X, 1, 30)
                End If
            End If
            JC!Name = Trim(JC!Name)
            JC.Update
            
            If JCJob.GetNext = False Then Exit Do
        Loop
    End If
        
    If JC.RecordCount = 0 Then
        MsgBox "No Job records found", vbExclamation
        GoBack
    End If
    
    JC.Sort = "Name"
    JobDrop = "|#0;NONE"
    JC.MoveFirst
    Do
        JobDrop = Trim(JobDrop) & "|#" & JC!JobID & ";" & Trim(JC!Name)
        JC.MoveNext
    Loop Until JC.EOF
    
    ' Dept Drop
    DeptDrop = ""
    SQLString = "SELECT * FROM PRDepartment ORDER BY DepartmentNumber"
    If PRDepartment.GetBySQL(SQLString) Then
        Do
            DeptDrop = Trim(DeptDrop) & "|#" & PRDepartment.DepartmentID & _
                       ";" & PRDepartment.Name
            If PRDepartment.GetNext = False Then Exit Do
        Loop
    End If
    
    ' city drop
    CityDrop = ""
    SQLString = "SELECT * FROM PRCity ORDER BY CityName"
    If PRCity.GetBySQL(SQLString) Then
        Do
            CityDrop = Trim(CityDrop) & "|#" & PRCity.CityID & ";" & PRCity.ShortName
            If PRCity.GetNext = False Then Exit Do
        Loop
    End If

End Sub

Private Sub PopEEFG()
    
    ' pop with all employees in PRTimeSheet for the week
    ' and all active employees
    
    On Error Resume Next
    EMP.Close
    Set EMP = Nothing
    fgEmp.DataMode = flexDMFree
    On Error GoTo 0
    
    EMP.CursorLocation = adUseClient
    EMP.Fields.Append "EmpID", adDouble
    EMP.Fields.Append "EmpNo", adDouble
    EMP.Fields.Append "Name", adVarChar, 50, adFldIsNullable
    EMP.Fields.Append "Hours", adSingle
    EMP.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate)
    rsInit SQLString, cn, TS
    If TS.RecordCount > 0 Then
        Do
            EMP.Find "EmpID = " & TS!EmployeeID, 0, adSearchForward, 1
            If EMP.EOF Then
                If PREmployee.GetByID(TS!EmployeeID) = False Then
                    MsgBox "Employee ID Not Found: " & TS!EmployeeID, vbExclamation
                    GoBack
                End If
                EMP.AddNew
                EMP!EmpID = PREmployee.EmployeeID
                EMP!EmpNo = PREmployee.EmployeeNumber
                EMP!Name = PREmployee.LFName
                EMP!Hours = 0
                EMP.Update
            End If
            TS.MoveNext
        Loop Until TS.EOF
    End If
    
    ' see if any active employees need to be added
    SQLString = "SELECT * FROM PREmployee WHERE Inactive = 0"
    If PREmployee.GetBySQL(SQLString) Then
        Do
            EMP.Find "EmpID = " & PREmployee.EmployeeID, 0, adSearchForward, 1
            If EMP.EOF Then
                EMP.AddNew
                EMP!EmpID = PREmployee.EmployeeID
                EMP!EmpNo = PREmployee.EmployeeNumber
                EMP!Name = PREmployee.LFName
                EMP!Hours = 0
                EMP.Update
            End If
            If Not PREmployee.GetNext Then Exit Do
        Loop
    End If
    
    EMP.Sort = "Name"
    SortCol = 2
    SortOrder = 0
    
    ' SetGrid EMP, Me.fgEmp
    fgEmp.FixedCols = 0                   ' see all cols selected by SQL
    fgEmp.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    fgEmp.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    fgEmp.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set fgEmp.DataSource = EMP.DataSource '
    fgEmp.DataMember = EMP.DataMember     '

    fgEmp.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fgEmp.TabBehavior = flexTabCells                       ' tab moves between cells
    fgEmp.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    With Me.fgEmp
        .FontSize = 9
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 4000
        .ColFormat(3) = HrFmt
        .ColWidth(3) = 1000
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    fgEmp.AutoSearch = flexSearchFromTop
    
    ' go to the top of the grid
    EMP.MoveFirst
    
End Sub

Private Sub PopFG()

Dim DayString As String
Dim rsERItem As New ADODB.Recordset
Dim ItemAbbrev, ItemTitle As String

    ' close it - don't bomb if not open
    On Error Resume Next
    TS.Close
    Set TS = Nothing
    fg.DataMode = flexDMFree
    On Error GoTo 0
    
    ' use billing rate?
    SQLString = "SELECT " & _
                " JobID, " & _
                " DepartmentID, " & _
                " ItemID, " & _
                " Note, " & _
                " SunHours, MonHours, TueHours, WedHours, ThuHours, FriHours, SatHours, TotalHours, " & _
                " EmployeeID, WEDate, BatchID" & _
                " FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate) & " " & _
                "AND EmployeeID = " & EMP!EmpID & " " & _
                "ORDER BY TimeSheetID"
    
    rsInit SQLString, cn, TS
    
    ' SetGrid TS, fg
    fg.FixedCols = 0                   ' see all cols selected by SQL
    fg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
    fg.DataMode = flexDMBound          ' Recordset cursor is maintained by grid
    fg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    Set fg.DataSource = TS.DataSource '
    fg.DataMember = TS.DataMember     '

    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.AllowSelection = False                          ' don't allow selection of ranges of cells
    
    With fg

        .FontSize = 8

        ' column headers
        .TextMatrix(0, 0) = "Customer:Job"
        .TextMatrix(0, 1) = "Work Cat"
        .TextMatrix(0, 2) = "Earng Type"
        .TextMatrix(0, 3) = "Note"

        ' Item Drop - ** PER EMPLOYEE ITEMS"
        ItemDrop = "|#99991;REG PAY|#99992;OVT PAY"
        SQLString = "SELECT * FROM PRItem WHERE " & _
                    "ItemType = " & PREquate.ItemTypeOE & " " & _
                    "AND EmployeeID = " & EMP!EmpID & " " & _
                    "ORDER BY ItemID"
        
        If PRItem.GetBySQL(SQLString) Then
            Do
                ' get the employer defn
                SQLString = "SELECT * FROM PRItem WHERE ItemID = " & PRItem.EmployerItemID
                rsInit SQLString, cn, rsERItem
                If rsERItem.RecordCount = 0 Then
                    MsgBox "Employer Item NF: " & PRItem.EmployerItemID, vbExclamation
                    End
                End If
                
                If IsNull(rsERItem!Title) Then
                    ItemTitle = ""
                Else
                    ItemTitle = Trim(rsERItem!Title) & ""
                End If
                
                If IsNull(rsERItem!Abbreviation) Then
                    ItemAbbrev = ""
                Else
                    ItemAbbrev = Trim(rsERItem!Abbreviation) & ""
                End If
                
                If ItemAbbrev <> "" Then
                    ItemDrop = Trim(ItemDrop) & "|#" & rsERItem!ItemID & ";" & Trim(ItemAbbrev)
                Else
                    ItemDrop = Trim(ItemDrop) & "|#" & rsERItem!ItemID & ";" & Trim(ItemTitle)
                End If
                If PRItem.GetNext = False Then Exit Do
            Loop
        End If

        ' assign drop downs
        .ColComboList(0) = JobDrop
        ' .ColComboList(3) = CityDrop
        .ColComboList(1) = DeptDrop
        .ColComboList(2) = ItemDrop

        ' right justify the job column?
        .ColAlignment(0) = flexAlignRightCenter

        ' other column widths
        .ColWidth(0) = 2500     ' job
        .ColWidth(1) = 1200     ' department
        .ColWidth(2) = 1200     ' item
        .ColWidth(3) = 1700     ' note

        ' show the date
        For i = 1 To 7
            DayString = Mid("SunMonTueWedThuFriSat", i * 3 - 2, 3)
            DayString = Trim(DayString) & " " & Day(WEDate - 7 + i)
            .TextMatrix(0, i + 3) = DayString
        Next i
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = "Total Hours"
        
        ' hour columns
        ' .SubtotalPosition = flexSTAbove
        For i = 4 To 11
            .ColFormat(i) = HrFmt
            .ColWidth(i) = 800
            ' .Subtotal flexSTSum, -1, i, , RGB(1, 1, 1), vbWhite, True
        Next i

        ' blank extra columns
        For i = 12 To .Cols - 1
            .ColWidth(i) = 0
        Next i
    
        .GridColor = vbBlack
    
    End With
    
    ' show the job name
    Me.lblJobName = ""
    On Error Resume Next
    If TS.RecordCount > 0 Then
        TS.MoveFirst
        If fg.TextMatrix(1, 0) = "" Then
        Else
            If JCJob.GetByID(fg.TextMatrix(1, 0)) = True Then
                Me.lblJobName = JCJob.FullName
            End If
        End If
    End If
    On Error GoTo 0
    
    CalcTotals 4
    
End Sub
Private Sub fg_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If LoadFlag Then Exit Sub
    CalcTotals Col
End Sub

Private Sub fgEmp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    ' employee name display
    On Error Resume Next
    Me.lblEEName = fgEmp.TextMatrix(fgEmp.Row, 2)
    On Error GoTo 0
    
    If LoadFlag = True Then Exit Sub
    CalcAll
    
    PopFG
     
    ' fg.SetFocus
    ' AddLine
    
End Sub

Private Sub AddLine()

Dim tsDepartmentID, tsJobID As Long

    If PREmployee.GetByID(EMP!EmpID) = False Then
        MsgBox "PREmployee Error: " & EMP!EmpID, vbExclamation
        GoBack
    End If
    
    ' handle if closed ???
    If TS.RecordCount > 0 Then
        TS.MoveFirst
        tsJobID = TS!JobID
        tsDepartmentID = TS!DepartmentID
    Else
        tsJobID = 0
        tsDepartmentID = 0
    End If
    
    TS.AddNew
    TS!EmployeeID = PREmployee.EmployeeID
    
    If tsJobID = 0 Then
        TS!JobID = PREmployee.DefaultJobID
    Else
        TS!JobID = tsJobID
    End If
    
    If tsDepartmentID = 0 Then
        TS!DepartmentID = PREmployee.DepartmentID
    Else
        TS!DepartmentID = tsDepartmentID
    End If
    
    TS!ItemID = 99991
    TS!WEDate = WEDate
    TS.Update
    TS.Requery
    
    SetGrid TS, fg, 1
    fg.Row = fg.Rows - 1
    fg.Col = 0
    fg.SetFocus

End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub
Private Sub cmdAddLine_Click()
    AddLine
End Sub

Private Sub cmdDelete_Click()
    If TS.RecordCount = 0 Then Exit Sub
    TS.Delete
    TS.Requery
    CalcTotals 7
    CalcAll
    PopFG
End Sub

Private Sub cmdDelAll_Click()
    If TS.RecordCount = 0 Then Exit Sub
    If MsgBox("OK to delete ALL entries for:" & vbCr & EMP!Name, _
              vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    TS.MoveFirst
    Do
        TS.Delete
        TS.MoveNext
    Loop Until TS.EOF
    TS.Requery
    CalcTotals 7
    CalcAll
    PopFG
End Sub

Private Sub cmdAddEmp_Click()
    
Dim SaveRow, SaveEmpID As Long
    
    ' save grid positions
    SaveRow = fgEmp.TopRow
    SaveEmpID = EMP!EmpID
    
    frmAddEmployee.Init
    frmAddEmployee.Show vbModal
    If frmAddEmployee.EmpID <= 0 Then Exit Sub
    If PREmployee.GetByID(frmAddEmployee.EmpID) = False Then
        Exit Sub
    End If
    
    ' don't add if already exists
    EMP.Find "EmpID = " & frmAddEmployee.EmpID, 0, adSearchForward, 1
    If EMP.EOF = False Then
        MsgBox "Employee already included!", vbInformation
        EMP.Find "EmpID = " & SaveEmpID, 0, adSearchForward, 1
        fgEmp.TopRow = SaveRow
    Else
        LoadFlag = True
        EMP.AddNew
        EMP!EmpID = PREmployee.EmployeeID
        EMP!EmpNo = PREmployee.EmployeeNumber
        EMP!Name = PREmployee.LFName
        EMP!Hours = 0
        EMP.Update
        
        ' point to the row just added????
        For fgRow = 1 To fgEmp.Rows - 1
            fgEmp.Row = fgRow
            If fgEmp.TextMatrix(fgRow, 0) = frmAddEmployee.EmpID Then
                Exit For
            End If
        Next fgRow
        fgEmp.TopRow = fgEmp.Row
        PopFG
        LoadFlag = False
    
    End If

End Sub

Private Sub CalcTotals(ByVal Col As Long)

Dim EETotal, Hrs As Single
Dim CellValue As String
Dim DayTotal(8) As Single
Dim SaveRow As Long
    
    If LoadFlag = True Then Exit Sub
    
    SaveRow = fg.Row
    
    ' only do if day column being changed
    If Col < 4 Then Exit Sub
    If Col > 10 Then Exit Sub
    
    EETotal = 0
    For fgCol = 1 To 8
        DayTotal(fgCol) = 0
    Next fgCol
    
    ' update row hour totals
    For fgRow = 1 To fg.Rows - 1
        Hrs = 0
        For fgCol = 4 To 10
            CellValue = fg.TextMatrix(fgRow, fgCol)
            If Not IsNull(CellValue) And CellValue <> "" Then
                Hrs = Hrs + CSng(fg.TextMatrix(fgRow, fgCol))
                DayTotal(fgCol - 3) = DayTotal(fgCol - 3) + CSng(fg.TextMatrix(fgRow, fgCol))
            End If
        Next fgCol
        fg.TextMatrix(fgRow, 11) = Format(Hrs, HrFmt)
        EETotal = EETotal + Hrs
    Next fgRow
    
    Me.lblSunHrs = Format(DayTotal(1), HrFmt)
    Me.lblMonHrs = Format(DayTotal(2), HrFmt)
    Me.lblTueHrs = Format(DayTotal(3), HrFmt)
    Me.lblWedHrs = Format(DayTotal(4), HrFmt)
    Me.lblThuHrs = Format(DayTotal(5), HrFmt)
    Me.lblFriHrs = Format(DayTotal(6), HrFmt)
    Me.lblSatHrs = Format(DayTotal(7), HrFmt)
    Me.lblTotHrs = Format(EETotal, HrFmt)
    
    EMP!Hours = EETotal
    Me.lblTotalHours = "TOTAL HOURS: " & Format(TotalHrsWithout + EETotal, HrFmt)
    
    fg.Row = SaveRow
    
End Sub

Private Sub CalcAll()
    
Dim LastID, EmpID As Long
Dim TSHrs, HrSubtl As Single
Dim SaveRow As Long
    
    On Error Resume Next
    TS.Close
    On Error GoTo 0
    
    ' save the employee ID pointed to
    EmpID = EMP!EmpID
    SaveRow = fgEmp.TopRow

    ' clear amounts
    LoadFlag = True
    TotalHrsWith = 0
    TotalHrsWithout = 0
    HrSubtl = 0
    LastID = 0
    If EMP.RecordCount > 0 Then
        EMP.MoveFirst
        Do
            EMP!Hours = 0
            EMP.Update
            EMP.MoveNext
        Loop Until EMP.EOF
    End If

    ' run thru all records for the WE Date
    ' get the total hours NOT INCLUDING the current employee selected
    ' update the totals in the Employee grid also
    SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate) & _
                " ORDER BY EmployeeID"
    rsInit SQLString, cn, TS
    If TS.RecordCount > 0 Then
        TS.MoveFirst
        Do
            If IsNull(TS!TotalHours) Then
                TSHrs = 0
            Else
                TSHrs = TS!TotalHours
            End If
            TotalHrsWith = TotalHrsWith + TSHrs
            If TS!EmployeeID <> EmpID Then
                TotalHrsWithout = TotalHrsWithout + TSHrs
            End If
            
            If LastID = 0 Or TS!EmployeeID <> LastID Then
                If LastID <> 0 Then
                    EMP!Hours = HrSubtl
                    EMP.Update
                End If
                EMP.Find "EmpID = " & TS!EmployeeID, 0, adSearchForward, 1
                If EMP.EOF Then
                    MsgBox "empid nf " & TS!EmployeeID
                    End ' ??????
                End If
                HrSubtl = 0
            End If
            LastID = EMP!EmpID
            HrSubtl = HrSubtl + TSHrs
            TS.MoveNext
        Loop Until TS.EOF
        ' update the last EE record
        EMP!Hours = HrSubtl
        EMP.Update
    
    End If

    ' go back to the original EMP
    EMP.Find "EmpID = " & EmpID, 0, adSearchForward, 1
    fgEmp.TopRow = SaveRow
    TS.Close

    LoadFlag = False

    Me.lblTotalHours = "TOTAL HOURS: " & Format(TotalHrsWith, HrFmt)

End Sub

Private Sub cmbWEDate_Click()
    If LoadFlag = True Then Exit Sub
    With Me.cmbWEDate
        WEDate = .ItemData(.ListIndex)
    End With
    LoadFlag = True
    PopEEFG
    PopFG
    CalcAll
    CalcTotals 7
    LoadFlag = False
    EMP.MoveFirst
    PopFG
End Sub

Private Sub cmdPrint_Click()
    CalcAll
    CalcTotals 7
    frmPRTSPrint.WEDate = WEDate
    frmPRTSPrint.Show vbModal
End Sub

Private Sub cmdSortHours_Click()
    fgEmpSort 3
End Sub

Private Sub cmdSortName_Click()
    fgEmpSort 2
End Sub

Private Sub cmdSortNum_Click()
    fgEmpSort 1
End Sub
Private Sub fgEmpSort(ByVal sCol As Byte)

    If sCol = SortCol Then
        If SortOrder = 0 Then
            SortOrder = 1
        Else
            SortOrder = 0
        End If
    End If
    SortCol = sCol
    SQLString = fgEmp.TextMatrix(0, sCol)
    If SortOrder = 1 Then
        SQLString = SQLString & " DESC"
    End If
    EMP.Sort = SQLString
    fgEmp.Refresh
    EMP.MoveFirst
    PopFG

End Sub

Private Sub fgEmp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
'' MsgBox fgEmp.MouseRow & vbCr & fgEmp.MouseCol & vbCr & SortCol & vbCr & SortOrder
'' Exit Sub
'
'    With fgEmp
'
'        ' clicking on a column header sorts based on that column
'        If Button = 1 And Shift = 0 And .MouseRow = 0 Then
'
'            If .MouseCol = SortCol Then
'                If SortOrder = 0 Then
'                    SortOrder = 1
'                Else
'                    SortOrder = 0
'                End If
'            Else
'                SortOrder = 0
'            End If
'
'            SortCol = .MouseCol
'
'            SQLString = .TextMatrix(0, SortCol)
'            If SortOrder = 1 Then
'                SQLString = SQLString & " DESC"
'            End If
'
'            EMP.Sort = SQLString
'
'            ' PopEEFG
'
'        End If
'
'    End With

End Sub

Private Sub cmd40Hours_Click()
    
Dim fgRow, fgCol As Long
    
    On Error Resume Next
    fgRow = fg.Row
    fgCol = fg.Col
    If fgRow < 1 Then Exit Sub
    If fg.Rows <= 1 Then Exit Sub
    fg.TextMatrix(fgRow, 4) = "0.00"
    fg.TextMatrix(fgRow, 5) = "8.00"
    fg.TextMatrix(fgRow, 6) = "8.00"
    fg.TextMatrix(fgRow, 7) = "8.00"
    fg.TextMatrix(fgRow, 8) = "8.00"
    fg.TextMatrix(fgRow, 9) = "8.00"
    fg.TextMatrix(fgRow, 10) = "0.00"
    fg.Refresh
    On Error GoTo 0
    
End Sub

Private Sub fgEmp_DblClick()
    On Error Resume Next
    If fg.Rows = 1 Then
        cmdAddLine_Click
        fg.SetFocus
    End If
    On Error GoTo 0
End Sub


