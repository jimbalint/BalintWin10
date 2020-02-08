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
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   13815
      _cx             =   24368
      _cy             =   8705
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
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddEmp 
      Caption         =   "&ADD INACTIVE EMPLOYEE"
      Height          =   855
      Left            =   11520
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   1935
      Left            =   5520
      TabIndex        =   4
      Top             =   720
      Width           =   5655
      _cx             =   9975
      _cy             =   3413
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
      Top             =   8520
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
      Left            =   7800
      TabIndex        =   10
      Top             =   2760
      Width           =   855
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
      Left            =   6840
      TabIndex        =   9
      Top             =   2760
      Width           =   855
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

Dim LoadFlag As Boolean
Dim MaxWeeks As Integer
Dim i, j, k As Long
Dim rw, Co As Long
Dim JC As New ADODB.Recordset
Dim EMP As New ADODB.Recordset
Dim TS As New ADODB.Recordset
Dim HrFmt As String

Dim JobDrop, DeptDrop, ItemDrop, CityDrop As String

Private Sub Form_Load()

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

    LoadFlag = False
    
    ' initial display and cursor setting
    CalcTotals 7
    Me.Show
    fg.SetFocus
    addline
    fg.Row = 1
    fg.Col = 2

    Me.KeyPreview = True

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
    JC.Fields.Append "Name", adVarChar, 30, adFldIsNullable
    JC.Open , , adOpenDynamic, adLockOptimistic

    ' Job Drop
    ' *** only jobs w/ City Rate filled in
    SQLString = "SELECT * FROM JCJob WHERE CityID <> 0"
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
                JC!Name = Mid(JCJob.FullName, 1, 30)
            ElseIf Trim(JCJob.Name) <> "" Then
                JC!Name = Mid(JCJob.Name, 1, 30)
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
    
    ' Item Drop
    ItemDrop = "|#99991;REG PAY|#99992;OVT PAY"
    SQLString = "SELECT * FROM PRItem WHERE " & _
                "ItemType = " & PREquate.ItemTypeOE & " " & _
                "AND EmployeeID = 0 " & _
                "ORDER BY ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do
            ItemDrop = Trim(ItemDrop) & "|#" & PRItem.ItemID & ";" & PRItem.Abbreviation
            If PRItem.GetNext = False Then Exit Do
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
    
    If LoadFlag = False Then EMP.Close
    
    EMP.CursorLocation = adUseClient
    EMP.Fields.Append "EmpID", adDouble
    EMP.Fields.Append "EmpNo", adDouble
    EMP.Fields.Append "Name", adVarChar, 50, adFldIsNullable
    EMP.Fields.Append "Hours", adSingle
    EMP.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PREmployee WHERE InActive = 0"
    If PREmployee.GetBySQL(SQLString) Then
        Do
            EMP.AddNew
            EMP!EmpID = PREmployee.EmployeeID
            EMP!EmpNo = PREmployee.EmployeeNumber
                        
            ' *********
            EMP!EmpNo = PREmployee.EmployeeID
                        
            EMP!Name = PREmployee.LFName
            EMP!Hours = 0
            EMP.Update
            If Not PREmployee.GetNext Then Exit Do
        Loop
    End If
    
    ' ++++++++++++++++++++++++++++++++++++++++
    ' add employees in Week not active
    ' ++++++++++++++++++++++++++++++++++++++++
    
    SetGrid EMP, Me.fgEmp
    
    With Me.fgEmp
        .FontSize = 9
        .ColWidth(0) = 0
        .ColWidth(2) = 3200
        .ColFormat(3) = HrFmt
        .ColWidth(3) = 1000
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
    End With
    
    ' go to the top of the grid
    EMP.MoveFirst
    fgEmp.Row = 1
    
End Sub

Private Sub PopFG()

Dim WEDate As Date
Dim DayString As String

    If LoadFlag = False Then TS.Close

    With Me.cmbWEDate
        WEDate = .ItemData(.ListIndex)
    End With
    
    SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate) & " " & _
                "AND EmployeeID = " & EMP!EmpID & " " & _
                "ORDER BY TimeSheetID"
                
    rsInit SQLString, cn, TS
    SetGrid TS, fg
    With fg

        .FontSize = 8

        ' hour columns
        ' .SubtotalPosition = flexSTAbove
        For i = 7 To 14
            .ColFormat(i) = HrFmt
            .ColWidth(i) = 800
            ' .Subtotal flexSTSum, -1, i, , RGB(1, 1, 1), vbWhite, True
        Next i

        ' column headers
        .TextMatrix(0, 2) = "Customer:Job"
        .TextMatrix(0, 4) = "Work Cat"
        .TextMatrix(0, 5) = "Earng Type"
        .TextMatrix(0, 6) = "Note"

        ' assign drop downs
        .ColComboList(2) = JobDrop
        .ColComboList(3) = CityDrop
        .ColComboList(4) = DeptDrop
        .ColComboList(5) = ItemDrop

        ' hidden columns
        .ColHidden(0) = True  ' TS ID
        .ColHidden(1) = True  ' emp id
        .ColHidden(3) = True  ' don't show the city???
        For i = 15 To 19
            .ColHidden(i) = True
        Next i

        ' other column widths
        .ColWidth(2) = 2500     ' job
        .ColWidth(4) = 1200     ' department
        .ColWidth(5) = 1200     ' item
        .ColWidth(6) = 1700     ' note

        ' show the date
        For i = 1 To 7
            DayString = Mid("SunMonTueWedThuFriSat", i * 3 - 2, 3)
            DayString = Trim(DayString) & " " & Day(WEDate - 7 + i)
            .TextMatrix(0, i + 6) = DayString
        Next i

    End With
    
'    For i = 1 To 10
'        TS.AddNew
'        TS!EmployeeID = EMP!EmpID
'        TS!WEDate = WEDate
'        TS.Update
'    Next i

End Sub
Private Sub fg_CellChanged(ByVal Row As Long, ByVal Col As Long)
    CalcTotals Col
End Sub

Private Sub CalcTotals(ByVal Col As Long)

Dim Hrs As Single
Dim CellValue As String
Dim DayTotal(8) As Single
    
    If LoadFlag = True Then Exit Sub
    
    ' only do if day column being changed
    If Col < 7 Then Exit Sub
    If Col > 13 Then Exit Sub
    
    For Co = 1 To 8
        DayTotal(Co) = 0
    Next Co
    
    ' update row hour totals
    For rw = 1 To fg.Rows - 1
        Hrs = 0
        For Co = 7 To 13
            CellValue = fg.TextMatrix(rw, Co)
            If Not IsNull(CellValue) And CellValue <> "" Then
                Hrs = Hrs + CSng(fg.TextMatrix(rw, Co))
                DayTotal(Co - 6) = DayTotal(Co - 6) + CSng(fg.TextMatrix(rw, Co))
            End If
        Next Co
        fg.TextMatrix(rw, 14) = Format(Hrs, HrFmt)
    Next rw
    
    Me.lblSunHrs = Format(DayTotal(1), HrFmt)
    Me.lblMonHrs = Format(DayTotal(2), HrFmt)
    
End Sub
Private Sub fgEmp_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If LoadFlag = True Then Exit Sub
    PopFG
    fg.SetFocus
    
    If fg.Rows = 1 Then Exit Sub
    fg.Row = 1
    fg.Col = 2
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub


