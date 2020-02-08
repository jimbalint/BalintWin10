VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAssignCity 
   Caption         =   "&H80000009&"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
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
   ScaleHeight     =   10485
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDateRange 
      Caption         =   "Date Range"
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
      Left            =   3240
      TabIndex        =   23
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CheckBox chkMonthRange 
      Caption         =   "Month Range"
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
      Left            =   3240
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CheckBox chkAllMonths 
      Caption         =   "All Months"
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
      Left            =   3240
      TabIndex        =   21
      Top             =   840
      Width           =   1575
   End
   Begin TDBDate6Ctl.TDBDate tdbStartDate 
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   1800
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calendar        =   "frmAssignCity.frx":0000
      Caption         =   "frmAssignCity.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAssignCity.frx":0164
      Keys            =   "frmAssignCity.frx":0182
      Spin            =   "frmAssignCity.frx":01E0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "06/30/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40359
      CenturyMode     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "   History Data Change   "
      Height          =   855
      Left            =   5520
      TabIndex        =   13
      Top             =   3120
      Width           =   4695
      Begin VB.OptionButton optDpt 
         Caption         =   "DEPARTMENT"
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optCity 
         Caption         =   "CITY"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
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
      Height          =   495
      Left            =   2625
      TabIndex        =   6
      Top             =   9840
      Width           =   1455
   End
   Begin VB.ComboBox cmbEYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1200
   End
   Begin VB.ComboBox cmbSYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid fgCity 
      Height          =   5535
      Left            =   5400
      TabIndex        =   5
      Top             =   4080
      Width           =   5055
      _cx             =   8916
      _cy             =   9763
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6825
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6135
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   4455
      _cx             =   7858
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "C&LEAR ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "&CHECK ALL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin TDBDate6Ctl.TDBDate tdbEndDate 
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   1800
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calendar        =   "frmAssignCity.frx":0208
      Caption         =   "frmAssignCity.frx":0308
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAssignCity.frx":036C
      Keys            =   "frmAssignCity.frx":038A
      Spin            =   "frmAssignCity.frx":03E8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "06/30/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40359
      CenturyMode     =   0
   End
   Begin VB.Label lblD2 
      Caption         =   "End:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblD1 
      Caption         =   " Start:"
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
      Left            =   5040
      TabIndex        =   17
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Select History range to change: (Uses Check Date)"
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
      Left            =   840
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblChangeType 
      Alignment       =   2  'Center
      Caption         =   "City Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   2640
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Employee Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblM2 
      Caption         =   "End:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblM1 
      Caption         =   " Start:"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   735
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
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   10095
   End
End
Attribute VB_Name = "frmAssignCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public YearFlag As Boolean
Public RSEmp As New ADODB.Recordset
Public RSCty As New ADODB.Recordset
Public RSDist As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Public i, ChangeCount As Long
Dim StateDrop As String
Dim LoadFlag As Boolean


Private Sub Form_Load()
    
    LoadFlag = True
    
    ' default to city change
    Me.optCity = True
    Me.optDpt = False
    
    Load_Emp_Grid
    Load_City_Grid
    
    Me.cmbEYear.ListIndex = 0
    fgCity.Row = 1
    
    Me.cmbSYear.Visible = False
    Me.cmbEYear.Visible = False
    Me.lblM1.Visible = False
    Me.lblM2.Visible = False
    
    tdbDateSet Me.tdbStartDate, Now()
    tdbDateSet Me.tdbEndDate, Now()
    
    Me.tdbStartDate.Visible = False
    Me.tdbEndDate.Visible = False
    Me.lblD1.Visible = False
    Me.lblD2.Visible = False
    
    Me.chkAllMonths = 1
    Me.chkMonthRange = 0
    Me.chkDateRange = 0
    
    Me.KeyPreview = True
    
    LoadFlag = False

End Sub

Private Sub Load_Emp_Grid()

    '  Loop Through Employee File for all Employees
    RSEmp.CursorLocation = adUseClient
    RSEmp.Fields.Append "Selected", adBoolean
    RSEmp.Fields.Append "EmpNo", adDouble
    RSEmp.Fields.Append "EmpName", adVarChar, 80, adFldIsNullable
    RSEmp.Fields.Append "EmpID", adDouble
    Me.lblCompanyName.Caption = PRCompany.Name
    
    RSEmp.Open , , adOpenDynamic, adLockOptimistic
    SQLString = "Select * from PREmployee ORDER BY LastName, FirstName"
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employee Records were Found: ", vbCritical
        GoBack
    End If
    
    Do
        RSEmp.AddNew
        RSEmp!Selected = False
        RSEmp!EmpNo = PREmployee.EmployeeNumber
        RSEmp!EmpName = PREmployee.LFName
        RSEmp!EmpID = PREmployee.EmployeeID
        RSEmp.Update
        
        If Not PREmployee.GetNext Then Exit Do
    
    Loop
   
    ' Populate Year/Month Start Dropdown
    RSDist.CursorLocation = adUseClient
    
    SQLString = "SELECT YearMonth from PRHist ORDER BY YearMonth DESC"
    rsInit SQLString, cn, RSDist
    If RSDist.RecordCount = 0 Then
        MsgBox "No History Found!", vbExclamation
        GoBack
    End If
    
    RSDist.MoveFirst
    
    Do
        YearFlag = False
        For i = 1 To Me.cmbSYear.ListCount
            Me.cmbSYear.ListIndex = i - 1
            If RSDist!YearMonth = Me.cmbSYear Then
                YearFlag = True
                i = Me.cmbSYear.ListCount + 1
            End If
        Next i
        If YearFlag = False Then
            Me.cmbSYear.AddItem RSDist!YearMonth
        End If
        RSDist.MoveNext
    Loop Until RSDist.EOF
    Me.cmbSYear.ListIndex = 0
    
    ' Populate Year/Month End Dropdown
    RSDist.MoveFirst
    Do
        YearFlag = False
        For i = 1 To Me.cmbEYear.ListCount
            Me.cmbEYear.ListIndex = i - 1
            If RSDist!YearMonth = Me.cmbEYear Then
                YearFlag = True
                i = Me.cmbEYear.ListCount + 1
            End If
        Next i
        If YearFlag = False Then
            Me.cmbEYear.AddItem RSDist!YearMonth
        End If
        RSDist.MoveNext
    Loop Until RSDist.EOF
    
    SetGrid RSEmp, fg

    fg.ScrollBars = flexScrollBarVertical
    
End Sub

Private Sub Load_City_Grid()
    
    ' Loop Through PRCity File for all Cities
    On Error Resume Next
    RSCty.Close
    rsDept.Close
    On Error GoTo 0
    RSCty.CursorLocation = adUseClient
    RSCty.Fields.Append "CityNo", adDouble
    RSCty.Fields.Append "CityName", adVarChar, 80, adFldIsNullable
    RSCty.Fields.Append "CityState", adSmallInt
    RSCty.Fields.Append "CityRate", adDouble
    RSCty.Fields.Append "CityID", adDouble

    RSCty.Open , , adOpenDynamic, adLockOptimistic
    SQLString = "SELECT * FROM PRcity ORDER BY CityName"
    If Not PRCity.GetBySQL(SQLString) Then
        MsgBox "No City Records were Found: ", vbExclamation
        GoBack
    End If

    Do
        RSCty.AddNew
        RSCty!CityNo = PRCity.CityNumber
        RSCty!CityName = PRCity.CityName
        RSCty!CityState = PRCity.StateID
        RSCty!CityRate = PRCity.CityRate
        RSCty!CityID = PRCity.CityID
        RSCty.Update

        If Not PRCity.GetNext Then Exit Do
    Loop
    
    SetGrid RSCty, fgCity

    ' get the string for state name
    ' state drop down
    StateDrop = ""
    SQLString = "SELECT * FROM PRState ORDER BY StateAbbrev"
    If PRState.GetBySQL(SQLString) Then
        Do
            StateDrop = Trim(StateDrop) & "|#" & CStr(PRState.StateID) & ";" & Trim(PRState.StateAbbrev)
            If Not PRState.GetNext Then Exit Do
        Loop
    End If
    
    With Me.fgCity
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ColFormat(3) = "##0.00"
        .ColComboList(2) = StateDrop
        .ScrollBars = flexScrollBarVertical
        .ColWidth(1) = 2200
        .ColWidth(4) = 0
    End With

    Me.lblChangeType = "CITY LISTING"

    RSCty.MoveFirst

End Sub

Private Sub Load_Dept_Grid()
    
    On Error Resume Next
    rsDept.Close
    RSCty.Close
    On Error GoTo 0
    rsDept.CursorLocation = adUseClient
    rsDept.Fields.Append "DeptNum", adDouble
    rsDept.Fields.Append "DeptName", adVarChar, 30, adFldIsNullable
    rsDept.Fields.Append "DeptID", adDouble
    rsDept.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRDepartment ORDER BY DepartmentNumber"
    If PRDepartment.GetBySQL(SQLString) = False Then
        MsgBox "No departments defined!", vbExclamation
        GoBack
    End If
    
    Do
        rsDept.AddNew
        rsDept!DeptNum = PRDepartment.DepartmentNumber
        rsDept!DeptName = PRDepartment.Name
        rsDept!DeptID = PRDepartment.DepartmentID
        rsDept.Update
        If PRDepartment.GetNext = False Then Exit Do
    Loop
    
    SetGrid rsDept, Me.fgCity

    Me.lblChangeType = "DEPT LISTING"

    rsDept.MoveFirst

    With Me.fgCity
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        .ColWidth(1) = 3000
        .ColWidth(2) = 0
        .ScrollBars = flexScrollBarVertical
    End With

End Sub

Private Sub cmdCheckAll_Click()
    RSEmp.MoveFirst
    Do
        RSEmp!Selected = True
        RSEmp.Update
        RSEmp.MoveNext
    Loop Until RSEmp.EOF
    RSEmp.MoveFirst
End Sub

Private Sub cmdClearAll_Click()
    RSEmp.MoveFirst
    Do
        RSEmp!Selected = False
        RSEmp.Update
        RSEmp.MoveNext
    Loop Until RSEmp.EOF
    RSEmp.MoveFirst
End Sub

Private Sub fg_GotFocus()
    fg.SelectionMode = flexSelectionListBox
End Sub

Private Sub fgcity_Click()
    fgCity.SelectionMode = flexSelectionByRow
    fgCity.Editable = flexEDNone
    fgCity.AllowSelection = False
End Sub

Private Sub cmdOK_Click()
    
    If Me.optCity = True Then
        If MsgBox("OK to change city designations to: " & fgCity.TextMatrix(fgCity.Row, 1), _
                  vbYesNo + vbQuestion, "Payroll City Assignment") = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("OK to change DEPT designations to: " & fgCity.TextMatrix(fgCity.Row, 1), _
                  vbYesNo + vbQuestion, "Payroll DEPT Assignment") = vbNo Then Exit Sub
    End If
    
    ChangeCount = 0
    
    frmAssignCity.Hide

    frmProgress.Show
    frmProgress.lblMsg1 = "Now Changing City Assignments ...."
    
    If Me.optCity = True Then
        If RSCty!CityNo = 0 Then
            MsgBox "Please select a City from the City Listing", vbExclamation
            GoBack
        End If
    End If

    RSEmp.MoveFirst
    Do
        
        If RSEmp!Selected = False Then GoTo NextRSEmp
            
        If Me.chkAllMonths = 1 Then
            SQLString = "SELECT * FROM PRDist WHERE EmployeeID = " & RSEmp!EmpID
        ElseIf Me.chkMonthRange = 1 Then
            SQLString = "SELECT * FROM PRDist WHERE YearMonth >= " & Me.cmbSYear & _
                        " AND YearMonth <= " & Me.cmbEYear & _
                        " AND EmployeeID = " & RSEmp!EmpID
        Else
            SQLString = "SELECT * FROM PRDist WHERE CheckDate >= " & CLng(Me.tdbStartDate) & _
                        " AND CheckDate <= " & CLng(Me.tdbEndDate) & _
                        " AND EmployeeID = " & RSEmp!EmpID
        End If
            
        If PRDist.GetBySQL(SQLString) = True Then
        
            Do
                
                If Me.optCity = True Then
                    PRDist.CityID = RSCty!CityID
                    PRDist.StateID = RSCty!CityState
                Else
                    PRDist.DepartmentID = rsDept!DeptID
                End If
                
                PRDist.Save (Equate.RecPut)
                
                ChangeCount = ChangeCount + 1
                
                If Me.optCity = True Then
                    ' check PRHist for different state
                    If PRHist.GetByID(PRDist.HistID) Then
                        If PRHist.StateID <> RSCty!CityState Then
                            PRHist.StateID = RSCty!CityState
                            PRHist.Save (Equate.RecPut)
                        End If
                    End If
                End If
                
                If Not PRDist.GetNext Then Exit Do
            
            Loop
    
        End If
    
        ' change in PRHist/PRItemHist also
        If Me.optDpt = True Then
                
            If Me.chkAllMonths = 1 Then
                SQLString = "SELECT * FROM PRHist WHERE EmployeeID = " & RSEmp!EmpID
            ElseIf Me.chkMonthRange = 1 Then
                SQLString = "SELECT * FROM PRHist WHERE YearMonth >= " & Me.cmbSYear & _
                            " AND YearMonth <= " & Me.cmbEYear & _
                            " AND EmployeeID = " & RSEmp!EmpID
            Else
                SQLString = "SELECT * FROM PRHist WHERE CheckDate >= " & CLng(Me.tdbStartDate) & _
                            " AND CheckDate <= " & CLng(Me.tdbEndDate) & _
                            " AND EmployeeID = " & RSEmp!EmpID
            End If
            
            If PRHist.GetBySQL(SQLString) = True Then
                Do
                    PRHist.DepartmentID = rsDept!DeptID
                    PRHist.Save (Equate.RecPut)
                    ChangeCount = ChangeCount + 1
                    If PRHist.GetNext = False Then Exit Do
                Loop
            End If

            If Me.chkAllMonths = 1 Then
                SQLString = "SELECT * FROM PRItemHist WHERE EmployeeID = " & RSEmp!EmpID
            ElseIf Me.chkMonthRange = 1 Then
                SQLString = "SELECT * FROM PRItemHist WHERE YearMonth >= " & Me.cmbSYear & _
                            " AND YearMonth <= " & Me.cmbEYear & _
                            " AND EmployeeID = " & RSEmp!EmpID
            Else
                SQLString = "SELECT * FROM PRItemHist WHERE CheckDate >= " & CLng(Me.tdbStartDate) & _
                            " AND CheckDate <= " & CLng(Me.tdbEndDate) & _
                            " AND EmployeeID = " & RSEmp!EmpID
            End If
            If PRItemHist.GetBySQL(SQLString) = True Then
                Do
                    PRItemHist.DepartmentID = rsDept!DeptID
                    PRItemHist.Save (Equate.RecPut)
                    If PRItemHist.GetNext = False Then Exit Do
                    ChangeCount = ChangeCount + 1
                Loop
            End If
        
        End If
        
NextRSEmp:
        RSEmp.MoveNext
    Loop Until RSEmp.EOF

    frmProgress.Hide
    frmAssignCity.Show
    If Me.optCity = True Then
        RSCty.MoveFirst
    Else
        rsDept.MoveFirst
    End If
    
    RSEmp.MoveFirst
    
    MsgBox "Number of Records Changed:  " & Format(ChangeCount, "#,##0"), vbInformation

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub chkAllMonths_Click()
    If LoadFlag = True Then Exit Sub
    LoadFlag = True
    If Me.chkAllMonths = 1 Then
        Me.chkMonthRange = 0
        Me.chkDateRange = 0
    Else
        Me.chkMonthRange = 1
        Me.chkDateRange = 0
    End If
    RangeDisplay
    LoadFlag = False
End Sub
Private Sub chkMonthRange_Click()
    If LoadFlag = True Then Exit Sub
    LoadFlag = True
    If Me.chkMonthRange = 1 Then
        Me.chkAllMonths = 0
        Me.chkDateRange = 0
    Else
        Me.chkAllMonths = 0
        Me.chkDateRange = 1
    End If
    RangeDisplay
    LoadFlag = False
End Sub
Private Sub chkDateRange_Click()
    If LoadFlag = True Then Exit Sub
    LoadFlag = True
    If Me.chkDateRange = 1 Then
        Me.chkAllMonths = 0
        Me.chkMonthRange = 0
    Else
        Me.chkAllMonths = 1
        Me.chkMonthRange = 0
    End If
    RangeDisplay
    LoadFlag = False
End Sub
Private Sub RangeDisplay()
    
    Me.lblD1.Visible = False
    Me.lblD2.Visible = False
    Me.tdbStartDate.Visible = False
    Me.tdbEndDate.Visible = False
    
    Me.lblM1.Visible = False
    Me.lblM2.Visible = False
    Me.cmbSYear.Visible = False
    Me.cmbEYear.Visible = False
    
    If Me.chkAllMonths = 1 Then
    ElseIf Me.chkMonthRange = 1 Then
        Me.lblM1.Visible = True
        Me.lblM2.Visible = True
        Me.cmbSYear.Visible = True
        Me.cmbEYear.Visible = True
    Else
        Me.lblD1.Visible = True
        Me.lblD2.Visible = True
        Me.tdbStartDate.Visible = True
        Me.tdbEndDate.Visible = True
    End If

End Sub

Private Sub optCity_Click()
    If LoadFlag = True Then Exit Sub
    If Me.optCity = True Then
        Load_City_Grid
    Else
        Load_Dept_Grid
    End If
End Sub

Private Sub optDpt_Click()
    If LoadFlag = True Then Exit Sub
    If Me.optCity = True Then
        Load_City_Grid
    Else
        Load_Dept_Grid
    End If
End Sub

