VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTaxWage 
   Caption         =   "Taxable Wage Edit"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9090
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fgPRHist 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   12975
      _cx             =   22886
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
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VSFlex8Ctl.VSFlexGrid fgPRDist 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   12975
      _cx             =   22886
      _cy             =   3201
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
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label lblMaxWage 
      Caption         =   "Max Wages"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblSUNMax 
      Caption         =   "SUNMax"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   11760
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblFUNMax 
      Caption         =   "FUNMax"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   10200
      TabIndex        =   8
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblSSMax 
      Caption         =   "SSMax"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Select Tax Year:"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblEmployeeName 
      Alignment       =   2  'Center
      Caption         =   "Employee Name"
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
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   11895
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
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   11895
   End
End
Attribute VB_Name = "frmTaxWage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsPRDist As New ADODB.Recordset
Dim rsPRHist As New ADODB.Recordset
Dim rsTaxYear As New ADODB.Recordset
Dim DescDrop, CityDrop As String
Dim FirstFlag As Boolean

Dim SSMax, FUNMax, SUNMax As Currency

Private Sub Form_Load()

    ' screen displays
    Me.lblCompanyName = PRCompany.Name
    
    If Not PREmployee.GetByID(EmpID) Then
        MsgBox "Employee ID Not Found: " & EmpID, vbExclamation
        Unload Me
    End If
    
    Me.lblEmployeeName = PREmployee.FLName

    ' get the tax years that the employee has
    rsTaxYear.CursorLocation = adUseClient
    rsTaxYear.Fields.Append "TaxYear", adInteger
    rsTaxYear.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRHist WHERE EmployeeID = " & EmpID & _
                " ORDER BY YearMonth DESC"
    If PRHist.GetBySQL(SQLString) Then
        Do
            SQLString = "TaxYear = " & Int(PRHist.YearMonth / 100)
            rsTaxYear.Find SQLString, 0, adSearchForward, 1
            If rsTaxYear.EOF Then
                rsTaxYear.AddNew
                rsTaxYear!TaxYear = Int(PRHist.YearMonth / 100)
                rsTaxYear.Update
                Me.cmbTaxYear.AddItem Int(PRHist.YearMonth / 100)
            End If
            If Not PRHist.GetNext Then Exit Do
        Loop
    End If
        
    ' get the city id's and names for the display
    ' state drop down
    CityDrop = ""
    SQLString = "SELECT * FROM PRCity ORDER BY CityID"
    If PRCity.GetBySQL(SQLString) Then
        Do
            CityDrop = Trim(CityDrop) & "|#" & CStr(PRCity.CityID) & ";" & Trim(PRCity.CityName)
            If Not PRCity.GetNext Then Exit Do
        Loop
    End If
        
    ' drop down for dist desc
    DescDrop = "|#0;Reg/Ovt Earng"
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0"
    If PRItem.GetBySQL(SQLString) Then
        Do
            DescDrop = Trim(DescDrop) & "|#" & PRItem.ItemID & ";" & PRItem.Title
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If
        
    FirstFlag = True
        
    If rsTaxYear.RecordCount > 0 Then
        Me.cmbTaxYear.ListIndex = 0
    End If
    Me.KeyPreview = True

    FirstFlag = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    rsTaxYear.Close
    Unload Me
End Sub

Private Sub cmbTaxYear_Click()
    
Dim StartYM, EndYM As Long
Dim Clock As Long

    If FirstFlag = False Then
        rsPRDist.Close
        Set rsPRDist = Nothing
        rsPRHist.Close
        Set rsPRHist = Nothing
    End If

    StartYM = Me.cmbTaxYear.Text * 100 + 1
    EndYM = Me.cmbTaxYear.Text * 100 + 12
    
    ' load the data for the year selected
    SQLString = "SELECT PEDate, CheckDate, EmployerItemID, Hours, Amount, CityTax, CityWage, CityID " & _
                " FROM PRDist WHERE EmployeeID = " & EmpID & _
                " AND YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                " ORDER BY CheckDate, HistID"
    rsInit SQLString, cn, rsPRDist
    
    ' SetGridFree rsPRDist, Me.fgPRDist
    SetGrid rsPRDist, Me.fgPRDist

    ' column headers
    Me.fgPRDist.TextMatrix(0, 0) = "PE Date"
    Me.fgPRDist.TextMatrix(0, 1) = "Check Date"
    Me.fgPRDist.TextMatrix(0, 2) = "Description"
    Me.fgPRDist.TextMatrix(0, 3) = "Hours"
    Me.fgPRDist.TextMatrix(0, 4) = "Gross Pay"
    Me.fgPRDist.TextMatrix(0, 5) = "City Tax"
    Me.fgPRDist.TextMatrix(0, 6) = "City Wage"
    Me.fgPRDist.TextMatrix(0, 7) = "City"
    
    Me.fgPRDist.ColComboList(2) = DescDrop
    Me.fgPRDist.ColComboList(7) = CityDrop
    Me.fgPRDist.ColWidth(7) = 3000

    Me.fgPRDist.ColWidth(0) = 1200
    Me.fgPRDist.ColWidth(1) = 1200
    Me.fgPRDist.ColWidth(2) = 2000
    Me.fgPRDist.ColWidth(3) = 800
    Me.fgPRDist.ColWidth(4) = 1450
    Me.fgPRDist.ColWidth(5) = 1450
    Me.fgPRDist.ColWidth(6) = 1450
    Me.fgPRDist.ColWidth(7) = 2200
    
    ' bold the editable column header
    Me.fgPRDist.Cell(flexcpFontBold, 0, 6) = True

    SQLString = "SELECT PEDate, CheckDate, Gross, SSWage, MedWage, FWTWage, SWTWage, FUNWage, SUNWage " & _
                " FROM PRHist WHERE EmployeeID = " & EmpID & _
                " AND YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                " ORDER BY CheckDate, HistID"
    rsInit SQLString, cn, rsPRHist
    ' SetGridFree rsPRHist, Me.fgPRHist
    SetGrid rsPRHist, Me.fgPRHist
    
    Me.fgPRHist.ColWidth(0) = 1200
    Me.fgPRHist.ColWidth(1) = 1200
    Me.fgPRHist.ColWidth(2) = 1450
    Me.fgPRHist.ColWidth(3) = 1450
    Me.fgPRHist.ColWidth(4) = 1450
    Me.fgPRHist.ColWidth(5) = 1450
    Me.fgPRHist.ColWidth(6) = 1450
    Me.fgPRHist.ColWidth(7) = 1450
    Me.fgPRHist.ColWidth(8) = 1450
    
    ' total columns
'    If FirstFlag = False Then
'        Clock = Timer
'        Do
'            If Timer - Clock > 1 Then Exit Do
'        Loop
'    End If

    ' bold the editable column header
    Me.fgPRHist.Cell(flexcpFontBold, 0, 3) = True
    Me.fgPRHist.Cell(flexcpFontBold, 0, 4) = True
    Me.fgPRHist.Cell(flexcpFontBold, 0, 5) = True
    Me.fgPRHist.Cell(flexcpFontBold, 0, 6) = True
    Me.fgPRHist.Cell(flexcpFontBold, 0, 7) = True
    Me.fgPRHist.Cell(flexcpFontBold, 0, 8) = True
    
    ' refresh the max wages
    SSMax = PRGlobal.GetAmount(PREquate.GlobalTypeSSMax, Me.cmbTaxYear.Text)
    FUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, Me.cmbTaxYear.Text)
    SUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeSUNMax, Me.cmbTaxYear.Text)
    
    Me.lblSSMax = Format(SSMax, "$###,##0.00")
    Me.lblFUNMax = Format(FUNMax, "$##,##0.00")
    Me.lblSUNMax = Format(SUNMax, "$##,##0.00")
    Me.lblMaxWage = "MAX WAGES FOR " & Me.cmbTaxYear.Text
    
    If FirstFlag = False Then
        MsgBox "Tax Year has been changed to: " & Me.cmbTaxYear, vbInformation, "Taxable Wage Maint"
    End If
    
    fgPRDist.BackColorAlternate = 0
    
    ' **********************************************
    
    ' total the city taxable wage
    fgPRDist.Subtotal flexSTSum, -1, 3, , RGB(1, 1, 1), vbWhite, True
    fgPRDist.Subtotal flexSTSum, -1, 4
    fgPRDist.Subtotal flexSTSum, -1, 5
    fgPRDist.Subtotal flexSTSum, -1, 6
    If fgPRDist.Rows >= 2 Then fgPRDist.Select 2, 6

    fgPRHist.Subtotal flexSTSum, -1, 2, , RGB(1, 1, 1), vbWhite, True
    fgPRHist.Subtotal flexSTSum, -1, 3
    fgPRHist.Subtotal flexSTSum, -1, 4
    fgPRHist.Subtotal flexSTSum, -1, 5
    fgPRHist.Subtotal flexSTSum, -1, 6
    fgPRHist.Subtotal flexSTSum, -1, 7
    fgPRHist.Subtotal flexSTSum, -1, 8
    If fgPRHist.Rows >= 2 Then fgPRHist.Select 2, 3

    ' **********************************************
    
    'fgPRDist.DataRefresh
    'fgPRHist.DataRefresh

End Sub

Private Sub fgPRDist_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' don't allow edits of total rows
    If fgPRDist.IsSubtotal(Row) Then
        Cancel = True
        Exit Sub
    End If
    
    ' only allow edit of city wage column
    ' **** allow edit of cities for Exec Temp
    If Mid(PRCompany.Name, 1, 9) = "EXECUTIVE" And User.Logon = "jim" Then
        If Col <> 6 And Col <> 7 Then
            Cancel = True
            Exit Sub
        End If
    Else
        If Col <> 6 Then
            Cancel = True
            Exit Sub
        End If
    End If

End Sub

Private Sub fgPRHist_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' don't allow edits of total rows
    If fgPRHist.IsSubtotal(Row) Then
        Cancel = True
        Exit Sub
    End If
    
    ' only allow edit of city wage column
    If Col <= 2 Then
        Cancel = True
        Exit Sub
    End If

End Sub

Private Sub fgPRDist_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    fgPRDist.Subtotal flexSTSum, -1, 3
    fgPRDist.Subtotal flexSTSum, -1, 4
    fgPRDist.Subtotal flexSTSum, -1, 5
    fgPRDist.Subtotal flexSTSum, -1, 6
End Sub
Private Sub fgPRHist_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    fgPRHist.Subtotal flexSTSum, -1, 2
    fgPRHist.Subtotal flexSTSum, -1, 3
    fgPRHist.Subtotal flexSTSum, -1, 4
    fgPRHist.Subtotal flexSTSum, -1, 5
    fgPRHist.Subtotal flexSTSum, -1, 6
    fgPRHist.Subtotal flexSTSum, -1, 7
    fgPRHist.Subtotal flexSTSum, -1, 8
End Sub


