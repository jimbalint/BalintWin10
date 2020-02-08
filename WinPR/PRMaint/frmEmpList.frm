VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEmpList 
   Caption         =   "Employee List"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7830
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActiveOnly 
      Caption         =   "Show Only Active Employees"
      Height          =   255
      Left            =   7440
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdTaxWage 
      Caption         =   "&TAXABLE WAGES"
      Height          =   735
      Left            =   10560
      TabIndex        =   5
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   10560
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      Default         =   -1  'True
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   9855
      _cx             =   17383
      _cy             =   11033
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmEmpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Dim X As String
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending

Dim dbFileName As String
Dim dbFields(4) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte
Dim LoadFlag As Boolean

Dim SQLStr As String

Private Sub Form_Load()
    
    LoadFlag = True
    
    Me.lblCompanyName = PRCompany.Name
        
    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
        
    ' default filtered?
    SQLString = "SELECT * FROM PRGlobal WHERE Description = 'EEMAINT' " & _
                " AND UserID = " & User.ID & _
                " AND Var2 = '" & PRCompany.CompanyID & "'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        If PRGlobal.Var1 = "1" Then Me.chkActiveOnly = 1
    End If
    GetEEData

    LoadFlag = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub


Private Sub cmdAdd_Click()
    
Dim EmployeeID As Long
    
    frmAddNewEmp.lblMsg1 = Trim(PRCompany.Name)
    frmAddNewEmp.Init
    frmAddNewEmp.Show vbModal
    EmployeeID = frmAddNewEmp.EmployeeID
    If EmployeeID = -1 Then Exit Sub
    
    ' re-init the grid
    ' rsInit GetSQLString, cn, rs
    ' Set fg.DataSource = rs.DataSource
    
    rs.Requery
    
    ' start the emp maint screen
    SelID = EmployeeID
    frmEmpForm.Show vbModal

    rs.Requery

    ' goto that line
    ' problem after new employee add
    '   err on fg.TopRow = 1 ???
    On Error Resume Next
    fg.Row = fg.FindRow(SelID, 0, 0)
    fg.TopRow = fg.Row
    rs.Find "EmployeeID = " & fg.TextMatrix(fg.Row, 0), 0, adSearchForward, 1
    On Error GoTo 0
    
End Sub

Private Sub cmdEdit_Click()

    If fg.Rows = 1 Then Exit Sub

    rw = 0
    If fg.Rows > 0 Then
        rw = fg.Row
    End If
     
    SelID = rs!EmployeeID
    frmEmpForm.Show vbModal
    Unload frmEmpForm
     
    rs.Close
    rsInit GetSQLString, cn, rs
    Set fg.DataSource = rs.DataSource
       
    rs.Find "EmployeeID = " & SelID, 0, adSearchForward, 1
    If rs.EOF = False Then
        rw = fg.FindRow(SelID, 0, 0)
        fg.TopRow = rw
        fg.Select rw, 0
    Else
        If fg.Rows > 0 And rw > 0 Then
            If rw > fg.Rows - 1 Then rw = fg.Rows - 1
            fg.TopRow = rw
            fg.Select rw, 0
        
            If rw = 1 Then rs.MoveFirst     ' ???
        
        End If
    End If
    
    fg.SetFocus

End Sub
     
Private Sub fg_DblClick()
    cmdEdit_Click
End Sub


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrID As Long
    
    CurrID = fg.TextMatrix(fg.Row, 0)
        
    rs.Close
    rsInit GetSQLString, cn, rs
    Set fg.DataSource = rs.DataSource
       
    rw = fg.FindRow(CurrID, 0, 0)
       
    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus
    
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Then     ' validates for number - must enter a value that does not already exist
       
       If fg.EditText = "" Or fg.EditText = "0" Then
          MsgBox "Enter a value!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
       rw = fg.FindRow(fg.EditText, 0, 0)
       If rw <> -1 Then
          MsgBox "Number already exists!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
    End If

End Sub

Private Sub cmdExit_Click()
    
    SQLString = "SELECT * FROM PRGlobal WHERE Description = 'EEMAINT' " & _
                " AND UserID = " & User.ID & _
                " AND Var2 = '" & PRCompany.CompanyID & "'"
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.Description = "EEMAINT"
        PRGlobal.UserID = User.ID
        PRGlobal.Var2 = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    PRGlobal.Var1 = Me.chkActiveOnly
    PRGlobal.Save (Equate.RecPut)
    
    GoBack

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
Dim trs As New ADODB.Recordset
    
    If fg.Rows = 1 Then Exit Sub
    
    ' no history can exist !!!
    SQLString = "SELECT * FROM PRHist WHERE PRHist.EmployeeID = " & fg.TextMatrix(fg.Row, 0)
    If PRHist.GetBySQL(SQLString) Then
        MsgBox UCase("History records exist - can't delete!"), vbExclamation
        Exit Sub
    End If
    
    ' what if no records left ????
        
    DelConfirm = MsgBox(fg.TextMatrix(fg.Row, 1) & vbCr & Trim(fg.TextMatrix(fg.Row, 2)) & ", " & Trim(fg.TextMatrix(fg.Row, 3)), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")
    
    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If
    
    ' delete records from related files
    SQLString = "DELETE * FROM PRItem WHERE PRItem.EmployeeID = " & fg.TextMatrix(fg.Row, 0)
    rsInit SQLString, cn, trs
    
    SQLString = "DELETE * FROM PREELists WHERE PREELists.EmployeeID = " & fg.TextMatrix(fg.Row, 0)
    rsInit SQLString, cn, trs
    
    rw = fg.Row
    ' DelAdo rs, fg, fg.TextMatrix(fg.Row, 0)
    DelAdo rs, fg
    
    If rw = fg.Rows Then rw = fg.Rows - 1
    
    fg.Select rw, 0
    fg.ShowCell rw, 0

End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    ' clicking on a column header sorts based on that column
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then

       ' toggle the sort order
       If fg.MouseCol = dbSortCol Then
          If dbSortDesc = False Then
             dbSortDesc = True
          Else
             dbSortDesc = False
          End If
       Else
          ' switch the column
          fg.Cell(flexcpFontBold, 0, fg.MouseCol) = True
          fg.Cell(flexcpFontBold, 0, dbSortCol) = False
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol)
          dbSortCol = fg.MouseCol
       End If
       
       If dbSortDesc Then
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "-"
       Else
          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "+"
       End If
    
       rs.Close
       
       rsInit GetSQLString, cn, rs
       Set fg.DataSource = rs.DataSource
       
       fg.ShowCell 1, 0

    End If
    
End Sub

Private Function GetSQLString() As String
    
Dim aa As Integer
    
' set the SQL string
'    x = "SELECT [Number],[Description] " & _
'        "FROM GLDescriptions ORDER BY [Number] DESC"

    GetSQLString = "SELECT "
    
    For aa = 0 To UBound(dbFields, 1)
        GetSQLString = GetSQLString & " [" & dbFields(aa) & "]"
        If aa <> UBound(dbFields, 1) Then GetSQLString = GetSQLString & ","
        GetSQLString = GetSQLString & " "
    Next aa
    
    GetSQLString = GetSQLString & "FROM " & dbFileName
    
    If Me.chkActiveOnly = 1 Then
        GetSQLString = Trim(GetSQLString) & " WHERE Inactive = 0"
    End If
    
    GetSQLString = Trim(GetSQLString) & " ORDER BY [" & dbFields(dbSortCol) & "]"
    
    If dbSortDesc Then
       GetSQLString = GetSQLString & " DESC"
    End If

End Function

Private Sub cmdTaxWage_Click()
    If rs.RecordCount = 0 Then Exit Sub
    EmpID = rs!EmployeeID
    frmTaxWage.Show vbModal
    Unload frmTaxWage
End Sub

Private Sub GetEEData()
    
    ' 11/06/2010 - continue if error
    
    ' get rid of nulls
    SQLString = "DELETE * FROM PREmployee WHERE PREmployee.EmployeeNumber = 0"
    ' rsInit SQLString, cn, rs
    On Error Resume Next
    cn.Execute SQLString
    On Error GoTo 0
    
    SQLString = "DELETE * FROM PREmployee WHERE IsNull(PREmployee.EmployeeNumber)"
    ' rsInit SQLString, cn, rs
    On Error Resume Next
    cn.Execute SQLString
    On Error GoTo 0
    
    ' set the constants for the file
    dbFileName = "PREmployee"
    dbFields(0) = "EmployeeID"
    dbFields(1) = "EmployeeNumber"
    dbFields(2) = "LastName"
    dbFields(3) = "FirstName"
    dbFields(4) = "InActive"
    dbSortCol = 1
    dbSortDesc = False
    
    X = GetSQLString
    
    rsInit X, cn, rs
    
    SetGrid rs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 0
    fg.ColWidth(1) = 1800
    fg.ColWidth(2) = 1800
    fg.ColWidth(3) = 3000
    
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    ' fg.HighLight = flexHighlightNever                   ' don't select ranges
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = dbFields(0) & "+"
    fg.Cell(flexcpFontBold, 0, 0) = True
    fg.AllowSelection = False
    fg.AutoSearch = flexSearchFromTop

End Sub

Private Sub chkActiveOnly_Click()
    If LoadFlag = True Then Exit Sub
    rs.Close
    GetEEData
End Sub


