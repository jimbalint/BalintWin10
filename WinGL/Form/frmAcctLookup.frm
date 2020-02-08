VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmAcctLookup 
   Caption         =   "GLAccount Look Up"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctLookup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbBranch 
      Height          =   615
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1085
      Calculator      =   "frmAcctLookup.frx":030A
      Caption         =   "frmAcctLookup.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAcctLookup.frx":0394
      Keys            =   "frmAcctLookup.frx":03B2
      Spin            =   "frmAcctLookup.frx":03FC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   7208965
      MinValueVT      =   6881285
   End
   Begin VB.ComboBox cmbTypes 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "All Types"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame fraSection 
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   5655
      Begin VB.OptionButton optExpense 
         Caption         =   "&Expense"
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optIncome 
         Caption         =   "&Income"
         Height          =   255
         Left            =   2840
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optBalSht 
         Caption         =   "&Bal Sht"
         Height          =   255
         Left            =   1480
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Nex&t"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin TDBText6Ctl.TDBText tdbDescFind 
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   1085
      Caption         =   "frmAcctLookup.frx":0424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAcctLookup.frx":04A0
      Key             =   "frmAcctLookup.frx":04BE
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber tdbNumberFind 
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1085
      Calculator      =   "frmAcctLookup.frx":0502
      Caption         =   "frmAcctLookup.frx":0522
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAcctLookup.frx":0594
      Keys            =   "frmAcctLookup.frx":05B2
      Spin            =   "frmAcctLookup.frx":05FC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   ""
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1900545
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   10215
      _cx             =   18018
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
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
End
Attribute VB_Name = "frmAcctLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelAcct As Long

Private mcn As ADODB.Connection
Private mrs As ADODB.Recordset
Dim x, Y, z As String
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending

Dim dbFileName As String
Dim dbFields(1) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte

Dim SQLStr As String

Dim xdbAcct As New XArrayDB
Dim AcctFlg As Byte
Dim FilterType As Byte
Dim i, j As Long

' To Do:
' blank file test
' search field select
' sort by header change - cell picture set

Private Sub cmdAdd_Click()
    LoadGrid 2, "B"
    fg.SetFocus
End Sub

Private Sub cmbTypes_LostFocus()

    ' set other filters
    Me.optAll = True
    Me.tdbBranch = 0

    If cmbTypes = "All Types" Then
       LoadGrid 0, ""
    Else
       LoadGrid 1, Mid(cmbTypes, 1, 1)
    End If
    
    fg.ShowCell 1, 1
    fg.SetFocus

End Sub

Private Sub cmdOK_Click()
    
    SelAcct = CLng(fg.TextMatrix(fg.Row, 0))
    Me.Hide

End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrNum As Long
    
    CurrNum = fg.TextMatrix(fg.Row, 0)
        
    mrs.Close
    rsInit GetSQLString, cnDes, mrs
    Set fg.DataSource = mrs.DataSource
       
    rw = fg.FindRow(CurrNum, 0, 0)
       
    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus
    
    
End Sub

Private Sub fg_DblClick()

    SelAcct = CLng(fg.TextMatrix(fg.Row, 0))
    Me.Hide

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

Private Sub Form_Load()
    
    SelAcct = 0
    
    ' use Flex Grid in unbound mode
    ' use GLDescriptions lookup where needed
    
    x = "SELECT [Account], [AcctType], [Description], [DescNumber] FROM GLAccount ORDER BY Account"
    
    rsInit x, cn, mrs
    
    ' no records ???
    If mrs.BOF And mrs.EOF Then Unload Me
          
    GLDescription.OpenRS
    
    ' set the title line
    fg.Cols = 3
    fg.Rows = 0
    fg.AddItem "Account +" & vbTab & "Type" & vbTab & "Description"
    fg.FixedRows = 1
    
    ' customize the grid
    fg.ColWidth(0) = 1300
    fg.ColWidth(1) = 700
    fg.ColAlignment(1) = flexAlignCenterCenter
    fg.ColWidth(2) = 2000
    
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.SelectionMode = flexSelectionByRow
    fg.AllowSelection = False                           ' don't select ranges
    
    ' loop thru record set and assign to the xArray
    ' col 0 = acct #
    '     1 = type
    '     2 = description
    '     3 = branch
    '     4 = AcctFlg
    
    ' AcctFlg
    '  0 = balance sheet
    '  1 = income
    '  2 = expense
    
    xdbAcct.ReDim 0, 0, 0, 4
    rw = 0
    AcctFlg = 0
    xdbAcct(0, 0) = "x"
  
    mrs.MoveFirst
    Do Until mrs.EOF

       If mrs!DescNumber = 0 Then
          x = mrs!Description
       Else
          If GLDescription.Find(mrs!DescNumber) Then
             x = GLDescription.Description
          Else
             x = mrs!DescNumber
          End If
       End If
       
       xdbAcct.AppendRows
       
       rw = rw + 1
       xdbAcct(rw, 0) = CStr(mrs!Account)
       xdbAcct(rw, 1) = CStr(mrs!AcctType)
       xdbAcct(rw, 2) = CStr(x)
       If GLCompany.SubDigits <> 0 Then
          xdbAcct(rw, 3) = CStr(mrs!Account Mod 10 ^ GLCompany.SubDigits)
       Else
          xdbAcct(rw, 3) = "0"
       End If
          
       If mrs!Account > GLCompany.FirstPAcct Then
          AcctFlg = 1       ' income account
       End If
       
       If AcctFlg <> 0 And mrs!AcctType = "I" Then
          AcctFlg = 1
       End If
       
       If AcctFlg <> 0 And mrs!AcctType = "E" Then
          AcctFlg = 2
       End If
       
       xdbAcct(rw, 4) = CStr(AcctFlg)
       
       mrs.MoveNext
    
    Loop
    
    ' clean up
    GLDescription.CloseRS
    mrs.Close
    Set mrs = Nothing
    
    ' set filter defaults
    optAll = True
    optBalSht = False
    optIncome = False
    optExpense = False
    
    If GLCompany.SubDigits = 0 Then
       tdbBranch.Enabled = False
    End If
    
    ' populate the AcctType Combo Box
    cmbTypes.AddItem "All Types"
    For i = 1 To 16
        cmbTypes.AddItem glTypeChar(i) & " " & glTypeName(i)
    Next i
    
    ' FilterType
    '   0 = None
    '   1 = AcctType
    '   2 = B I E
    '   3 = Branch
    FilterType = 0
    
'    LoadGrid 1, "T"
    LoadGrid 0, ""
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = "Account +"
    fg.Cell(flexcpFontBold, 0, 0) = True
    fg.Col = 0
    fg.Sort = flexSortNumericAscending
    fg.ShowCell 1, 0
    fg.Select 1, 0
    fg.Editable = flexEDNone
    
    ' sort by acct ascending
    SortCol = 0
    SortType = 0
    
'    Me.cmdFindNext.Enabled = False
    
End Sub

Private Sub CmdExit_Click()
    
    fg.SetFocus   ' set the cursor back for next time
    SelAcct = 0
    Me.Hide

End Sub

Private Sub optAll_Click()
    LoadGrid 0, ""
    
    If fg.Visible Then
       fg.ShowCell 1, 0
       fg.SetFocus
    End If
End Sub

Private Sub optBalSht_Click()
    LoadGrid 2, 0
    fg.ShowCell 1, 0
    fg.SetFocus
End Sub

Private Sub optExpense_Click()
    LoadGrid 2, 2
    fg.ShowCell 1, 0
    fg.SetFocus
End Sub

Private Sub optIncome_Click()
    LoadGrid 2, 1
    fg.ShowCell 1, 0
    fg.SetFocus
End Sub

Private Sub tdbBranch_LostFocus()
    
    Me.cmbTypes.ListIndex = 0
    Me.optAll = True
    Me.tdbNumberFind = 0
    
    If tdbBranch.Value <> 0 Then
       LoadGrid 3, tdbBranch.Value
    Else
       LoadGrid 0, ""
    End If

    fg.ShowCell 1, 0
    fg.SetFocus

End Sub

Private Sub tdbNumberFind_lostfocus()
    
    If IsNull(tdbNumberFind.Value) Then Exit Sub
    If tdbNumberFind.Value = 0 Then Exit Sub
    
    Me.tdbDescFind = ""
    
    rw = fg.FindRow(Me.tdbNumberFind.Value, 0, 0)

    If rw = -1 Then
       MsgBox "Not Found: " & Me.tdbNumberFind, vbExclamation + vbOKOnly
       fg.TopRow = 1
       fg.Select 1, 0
       fg.SetFocus
       Exit Sub
    End If

    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus

End Sub

Private Sub tdbDescFind_lostfocus()
    
    If fg.Rows = 1 Then Exit Sub
    
    If IsNull(tdbDescFind.Text) Then Exit Sub
    If tdbDescFind.Text = "" Then Exit Sub
    
    rw = fg.FindRow(Me.tdbDescFind.Text, 0, 2, False, False)
  
    If rw = -1 Then
       MsgBox "Not Found: " & Me.tdbDescFind.Text, vbExclamation + vbOKOnly
       fg.SetFocus
       fg.TopRow = 1
       fg.Select 1, 0
       Exit Sub
    End If

    fg.TopRow = rw
    fg.Select rw, 1
    Me.cmdFindNext.Enabled = True
    fg.SetFocus

End Sub

Private Sub cmdFindNext_Click()

    If Me.tdbDescFind.Text = "" Then Exit Sub

    rw = fg.FindRow(Me.tdbDescFind.Text, fg.Row + 1, 2, False, False)
  
    If rw = -1 Then
       MsgBox "Not Found: " & Me.tdbDescFind.Text, vbExclamation + vbOKOnly
       fg.SetFocus
       fg.TopRow = 1
       fg.Select 1, 0
       Exit Sub
    End If

    fg.SetFocus
    fg.TopRow = rw
    fg.Select rw, 1
    Me.cmdFindNext.Enabled = True


End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x1 As Single, ByVal y1 As Single, Cancel As Boolean)

    ' clicking on a column header sorts based on that column
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then

       If fg.MouseCol = SortCol Then
          
          ' toggle the sort order
          fg.Col = SortCol
          If SortType = 0 Then    ' switch from ascending to descending
             SortType = 1
             fg.Sort = flexSortGenericDescending
          Else                    ' switch from desc to ascending
             SortType = 0
             fg.Sort = flexSortGenericAscending
          End If
       
       Else       ' switch the sort column
       
          SortType = 0
          SortCol = fg.MouseCol
          fg.Col = SortCol
          fg.Sort = flexSortGenericAscending
       
       End If
       
       ' set the column header
       If SortType = 0 Then
          Y = " +"
       Else
          Y = " -"
       End If
       
       For i = 0 To 2
           
           If i = 0 Then x = "Account"
           If i = 1 Then x = "Type"
           If i = 2 Then x = "Description"
           
           If i = SortCol Then
              fg.TextMatrix(0, i) = x & Y
              fg.Cell(flexcpFontBold, 0, i) = True
           Else
              fg.TextMatrix(0, i) = x
              fg.Cell(flexcpFontBold, 0, i) = False
           End If
       
       Next i
       
       fg.ShowCell 1, SortCol
       
    End If
    
End Sub

Private Function GetSQLString() As String
    
Dim aa As Integer
    
' set the SQL string
'    x = "SELECT [Number],[Description] " & _
'        "FROM GLDescriptions ORDER BY [Number] DESC"

    GetSQLString = "SELECT"
    For aa = 0 To UBound(dbFields, 1)
        GetSQLString = GetSQLString & " [" & dbFields(aa) & "]"
        If aa <> UBound(dbFields, 1) Then GetSQLString = GetSQLString & ","
        GetSQLString = GetSQLString & " "
    Next aa
    
    GetSQLString = GetSQLString & "FROM " & dbFileName & " ORDER BY [" & dbFields(dbSortCol) & "]"
    If dbSortDesc Then
       GetSQLString = GetSQLString & " DESC"
    End If

End Function

Private Sub LoadGrid(ByVal FType As Byte, ByVal FValue As String)

Dim xCt As Long

    ' make sure records exist
    j = xdbAcct.UpperBound(1)
    xCt = 0
    For i = 1 To j
        Select Case FType
           Case 1    ' filter by acct type
              If xdbAcct(i, 1) <> FValue Then GoTo NxtI1
              
           Case 2    ' filter by section
              If xdbAcct(i, 4) <> FValue Then GoTo NxtI1
                   
           Case 3    ' filter by branch
              If xdbAcct(i, 3) <> CStr(FValue) Then GoTo NxtI1
        
        End Select
        xCt = xCt + 1
NxtI1:
    Next i

    If xCt = 0 Then
       MsgBox "No records found !!!", vbExclamation + vbOKOnly, "GL Account Search"
       GoTo LeaveIt
    End If

    fg.Clear flexClearScrollable, flexClearData
    fg.Clear flexClearScrollable, flexClearText
    fg.Rows = 1
    
    j = xdbAcct.UpperBound(1)
    For i = 1 To j
        Select Case FType
           Case 1    ' filter by acct type
              If xdbAcct(i, 1) <> FValue Then GoTo NxtI
              
           Case 2    ' filter by section
              If xdbAcct(i, 4) <> FValue Then GoTo NxtI
                   
           Case 3    ' filter by branch
              If xdbAcct(i, 3) <> CStr(FValue) Then GoTo NxtI
        
        End Select
        fg.AddItem xdbAcct.Value(i, 0) & vbTab & xdbAcct.Value(i, 1) & vbTab & xdbAcct.Value(i, 2)
NxtI:
    Next i
    
LeaveIt:
    fg.ShowCell 1, 0
    fg.Select 1, 0

End Sub

