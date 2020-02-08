VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmDescriptions 
   Caption         =   " General Ledger Descriptions for ALL Clients"
   ClientHeight    =   7380
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "N&ext"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin TDBText6Ctl.TDBText tdbDescFind 
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   360
      Width           =   4215
      _Version        =   65536
      _ExtentX        =   7435
      _ExtentY        =   1085
      Caption         =   "Descriptions.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Descriptions.frx":007C
      Key             =   "Descriptions.frx":009A
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
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   1085
      Calculator      =   "Descriptions.frx":00DE
      Caption         =   "Descriptions.frx":00FE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Descriptions.frx":0170
      Keys            =   "Descriptions.frx":018E
      Spin            =   "Descriptions.frx":01D8
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
      ValueVT         =   27262977
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
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
      Left            =   4620
      TabIndex        =   3
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   6600
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      _cx             =   18018
      _cy             =   9128
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
      Left            =   7320
      TabIndex        =   2
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "frmDescriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcn As ADODB.Connection
Private mrs As ADODB.Recordset
Dim x As String
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending

Dim dbFileName As String
Dim dbFields(1) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte

Dim SQLStr As String

' To Do:
' blank file test
' search field select
' sort by header change - cell picture set

Private Sub cmdAdd_Click()
    AddAdo mrs, fg
End Sub


Private Sub cmdSort_Click()
    mrs.Close
    Set fg.DataSource = Nothing
    SetAdo cnDes, mrs, x
    SetGrid mrs, fg
    fg.Row = 1
    fg.SetFocus

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
    
    CNDesOpen ("\Balint\Data\GLSystem.mdb")
    
    ' get rid of nulls
    rsInit "DELETE * FROM glDescriptions WHERE Isnull(Number)", cnDes, mrs
    rsInit "DELETE * FROM glDescriptions WHERE Number = 0", cnDes, mrs
    
    ' set the constants for the file
    dbFileName = "GLDescriptions"
    dbFields(0) = "Number"
    dbFields(1) = "Description"
    dbSortCol = 0
    dbSortDesc = False
    
    ' set the SQL string
    x = "SELECT [Number],[Description] " & _
        "FROM GLDescriptions WHERE Number > 0 ORDER BY [Number]"
    
    rsInit GetSQLString, cnDes, mrs
    SetGrid mrs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 1300
    fg.ColWidth(1) = 8500
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.HighLight = flexHighlightNever                   ' don't select ranges
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = dbFields(0) & "+"
    fg.Cell(flexcpFontBold, 0, 0) = True
    
    Me.cmdFindNext.Enabled = False
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
    
    If fg.Rows = 1 Then Exit Sub
    
    ' what if no records left ????
        
    DelConfirm = MsgBox(fg.TextMatrix(fg.Row, 0) & vbCr & fg.TextMatrix(fg.Row, 1), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")
    
    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If
    
    
    rw = fg.Row
    DelAdo mrs, fg, fg.TextMatrix(fg.Row, 0)
    
    If rw = fg.Rows Then rw = fg.Rows - 1
    
    fg.Select rw, 0
    fg.ShowCell rw, 0

End Sub

Private Sub tdbNumberFind_lostfocus()
    
    If fg.Rows = 1 Then Exit Sub
    
    If IsNull(tdbNumberFind.Value) Then Exit Sub
    
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
    
    rw = fg.FindRow(Me.tdbDescFind.Text, 0, 1, False, False)
  
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

    rw = fg.FindRow(Me.tdbDescFind.Text, fg.Row + 1, 1, False, False)
  
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

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)

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
    
       mrs.Close
       
       rsInit GetSQLString, cnDes, mrs
       Set fg.DataSource = mrs.DataSource
       
       fg.ShowCell 1, 0

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
