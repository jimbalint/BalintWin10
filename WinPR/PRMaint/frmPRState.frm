VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPRState 
   Caption         =   "State File Maintenance"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
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
   ScaleHeight     =   7275
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   10200
      TabIndex        =   3
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   615
      Left            =   10200
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   10200
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9615
      _cx             =   16960
      _cy             =   11456
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
End
Attribute VB_Name = "frmPRState"
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
Dim dbFields(3) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte

Dim SQLStr As String


Private Sub Form_Load()
    
    ' get rid of nulls
    rsInit "DELETE * FROM PRState WHERE Isnull(StateAbbrev)", cnDes, rs
    rsInit "DELETE * FROM PRState WHERE Isnull(StateName)", cnDes, rs
    
    ' set the constants for the file
    dbFileName = "PRState"
    dbFields(0) = "StateID"
    dbFields(1) = "StateAbbrev"
    dbFields(2) = "StateName"
    dbFields(3) = "UnEmpMax"
    dbSortCol = 1
    dbSortDesc = False
    
    GetSQLString
    rsInit GetSQLString, cnDes, rs
    SetGrid rs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 0
    fg.ColWidth(1) = 2000
    fg.ColWidth(2) = 3500
    fg.ColWidth(3) = 1500
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.HighLight = flexHighlightNever                   ' don't select ranges
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 1) = dbFields(1) & "+"
    fg.Cell(flexcpFontBold, 0, 1) = True
    
    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    
    fg.AllowSelection = False
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub


Private Sub cmdAdd_Click()
    AddAdo rs, fg
End Sub


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrID As Long
    
    CurrID = fg.TextMatrix(fg.Row, 0)
        
    rs.Close
    rsInit GetSQLString, cnDes, rs
    Set fg.DataSource = rs.DataSource
       
    rw = fg.FindRow(CurrID, 0, 0)
       
    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus
    
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 1 Then     ' validates for number - must enter a value that does not already exist
       
       If fg.EditText = "" Or fg.EditText = "0" Then
          MsgBox "Enter a value!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
       rw = fg.FindRow(UCase(fg.EditText), 0, 1)
       If rw <> -1 Then
          MsgBox "State already exists!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
    End If
    
    If Col = 1 Or Col = 2 Then
        fg.EditText = Trim(UCase(fg.EditText))
    End If

End Sub

Private Sub cmdExit_Click()
    
    GoBack

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
       
       rsInit GetSQLString, cnDes, rs
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
    
    GetSQLString = GetSQLString & "FROM " & dbFileName & " ORDER BY [" & dbFields(dbSortCol) & "]"
    If dbSortDesc Then
       GetSQLString = GetSQLString & " DESC"
    End If

End Function


