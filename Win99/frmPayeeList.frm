VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPayeeList 
   Caption         =   "1099 Payee List"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayeeList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewYear 
      Caption         =   "&Init Tax Year"
      Height          =   615
      Left            =   12000
      TabIndex        =   7
      Top             =   7920
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   8415
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   11175
      _cx             =   19711
      _cy             =   14843
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
      Caption         =   "&PRINT LIST"
      Height          =   615
      Left            =   12000
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   615
      Left            =   12000
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      Height          =   615
      Left            =   12000
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   12000
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   12000
      TabIndex        =   0
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
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
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmPayeeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean

Dim dbFileName As String
Dim dbFields(3) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte
Dim SelID As Long

Dim Rw As Long

Private rs As ADODB.Recordset
Public PayeeID As Long

Private Sub cmdAdd_Click()
    
    frmPayeeEdit.ScreenMode = 1
    frmPayeeEdit.PayeeID = 0
    frmPayeeEdit.Show vbModal
    
    rs.Close
    rsInit GetSQLString, cn, rs
    Set fg.DataSource = rs.DataSource
    
    rs.Find "PayeeID = " & PayeeID, 0, adSearchForward, 1
    If rs.EOF = False Then
        Rw = fg.FindRow(PayeeID, 0, 0)
        fg.TopRow = Rw
        fg.Select Rw, 0
    Else
        If fg.Rows > 0 And Rw > 0 Then
            If Rw > fg.Rows - 1 Then Rw = fg.Rows - 1
            fg.TopRow = Rw
            fg.Select Rw, 0
        
            If Rw = 1 Then rs.MoveFirst     ' ???
        
        End If
    End If
    
    fg.SetFocus
End Sub

Private Sub cmdEdit_Click()
    
    If fg.Rows = 1 Then Exit Sub

    Rw = 0
    If fg.Rows > 0 Then
        Rw = fg.Row
    End If
     
    SelID = rs!PayeeID
    frmPayeeEdit.PayeeID = SelID
    frmPayeeEdit.ScreenMode = 2
    frmPayeeEdit.Show vbModal
    
    rs.Close
    rsInit GetSQLString, cn, rs
    Set fg.DataSource = rs.DataSource
       
    rs.Find "PayeeID = " & SelID, 0, adSearchForward, 1
    If rs.EOF = False Then
        Rw = fg.FindRow(SelID, 0, 0)
        fg.TopRow = Rw
        fg.Select Rw, 0
    Else
        If fg.Rows > 0 And Rw > 0 Then
            If Rw > fg.Rows - 1 Then Rw = fg.Rows - 1
            fg.TopRow = Rw
            fg.Select Rw, 0
        
            If Rw = 1 Then rs.MoveFirst     ' ???
        
        End If
    End If
    
    fg.SetFocus

End Sub

Private Sub cmdNewYear_Click()
    Dim txyr
    txyr = InputBox("Enter tax year to init for")
    If txyr = "" Then Exit Sub
    
    ' 2022-01-10 clear forms first
    ClearFormRecs "NEC", txyr
    ClearFormRecs "MISC", txyr
    ClearFormRecs "R", txyr
    ClearFormRecs "INT", txyr
    ClearFormRecs "DIV", txyr
    ClearFormRecs "1096", txyr
    
'SQLString = "SELECT distinct(FormType) as FormType from Form99 where TaxYear = 2021"
'rsInit SQLString, cn99, rs
'Do While Not rs.EOF
'    MsgBox (rs("FormType"))
'    rs.MoveNext
'Loop
'rs.Close
'End

    
    CopyForms CInt(txyr)
    
    If txyr = "2020" Then
        Create2020Forms ("MISC")
        Create2020Forms ("NEC")
    End If
    
    If txyr = "2021" Then
        Create2021Forms ("NEC")
    End If
    
End Sub

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.KeyPreview = True

    GetPayeeData

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub GetPayeeData()

    ' get rid of nulls
    SQLString = "DELETE * FROM Payee99 WHERE PayeeNumber = 0"
    ' rsInit SQLString, cn, rs
    On Error Resume Next
    cn.Execute SQLString
    On Error GoTo 0
    
    SQLString = "DELETE * FROM Payee99 WHERE IsNull(PayeeNumber)"
    ' rsInit SQLString, cn, rs
    On Error Resume Next
    cn.Execute SQLString
    On Error GoTo 0
    
    dbFileName = "Payee99"
    dbFields(0) = "PayeeID"
    dbFields(1) = "PayeeNumber"
    dbFields(2) = "PayeeName"
    dbFields(3) = "Inactive"
    dbSortCol = 1

    X = GetSQLString
    
    rsInit X, cn, rs
    
'    If mod99Global.NewADO Then
'        rs.MoveFirst
'        Do While Not rs.EOF
'            rs!FederalID = RC4Decrypt(rs!FederalID, rc4Key)
'            rs.Update
'            rs.MoveNext
'        Loop
'    End If
    
    SetGrid rs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 0
    fg.ColWidth(1) = 1800
    fg.ColWidth(2) = 6000
    fg.ColWidth(3) = 2000
    
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
       
    Rw = fg.FindRow(CurrID, 0, 0)
       
    fg.TopRow = Rw
    fg.Select Rw, 0
    fg.SetFocus
    
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Then     ' validates for number - must enter a value that does not already exist
       
       If fg.EditText = "" Or fg.EditText = "0" Then
          MsgBox "Enter a value!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
       Rw = fg.FindRow(fg.EditText, 0, 0)
       If Rw <> -1 Then
          MsgBox "Number already exists!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
    End If

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
Dim trs As New ADODB.Recordset

    If fg.Rows = 1 Then Exit Sub

    ' what if no records left ????

    DelConfirm = MsgBox(fg.TextMatrix(fg.Row, 1) & vbCr & Trim(fg.TextMatrix(fg.Row, 2)) & ", " & Trim(fg.TextMatrix(fg.Row, 3)), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")

    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If

    ' delete records from related files
    SQLString = "DELETE * FROM Payee99 WHERE PayeeID = " & fg.TextMatrix(fg.Row, 0)
    cn.Execute SQLString

    Rw = fg.Row
    ' DelAdo rs, fg, fg.TextMatrix(fg.Row, 0)
    DelAdo rs, fg

    If Rw = fg.Rows Then Rw = fg.Rows - 1

    fg.Select Rw, 0
    fg.ShowCell Rw, 0

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
    
    GetSQLString = Trim(GetSQLString) & " ORDER BY [" & dbFields(dbSortCol) & "]"
    
    If dbSortDesc Then
       GetSQLString = GetSQLString & " DESC"
    End If

End Function


