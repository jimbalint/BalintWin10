VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUsers 
   Caption         =   "General Ledger User Maintenance"
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
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      _cx             =   18018
      _cy             =   10186
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
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private CnDes As ADODB.Connection

Private mrs As ADODB.Recordset
Dim x As String
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending

Dim dbFileName As String
Dim dbFields(2) As String
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

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrLogon As String
    
    CurrLogon = fg.TextMatrix(fg.Row, 0)
        
    mrs.Close
    rsInit GetSQLString, cnDes, mrs
    Set fg.DataSource = mrs.DataSource
       
    rw = fg.FindRow(CurrLogon, 0, 0)
       
    fg.TopRow = rw
    fg.Select rw, 0
    fg.SetFocus
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Then     ' validates for number - must enter a value that does not already exist
       
       If fg.EditText = "" Then
          MsgBox "Enter a value!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
       rw = fg.FindRow(fg.EditText, 0, 0)
       If rw <> -1 Then
          MsgBox "User already exists!", vbExclamation + vbOKOnly
          Cancel = True
       End If
    
    End If

End Sub

Private Sub Form_Load()
    
    ' get rid of nulls
    rsInit "DELETE * FROM Users WHERE IsNull(Logon)", cnDes, mrs
    rsInit "DELETE * FROM Users WHERE Logon=''", cnDes, mrs
    
    ' set the constants for the file
    dbFileName = "Users"
    dbFields(0) = "Logon"
    dbFields(1) = "Name"
    dbFields(2) = "Password"
    dbSortCol = 0
    dbSortDesc = False
    
    rsInit GetSQLString, cnDes, mrs
    SetGrid mrs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 1300
    fg.ColWidth(1) = 5000
    fg.ColWidth(2) = 1300
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.HighLight = flexHighlightNever                   ' don't select ranges
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = dbFields(0) & "+"
    fg.Cell(flexcpFontBold, 0, 0) = True
    
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
    ' DelAdo mrs, fg, fg.TextMatrix(fg.Row, 0)
    DelAdo mrs, fg
    
    If rw = fg.Rows Then rw = fg.Rows - 1
    
    fg.Select rw, 0
    fg.ShowCell rw, 0

End Sub


Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

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
