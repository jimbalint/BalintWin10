VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmContactType 
   Caption         =   "Contact Type Maintenance"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      _cx             =   8281
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
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
      FixedCols       =   0
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
Attribute VB_Name = "frmContactType"
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

' set to 1 if form used for lookup
Public LookUp As Byte
Public SelectedNumber As Long
Public SelectedDescription As String

Private Sub Form_Load()

    ' get rid of nulls
    rsInit "DELETE * FROM PRGlobal WHERE TypeCode = 0", cnDes, mrs
    rsInit "DELETE * FROM PRGlobal WHERE IsNull(Description)", cnDes, mrs
    
    ' set the constants for the file
    dbFileName = "PRGlobal"
    dbFields(0) = "Description"
    dbFields(1) = "TypeCode"
    dbSortCol = 0
    dbSortDesc = False
    
    x = "SELECT [Description], [TypeCode] FROM PRGlobal WHERE " & _
        "[TypeCode] = " & CStr(PREquate.GlobalTypeContact) & " " & _
        "ORDER BY [Description]"
    
    rsInit x, cnDes, mrs
    SetGrid mrs, fg
    
    ' customize the grid
    fg.ColWidth(0) = 4500
    
    ' don't show PRGlobal.TypeCode
    fg.ColWidth(1) = 0
    
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fg.TabBehavior = flexTabCells                       ' tab moves between cells
    fg.HighLight = flexHighlightNever                   ' don't select ranges
    
    ' sorted by first col ascending
    fg.TextMatrix(0, 0) = dbFields(0) & "+"
    fg.Cell(flexcpFontBold, 0, 0) = True
    
    ' not for lookup - hide the select button
    If LookUp = 0 Then
       
       cmdSelect.Visible = False
    
    Else
       
       SelectedNumber = 0
       SelectedDescription = ""
       
        ' *** allow edits and deletes in lookup mode
        ' add/delete on the fly - fix later ....
        '   issue with Editable property
        ' Me.cmdAdd.Visible = False
        ' Me.cmdDelete.Visible = False
        ' fg.Editable = flexEDNone
    
    End If
       
    ' add/delete on the fly - fix later ....
    '   issue with Editable property
    ' Me.cmdAdd.Visible = False
    ' Me.cmdDelete.Visible = False

End Sub

Private Sub cmdExit_Click()
        
    ' don't leave if being use for lookup
    If LookUp = 0 Then
       ' GoBack
        End
    Else
        Me.Hide
    End If

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
        
    If fg.Rows = 1 Then Exit Sub
    
    ' what if no records left ????
        
    DelConfirm = MsgBox(fg.TextMatrix(fg.Row, 0), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")
    
    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If
    
    rw = fg.Row
    DelAdo mrs, fg
    
    If rw = fg.Rows Then rw = fg.Rows - 1
    
    fg.Select rw, 0
    fg.ShowCell rw, 0

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

Private Sub cmdAdd_Click()
    AddAdo mrs, fg
    fg.Cell(flexcpText, 0, 1, fg.Rows - 1, 1) = CStr(PREquate.GlobalTypeContact)
End Sub

Private Sub cmdSelect_Click()
    fg_DblClick
End Sub

Private Sub fg_DblClick()

    If LookUp = 1 Then
       SelectedDescription = fg.TextMatrix(fg.Row, 0)
       Me.Hide
    End If

End Sub

