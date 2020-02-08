VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAccount 
   Caption         =   " Accounts"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DEL"
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   9240
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdSortDesc 
      Caption         =   "Desc."
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdSortType 
      Caption         =   "Type"
      Height          =   495
      Left            =   9240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdSortAcct 
      Caption         =   "Acct."
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6135
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _cx             =   9340
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
      FocusRect       =   3
      HighLight       =   2
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
      FormatString    =   $"frmAccount.frx":0000
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
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdUnccheck 
      Caption         =   "&UnCheck All"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check All"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   5640
      Width           =   735
   End
   Begin VB.ListBox lstType 
      Height          =   3660
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "FILE"
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SORT"
      Height          =   255
      Left            =   9240
      TabIndex        =   12
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mra As ADODB.Recordset
Private GotFocusColor As Long
Private LostFocusColor As Long

Private Sub cmdAdd_Click()

'    fg.Rows = fg.Rows + 1
'    fg.Row = fg.Rows - 1
'    fg.Col = 1
'    fg.TextMatrix(fg.Row, 0) = "0"
'    fg.TextMatrix(fg.Row, 1) = "4444"
'    fg.TextMatrix(fg.Row, 2) = ""
'    fg.ShowCell fg.Row, 1
'    fg.SetFocus
    
    Dim mradd As ADODB.Recordset
    SetAdo cn, mradd, "Select * from glAccount"
    
    mradd!AcctType = "0"
    mradd!Account = 0
    mradd!Description = "New Account"
    mradd!AllSchedules = False
    mradd!AllStatements = False
    mradd!BranchAcct = False
    mradd!BSColumn = 0
    mradd!ConsAcct = False
    mradd!Date1 = 0
    mradd!Date2 = 0
    mradd!DollarSign = False
    mradd!LineFeeds = 0
    mradd!PrintTab = 0
    mradd!SignRevSched = False
    mradd!SignRevStmt = False
    mradd!TotalLevel = 0
    mradd!TotalOnLedger = False
    mradd.AddNew
    mradd.Update
    mradd.Close
    
    fg.Rows = fg.Rows + 1
    fg.Row = fg.Rows - 1
    fg.TextMatrix(fg.Row, 0) = mra!AcctType
    fg.TextMatrix(fg.Row, 1) = mra!Account
    fg.TextMatrix(fg.Row, 2) = mra!Description
    fg.AutoSize (2)
    fg.Col = 1
    fg.ShowCell fg.Row, 1
    fg.SetFocus
'    cmdEdit_Click
End Sub

Private Sub cmdCheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstType.ListCount - 1
        lstType.Selected(ndx) = True
    Next ndx
End Sub

Private Sub cmdEdit_Click()
    If glConnect() Then
        frmAccountForm.strAccount = fg.TextMatrix(fg.Row, 1)
        frmAccountForm.Show vbModal, Me
        fg.TextMatrix(fg.Row, 0) = mra!AcctType
        fg.TextMatrix(fg.Row, 1) = mra!Account
        fg.TextMatrix(fg.Row, 2) = mra!Description
        fg.AutoSize (2)
    End If
    fg.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    LoadGrid
    fg.SetFocus
End Sub

Private Sub cmdSortAcct_Click()
    fg.Select 1, 1
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdSortDesc_Click()
    fg.Select 1, 2
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdSortType_Click()
    fg.Select 1, 0
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdUnccheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstType.ListCount - 1
        lstType.Selected(ndx) = False
    Next ndx
End Sub

Private Sub fg_EnterCell()
    If fg.Row = 0 Then Exit Sub
    If IsNumeric(fg.TextMatrix(fg.Row, 1)) Then
        If glConnect() Then
            SetAdo cn, mra, "Select * from glAccount where Account = " & fg.TextMatrix(fg.Row, 1)
        Else
        End If
    Else
    End If
End Sub

Private Sub fg_GotFocus()
    fg.BackColor = GotFocusColor
End Sub

Private Sub fg_LostFocus()
    fg.BackColor = LostFocusColor
End Sub

Private Sub Form_Load()
    frmAccount.Caption = " Accounts for " & glCompanyName
    Dim ndx As Integer
    For ndx = 0 To 15
        lstType.AddItem glTypeChar(ndx) & "   " & glTypeName(ndx)
    Next ndx
    cmdCheck_Click
    LoadGrid
    LostFocusColor = frmAccount.BackColor
    GotFocusColor = fg.BackColor
    lstType.BackColor = LostFocusColor
End Sub

Private Sub LoadGrid()
    On Error GoTo glErr
    Dim mrs As ADODB.Recordset
    If glConnect() Then
        SetAdo cn, mrs, "Select * from glAccount"
    End If
    fg.Redraw = False
    fg.Rows = 1
    fg.Cols = 3
    fg.ColDataType(2) = flexDTString
    mrs.MoveFirst
    While mrs.EOF = False
        If lstType.Selected(glTypeByte(mrs!AcctType)) = True Then
            fg.Rows = fg.Rows + 1
            fg.TextMatrix(fg.Rows - 1, 0) = mrs!AcctType
            fg.TextMatrix(fg.Rows - 1, 1) = mrs!Account
            fg.TextMatrix(fg.Rows - 1, 2) = mrs!Description
        End If
        mrs.MoveNext
    Wend
    fg.AutoSize (2)
    fg.Row = 1
    fg.Col = 2
    fg.Redraw = True
    mrs.Close
glErr:
End Sub

Private Sub lstType_GotFocus()
    lstType.BackColor = GotFocusColor
End Sub

Private Sub lstType_LostFocus()
    lstType.BackColor = LostFocusColor
End Sub
