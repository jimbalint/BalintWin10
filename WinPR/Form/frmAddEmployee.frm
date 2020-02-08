VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAddEmployee 
   Caption         =   "Select Employee to add"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
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
   ScaleHeight     =   7440
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optInactive 
      Caption         =   "&INACTIVE"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optActive 
      Caption         =   "A&CTIVE"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame frmEmpSel 
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   6255
      Begin VB.OptionButton optAll 
         Caption         =   "&ALL"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   8640
      TabIndex        =   3
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Default         =   -1  'True
      Height          =   1455
      Left            =   8640
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5655
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   8055
      _cx             =   14208
      _cy             =   9975
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmAddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Public EmpID As Long
Dim InitFlag As Boolean

Private Sub Form_Load()

    InitFlag = True

    Me.optAll = True
    Me.lblCompanyName = PRCompany.Name

    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    Me.fg.AllowSelection = False

    InitFlag = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    EmpID = -1
    SelReturn
End Sub

Private Sub SelReturn()
    rs.Close
    Me.Hide
End Sub

Private Sub fg_DblClick()
    EmpID = rs!EmployeeID
    SelReturn
End Sub
Private Sub cmdSelect_Click()
    If rs.RecordCount = 0 Then Exit Sub
    EmpID = rs!EmployeeID
    SelReturn
End Sub

Public Sub Init()

    rs.CursorLocation = adUseClient
    
    SQLString = "SELECT EmployeeID, EmployeeNumber, LastName, FirstName FROM PREmployee"
    If Me.optActive = True Then
        SQLString = Trim(SQLString) & " WHERE Inactive = 0"
    ElseIf Me.optInactive = True Then
        SQLString = Trim(SQLString) & " WHERE Inactive = 1"
    End If
    SQLString = Trim(SQLString) & " ORDER BY EmployeeNumber"
    
    rsInit SQLString, cn, rs
    
    If rs.RecordCount = 0 Then
        MsgBox "No Employees Found!", vbExclamation
    End If

    ' set the grid
    SetGrid rs, fg
    
    fg.ScrollBars = flexScrollBarVertical
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    
    fg.ColHidden(0) = True
    fg.ColWidth(1) = 1700
    fg.ColWidth(2) = 3000
    fg.ColWidth(3) = 3000

End Sub

Private Sub optAll_Click()
    If InitFlag Then Exit Sub
    Me.optActive = False
    Me.optInactive = False
    rs.Close
    Init
End Sub
Private Sub optActive_Click()
    Me.optAll = False
    Me.optInactive = False
    rs.Close
    Init
End Sub
Private Sub optInActive_Click()
    Me.optAll = False
    Me.optActive = False
    rs.Close
    Init
End Sub


