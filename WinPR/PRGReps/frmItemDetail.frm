VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmItemDetail 
   Caption         =   "Item Detail Report"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10500
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTotalsOnly 
      Caption         =   "Totals Only"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CheckBox chkRecall 
      Caption         =   "Use Recall Date if entered"
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   8640
      Width           =   2655
   End
   Begin VB.CheckBox chkAnniv 
      Caption         =   "Rollover on hired anniv"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   8040
      Width           =   2655
   End
   Begin VB.CheckBox chkPrintSS 
      Caption         =   "Print SS Number"
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   9000
      Width           =   3015
   End
   Begin TDBNumber6Ctl.TDBNumber tdbMaxPct 
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   8040
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   661
      Calculator      =   "frmItemDetail.frx":0000
      Caption         =   "frmItemDetail.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmItemDetail.frx":0094
      Keys            =   "frmItemDetail.frx":00B2
      Spin            =   "frmItemDetail.frx":00FC
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
      HighlightText   =   0
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
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CheckBox chkMatching 
      Caption         =   "Matching Contribution"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   7440
      Width           =   2535
   End
   Begin VB.CheckBox chkNoGross 
      Caption         =   "Don't display Gross Wage"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   7440
      Width           =   2895
   End
   Begin VB.CheckBox chkShowRemain 
      Caption         =   "Show Remaining Value"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "&DATE RANGE"
      Height          =   615
      Left            =   1523
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtDisplay 
      Height          =   615
      Left            =   2723
      TabIndex        =   32
      Top             =   600
      Width           =   5775
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   5483
      TabIndex        =   21
      Top             =   2160
      Width           =   3495
      Begin VB.CommandButton cmdSelDed 
         Caption         =   "Select &All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   650
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearDed 
         Caption         =   "C&lear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   640
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "OE && Deduction Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   1950
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Please Select Up To FIVE (5)  Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   2940
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   735
      Left            =   960
      TabIndex        =   28
      Top             =   1440
      Width           =   5535
      Begin VB.OptionButton optName 
         Caption         =   "&Name"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optEmpNo 
         BackColor       =   &H80000016&
         Caption         =   "&Employee Number"
         Height          =   245
         Left            =   240
         TabIndex        =   1
         Top             =   370
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optChkDate 
         BackColor       =   &H80000016&
         Caption         =   "Check &Date"
         Height          =   245
         Left            =   2520
         TabIndex        =   2
         Top             =   370
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Order By"
         Height          =   255
         Left            =   1560
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Height          =   855
      Left            =   1043
      TabIndex        =   24
      Top             =   2280
      Width           =   3015
      Begin VB.CommandButton cmdClearAll 
         BackColor       =   &H80000014&
         Caption         =   "&Clear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectAll 
         BackColor       =   &H80000016&
         Caption         =   "&Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Employee Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Height          =   735
      Left            =   7320
      TabIndex        =   23
      Top             =   1320
      Width           =   2295
      Begin VB.Label lblEmpCount 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   200
         Left            =   480
         TabIndex        =   27
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Employees Selected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   120
         Width           =   1740
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   19
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9720
      Width           =   2175
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   3975
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   4335
      _cx             =   7646
      _cy             =   7011
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
   Begin VSFlex8Ctl.VSFlexGrid fgItem 
      Height          =   3975
      Left            =   5160
      TabIndex        =   9
      Top             =   3240
      Width           =   4335
      _cx             =   7646
      _cy             =   7011
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
   Begin TDBNumber6Ctl.TDBNumber tdbMatchPct 
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   8520
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   661
      Calculator      =   "frmItemDetail.frx":0124
      Caption         =   "frmItemDetail.frx":0144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmItemDetail.frx":01BA
      Keys            =   "frmItemDetail.frx":01D8
      Spin            =   "frmItemDetail.frx":0222
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
      HighlightText   =   0
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
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "COMPANY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   220
      Left            =   600
      TabIndex        =   22
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmItemDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsEmp As New ADODB.Recordset
Public rsItem As New ADODB.Recordset

Dim rs As New ADODB.Recordset
Dim rsDedExcl As New ADODB.Recordset
Dim rsd As New ADODB.Recordset

Public PEDate As Long
Public CheckDt As Long
Public EmpCount As Long
Public NoItems As Long

Dim i, j As Long
Dim x, Y, Z As String
Dim MatchID As Long
Dim P1 As Currency
Dim LastEmpName As String
Dim GrossPayTl(2), DedBasisTl(2), DedAmtTl(2), DedMatchTl(2), TotalContrTl(2) As Currency
Dim ItemCount As Long
Dim GlobalID As Long
Dim boo As Boolean

Private Sub Form_Load()
    
    ' screen defaults
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeScreenDefault & _
                " AND Description = 'ItemDetail'" & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeScreenDefault
        PRGlobal.Description = "ItemDetail"
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    GlobalID = PRGlobal.GlobalID
    
    SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeOE & _
                " OR ItemType = " & PREquate.ItemTypeDED
    If PRItem.GetBySQL(SQLString) = False Then
        MsgBox "No OE or DED items found!", vbExclamation
        GoBack
    End If
    
    Me.tdbMaxPct.Format = "##0.00 %"
    Me.tdbMaxPct.DisplayFormat = ""
    Me.tdbMaxPct.Visible = False
    
    Me.tdbMatchPct.Format = "##0.00 %"
    Me.tdbMatchPct.DisplayFormat = ""
    Me.tdbMatchPct.Visible = False
    Me.chkPrintSS.Visible = False
    
    Me.chkAnniv.Visible = False
    Me.chkRecall.Visible = False
    
    Load_Grids
    Me.lblCompanyName.Caption = PRCompany.Name
    
    If PRBatchID > 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "PRBatch Not Found: " & PRBatchID, vbCritical
            End
        End If
        PEDate = PRBatch.PEDate
        CheckDt = PRBatch.CheckDate
        OptDate = " "
        txtDisplay = "Batch: " & PRBatchID & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        RangeType = PREquate.RangeTypeBatch
        Me.optEmpNo = False
        Me.optChkDate = True
    End If
    
    ' defaulted check boxes
    boo = PRGlobal.GetByID(GlobalID)
    Me.chkNoGross = PRGlobal.Byte1
    Me.chkMatching = PRGlobal.Byte2
    Me.chkPrintSS = PRGlobal.Byte3
    Me.chkShowRemain = PRGlobal.Byte4
    Me.chkAnniv = PRGlobal.Byte5
    Me.chkRecall = PRGlobal.Byte6
    Me.chkTotalsOnly = PRGlobal.Byte8
    
    Me.optEmpNo = True
    Me.optChkDate = False
    Me.optName = False
    If PRGlobal.Byte7 = 2 Then Me.optChkDate = True
    If PRGlobal.Byte7 = 3 Then Me.optName = True
        
    Me.KeyPreview = True
    
End Sub

Private Sub fgEmp_LostFocus()
    EmpCount = 0
    rsEmp.MoveFirst
    Do
        If rsEmp!Selected = True Then
            EmpCount = EmpCount + 1
        End If
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
    
End Sub

Private Sub fgItem_LostFocus()
    rsItem.MoveFirst
    NoItems = 0
    Do
        If rsItem!Selected = True Then
            NoItems = NoItems + 1
        End If
        rsItem.MoveNext
    Loop Until rsItem.EOF
    rsItem.MoveFirst
    If NoItems > 5 Then
        MsgBox "Please select ONLY FIVE (5) Items", vbCritical, "Item Detail Report"
        GoBack
    End If

End Sub

Private Sub cmdDateRange_Click()
    frmDateRange.lblProgram = "Item Detail"
    frmDateRange.Show vbModal
        
    If frmDateRange.optCheckDate = True Then
        OptDate = "CHECK DATE"
    ElseIf frmDateRange.optPEDate = True Then
        OptDate = "P/E DATE"
    End If
        
    If InitFlag = False Then Exit Sub   ' user exited
    
    If BatchNumbr > 0 Then
        If Not PRBatch.GetByID(BatchNumbr) Then
            MsgBox "PRBatch Not Found: " & BatchNumbr, vbCritical
            End
        End If
        PEDate = PRBatch.PEDate
        CheckDt = PRBatch.CheckDate
        OptDate = " "
        txtDisplay = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        RangeType = PREquate.RangeTypeBatch

        Me.optChkDate = True
        Me.optEmpNo = False

    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    PRBatchID = BatchNumbr
    Me.Refresh
End Sub

Private Sub cmdClearAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Selected = False
        rsEmp.Update
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
End Sub

Private Sub cmdSelectAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Selected = True
        rsEmp.Update
        rsEmp.MoveNext
        
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
    lblEmpCount = "All Employees"
End Sub

Public Sub Load_Grids()
'  Loop Through Employee File for all Employees
    rsEmp.CursorLocation = adUseClient
    rsEmp.Fields.Append "Selected", adBoolean
    rsEmp.Fields.Append "EmpNo", adDouble
    rsEmp.Fields.Append "EmpName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "EmpID", adDouble
    Me.lblCompanyName.Caption = PRCompany.Name
    
    rsEmp.Open , , adOpenDynamic, adLockOptimistic
    SQLString = "Select * from PREmployee ORDER BY LastName, FirstName"
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employee Records were Found: ", vbCritical
        GoBack
    End If
    
    Do
        rsEmp.AddNew
        rsEmp!Selected = True
        rsEmp!EmpNo = PREmployee.EmployeeNumber
        rsEmp!EmpName = PREmployee.LFName
        rsEmp!EmpID = PREmployee.EmployeeID
        rsEmp.Update
        
        If Not PREmployee.GetNext Then Exit Do
    Loop
    SetGrid rsEmp, fgEmp
    fgEmp.ScrollBars = flexScrollBarVertical
    
'  Loop Through PRItemHist for all Items    **********************************
    rsItem.CursorLocation = adUseClient
    rsItem.Fields.Append "Selected", adBoolean
    rsItem.Fields.Append "Type", adVarChar, 20
    rsItem.Fields.Append "Description", adVarChar, 80, adFldIsNullable
    rsItem.Fields.Append "ItemID", adDouble
    rsItem.Fields.Append "IsItHours", adBoolean
    rsItem.Fields.Append "MaxAmount", adCurrency
    
    rsItem.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " OR PRItem.itemtype = " & PREquate.ItemTypeSDTax & ")" & _
                " ORDER BY PRItem.ItemType, PRItem.ItemID"

    If Not PRItem.GetBySQL(SQLString) Then
'        Me.fgitem.Visible = False
    Else
        Do
            rsItem.AddNew
            rsItem!Selected = False
            rsItem.Fields("Type") = PRItem.ItemType
            rsItem.Fields("ItemID") = Trim(PRItem.ItemID)
            rsItem.Fields("Description") = PRItem.Abbreviation
            
            ' selected last time?
            If String_To_Long(PRGlobal.Var1, "A") = PRItem.ItemID Then rsItem!Selected = True
            If String_To_Long(PRGlobal.Var2, "A") = PRItem.ItemID Then rsItem!Selected = True
            If String_To_Long(PRGlobal.Var3, "A") = PRItem.ItemID Then rsItem!Selected = True
            If String_To_Long(PRGlobal.Var4, "A") = PRItem.ItemID Then rsItem!Selected = True
            If String_To_Long(PRGlobal.Var5, "A") = PRItem.ItemID Then rsItem!Selected = True
            
            ' HOURS
            rsItem.Update
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                rsItem.AddNew
                rsItem!Selected = False
                rsItem.Fields("Type") = 6
                rsItem.Fields("ItemID") = Trim(PRItem.ItemID)
                rsItem.Fields("Description") = PRItem.Abbreviation
                rsItem.Fields("IsItHours") = True
            
                ' selected last time?
                If String_To_Long(PRGlobal.Var1, "H") = PRItem.ItemID Then rsItem!Selected = True
                If String_To_Long(PRGlobal.Var2, "H") = PRItem.ItemID Then rsItem!Selected = True
                If String_To_Long(PRGlobal.Var3, "H") = PRItem.ItemID Then rsItem!Selected = True
                If String_To_Long(PRGlobal.Var4, "H") = PRItem.ItemID Then rsItem!Selected = True
                If String_To_Long(PRGlobal.Var5, "H") = PRItem.ItemID Then rsItem!Selected = True
                
                rsItem.Update
            
            End If
            rsItem.Fields("MaxAmount") = PRItem.MaxAmount
            If Not PRItem.GetNext Then Exit Do
        Loop

    End If
    
    frmDateRange.lblClient = PRCompany.Name
    Me.KeyPreview = True
    SetGrid rsItem, fgItem
    fgItem.ScrollBars = flexScrollBarVertical
    fgItem.ColComboList(1) = "|#3;OTH EARN|#4;DEDUCT|#6;OTH HRS|#5;SD TAX"
    fgItem.ColWidth(1) = 1000
    fgItem.ColWidth(2) = 2800
    
End Sub

Private Sub cmdClearDed_Click()
    rsItem.MoveFirst
    Do
        rsItem!Selected = False
        rsItem.Update
        rsItem.MoveNext
    Loop Until rsItem.EOF
    rsItem.MoveFirst
End Sub

Private Sub cmdOK_Click()
    
Dim DedCt As Long
Dim ItmCt As Byte
    
    fgEmp_LostFocus
    fgItem_LostFocus
    If chkShowRemain = 1 Then
        If NoItems > 1 Then
            MsgBox "Please select ONLY ONE Item", vbExclamation, "Item Detail Report"
            Exit Sub
        End If
    End If
    
    ' only one deduction can be chosen for this
    If Me.chkMatching = 1 Then
        DedCt = 0
        rsItem.MoveFirst
        Do
            If rsItem!Selected = True And rsItem!Type <> PREquate.ItemTypeDED Then
                MsgBox "Deductions only for the match option!", vbExclamation
                Exit Sub
            End If
            If rsItem!Selected = True And rsItem!Type = PREquate.ItemTypeDED Then
                DedCt = DedCt + 1
                MatchID = rsItem!ItemID
            End If
            rsItem.MoveNext
        Loop Until rsItem.EOF
        If DedCt <> 1 Then
            MsgBox "You must pick ONE deduction for the match option!", vbExclamation
            Exit Sub
        End If
    End If
        
    If PRBatchID = 0 And StartDate = 0 And EndDate = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbExclamation, "Item Detail Report"
        Exit Sub
    End If
    
    If EmpCount = 0 Then
        MsgBox "Please select AT LEAST ONE Employee", vbExclamation, "Item Detail Report"
        Exit Sub
    End If
    
    If NoItems = 0 Then
        MsgBox "Please select AT LEAST ONE Item", vbExclamation, "Item Detail Report"
        Exit Sub
    End If
        
    ' *******************************************************
    ' store the screen defaults
    If PRGlobal.GetByID(GlobalID) = True Then
        PRGlobal.Var1 = ""
        PRGlobal.Var2 = ""
        PRGlobal.Var3 = ""
        PRGlobal.Var4 = ""
        PRGlobal.Var5 = ""
        
        PRGlobal.Byte1 = Me.chkNoGross
        PRGlobal.Byte2 = Me.chkMatching
        PRGlobal.Byte3 = Me.chkPrintSS
        PRGlobal.Byte4 = Me.chkShowRemain
        PRGlobal.Byte5 = Me.chkAnniv
        PRGlobal.Byte6 = Me.chkRecall
        
        If Me.optEmpNo = True Then PRGlobal.Byte7 = 1
        If Me.optChkDate = True Then PRGlobal.Byte7 = 2
        If Me.optName = True Then PRGlobal.Byte7 = 3
        
        PRGlobal.Byte8 = Me.chkTotalsOnly
        
        ItmCt = 0
        rsItem.MoveFirst
        Do
            If rsItem!Selected = True Then
                ItmCt = ItmCt + 1
                If rsItem!IsItHours = True Then
                    x = "H" & rsItem!ItemID
                Else
                    x = "A" & rsItem!ItemID
                End If
                If ItmCt = 1 Then PRGlobal.Var1 = x
                If ItmCt = 2 Then PRGlobal.Var2 = x
                If ItmCt = 3 Then PRGlobal.Var3 = x
                If ItmCt = 4 Then PRGlobal.Var4 = x
                If ItmCt = 5 Then PRGlobal.Var5 = x
            End If
            rsItem.MoveNext
        Loop Until rsItem.EOF
        PRGlobal.Save (Equate.RecPut)
    End If
    ' *******************************************************
    
    InitFlag = True
    Me.Hide
    If Me.chkMatching Then
        MatchingReport
    ElseIf Me.chkAnniv Then
        AnnivRemain Me.chkRecall, Me.rsEmp, Me.rsItem
    Else
        ItemDetail RangeType, PRBatchID, CLng(Int(PEDate)), CLng(Int(CheckDt)), _
                   CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate
    End If

End Sub


Private Sub cmdSelDed_Click()
    rsItem.MoveFirst
    Do
        rsItem!Selected = True
        rsItem.Update
        rsItem.MoveNext
    Loop Until rsItem.EOF
    rsItem.MoveFirst
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub optChkDate_Click()
    
    ' can't show remaining if by date
    If optChkDate = True Then
        Me.chkShowRemain = 0
        Me.chkShowRemain.Enabled = False
    Else
        Me.chkShowRemain.Enabled = True
    End If

End Sub

Private Sub optEmpNo_Click()
    If optEmpNo = True Then
        Me.chkShowRemain.Enabled = True
    End If
End Sub
Private Sub chkMatching_Click()
    
Dim DedCount As Integer
Dim ItemID As Long
    
    If chkMatching Then
        
        ' only one deduction can be chosen
        DedCount = 0
        rsItem.MoveFirst
        Do
            If rsItem!Selected = True And rsItem!Type = PREquate.ItemTypeDED Then
                DedCount = DedCount + 1
                ItemID = rsItem!ItemID
            End If
            rsItem.MoveNext
        Loop Until rsItem.EOF
        
        If DedCount <> 1 Then
            Me.chkMatching = 0
            MsgBox "One deduction must be chosen to use this option!", vbExclamation
            Exit Sub
        End If
        
        Me.tdbMatchPct.Visible = True
        Me.tdbMaxPct.Visible = True
        Me.chkPrintSS.Visible = True
    
        If PRItem.GetByID(ItemID) = False Then
            MsgBox "Item not found: " & ItemID, vbExclamation
            Me.chkMatching = 0
            Exit Sub
        End If
    
        Me.tdbMatchPct = PRItem.MatchPct
        Me.tdbMaxPct = PRItem.MaxPct
        ' Me.tdbMaxPct.SetFocus
    
    Else
        Me.tdbMatchPct.Visible = False
        Me.tdbMaxPct.Visible = False
        Me.chkPrintSS.Visible = False
    End If

End Sub

Private Sub MatchingReport()

Dim LastID, CurrID As Long
Dim DedBasis As Currency
Dim EmpFilter As Boolean
Dim RecCount, RecTotal As Long
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "EmpNum", adDouble
    rs.Fields.Append "SSN", adDouble
    rs.Fields.Append "EmpName", adVarChar, 50, adFldIsNullable
    rs.Fields.Append "PEDate", adDate
    rs.Fields.Append "ChkDate", adDate
    rs.Fields.Append "GrossPay", adCurrency
    rs.Fields.Append "DedBasis", adCurrency
    rs.Fields.Append "DedAmt", adCurrency
    rs.Fields.Append "DedMatch", adCurrency
    rs.Fields.Append "TotalContr", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic

    ' find the deduction item being reported
    If PRItem.GetByID(MatchID) = False Then
        MsgBox "Item not found: " & MatchID, vbExclamation
        GoBack
    End If

    ' gather the data
    frmProgress.Caption = Trim(PRCompany.Name) & " Payroll Matching Item Report"
    frmProgress.lblMsg1 = txtDisplay
    frmProgress.Show

    If PRBatchID > 0 Then
        SQLString = "SELECT * FROM PRItemHist WHERE BatchID = " & PRBatchID & _
                    " AND EmployerItemID = " & MatchID
    Else
        If OptDate = "CHECK DATE" Then
            SQLString = "SELECT * FROM PRItemHist WHERE CheckDate >= " & CLng(StartDate) & _
                        " AND CheckDate <= " & CLng(EndDate) & _
                        " AND EmployerItemID = " & MatchID
                        
        Else
            SQLString = "SELECT * FROM PRItemHist WHERE PEDate >= " & CLng(StartDate) & _
                        " AND PEDate <= " & CLng(EndDate) & _
                        " AND EmployerItemID = " & MatchID
        End If
    End If
    
    SQLString = Trim(SQLString) & " ORDER BY EmployeeID"
    
    If PRItemHist.GetBySQL(SQLString) = False Then
        MsgBox "No item detail data found for this deduction!", vbExclamation
        GoBack
    End If
    
    LastID = 0
    RecTotal = PRItemHist.Records
    
    ' employee filter?
    EmpFilter = False
    rsEmp.MoveFirst
    Do
        If rsEmp!Selected = False Then
            EmpFilter = True
            Exit Do
        End If
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    
    Do
                        
        RecCount = RecCount + 1
        If RecCount Mod 10 = 1 Then
            frmProgress.lblMsg2 = "On Record: " & Format(RecCount, "#,###,##0") & _
                                  " of: " & Format(RecTotal, "#,###,##0")
            frmProgress.Refresh
        End If
                        
        ' employee filter?
        If EmpFilter Then
            rsEmp.Find "EmpID = " & PRItemHist.EmployeeID, 0, adSearchForward, 1
            If rsEmp.EOF Then GoTo NxtItemHist
            If rsEmp!Selected = False Then GoTo NxtItemHist
        End If
                        
        If LastID = 0 Or PRItemHist.EmployeeID <> LastID Then
            If PREmployee.GetByID(PRItemHist.EmployeeID) = False Then
                MsgBox "Employee not found: " & PRItemHist.EmployeeID, vbExclamation
                GoBack
            End If
        End If
        LastID = PRItemHist.EmployeeID
            
        If PRHist.GetByID(PRItemHist.HistID) = False Then
            MsgBox "PRHist not found: " & PRItemHist.HistID, vbExclamation
            GoBack
        End If
            
        If frmItemDetail.chkTotalsOnly = 0 Then     ' a record for each
            rs.AddNew
            rs!EmpNum = PREmployee.EmployeeNumber
            rs!SSN = PREmployee.SSN
            rs!EmpName = PREmployee.LFName
            rs!PEDate = PRHist.PEDate
            rs!ChkDate = PRHist.CheckDate
            rs!GrossPay = 0
            rs!DedBasis = 0
            rs!DedAmt = 0
            rs!DedMatch = 0
            rs!TotalContr = 0
        Else                                        ' total per employee
            ' find - create if dne
            rs.Find "EmpNum = " & PREmployee.EmployeeNumber, 0, adSearchForward, 1
            If rs.EOF Then
                rs.AddNew
                rs!EmpNum = PREmployee.EmployeeNumber
                rs!SSN = PREmployee.SSN
                rs!EmpName = PREmployee.LFName
                rs!PEDate = 0
                rs!ChkDate = 0
                rs!GrossPay = 0
                rs!DedBasis = 0
                rs!DedAmt = 0
                rs!DedMatch = 0
                rs!TotalContr = 0
            End If
        End If
        
        rs!GrossPay = rs!GrossPay + PRHist.Gross
        rs!DedBasis = rs!DedBasis + PRHist.Gross - PRItemHist.WageExcluded
        rs!DedAmt = rs!DedAmt + PRItemHist.Amount
        
        ' DedMatch calc......
        P1 = Round((PRHist.Gross - PRItemHist.WageExcluded) * Me.tdbMaxPct / 100, 2) ' wage base x max pct
        If P1 <= PRItemHist.Amount Then
            rs!DedMatch = rs!DedMatch + Round(P1 * Me.tdbMatchPct / 100, 2)
        Else
            rs!DedMatch = rs!DedMatch + Round(PRItemHist.Amount * Me.tdbMatchPct / 100, 2)
        End If
        
        rs!TotalContr = rs!DedAmt + rs!DedMatch
        rs.Update
        
NxtItemHist:
        If PRItemHist.GetNext = False Then Exit Do
    
    Loop
    
    If rs.RecordCount = 0 Then
        MsgBox "No data found", vbExclamation
        GoBack
    End If
    
    If Me.optEmpNo = True Then
        rs.Sort = "EmpNum, ChkDate"
    ElseIf Me.optChkDate = True Then
        rs.Sort = "ChkDate, EmpNum"
    Else
        rs.Sort = "EmpName, ChkDate"
    End If
    
    PrtInit "Land"
    SetFont 8, Equate.LandScape
    Columns = Columns - 14
    
    MatchHeader
    
    LastID = 0
    
    rs.MoveFirst
    Do
        
        ' subtotal
        If Me.optChkDate = True Then
            CurrID = rs!ChkDate
        Else
            CurrID = rs!EmpNum      ' works for # or name order
        End If
        If LastID <> 0 And LastID <> CurrID Then
            If Ln >= MaxLines - 2 Then
                FormFeed
                MatchHeader
            End If
            If ItemCount > 1 Then
                If Me.optChkDate = True Then
                    MatchSubTl "Check Date: " & Format(LastID, "mm/dd/yy")
                Else
                    MatchSubTl LastEmpName
                End If
            Else
                Ln = Ln + 1
            End If
            GrossPayTl(1) = 0
            DedBasisTl(1) = 0
            DedAmtTl(1) = 0
            DedMatchTl(1) = 0
            TotalContrTl(1) = 0
            ItemCount = 0
        End If
        LastID = CurrID
                    
        PrintValue(1) = rs!EmpNum:          FormatString(1) = "n9"
        PrintValue(2) = " ":                FormatString(2) = "a1"
        If Me.chkPrintSS = 1 Then
            LastEmpName = Format(rs!SSN, "###-##-####") & " " & rs!EmpName
        Else
            LastEmpName = rs!EmpName
        End If
        PrintValue(3) = LastEmpName:                        FormatString(3) = "a35"
        
        If rs!PEDate <> 0 Then
            PrintValue(4) = Format(rs!PEDate, " mm/dd/yy "):    FormatString(4) = "a10"
            PrintValue(5) = Format(rs!ChkDate, " mm/dd/yy "):   FormatString(5) = "a10"
        Else
            PrintValue(4) = "":     FormatString(4) = "a10"
            PrintValue(5) = "":     FormatString(5) = "a10"
        End If
        
        If Me.chkNoGross = 0 Then
            PrintValue(6) = rs!GrossPay:                        FormatString(6) = "d14"
            PrintValue(7) = rs!DedBasis:                        FormatString(7) = "d14"
        Else
            PrintValue(6) = " ":            FormatString(6) = "a14"
            PrintValue(7) = " ":            FormatString(7) = "a14"
        End If
        
        PrintValue(8) = rs!DedAmt:                          FormatString(8) = "d14"
        PrintValue(9) = rs!DedMatch:                        FormatString(9) = "d14"
        PrintValue(10) = rs!TotalContr:                     FormatString(10) = "d14"
        PrintValue(11) = " ":                               FormatString(11) = "~"
        FormatPrint
        Ln = Ln + 1
        
        ItemCount = ItemCount + 1
        
        ' update totals
        For i = 1 To 2
            GrossPayTl(i) = GrossPayTl(i) + rs!GrossPay
            DedBasisTl(i) = DedBasisTl(i) + rs!DedBasis
            DedAmtTl(i) = DedAmtTl(i) + rs!DedAmt
            DedMatchTl(i) = DedMatchTl(i) + rs!DedMatch
            TotalContrTl(i) = TotalContrTl(i) + rs!TotalContr
        Next i
        
        If Ln >= MaxLines Then
            FormFeed
            MatchHeader
        End If
        
        rs.MoveNext
    
    Loop Until rs.EOF
    
    ' print last subtl
    If ItemCount > 1 Then
        If Me.optChkDate = True Then
            If LastID <> 0 Then
                MatchSubTl "Check Date: " & Format(LastID, "mm/dd/yy")
            Else
                MatchSubTl "Total: "
            End If
        Else
            MatchSubTl LastEmpName
        End If
    Else
        Ln = Ln + 1
    End If
    
    ' grand totals
    GrossPayTl(1) = GrossPayTl(2)
    DedBasisTl(1) = DedBasisTl(2)
    DedAmtTl(1) = DedAmtTl(2)
    DedMatchTl(1) = DedMatchTl(2)
    TotalContrTl(1) = TotalContrTl(2)
    If Ln >= MaxLines - 2 Then
        FormFeed
        MatchHeader
    End If
    MatchSubTl "Final Total"
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    
End Sub

Private Sub MatchSubTl(ByVal SubString As String)
            
    PrintValue(1) = " ":                                FormatString(1) = "a10"
    PrintValue(2) = SubString:                          FormatString(2) = "a35"
    PrintValue(3) = " ":                                FormatString(3) = "a10"
    PrintValue(4) = " ":                                FormatString(4) = "a10"
    If Me.chkNoGross = 0 Then
        PrintValue(5) = GrossPayTl(1):                      FormatString(5) = "d14"
        PrintValue(6) = DedBasisTl(1):                      FormatString(6) = "d14"
    Else
        PrintValue(5) = " ":            FormatString(5) = "a14"
        PrintValue(6) = " ":            FormatString(6) = "a14"
    End If
    
    PrintValue(7) = DedAmtTl(1):                        FormatString(7) = "d14"
    PrintValue(8) = DedMatchTl(1):                      FormatString(8) = "d14"
    PrintValue(9) = TotalContrTl(1):                    FormatString(9) = "d14"
    PrintValue(10) = " ":                               FormatString(10) = "~"
    FormatPrint
    Ln = Ln + 2

End Sub

Private Sub MatchHeader()

    PageHeader "ITEM DETAIL MATCHING REPORT FOR: " & PRItem.Title, _
               Format(Me.tdbMaxPct / 100, "##0.00 %") & " of Wage Max / " & _
               Format(Me.tdbMatchPct / 100, "##0.00 %") & " Match", _
               txtDisplay
               
    PrintValue(1) = "  EMP NUM ":       FormatString(1) = "a10"
    PrintValue(2) = "EMPLOYEE NAME":    FormatString(2) = "a35"
    PrintValue(3) = " P/E DATE":        FormatString(3) = "a10"
    PrintValue(4) = "CHECK DATE":       FormatString(4) = "a10"
    If Me.chkNoGross = 0 Then
        PrintValue(5) = "GROSS PAY ":       FormatString(5) = "r14"
        PrintValue(6) = "DEDUCT BASIS ":    FormatString(6) = "r14"
    Else
        PrintValue(5) = " ":                FormatString(5) = "r14"
        PrintValue(6) = " ":                FormatString(6) = "r14"
    End If
    PrintValue(7) = "DEDUCT AMT ":      FormatString(7) = "r14"
    PrintValue(8) = "MATCH AMT ":       FormatString(8) = "r14"
    PrintValue(9) = "TOTAL CONTRIB ":   FormatString(9) = "r14"
    PrintValue(10) = " ":               FormatString(10) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = String(Columns - 5, "="):     FormatString(1) = "a" & Columns
    PrintValue(2) = " ":                        FormatString(2) = "~"
    FormatPrint
    Ln = Ln + 1
    
End Sub
Private Sub fgEmp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
Private Sub fgItem_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Function String_To_Long(ByVal InString As String, vType As String) As Long

    String_To_Long = 0
    
    If IsNull(InString) Then Exit Function
    If InString = "" Then Exit Function
    
    ' A for Amt / H for Hrs
    If Mid(InString, 1, 1) <> vType Then Exit Function
    
    InString = Mid(InString, 2, Len(InString) - 1)
    
    If IsNumeric(InString) = False Then Exit Function
 
    On Error Resume Next
    String_To_Long = CLng(InString)
    If Err.Number <> 0 Then String_To_Long = 0
    On Error GoTo 0

End Function

Private Sub chkShowRemain_Click()
    If Me.chkShowRemain Then
        Me.chkAnniv.Visible = True
        Me.chkRecall.Visible = True
        Me.optEmpNo = True
        Me.optChkDate.Enabled = False
    Else
        Me.chkAnniv.Visible = False
        Me.chkRecall.Visible = False
        Me.chkAnniv = 0
        Me.txtDisplay.Enabled = True
        Me.optChkDate.Enabled = True
        ' Me.cmdDateRange.Enabled = True
    End If
End Sub

Private Sub chkAnniv_Click()
'    If chkAnniv Then
'        Me.cmdDateRange.Enabled = False
'        Me.txtDisplay.Enabled = False
'    Else
'        Me.cmdDateRange.Enabled = True
'        Me.txtDisplay.Enabled = True
'    End If
End Sub


