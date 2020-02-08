VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmW2Print 
   Caption         =   "W2 Totals / Print"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9120
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHCState 
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdPrintW3 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   8040
      TabIndex        =   23
      Top             =   6960
      Width           =   1215
   End
   Begin TDBText6Ctl.TDBText tdbtxtCityOver 
      Height          =   375
      Left            =   8160
      TabIndex        =   21
      Top             =   1560
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "frmW2Print.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":0064
      Key             =   "frmW2Print.frx":0082
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
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
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CheckBox chkCityOver 
      Caption         =   "Override City Name"
      Height          =   495
      Left            =   6360
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Frame fraCitiesPer 
      Caption         =   "  Cities Per W2  "
      Height          =   975
      Left            =   7560
      TabIndex        =   17
      Top             =   2160
      Width           =   2895
      Begin VB.OptionButton optTwoCities 
         Caption         =   "TWO"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optOneCity 
         Caption         =   "ONE"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CheckBox chkDist 
      Caption         =   "Print City/State/SD Tax Distribution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   1080
      Width           =   4575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumL4Between 
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   5880
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":00C6
      Caption         =   "frmW2Print.frx":00E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":0156
      Keys            =   "frmW2Print.frx":0174
      Spin            =   "frmW2Print.frx":01BE
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdPrintTotals 
      Caption         =   "PRINT TOTALS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   8640
      Width           =   2775
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumL2Vert 
      Height          =   375
      Left            =   9840
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":01E6
      Caption         =   "frmW2Print.frx":0206
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":026C
      Keys            =   "frmW2Print.frx":028A
      Spin            =   "frmW2Print.frx":02D4
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintL4 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintL2 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      _cx             =   10186
      _cy             =   13150
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumL2Horz 
      Height          =   375
      Left            =   9840
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":02FC
      Caption         =   "frmW2Print.frx":031C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":0386
      Keys            =   "frmW2Print.frx":03A4
      Spin            =   "frmW2Print.frx":03EE
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumL4Vert 
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":0416
      Caption         =   "frmW2Print.frx":0436
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":049C
      Keys            =   "frmW2Print.frx":04BA
      Spin            =   "frmW2Print.frx":0504
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumL4Horz 
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":052C
      Caption         =   "frmW2Print.frx":054C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":05B6
      Keys            =   "frmW2Print.frx":05D4
      Spin            =   "frmW2Print.frx":061E
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumW3Vert 
      Height          =   375
      Left            =   9840
      TabIndex        =   24
      Top             =   7200
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":0646
      Caption         =   "frmW2Print.frx":0666
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":06CC
      Keys            =   "frmW2Print.frx":06EA
      Spin            =   "frmW2Print.frx":0734
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumW3Horz 
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   6720
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmW2Print.frx":075C
      Caption         =   "frmW2Print.frx":077C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2Print.frx":07E6
      Keys            =   "frmW2Print.frx":0804
      Spin            =   "frmW2Print.frx":084E
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
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6360
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6360
      X2              =   11880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      Caption         =   "FORM W3"
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblW2Count 
      Caption         =   "W2 Count"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label lbl4 
      Caption         =   "FOUR PER PAGE"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "TWO PER PAGE"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblTaxYear 
      Caption         =   "Tax Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Height          =   735
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "T O T A L S:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmW2Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TaxYear As Long
Dim rs As New ADODB.Recordset
Dim rsState As New ADODB.Recordset
Dim Box As Integer
Dim BoxTitle As String
Dim i, j As Integer
Dim ID, W2Count As Long
Dim Amt As Currency
Dim x As String

Dim W2Total(19, 4) As Currency

Dim W2BX(20, 4) As Variant
Dim BoxA, BoxB As String
Dim BoxC(5) As String
Dim BoxD As Long
Dim BoxE(5) As String
Dim Box13(3), Box12Code(4), Box14Code(4) As String
Dim W2Type As String
Dim W2CT, BottomCount, FormCount As Integer
Dim TotalFlag As Boolean
Dim FmtA As String
Dim VSpace As Long
Dim CurX, CurY As Long
Dim Box12String(4), Box14String(4) As String
Dim CPP As Byte
Dim FinalCity As Boolean
Dim CityCount As Long
Dim StateFlag As Boolean
Dim Cities As Long
Dim Box12Total As Currency
Dim LastStateID As Long

' store box 12 position (a=1 / b=2 ... for Box12 code = "CC")
Dim HIREBox As Byte
Dim HIREAmt As Currency

Private Sub Form_Load()
 
    If LCase(Mid(PRCompany.Name, 1, 9)) <> "hernandez" Then
        Me.txtHCState.Visible = False
    Else
        Me.txtHCState.Text = "STATE"
    End If
 
    HIREBox = 0
    HIREAmt = 0
    
    Me.lblCompanyName = PRCompany.Name
    Me.lblTaxYear = "Tax Year: " & TaxYear
    
    ' ------------------------------------------------------
    ' recordset to track multi-state ID and totals
    On Error Resume Next
    rsState.Close
    On Error GoTo 0
    rsState.CursorLocation = adUseClient
    rsState.Fields.Append "StateID", adDouble
    rsState.Fields.Append "ERStateID", adVarChar, 20, adFldIsNullable
    rsState.Fields.Append "StateWage", adCurrency
    rsState.Fields.Append "StateTax", adCurrency
    rsState.Fields.Append "CityName", adVarChar, 10, adFldIsNullable
    rsState.Fields.Append "CityWage", adCurrency
    rsState.Fields.Append "CityTax", adCurrency
    rsState.Open , , adOpenDynamic, adLockOptimistic
    ' ------------------------------------------------------
    
    Me.optOneCity = True
    Me.optOneCity.Enabled = False
    Me.optTwoCities.Enabled = False
    Me.fraCitiesPer.Enabled = False
    
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "W2Box", adDouble
    rs.Fields.Append "Desc", adVarChar, 50, adFldIsNullable
    rs.Fields.Append "Amount", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    CalcTotals
    
    ' 4 per page - between form adj
    SetNudge Me.tdbnumL4Between
    Me.tdbnumL4Between.ToolTipText = "SPACE BETWEEN TOP AND BOTTOM FORMS"
    GetNudge User.ID, "W2L4B"
    Me.tdbnumL4Between = VertNudge
    
    SetNudge Me.tdbnumL2Horz
    Me.tdbnumL2Horz.ToolTipText = "MOVE TEXT TO THE RIGHT"
    SetNudge Me.tdbnumL2Vert
    Me.tdbnumL2Vert.ToolTipText = "MOVE TEXT DOWN"
    
    GetNudge User.ID, "W2L2"
    Me.tdbnumL2Horz = HorzNudge
    Me.tdbnumL2Vert = VertNudge
    
    SetNudge Me.tdbnumL4Horz
    Me.tdbnumL4Horz.ToolTipText = "MOVE TEXT TO THE RIGHT"
    SetNudge Me.tdbnumL4Vert
    Me.tdbnumL4Vert.ToolTipText = "MOVE TEXT DOWN"
    
    GetNudge User.ID, "W2L4"
    Me.tdbnumL4Horz = HorzNudge
    Me.tdbnumL4Vert = VertNudge
    
    GetNudge User.ID, "W2L4B"
    Me.tdbnumL4Between = VertNudge
    
    GetNudge User.ID, "W3"
    Me.tdbnumW3Horz = HorzNudge
    Me.tdbnumW3Vert = VertNudge
    
    ' only one city - hide options
    If Cities = 1 Then
        Me.chkDist.Visible = False
        Me.chkDist.Value = 0
        Me.chkCityOver.Visible = False
        Me.tdbtxtCityOver.Visible = False
        Me.fraCitiesPer.Visible = False
        Me.optOneCity.Visible = False
        Me.optOneCity = True
        Me.optTwoCities.Visible = False
    End If
    
    Me.KeyPreview = True

End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    rs.Close
    rsState.Close
    On Error GoTo 0
    Unload Me
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub CalcTotals()
    
    frmProgress.Caption = "Calculating tax totals for: " & TaxYear
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.Show
    
    W2Count = 0
    Cities = 0      ' number of cities for the company
    
    For i = 0 To 19
        For j = 0 To 4
            W2Total(i, j) = 0
        Next j
    Next i
    
    For Box = 1 To 19
        
        BoxTitle = ""
        If Box = 1 Then BoxTitle = "Box  1 Wages"
        If Box = 2 Then BoxTitle = "Box  2 Fed Inc Tax"
        If Box = 3 Then BoxTitle = "Box  3 SS Wages"
        If Box = 4 Then BoxTitle = "Box  4 SS Tax"
        If Box = 5 Then BoxTitle = "Box  5 Med Wages"
        If Box = 6 Then BoxTitle = "Box  6 Med Tax"
        If Box = 7 Then BoxTitle = "Box  7 SS Tips"
        If Box = 8 Then BoxTitle = "Box  8 Alloc Tips"
        If Box = 9 Then BoxTitle = "Box  9 EIC Payment"
        If Box = 10 Then BoxTitle = "Box 10 Depend Care"
        If Box = 11 Then BoxTitle = "Box 11 NonQual Plans"
        If Box = 16 Then BoxTitle = "Box 16 State Wages"
        If Box = 17 Then BoxTitle = "Box 17 State Tax"
        If Box = 18 Then BoxTitle = "Box 18 Local Wages"
        If Box = 19 Then BoxTitle = "Box 19 Local Tax"
            
        If BoxTitle <> "" Then
            rs.AddNew
            rs!W2Box = Box
            rs!Desc = BoxTitle
            rs!Amount = 0
            rs.Update
        End If
    
    Next Box
    
    Box12Total = 0
    
    SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear & " AND " & _
                "Skip = 0"
    If PRW2.GetBySQL(SQLString) = False Then
        MsgBox "No W2 data found!", vbExclamation
        GoBack
    End If
    
    Do
        
        frmProgress.lblMsg2 = Trim(PRW2.BoxE_EEFirstName) & " " & Trim(PRW2.BoxE_EELastName)
        frmProgress.Refresh
        
        BoxUpdate PRW2.Box1_Wages, 1
        BoxUpdate PRW2.Box2_FedTax, 2
        BoxUpdate PRW2.Box3_SSWages, 3
        BoxUpdate PRW2.Box4_SSTax, 4
        BoxUpdate PRW2.Box5_MedWages, 5
        BoxUpdate PRW2.Box6_MedTax, 6
        BoxUpdate PRW2.Box7_SSTips, 7
        BoxUpdate PRW2.Box8_AllocTips, 8
        BoxUpdate PRW2.Box9_EIC, 9
        BoxUpdate PRW2.Box10_DCBen, 10
        BoxUpdate PRW2.Box11_NQPlans, 11
        
        For Box = 1 To 4
            If Box = 1 Then
                Amt = PRW2.Box12A_Amount
                ID = PRW2.Box12A_ID
                x = "a"
            End If
            If Box = 2 Then
                Amt = PRW2.Box12B_Amount
                ID = PRW2.Box12B_ID
                x = "b"
            End If
            If Box = 3 Then
                Amt = PRW2.Box12C_Amount
                ID = PRW2.Box12C_ID
                x = "c"
            End If
            If Box = 4 Then
                Amt = PRW2.Box12D_Amount
                ID = PRW2.Box12D_ID
                x = "d"
            End If
            If Amt <> 0 Then
                SQLString = "W2Box = " & 12 + Box / 10
                rs.Find SQLString, 0, adSearchForward, 1
                If rs.EOF Then
                    rs.AddNew
                    rs!W2Box = 12 + Box / 10
                    rs.Update
                End If
                rs!Amount = rs!Amount + Amt
                If PRGlobal.GetByID(ID) = False Then
                    MsgBox "W2 Box 12 Code err", vbExclamation
                    GoBack
                End If
                rs!Desc = "Box 12" & Trim(x) & " " & PRGlobal.Description
                rs.Update
                Box12Code(Box) = InParen(PRGlobal.Description)
                Box12Total = Box12Total + Amt
                            
                ' store usage of Box 12 code CC for HIRE
                If Box12Code(Box) = "CC" Then
                    HIREBox = Box
                    HIREAmt = HIREAmt + Amt
                End If
            
            End If
        Next Box
        
        For Box = 1 To 4
            If Box = 1 Then
                Amt = PRW2.Box14A_Amount
                ID = PRW2.Box14A_ID
                x = "a"
            End If
            If Box = 2 Then
                Amt = PRW2.Box14B_Amount
                ID = PRW2.Box14B_ID
                x = "b"
            End If
            If Box = 3 Then
                Amt = PRW2.Box14C_Amount
                ID = PRW2.Box14C_ID
                x = "c"
            End If
            If Box = 4 Then
                Amt = PRW2.Box14D_Amount
                ID = PRW2.Box14D_ID
                x = "d"
            End If
            If Amt <> 0 Then
                SQLString = "W2Box = " & 14 + Box / 10
                rs.Find SQLString, 0, adSearchForward, 1
                If rs.EOF Then
                    rs.AddNew
                    rs!W2Box = 14 + Box / 10
                    rs.Update
                End If
                rs!Amount = rs!Amount + Amt
                If PRGlobal.GetByID(ID) = False Then
                    MsgBox "W2 Box 14 Code err " & ID, vbExclamation
                    GoBack
                End If
                rs!Desc = "Box 14" & Trim(x) & " " & PRGlobal.Description
                rs.Update
                Box14Code(Box) = PRGlobal.Description
            End If
        Next Box

        ' State file
        SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & _
                    " AND TaxYear = " & TaxYear
        If PRW2State.GetBySQL(SQLString) Then
            Do
                
                ' update all states totals
                BoxUpdate PRW2State.StateWage, 16
                BoxUpdate PRW2State.StateTax, 17
                
                ' update individual state
                SQLString = "W2Box = " & 1600000 + PRW2State.StateID
                rs.Find SQLString, 0, adSearchForward, 1
                If rs.EOF Then
                    rs.AddNew
                    rs!W2Box = 1600000 + PRW2State.StateID
                    rs!Amount = 0
                    If PRState.GetByID(PRW2State.StateID) = False Then
                        MsgBox "State NF: " & PRW2State.StateID, vbExclamation
                        GoBack
                    End If
                    rs!Desc = Trim(PRState.StateAbbrev) & " Wage"
                    rs.Update
                End If
                rs!Amount = rs!Amount + PRW2State.StateWage
                rs.Update
                
                ' update individual state
                SQLString = "W2Box = " & 1700000 + PRW2State.StateID
                rs.Find SQLString, 0, adSearchForward, 1
                If rs.EOF Then
                    rs.AddNew
                    rs!W2Box = 1700000 + PRW2State.StateID
                    rs!Amount = 0
                    If PRState.GetByID(PRW2State.StateID) = False Then
                        MsgBox "State NF: " & PRW2State.StateID, vbExclamation
                        GoBack
                    End If
                    rs!Desc = Trim(PRState.StateAbbrev) & " Tax"
                    rs.Update
                End If
                rs!Amount = rs!Amount + PRW2State.StateTax
                rs.Update
                
                If PRW2State.GetNext = False Then Exit Do
            Loop
        
        End If
        
        ' city file
        SQLString = "SELECT * FROM PRW2City WHERE W2ID = " & PRW2.W2ID & _
                    " AND TaxYear = " & TaxYear
        If PRW2City.GetBySQL(SQLString) Then
            Do
                
                ' update individual City
                If PRW2City.SDTax = 0 Then
                    
                    ' update all Citys totals
                    If PRW2City.Courtesy = 0 Then
                        BoxUpdate PRW2City.CityWage, 18
                    End If
                    BoxUpdate PRW2City.CityTax, 19
                    
                    SQLString = "W2Box = " & 1900000 + PRW2City.CityID * 10 + 1
                    rs.Find SQLString, 0, adSearchForward, 1
                    If rs.EOF Then
                        rs.AddNew
                        rs!W2Box = 1900000 + PRW2City.CityID * 10 + 1
                        rs!Amount = 0
                        If PRCity.GetByID(PRW2City.CityID) = False Then
                            MsgBox "City NF: " & PRW2City.CityID, vbExclamation
                            GoBack
                        End If
                        rs!Desc = Trim(PRCity.ShortName) & " Tax"
                        rs.Update
                    End If
                Else        ' SD tax - State ID = 0 ...
                    SQLString = "W2Box = " & 2000000 + PRW2City.CityID * 10 + 1
                    rs.Find SQLString, 0, adSearchForward, 1
                    If rs.EOF Then
                        rs.AddNew
                        rs!W2Box = 2000000 + PRW2City.CityID * 10 + 1
                        rs!Amount = 0
                        If PRItem.GetByID(PRW2City.CityID) = False Then
                            ' ???
                        End If
                        rs!Desc = Trim(PRItem.Abbreviation) & " Tax"
                        rs.Update
                    End If
                End If
                rs!Amount = rs!Amount + PRW2City.CityTax
                rs.Update
                
                If PRW2City.GetNext = False Then Exit Do
            Loop
        
        End If
        
        W2Count = W2Count + 1
        If PRW2.GetNext = False Then Exit Do
    
    Loop
            
    ' count for the total W2
    W2Count = W2Count + 1
            
    ' create/update W3 PRGlobal records
    For i = 1 To 4
        
        If i = 1 Then j = PREquate.GlobalTypeW3A
        If i = 2 Then j = PREquate.GlobalTypeW3B
        If i = 3 Then j = PREquate.GlobalTypeW3C
        If i = 4 Then j = PREquate.GlobalTypeW3D
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & j & " AND " & _
                    "Year = " & TaxYear & " AND " & _
                    "UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.Year = TaxYear
            PRGlobal.UserID = PRCompany.CompanyID
            PRGlobal.TypeCode = j
            PRGlobal.Save (Equate.RecAdd)
        End If
    
        If i = 1 Then
            For j = 1 To 10
                rs.Find "W2Box = " & j, 0, adSearchForward, 1
                If rs.EOF Then
                    MsgBox "W2 Box " & j & " total not found!", vbExclamation
                    GoBack
                End If
                If j = 1 Then PRGlobal.Var1 = rs!Amount
                If j = 2 Then PRGlobal.Var2 = rs!Amount
                If j = 3 Then PRGlobal.Var3 = rs!Amount
                If j = 4 Then PRGlobal.Var4 = rs!Amount
                If j = 5 Then PRGlobal.Var5 = rs!Amount
                If j = 6 Then PRGlobal.Var6 = rs!Amount
                If j = 7 Then PRGlobal.Var7 = rs!Amount
                If j = 8 Then PRGlobal.Var8 = rs!Amount
                If j = 9 Then PRGlobal.Var9 = rs!Amount
                If j = 10 Then PRGlobal.Var10 = rs!Amount
            Next j
        End If
        
        If i = 2 Then
            For j = 11 To 20
                
                If j = 11 Or (j >= 16 And j <= 19) Then
                    rs.Find "W2Box = " & j, 0, adSearchForward, 1
                    If rs.EOF Then
                        MsgBox "W2 Box " & j & " total not found!", vbExclamation
                        GoBack
                    End If
                    If j = 11 Then PRGlobal.Var1 = rs!Amount
                    If j = 12 Then PRGlobal.Var2 = rs!Amount
                    If j = 13 Then PRGlobal.Var3 = rs!Amount
                    If j = 14 Then PRGlobal.Var4 = rs!Amount
                
                    If j = 16 Then PRGlobal.Var6 = rs!Amount
                    If j = 17 Then PRGlobal.Var7 = rs!Amount
                    If j = 18 Then PRGlobal.Var8 = rs!Amount
                    If j = 19 Then PRGlobal.Var9 = rs!Amount
                
                End If
                
                If j = 12 Then PRGlobal.Var2 = Box12Total
                
                ' j = 13 3rd party sick pay
                ' j=14 3rd party sick pay tax wh
                
                If j = 15 And PRGlobal.Var5 = "" Then
                    If PRState.GetByID(PRCompany.AddrStateID) Then
                        PRGlobal.Var5 = PRState.StateAbbrev
                    Else
                        PRGlobal.Var5 = "OH"
                    End If
                End If
                
                If j = 20 And PRGlobal.Var10 = "" Then
                    PRGlobal.Var10 = PRCompany.StateID
                End If
            
            Next j
        End If
    
        If i = 3 Then
            
            If PRGlobal.Var1 = "" Then PRGlobal.Var1 = "0"
            PRGlobal.Var2 = W2Count
            ' Var3 - box d est number
            If PRGlobal.Var4 = "" Then PRGlobal.Var4 = PRCompany.FederalID
            If PRGlobal.Var5 = "" Then PRGlobal.Var5 = PRCompany.Name
            If PRGlobal.Var6 = "" Then PRGlobal.Var6 = PRCompany.Address1
            If PRGlobal.Var7 = "" Then PRGlobal.Var7 = PRCompany.Address2
            If PRGlobal.Var8 = "" Then PRGlobal.Var8 = PRCompany.CSZ
            ' Var 9 - addr line 4
            ' var 10 - Other EIN
        End If
    
        If i = 4 Then
            PRGlobal.Var1 = HIREAmt
        End If
    
        PRGlobal.Save (Equate.RecPut)
    
    Next i
            
    frmProgress.Hide
        
    rs.Sort = "W2Box"
        
    SetGrid rs, fg
    fg.ColWidth(0) = 0
    
    fg.ColWidth(1) = 3500
    fg.TextMatrix(0, 1) = "Description"
    
    fg.ColWidth(2) = 2000
        
    fg.Editable = flexEDNone
        
    Me.lblW2Count = "W2 Count: " & Format(W2Count, "###,##0")
        
End Sub

Private Function InParen(ByVal InString As String) As String
    
Dim StringFlag As Boolean
    
    InParen = ""
    If IsNull(InString) Then Exit Function
    If InString = "" Then Exit Function
    
    StringFlag = False
    InString = Trim(UCase(InString))
    j = Len(InString)
    For i = 1 To j
        If StringFlag = False And Mid(InString, i, 1) = "(" Then
            StringFlag = True
            i = i + 1
        End If
        If StringFlag = True And Mid(InString, i, 1) = ")" Then Exit Function
        If StringFlag = True Then
            InParen = Trim(InParen) & Mid(InString, i, 1)
        End If
    Next i
    
End Function

Private Sub BoxUpdate(ByVal W2Amount As Currency, ByVal W2Box As Double)
    SQLString = "W2Box = " & W2Box
    rs.Find SQLString, 0, adSearchForward, 1
    If rs.EOF Then
        MsgBox "Box Not Found: " & W2Box, vbExclamation
        GoBack
    End If
    rs!Amount = rs!Amount + W2Amount
    rs.Update
End Sub

Private Sub PrintLoop()

    Dim StateFlag As Boolean
    
    ' clear the rsState record set
    rsDelAll rsState
    
    ' cities per page
    If Me.optOneCity = True Then
        CPP = 1
    Else
        CPP = 2
    End If
    
    BoxD = 0
    TotalFlag = False

    SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear
    If frmW2.optOrderName = True Then
        SQLString = Trim(SQLString) & " ORDER BY BoxE_EELastName, BoxE_EEFirstName, BoxE_EEMidInit"
    Else
        SQLString = Trim(SQLString) & " ORDER BY EmployeeNumber"
    End If
        
    If PRW2.GetBySQL(SQLString) = False Then
        MsgBox "PRW2 data error", vbExclamation
        GoBack
    End If

'    frmProgress.Caption = "W2 Printing"
'    frmProgress.lblMsg1 = Trim(PRCompany.Name) & " Tax Year: " & TaxYear
'    frmProgress.Show

    W2CT = 0
    FormCount = 0
    BottomCount = 0
    Ln = 0
    
    Do
    
        If PRW2.Skip = 1 Then GoTo NextPRW2
            
        BoxA = Format(PRW2.BoxA_SSNumber, "000-00-0000")
        BoxB = PRW2.BoxB_FedID
        
        For i = 1 To 4
            BoxC(i) = ""
            BoxE(i) = ""
        Next i
        
        j = 0
        For i = 1 To 4
            If i = 1 Then x = PRW2.BoxC_ERName
            If i = 2 Then x = PRW2.BoxC_ERAddr1
            If i = 3 Then x = PRW2.BoxC_ERAddr2
            If i = 4 Then x = PRW2.BoxC_ERCity & ", " & PRW2.BoxC_ERState & "  " & PRW2.BoxC_ERZip
            If x <> "" Then
                j = j + 1
                BoxC(j) = x
            End If
        Next i
        
        j = 0
        For i = 1 To 4
            If i = 1 Then
                x = PRW2.BoxE_EEFirstName
                If PRW2.BoxE_EEMidInit <> "" Then
                    x = x & " " & PRW2.BoxE_EEMidInit
                End If
                x = Trim(x) & " " & PRW2.BoxE_EELastName
            End If
            If i = 2 Then x = PRW2.BoxE_EEAddr1
            If i = 3 Then x = PRW2.BoxE_EEAddr2
            If i = 4 Then x = PRW2.BoxE_EECity & ", " & PRW2.BoxE_EEState & "  " & PRW2.BoxE_EEZip
            If x <> "" Then
                j = j + 1
                BoxE(j) = x
            End If
        Next i
    
        W2BX(1, 0) = PRW2.Box1_Wages
        W2BX(2, 0) = PRW2.Box2_FedTax
        W2BX(3, 0) = PRW2.Box3_SSWages
        W2BX(4, 0) = PRW2.Box4_SSTax
        W2BX(5, 0) = PRW2.Box5_MedWages
        W2BX(6, 0) = PRW2.Box6_MedTax
        W2BX(7, 0) = PRW2.Box7_SSTips
        W2BX(8, 0) = PRW2.Box8_AllocTips
        W2BX(9, 0) = PRW2.Box9_EIC
        W2BX(10, 0) = PRW2.Box10_DCBen
        W2BX(11, 0) = PRW2.Box11_NQPlans
        
        W2BX(12, 1) = PRW2.Box12A_Amount
        W2BX(12, 2) = PRW2.Box12B_Amount
        W2BX(12, 3) = PRW2.Box12C_Amount
        W2BX(12, 4) = PRW2.Box12D_Amount
        
        For i = 1 To 3
            If i = 1 Then j = PRW2.Box13_StatEmp
            If i = 2 Then j = PRW2.Box13_RetirePlan
            If i = 3 Then j = PRW2.Box13_3rdParty
            If j = 1 Then
                Box13(i) = "X"
            Else
                Box13(i) = " "
            End If
        Next i
        
        W2BX(14, 1) = PRW2.Box14A_Amount
        W2BX(14, 2) = PRW2.Box14B_Amount
        W2BX(14, 3) = PRW2.Box14C_Amount
        W2BX(14, 4) = PRW2.Box14D_Amount
    
        For i = 1 To 4
            
            'clear
            W2BX(15, i) = ""
            W2BX(16, i) = 0
            W2BX(17, i) = 0
            W2BX(18, i) = 0
            W2BX(19, i) = 0
            W2BX(20, i) = ""
        
            ' box 12/14 strings
            If W2BX(12, i) > 0 Then
                x = Mid(Box12Code(i), 1, 2)
                Box12String(i) = x & Space(3 - Len(x)) & _
                                 PadRight(Format(W2BX(12, i), "#,###,##0.00"), 12)
            Else
                Box12String(i) = ""
            End If
        
            If W2BX(14, i) > 0 Then
                Box14String(i) = Mid(Box14Code(i), 1, 10) & _
                                 PadRight(Format(W2BX(14, i), "##,##0.00"), 9)
            
                x = Trim(Mid(Box14Code(i), 1, 8))
                Box14String(i) = x & Space(8 - Len(x)) & _
                                 PadRight(Format(W2BX(14, i), "#####0.00"), 9)
            Else
                Box14String(i) = ""
            End If
        
            ' test pattern
            ' Box12String(i) = "X  " & PadRight(Format(i * 100, "#,###,##0.00"), 12)
            ' Box14String(i) = Mid("Box 14-" & i, 1, 10) & PadRight(Format(i * 100, "#,###,##0.00"), 9)
        
        Next i
        
        ' 4 per page - print W2 Top if on new state
        StateFlag = False
        
        ' loop for cities of each state
        LastStateID = 0
        
        ' hernandez multi state patch
        If Me.txtHCState.Visible = False Or Me.txtHCState = "" Or Me.txtHCState = "FED" Then
            SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & " " & _
                        " AND TaxYear = " & TaxYear & " " & _
                        "ORDER BY StateID"
        Else
            
            SQLString = "SELECT * FROM PRState WHERE StateAbbrev = '" & Me.txtHCState & "'"
            If PRState.GetBySQL(SQLString) = False Then
                MsgBox "State NF: " & Me.txtHCState, vbExclamation
                End
            End If
            
            SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & " " & _
                        " AND TaxYear = " & TaxYear & " " & _
                        " AND StateID = " & PRState.StateID & _
                        " ORDER BY StateID"
        End If
        
        If PRW2State.GetBySQL(SQLString) = False Then
            ' ?????
            GoTo NextPRW2
        End If
        
        PrintW2Top

        Do
            
            ' ---------------------------------------------------------------------
            ' accum totals by state
            rsState.Find "StateID = " & PRW2State.StateID, 0, adSearchForward, 1
            If rsState.EOF Then
                rsState.AddNew
                rsState!StateID = PRW2State.StateID
                rsState!ERStateID = PRW2State.ERStateID
                rsState!StateWage = 0
                rsState!StateTax = 0
                rsState!CityName = ""
                rsState!CityWage = 0
                rsState!CityTax = 0
                rsState.Update
            End If
            
            rsState!StateWage = rsState!StateWage + PRW2State.StateWage
            rsState!StateTax = rsState!StateTax + PRW2State.StateTax
            ' ---------------------------------------------------------------------
            
            ' assign the state boxes
            If PRState.GetByID(PRW2State.StateID) = False Then
                W2BX(15, 1) = "??"
            Else
                W2BX(15, 1) = PRState.StateAbbrev
            End If
            
            W2BX(15, 2) = PRW2State.ERStateID
            W2BX(16, 1) = PRW2State.StateWage
            W2BX(17, 1) = PRW2State.StateTax
            
            ' employee subtl
            W2BX(16, 4) = W2BX(16, 4) + PRW2State.StateWage
            W2BX(17, 4) = W2BX(17, 4) + PRW2State.StateTax
            
            ' dist W2 - second or more state for the employee
            If StateFlag = True Then
                PrintW2Top
            End If
            
            ' how many cities for this state?
            CityCount = 0
            SQLString = "SELECT * FROM PRW2City WHERE W2ID = " & PRW2.W2ID & _
                        " AND StateID = " & PRW2State.StateID & _
                        " AND TaxYear = " & TaxYear
            If PRW2City.GetBySQL(SQLString) = False Then
                W2BX(18, 1) = 0
                W2BX(19, 1) = 0
                W2BX(20, 1) = ""
                If frmW2Print.chkDist = 1 Then
                    PrintW2Bottom
                End If
                            
                ' 140118 - fix for no city wh W2s
                FinalCity = True
                            
            Else
                Do
        
                    CityCount = CityCount + 1
                    If CityCount = PRW2City.Records Then
                        FinalCity = True
                    Else
                        FinalCity = False
                    End If
                    
                    W2BX(18, 1) = PRW2City.CityWage
                    W2BX(19, 1) = PRW2City.CityTax
                    W2BX(20, 1) = PRW2City.CityName
        
                    ' employee subtl
                    If PRW2City.SDTax = 0 And PRW2City.Courtesy = 0 Then
                        W2BX(18, 4) = W2BX(18, 4) + PRW2City.CityWage
                        W2BX(19, 4) = W2BX(19, 4) + PRW2City.CityTax
                    End If
        
                    ' print each distn
                    If frmW2Print.chkDist = 1 Then
                        PrintW2Bottom
                    End If
            
                    ' update totals by state
                    If PRW2City.SDTax = 0 And PRW2City.Courtesy = 0 Then
                        rsState!CityWage = rsState!CityWage + PRW2City.CityWage
                        rsState!CityTax = rsState!CityTax + PRW2City.CityTax
                        If rsState!CityName = "" Then
                            rsState!CityName = Mid(PRW2City.CityName, 1, 10)
                        End If
                    End If
                    
                    If PRW2City.GetNext = False Then Exit Do
                    
                Loop
            End If
        
            rsState.Update
        
            ' one bottom per state
            If frmW2Print.chkDist = 0 Then
                For i = 16 To 19
                    W2BX(i, 1) = W2BX(i, 4)
                    W2BX(16, 4) = 0
                Next i
                ' ??? box 15 / 20
                PrintW2Bottom
            Else
                StateFlag = True    ' signal to print W2 top if another state exists
                                    ' for the employee
            End If
            
            If PRW2State.GetNext = False Then Exit Do
        
            ' multi state fed pass
            If Me.txtHCState = "FED" Then Exit Do
        
        Loop
        
NextPRW2:
        If PRW2.GetNext = False Then Exit Do
    
    Loop
    
    ' -------------------------------------------------
    ' ER totals - set up boxes 1 thru 14
    TotalFlag = True
    BoxA = ""
    rs.MoveFirst
    Do
        For i = 1 To 11
            If rs!W2Box = i Then W2BX(i, 0) = rs!Amount
        Next i
        For i = 1 To 4
            If rs!W2Box = 12 + i / 10 Then W2BX(12, i) = rs!Amount
            If rs!W2Box = 14 + i / 10 Then W2BX(14, i) = rs!Amount
        Next i
        rs.MoveNext
    Loop Until rs.EOF
    
    For i = 1 To 4
        BoxE(i) = ""
        
        ' box 12/14 strings
        If W2BX(12, i) > 0 Then
            x = Mid(Box12Code(i), 1, 2)
            Box12String(i) = x & Space(3 - Len(x)) & _
                             PadRight(Format(W2BX(12, i), "#,###,##0.00"), 12)
        Else
            Box12String(i) = ""
        End If
    
        If W2BX(14, i) > 0 Then
            Box14String(i) = Mid(Box14Code(i), 1, 10) & _
                             PadRight(Format(W2BX(14, i), "##,##0.00"), 9)
        
            x = Trim(Mid(Box14Code(i), 1, 8))
            Box14String(i) = x & Space(8 - Len(x)) & _
                             PadRight(Format(W2BX(14, i), "#####0.00"), 9)
        
        Else
            Box14String(i) = ""
        End If
    
    Next i
    
    ' -------------------------------------------------
    ' print ER totals - loop each state
    If rsState.RecordCount = 0 Then
        MsgBox "No state info?", vbExclamation
        GoBack
    End If

    rsState.MoveFirst
    Do
        
        ' get the state
        If PRState.GetByID(rsState!StateID) = False Then
            MsgBox "State not found: " & rsState!StateID, vbExclamation
            GoBack
        End If
        W2BX(15, 1) = PRState.StateAbbrev
        W2BX(15, 2) = rsState!ERStateID
        W2BX(16, 1) = rsState!StateWage
        W2BX(17, 1) = rsState!StateTax
        W2BX(18, 1) = rsState!CityWage
        W2BX(19, 1) = rsState!CityTax
        W2BX(20, 1) = rsState!CityName
        
        PrintW2Top
        PrintW2Bottom
        
        rsState.MoveNext
    
    Loop Until rsState.EOF

'    ' print ER totals
'    TotalFlag = True
'    BoxA = ""
'    rs.MoveFirst
'    Do
'        For i = 1 To 11
'            If rs!W2Box = i Then W2BX(i, 0) = rs!Amount
'        Next i
'        For i = 1 To 4
'            If rs!W2Box = 12 + i / 10 Then W2BX(12, i) = rs!Amount
'            If rs!W2Box = 14 + i / 10 Then W2BX(14, i) = rs!Amount
'        Next i
'        For i = 16 To 19
'            If rs!W2Box = i Then W2BX(i, 1) = rs!Amount
'        Next i
'        rs.MoveNext
'    Loop Until rs.EOF
'
'    For i = 1 To 4
'        BoxE(i) = ""
'
'        ' box 12/14 strings
'        If W2BX(12, i) > 0 Then
'            Box12String(i) = Mid(Box12Code(i), 1, 1) & "  " & _
'                             PadRight(Format(W2BX(12, i), "#,###,##0.00"), 12)
'        Else
'            Box12String(i) = ""
'        End If
'
'        If W2BX(14, i) > 0 Then
'            Box14String(i) = Mid(Box14Code(i), 1, 10) & _
'                             PadRight(Format(W2BX(14, i), "##,##0.00"), 9)
'        Else
'            Box14String(i) = ""
'        End If
'
'    Next i
'
'    PrintW2Top
'    PrintW2Bottom
        
    Prvw.vsp.EndDoc
    PrvwReturn = True
    Prvw.Show vbModal

End Sub
Private Sub PrintW2Top()
    
    If W2Type = "L2" Then PrintL2Top
    
    If W2Type = "L4" Then
        W2CT = W2CT + 1
        PrintL4Top 100, 790:    PrintL4Top 6000, 790
        PrintL4Top 100, 8730:   PrintL4Top 6000, 8730
    End If
        
End Sub
Private Sub PrintW2Bottom()
    
    ' city name override
    With Me.chkCityOver
        If .Visible = True And .Value = 1 Then
            W2BX(20, 1) = Me.tdbtxtCityOver
        End If
    End With
    
    BottomCount = BottomCount + 1
    
    If BottomCount > CPP Then
        If W2Type = "L4" And CPP = 2 Then FormFeed
        PrintW2Top
        BottomCount = 1
    End If
    
    If W2Type = "L2" Then
        
        PrintL2Bottom
    
    Else        ' EE copy - 4 per sheet
        
        ' one state per sheet
        If BottomCount = 1 Then
            PrintL4State 200, 6200:     PrintL4State 6130, 6200
            PrintL4State 200, 14150:    PrintL4State 6130, 14150
        End If
                    
        If CPP = 1 Then     ' one city per page
            PrintL4City 200, 6800:      PrintL4City 6130, 6800
            PrintL4City 200, 14685:     PrintL4City 6130, 14685
            FormFeed
        Else
            If BottomCount = 1 Then
                PrintL4City 200, 6800:      PrintL4City 6130, 6800
                PrintL4City 200, 14670:     PrintL4City 6130, 14670
            Else
                PrintL4City 200, 7000:      PrintL4City 6130, 7000
                PrintL4City 200, 14870:     PrintL4City 6130, 14870
            End If
            
            If FinalCity = True Then FormFeed
        End If
            
    End If

End Sub

Private Sub PrintL2Top()
    
    YUnits = 240
    
    W2CT = W2CT + 1
    FormCount = FormCount + 1
    If FormCount = 3 Then
        FormFeed
        FormCount = 1
    End If
    
    BottomCount = 0
    
    If FormCount = 1 Then
        Ln = Ln + 4
    Else
        Ln = 37
    End If
    
    ' void box for the total w2
    If TotalFlag = True Or PRW2.Void = 1 Then
        x = "X"
    Else
        x = ""
    End If
    If FormCount = 1 Then
        PosPrint 2220, 850, x
    Else
        PosPrint 2220, 8700, x
    End If
    
    x = ""      ' void box printed above
    
'    PrintValue(1) = " ":                        FormatString(1) = "a18"
'    PrintValue(2) = x:                          FormatString(2) = "a9"
'    PrintValue(3) = BoxA:                       FormatString(3) = "a11"
'    PrintValue(4) = " ":                        FormatString(4) = "~"
'    FormatPrint
'    Ln = Ln + 2
'
'    PrintValue(1) = " ":                        FormatString(1) = "a3"
'    PrintValue(2) = BoxB:                       FormatString(2) = "a10"
'    PrintValue(3) = " ":                        FormatString(3) = "a46"
'    PrintValue(4) = W2BX(1, 0):                 FormatString(4) = "d12"
'    PrintValue(5) = " ":                        FormatString(5) = "a7"
'    PrintValue(6) = W2BX(2, 0):                 FormatString(6) = "d12"
'    PrintValue(7) = " ":                        FormatString(7) = "~"
'    FormatPrint
'    Ln = Ln + 2
'
'    PrintValue(1) = " ":                        FormatString(1) = "a59"
'    PrintValue(2) = W2BX(3, 0):                 FormatString(2) = "d12"
'    PrintValue(3) = " ":                        FormatString(3) = "a7"
'    PrintValue(4) = W2BX(4, 0):                 FormatString(4) = "d12"
'    PrintValue(5) = " ":                        FormatString(5) = "~"
'    FormatPrint
'    Ln = Ln + 1
'
'    PrintValue(1) = " ":                        FormatString(1) = "a3"
'    PrintValue(2) = BoxC(1):                    FormatString(2) = "a50"
'    PrintValue(3) = " ":                        FormatString(3) = "~"
'    FormatPrint
'    Ln = Ln + 1
'
'    PrintValue(1) = " ":                        FormatString(1) = "a3"
'    PrintValue(2) = BoxC(2):                    FormatString(2) = "a50"
'    PrintValue(3) = " ":                        FormatString(3) = "a6"
'    PrintValue(4) = W2BX(5, 0):                 FormatString(4) = "d12"
'    PrintValue(5) = " ":                        FormatString(5) = "a7"
'    PrintValue(6) = W2BX(6, 0):                 FormatString(6) = "d12"
'    PrintValue(7) = " ":                        FormatString(7) = "~"
'    FormatPrint
'    Ln = Ln + 1
'
'    PrintValue(1) = " ":                        FormatString(1) = "a3"
'    PrintValue(2) = BoxC(3):                    FormatString(2) = "a50"
'    PrintValue(3) = " ":                        FormatString(3) = "~"
'    FormatPrint
'    Ln = Ln + 1
'
'    PrintValue(1) = " ":                        FormatString(1) = "a3"
'    PrintValue(2) = BoxC(4):                    FormatString(2) = "a50"
'    PrintValue(3) = " ":                        FormatString(3) = "a6"
'    PrintValue(4) = W2BX(7, 0):                 FormatString(4) = "d12"
'    PrintValue(5) = " ":                        FormatString(5) = "a7"
'    PrintValue(6) = W2BX(8, 0):                 FormatString(6) = "d12"
'    PrintValue(7) = " ":                        FormatString(7) = "~"
'    FormatPrint
'    Ln = Ln + 2
    
    
    
    PrintValue(1) = " ":                        FormatString(1) = "a18"
    PrintValue(2) = x:                          FormatString(2) = "a9"
    PrintValue(3) = BoxA:                       FormatString(3) = "a11"
    PrintValue(4) = " ":                        FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxB:                       FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "a45"
    PrintValue(4) = W2BX(1, 0):                 FormatString(4) = "d13"
    PrintValue(5) = " ":                        FormatString(5) = "a6"
    PrintValue(6) = W2BX(2, 0):                 FormatString(6) = "d13"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = " ":                        FormatString(1) = "a58"
    PrintValue(2) = W2BX(3, 0):                 FormatString(2) = "d13"
    PrintValue(3) = " ":                        FormatString(3) = "a6"
    PrintValue(4) = W2BX(4, 0):                 FormatString(4) = "d13"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxC(1):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1


    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxC(2):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "a5"
    PrintValue(4) = W2BX(5, 0):                 FormatString(4) = "d13"
    PrintValue(5) = " ":                        FormatString(5) = "a6"
    PrintValue(6) = W2BX(6, 0):                 FormatString(6) = "d13"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxC(3):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxC(4):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "a5"
    PrintValue(4) = W2BX(7, 0):                 FormatString(4) = "d13"
    PrintValue(5) = " ":                        FormatString(5) = "a6"
    PrintValue(6) = W2BX(8, 0):                 FormatString(6) = "d13"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    BoxD = BoxD + 1
    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxD:                       FormatString(2) = "n4":
    PrintValue(3) = " ":                        FormatString(3) = "a52"
    
    
    ' 2011 - no box 9
    If TaxYear = 2010 Then
        PrintValue(4) = W2BX(9, 0): FormatString(4) = "d12"
    ElseIf TaxYear >= 2011 Then
        PrintValue(4) = "": FormatString(4) = "a12"
    End If
    
    PrintValue(5) = " ":                        FormatString(5) = "a7"
    PrintValue(6) = W2BX(10, 0):                FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a59"
    PrintValue(2) = W2BX(11, 0):                FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a4"
    PrintValue(4) = Box12String(1):             FormatString(4) = "a15"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxE(1):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "a1"
    PrintValue(4) = Box13(1):                   FormatString(4) = "a1"
    PrintValue(5) = " ":                        FormatString(5) = "a5"
    PrintValue(6) = Box13(2):                   FormatString(6) = "a1"
    PrintValue(7) = " ":                        FormatString(7) = "a6"
    PrintValue(8) = Box13(3):                   FormatString(8) = "a1"
    PrintValue(9) = " ":                        FormatString(9) = "a7"
    PrintValue(10) = Box12String(2):            FormatString(10) = "a15"
    PrintValue(11) = " ":                       FormatString(11) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxE(2):                    FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxE(3):                    FormatString(2) = "a50"
    PrintValue(3) = Box14String(1):             FormatString(3) = "a19"
    PrintValue(4) = " ":                        FormatString(4) = "a3"
    PrintValue(5) = Box12String(3):             FormatString(5) = "a15"
    PrintValue(6) = " ":                        FormatString(6) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a3"
    PrintValue(2) = BoxE(4):                    FormatString(2) = "a50"
    PrintValue(3) = Box14String(2):             FormatString(3) = "a19"
    PrintValue(4) = " ":                        FormatString(4) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a53"
    PrintValue(2) = Box14String(3):             FormatString(2) = "a19"
    PrintValue(3) = " ":                        FormatString(3) = "a3"
    PrintValue(4) = Box12String(4):             FormatString(4) = "a15"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a53"
    PrintValue(2) = Box14String(4):             FormatString(2) = "a19"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2

        
    
End Sub

Private Sub PrintL2Bottom()
        
    If BottomCount = 1 Then
        PrintValue(1) = " ":                        FormatString(1) = "a2"
        PrintValue(2) = W2BX(15, 1):                FormatString(2) = "a3"
        PrintValue(3) = " ":                        FormatString(3) = "a5"
        PrintValue(4) = W2BX(15, 2):                FormatString(4) = "a15"
        PrintValue(5) = " ":                        FormatString(5) = "a5"
        PrintValue(6) = W2BX(16, 1):                FormatString(6) = "d13"
        PrintValue(7) = " ":                        FormatString(7) = "a1"
        PrintValue(8) = W2BX(17, 1):                FormatString(8) = "d12"
        PrintValue(9) = " ":                        FormatString(9) = "a1"
    Else        ' only print the state once per form
        PrintValue(1) = " ":                        FormatString(1) = "a2"
        PrintValue(2) = " ":                        FormatString(2) = "a3"
        PrintValue(3) = " ":                        FormatString(3) = "a5"
        PrintValue(4) = " ":                        FormatString(4) = "a15"
        PrintValue(5) = " ":                        FormatString(5) = "a6"
        PrintValue(6) = " ":                        FormatString(6) = "a12"
        PrintValue(7) = " ":                        FormatString(7) = "a1"
        PrintValue(8) = " ":                        FormatString(8) = "a12"
        PrintValue(9) = " ":                        FormatString(9) = "a1"
    End If
    
    PrintValue(10) = W2BX(18, 1):               FormatString(10) = "d13"
    PrintValue(11) = " ":                       FormatString(11) = "a2"
    PrintValue(12) = W2BX(19, 1):               FormatString(12) = "d12"
    PrintValue(13) = " ":                       FormatString(13) = "a0"
    PrintValue(14) = W2BX(20, 1):               FormatString(14) = "a7"
    PrintValue(15) = " ":                       FormatString(15) = "~"
    FormatPrint
    Ln = Ln + 2
    
End Sub

Private Sub PrintL4Top(ByVal StartX As Long, ByVal StartY As Long)

Dim L4 As Byte

    YUnits = 190
    VSpace = 100

    Prvw.vsp.Font.Bold = True

    FmtA = "##,###,##0.00"
    
    BottomCount = 0

    CurX = StartX
    CurY = StartY
    
    ' between forms nudge
    If StartY > 7000 Then
        CurY = CurY + Me.tdbnumL4Between * 20
    End If
    
    CurX = StartX + 2180
    PosPrint CurX, CurY, L4Fmt(W2BX(1, 0))
    CurX = CurX + 1960
    PosPrint CurX, CurY, L4Fmt(W2BX(2, 0))

    CurX = StartX + 400
    CurY = CurY + YUnits
    PosPrint CurX, CurY, BoxA

    CurY = CurY + YUnits
    CurX = StartX + 2180
    PosPrint CurX, CurY, L4Fmt(W2BX(3, 0))
    CurX = CurX + 1960
    PosPrint CurX, CurY, L4Fmt(W2BX(4, 0))

    CurY = CurY + YUnits * 2
    CurX = StartX + 160
    PosPrint CurX, CurY, BoxB
    CurX = StartX + 2180
    PosPrint CurX, CurY, L4Fmt(W2BX(5, 0))
    CurX = CurX + 1960
    PosPrint CurX, CurY, L4Fmt(W2BX(6, 0))
        
    CurY = CurY + YUnits * 2
    CurX = StartX + 230
    For i = 1 To 4
        PosPrint CurX, CurY, BoxC(i)
        CurY = CurY + YUnits
    Next i

    ' Control #
    BoxD = W2CT
    CurY = CurY + YUnits * 1
    CurX = StartX + 2400
    PosPrint CurX, CurY, BoxD

    CurY = CurY + YUnits * 3
    CurX = StartX + 230
    For i = 1 To 4
        PosPrint CurX, CurY, BoxE(i)
        CurY = CurY + YUnits
    Next i
    
    CurY = CurY + YUnits * 1
    CurX = StartX + 340
    PosPrint CurX, CurY, L4Fmt(W2BX(7, 0))
    CurX = CurX + 1890
    PosPrint CurX, CurY, L4Fmt(W2BX(8, 0))
    CurX = CurX + 1890
    
    ' 2011 - no box 9
    If TaxYear = 2010 Then
        PosPrint CurX, CurY, L4Fmt(W2BX(9, 0))
    ElseIf TaxYear >= 2011 Then
        PosPrint CurX, CurY, ""
    End If
    
    CurY = CurY + YUnits * 2
    CurX = StartX + 350
    PosPrint CurX, CurY, L4Fmt(W2BX(10, 0))
    CurX = CurX + 1900
    PosPrint CurX, CurY, L4Fmt(W2BX(11, 0))
    CurX = CurX + 1800
    PosPrint CurX, CurY, Box12String(1)
    
    CurY = CurY + YUnits * 2
    CurX = StartX + 600
    PosPrint CurX, CurY, Box13(1)
    CurX = StartX + 1600
    PosPrint CurX, CurY, Box14String(1)
    CurX = StartX + 4040
    PosPrint CurX, CurY, Box12String(2)
    
    CurY = CurY + YUnits
    CurX = StartX + 1600
    PosPrint CurX, CurY, Box14String(2)
    
    CurY = CurY + YUnits
    CurX = StartX + 600
    PosPrint CurX, CurY, Box13(2)
    CurX = StartX + 1600
    PosPrint CurX, CurY, Box14String(3)
    CurX = StartX + 4040
    PosPrint CurX, CurY, Box12String(3)
    
    CurY = CurY + YUnits
    CurX = StartX + 1600
    PosPrint CurX, CurY, Box14String(4)
    
    CurY = CurY + YUnits
    CurX = StartX + 600
    PosPrint CurX, CurY, Box13(3)
    CurX = StartX + 4040
    PosPrint CurX, CurY, Box12String(4)
    
End Sub

Private Sub PrintL4State(ByVal StartX As Long, StartY As Long)

    CurX = StartX
    CurY = StartY
    
    ' between forms nudge
    If StartY > 7000 Then
        CurY = CurY + Me.tdbnumL4Between * 20
    End If
    
    CurX = StartX + 100
    PosPrint CurX, CurY, W2BX(15, 1)
    CurX = StartX + 550
    PosPrint CurX, CurY, W2BX(15, 2)
    CurX = StartX + 2350
    PosPrint CurX, CurY, L4Fmt(W2BX(16, 1))
    ' CurX = StartX + 4100
    CurX = StartX + 4000
    PosPrint CurX, CurY, L4Fmt(W2BX(17, 1))
    
    
End Sub

Private Sub PrintL4City(ByVal StartX As Long, StartY As Long)

    CurX = StartX
    CurY = StartY
    
    ' between forms nudge
    If StartY > 7000 Then
        CurY = CurY + Me.tdbnumL4Between * 20
    End If
    
    CurX = StartX + 335
    PosPrint CurX, CurY, L4Fmt(W2BX(18, 1))
    CurX = StartX + 2220
    PosPrint CurX, CurY, L4Fmt(W2BX(19, 1))
    CurX = StartX + 3800
    PosPrint CurX, CurY, W2BX(20, 1)
    
End Sub
Private Sub cmdPrintL2_Click()
        
    W2Type = "L2"
    HorzNudge = Me.tdbnumL2Horz
    VertNudge = Me.tdbnumL2Vert
    SaveNudge User.ID, "W2L2"
    
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    
    PrintLoop

End Sub
Private Sub cmdPrintL4_Click()

    W2Type = "L4"
    
    VertNudge = Me.tdbnumL4Between
    SaveNudge User.ID, "W2L4B"
    
    HorzNudge = Me.tdbnumL4Horz
    VertNudge = Me.tdbnumL4Vert
    SaveNudge User.ID, "W2L4"

    PrtInit ("Port")
    SetFont 8, Equate.Portrait

    PrintLoop

End Sub

Private Function L4Fmt(ByVal Amount As Currency) As String
    L4Fmt = PadRight(Format(Amount, FmtA), 14)
End Function

Private Sub chkDist_Click()
    If Me.chkDist = 1 Then
        Me.fraCitiesPer.Enabled = True
        Me.optOneCity.Enabled = True
        Me.optTwoCities.Enabled = True
        Me.chkCityOver.Visible = False
        Me.tdbtxtCityOver.Visible = False
        Me.optTwoCities = True
    Else
        Me.fraCitiesPer.Enabled = False
        Me.optOneCity.Enabled = False
        Me.optTwoCities.Enabled = False
        Me.optOneCity = True
        Me.chkCityOver.Visible = True
        Me.tdbtxtCityOver.Visible = True
    End If
End Sub
Private Sub chkCityOver_Click()
    If Me.chkCityOver = 1 Then
        Me.tdbtxtCityOver = "VARIOUS"
    Else
        Me.tdbtxtCityOver = ""
    End If
End Sub

Private Sub cmdPrintW3_Click()
    
    HorzNudge = Me.tdbnumW3Horz
    VertNudge = Me.tdbnumW3Vert
    SaveNudge User.ID, "W3"
    
    If TaxYear = 2010 Then
        frmW3Print2010.Show vbModal
    ElseIf TaxYear >= 2011 Then
        frmW3Print2011.Show vbModal
    Else
        frmW3Print.Show vbModal
    End If

End Sub



