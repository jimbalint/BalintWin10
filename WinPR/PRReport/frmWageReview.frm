VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOHBUC 
   Caption         =   "Ohio BUC Report"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
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
   ScaleHeight     =   9855
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbEmpCount2 
      Height          =   375
      Left            =   6120
      TabIndex        =   20
      Top             =   8400
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calculator      =   "frmWageReview.frx":0000
      Caption         =   "frmWageReview.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":0084
      Keys            =   "frmWageReview.frx":00A2
      Spin            =   "frmWageReview.frx":00EC
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
      MaxValueVT      =   6881285
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbEmpCount1 
      Height          =   375
      Left            =   2640
      TabIndex        =   19
      Top             =   8400
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmWageReview.frx":0114
      Caption         =   "frmWageReview.frx":0134
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":01B4
      Keys            =   "frmWageReview.frx":01D2
      Spin            =   "frmWageReview.frx":021C
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
   Begin VB.Frame fraFormVersion 
      Caption         =   "  Form Version  "
      Height          =   1095
      Left            =   360
      TabIndex        =   36
      Top             =   8400
      Width           =   2055
      Begin VB.OptionButton optFormSep2010 
         Caption         =   "Sept 2010"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optFormOrig 
         Caption         =   "Original"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumStartPageNum 
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calculator      =   "frmWageReview.frx":0244
      Caption         =   "frmWageReview.frx":0264
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":02D2
      Keys            =   "frmWageReview.frx":02F0
      Spin            =   "frmWageReview.frx":033A
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
   Begin VB.CheckBox chkRed 
      Caption         =   "Employer &Red form is first page in printer"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Frame fraNudge 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1920
      TabIndex        =   35
      Top             =   2520
      Width           =   7815
      Begin TDBNumber6Ctl.TDBNumber tdbnumHorzNudge 
         Height          =   300
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frmWageReview.frx":0362
         Caption         =   "frmWageReview.frx":0382
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmWageReview.frx":0408
         Keys            =   "frmWageReview.frx":0426
         Spin            =   "frmWageReview.frx":0470
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
      Begin TDBNumber6Ctl.TDBNumber tdbnumVertNudge 
         Height          =   300
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   529
         Calculator      =   "frmWageReview.frx":0498
         Caption         =   "frmWageReview.frx":04B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmWageReview.frx":053C
         Keys            =   "frmWageReview.frx":055A
         Spin            =   "frmWageReview.frx":05A4
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
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Select Supplemental Form Output"
      Height          =   705
      Left            =   4200
      TabIndex        =   27
      Top             =   1800
      Width           =   4095
      Begin VB.OptionButton optPurple 
         Caption         =   "P&urple Form"
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optPlain 
         Caption         =   "&Plain Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraReport 
      Caption         =   "Select Report"
      Height          =   675
      Left            =   1223
      TabIndex        =   25
      Top             =   900
      Width           =   3855
      Begin VB.OptionButton optReviewJournal 
         Caption         =   "Re&view Journal"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   270
         Width           =   1815
      End
      Begin VB.OptionButton optSupplement 
         Caption         =   "&Supplemental"
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumEmployeeCount 
      Height          =   300
      Left            =   1920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   529
      Calculator      =   "frmWageReview.frx":05CC
      Caption         =   "frmWageReview.frx":05EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":065E
      Keys            =   "frmWageReview.frx":067C
      Spin            =   "frmWageReview.frx":06C6
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   7755
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3375
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   10095
      _cx             =   17806
      _cy             =   5953
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
      ScrollBars      =   2
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
   Begin VB.Frame fraSelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   960
      TabIndex        =   29
      Top             =   3240
      Width           =   9015
      Begin VB.ComboBox cmbYear 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   360
         ItemData        =   "frmWageReview.frx":06EE
         Left            =   1080
         List            =   "frmWageReview.frx":06F0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   420
         Width           =   855
      End
      Begin VB.ComboBox cmbQtr 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   12
         Top             =   420
         Width           =   5175
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   360
         Left            =   7440
         TabIndex        =   13
         Top             =   420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   635
         Calendar        =   "frmWageReview.frx":06F2
         Caption         =   "frmWageReview.frx":07F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmWageReview.frx":0856
         Keys            =   "frmWageReview.frx":0874
         Spin            =   "frmWageReview.frx":08D2
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7740
         TabIndex        =   33
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1215
         TabIndex        =   31
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Qtr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   30
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort By:"
      Height          =   675
      Left            =   5543
      TabIndex        =   26
      Top             =   900
      Width           =   3975
      Begin VB.OptionButton optEmployee 
         Caption         =   "&Employee Number"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optSSN 
         Caption         =   "SS &Number"
         Height          =   240
         Left            =   2400
         TabIndex        =   3
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   8520
      TabIndex        =   23
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6480
      TabIndex        =   22
      Top             =   9000
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumWageTotal 
      Height          =   300
      Left            =   6000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   529
      Calculator      =   "frmWageReview.frx":08FA
      Caption         =   "frmWageReview.frx":091A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":0984
      Keys            =   "frmWageReview.frx":09A2
      Spin            =   "frmWageReview.frx":09EC
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
   Begin TDBNumber6Ctl.TDBNumber tdbEmpCount3 
      Height          =   375
      Left            =   8160
      TabIndex        =   21
      Top             =   8400
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calculator      =   "frmWageReview.frx":0A14
      Caption         =   "frmWageReview.frx":0A34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmWageReview.frx":0A98
      Keys            =   "frmWageReview.frx":0AB6
      Spin            =   "frmWageReview.frx":0B00
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
      MaxValueVT      =   6881285
      MinValueVT      =   5
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   320
      Left            =   353
      TabIndex        =   34
      Top             =   120
      Width           =   10035
   End
   Begin VB.Label lblTitle 
      Caption         =   "Ohio BUC Report"
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
      Left            =   4185
      TabIndex        =   28
      Top             =   495
      Width           =   2370
   End
End
Attribute VB_Name = "frmOHBUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As New ADODB.Recordset
Dim TempID As Double
Dim WageTotal As Currency
Dim EmpCount As Long
Dim LoadFlag As Boolean

Private Sub Form_Load()
    
    LoadFlag = True
    
    Me.lblCompany.Caption = PRCompany.Name

    CurrDate = Now()
       
'    cmbQtr.AddItem "1"
'    cmbQtr.AddItem "2"
'    cmbQtr.AddItem "3"
'    cmbQtr.AddItem "4"
'    cmbQtr.ListIndex = 0
'    CurrYear = Year(Now())
'    cmbYear.AddItem CurrYear
'    cmbYear.AddItem CurrYear - 1
'    cmbYear.AddItem CurrYear - 2
'    cmbYear.AddItem CurrYear - 3
'    cmbYear.AddItem CurrYear - 4
'    cmbYear.AddItem CurrYear - 5
'    cmbYear.AddItem CurrYear - 6
'    cmbYear.AddItem CurrYear - 7
'    cmbYear.AddItem CurrYear - 8
'    cmbYear.AddItem CurrYear - 9
'    cmbYear.AddItem CurrYear - 10
'
'    cmbYear.ListIndex = 0
    
    ' init year / qtr combo even if no data exists
    Dim I As Integer
    Dim J As Integer
    With Me.cmbQtr
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    With Me.cmbYear
        J = Year(Now()) + 1
        For I = 1 To 5
            .AddItem J
            J = J - 1
        Next I
        .ListIndex = 1
    End With
    
    ' select the default qtr
    Select Case Month(Now())
        Case 1
            cmbQtr.ListIndex = 3    ' Q4
            If cmbYear.ListCount > 1 Then cmbYear.ListIndex = 2
        Case 2 To 4
            cmbQtr.ListIndex = 0    ' Q1
        Case 5 To 7
            cmbQtr.ListIndex = 1    ' Q2
        Case Else
            cmbQtr.ListIndex = 2    ' Q3
    End Select
    
    
    TDBDate1 = CurrDate
    PrtDate = CurrDate
    
    tdbIntegerSet Me.tdbnumEmployeeCount
    tdbIntegerSet Me.tdbEmpCount1
    tdbIntegerSet Me.tdbEmpCount2
    tdbIntegerSet Me.tdbEmpCount3
    
    Me.tdbnumEmployeeCount.ReadOnly = True
    
    tdbAmountSet Me.tdbnumWageTotal
    Me.tdbnumWageTotal.ReadOnly = True
    Me.tdbnumEmployeeCount.MinValue = 1
    Me.tdbnumEmployeeCount.MaxValue = 99999999

    ' **** nudge setup ****
    SetNudge Me.tdbnumHorzNudge
    Me.tdbnumHorzNudge.ToolTipText = "MOVE TEXT TO THE RIGHT"
    SetNudge Me.tdbnumVertNudge
    Me.tdbnumVertNudge.ToolTipText = "MOVE TEXT DOWN"
    
    GetNudge User.ID, "OHBUC"
    Me.tdbnumHorzNudge = HorzNudge
    Me.tdbnumVertNudge = VertNudge
    
    ' **** nudge setup ****
    
    ' start page number field
    tdbIntegerSet Me.tdbnumStartPageNum
    With Me.tdbnumStartPageNum
        .Format = "0"
        .DisplayFormat = ""
        .Spin.Enabled = True
        .Spin.Visible = dbiShowAlways
        .Value = 1
        .MinValue = 1
    End With
    
    ' default form type and other defaults
    Me.optPlain = True
    Me.optFormSep2010 = True
    Me.optPurple = False
    Me.chkRed = 0
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeBUCForm & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) Then
        If PRGlobal.Var1 = "1" Then Me.chkRed = 1
        If PRGlobal.Var2 = "2" Then
            Me.optPlain = False
            Me.optPurple = True
        End If
        Me.txtTitle = PRGlobal.Var3
        If PRGlobal.Var4 <> "" Then Me.tdbnumStartPageNum.Value = PRGlobal.Var4
        
        If PRGlobal.Var5 <> "" Then
            Me.optFormOrig = True
        End If
        
    End If
    
    GetGridData

    Me.KeyPreview = True

    LoadFlag = False
    PageNumSet
    
    If Me.optFormSep2010 = True Then Me.fraOutput.Visible = False
    
    ' set cursor to first field
    Me.Show
    Me.optSupplement.SetFocus
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub optEmployee_Click()
    rs.Sort = "EmpNo"
End Sub


Private Sub optFormOrig_Click()
    If Me.optFormSep2010 = True Then
        Me.fraOutput.Visible = False
    Else
        Me.fraOutput.Visible = True
    End If
End Sub

Private Sub optFormSep2010_Click()
    If Me.optFormSep2010 = True Then
        Me.fraOutput.Visible = False
    Else
        Me.fraOutput.Visible = True
    End If
End Sub

Private Sub optReviewJournal_Click()
   fraNudge.Visible = False
   fraOutput.Visible = False
End Sub

Private Sub optSSN_Click()
    rs.Sort = "SSN"
End Sub

Private Sub optSupplement_Click()
   fraNudge.Visible = True
   fraOutput.Visible = True
End Sub

Private Sub txtDate_Change()
   PrtDate = TxtDate
End Sub

Private Sub cmdOK_Click()

    If rs.RecordCount = 0 Then
        MsgBox "No Payroll Data to report", vbInformation, "Ohio BUC Form"
        Exit Sub
    End If

    qYear = cmbYear
    qQuarter = cmbQtr

    If Me.optSupplement Then
        
        ' save the form type and title per company
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeBUCForm & _
                    " AND UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = False Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeBUCForm
            PRGlobal.UserID = PRCompany.CompanyID
            PRGlobal.Save (Equate.RecAdd)
        End If
        If Me.chkRed = 1 Then
            PRGlobal.Var1 = "1"
        Else
            PRGlobal.Var1 = ""
        End If
        
        If Me.optPlain = True Then
            FormColor = "Plain"
            PRGlobal.Var2 = "1"
        Else
            FormColor = "Purple"
            PRGlobal.Var2 = "2"
        End If
        PRGlobal.Var3 = Me.txtTitle & ""
        PRGlobal.Var4 = Me.tdbnumStartPageNum.Value
        If Me.optFormOrig = True Then
            PRGlobal.Var5 = "Orig"
        Else
            PRGlobal.Var5 = ""
        End If
        
        PRGlobal.Save (Equate.RecPut)
        
        HorzNudge = Me.tdbnumHorzNudge
        VertNudge = Me.tdbnumVertNudge
        SaveNudge User.ID, "OHBUC"
        
        If Me.optFormSep2010 = True Then
            OHBUCRed201009
        Else
            If Me.chkRed = 1 Then
                If Me.optFormOrig = True Then
                    OHBUCRed FormColor
                End If
            Else
                OHBUCPurple FormColor, False
            End If
        End If
        
    Else
        Nudge = 0
        OHBUCJournal
    End If

End Sub

Private Sub cmdExit_Click()
   GoBack
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 7 Then
        Cancel = True
    End If
End Sub

Private Function GetEECount(ByVal YrMth As String) As Integer

Dim rsEECount As New ADODB.Recordset

    SQLString = " SELECT DISTINCT(EmployeeID) FROM PRHist WHERE " & _
                " YearMonth = " & YrMth & _
                " AND Gross > 0 "
    rsInit SQLString, cn, rsEECount
    GetEECount = rsEECount.RecordCount()

End Function

Public Sub GetGridData()
Dim YM1, YM2 As Long
Dim LastEmp As Double
Dim GrossAmt As Currency
        
    rs.CursorLocation = adUseClient
    rs.Fields.Append "EmpNo", adDouble
    rs.Fields.Append "SSN", adDouble
    rs.Fields.Append "EmpName", adVarChar, 80, adFldIsNullable
    rs.Fields.Append "LastName", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "FirstName", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "MidInit", adVarChar, 1, adFldIsNullable
    rs.Fields.Append "Gross", adCurrency
    rs.Fields.Append "NoWeeks", adDouble
    rs.Fields.Append "EmpID", adDouble
    rs.Fields.Append "PayCount", adDouble
    rs.Fields.Append "PaysPerYear", adDouble

    rs.Open , , adOpenDynamic, adLockOptimistic
    YM1 = Me.cmbYear * 100
    If Me.cmbQtr = 1 Then YM1 = YM1 + 1
    If Me.cmbQtr = 2 Then YM1 = YM1 + 4
    If Me.cmbQtr = 3 Then YM1 = YM1 + 7
    If Me.cmbQtr = 4 Then YM1 = YM1 + 10
    YM2 = YM1 + 2
    
    ' get employee counts
    Me.tdbEmpCount1 = GetEECount(YM1)
    Me.tdbEmpCount2 = GetEECount(YM1 + 1)
    Me.tdbEmpCount3 = GetEECount(YM2)
    
    ' *** OHIO ONLY *** - StateID = 36
    SQLString = "SELECT * FROM PRHist WHERE PRHist.YearMonth >= " & YM1 & _
                " AND PRHist.YearMonth <= " & YM2 & _
                " AND StateID = 36"

    If Not PRHist.GetBySQL(SQLString) Then
        If LoadFlag = False Then
            MsgBox "No payroll history found!!", vbExclamation
        End If
        Exit Sub
    End If
    
    Do

        SQLString = "EmpID= " & PRHist.EmployeeID
        rs.Find SQLString, 0, adSearchForward, 1
                
        If rs.EOF Then
            
            If Not PREmployee.GetByID(PRHist.EmployeeID) Then
                MsgBox "Employee Not Found: " & PRHist.EmployeeID, vbCritical
                GoBack
            End If
        
            rs.AddNew
'            rs!SSN = Format(PREmployee.SSN, "000-00-0000")
            rs!EmpNo = PREmployee.EmployeeNumber
            rs!SSN = PREmployee.SSN
            rs!EmpName = PREmployee.LFName
            rs!LastName = Mid(PREmployee.LastName, 1, 40)
            rs!FirstName = Mid(PREmployee.FirstName, 1, 40)
            rs!MidInit = Mid(PREmployee.MidInit, 1, 1)
            rs!Gross = 0
            rs!EmpID = PRHist.EmployeeID
            rs!PayCount = 0
            rs!PaysPerYear = PREmployee.PaysPerYear
            rs.Update
        
        End If
            
        rs!PayCount = rs!PayCount + 1
        rs!Gross = rs!Gross + PRHist.SUNWageBase
        rs.Update
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop
    
    ' calc the weeks worked
    rs.MoveFirst
    Do
        If rs!Gross = 0 Then
            rs.Delete
        Else
            If rs!PaysPerYear = 52 Then
                rs!NoWeeks = rs!PayCount * 1.09
            ElseIf rs!PaysPerYear = 26 Then
                rs!NoWeeks = rs!PayCount * 2.17
            ElseIf rs!PaysPerYear = 24 Then
                rs!NoWeeks = rs!PayCount * 2.17
            ElseIf rs!PaysPerYear = 12 Then
                rs!NoWeeks = rs!PayCount * 4.34
            Else
                rs!NoWeeks = 13
            End If
            rs!NoWeeks = Round(rs!NoWeeks, 0)
            If rs!NoWeeks > 13 Then rs!NoWeeks = 13
            rs.Update
        End If
        rs.MoveNext
    Loop Until rs.EOF
    
    If Me.optEmployee Then
        rs.Sort = "EmpID"
    Else
        rs.Sort = "SSN"
    End If
    
    SetGrid rs, fg
    
    fg.ColFormat(0) = "999-99-9999"
    fg.ColWidth(2) = 2600
    fg.ColWidth(3) = 1200
    fg.ColWidth(4) = 1600
    fg.ColWidth(5) = 0
    
    fg.TextMatrix(0, 4) = "Weeks Worked"

    If fg.Rows >= 2 Then fg.Select 1, 4

    CalcTotals

End Sub

Private Sub cmdDelete_Click()
    If MsgBox("OK to delete: " & rs!EmpName, vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    Else
        rs.Delete
        CalcTotals
    End If
End Sub

Private Sub CalcTotals()
    
    WageTotal = 0
    EmpCount = 0
        
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    Do
        EmpCount = EmpCount + 1
        WageTotal = WageTotal + rs!Gross
        rs.MoveNext
    Loop Until rs.EOF
    rs.MoveFirst

    Me.tdbnumEmployeeCount.Value = EmpCount
    
    ' calculated above...
    ' Me.tdbEmpCount1 = EmpCount
    ' Me.tdbEmpCount2 = EmpCount
    ' Me.tdbEmpCount3 = EmpCount
    
    Me.tdbnumWageTotal.Value = WageTotal

End Sub

Private Sub cmbQtr_Click()
    If LoadFlag Then Exit Sub
    rs.Close
    Set rs = Nothing
    GetGridData
End Sub

Private Sub cmbYear_Click()
    If LoadFlag Then Exit Sub
    rs.Close
    Set rs = Nothing
    GetGridData
End Sub

Private Sub PageNumSet()
    If Me.chkRed = 1 Then
        Me.tdbnumStartPageNum = 1
        Me.tdbnumStartPageNum.Enabled = False
    Else
        Me.tdbnumStartPageNum.Enabled = True
    End If
End Sub
Private Sub chkRed_Click()
    PageNumSet
End Sub


Private Sub TDBNumber2_Change()

End Sub
