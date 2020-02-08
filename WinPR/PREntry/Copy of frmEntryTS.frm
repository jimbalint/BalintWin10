VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmEntryTS 
   Caption         =   "Payroll Data Entry"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15165
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
   ScaleHeight     =   9900
   ScaleWidth      =   15165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave2 
      Caption         =   "SAVE F10"
      Height          =   735
      Left            =   11880
      Picture         =   "frmEntryTS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "&REVERSE"
      Height          =   375
      Left            =   3000
      TabIndex        =   34
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   9360
      Width           =   1215
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumBChecks 
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   600
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":030A
      Caption         =   "frmEntryTS.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0396
      Keys            =   "frmEntryTS.frx":03B4
      Spin            =   "frmEntryTS.frx":03FE
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
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumBRegHrs 
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   120
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0426
      Caption         =   "frmEntryTS.frx":0446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":04AA
      Keys            =   "frmEntryTS.frx":04C8
      Spin            =   "frmEntryTS.frx":0512
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
   Begin VB.CommandButton cmdAddEarn 
      Caption         =   "ADD"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "SKIP F11"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   9360
      Width           =   1215
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumCheckNum 
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   1440
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":053A
      Caption         =   "frmEntryTS.frx":055A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":05BC
      Keys            =   "frmEntryTS.frx":05DA
      Spin            =   "frmEntryTS.frx":0624
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE F10"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdReCalc 
      Caption         =   "RE-CALC"
      Height          =   375
      Left            =   13080
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClearManual 
      Caption         =   "CLR MANUAL"
      Height          =   375
      Left            =   11400
      TabIndex        =   19
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSetManual 
      Caption         =   "SET MANUAL"
      Height          =   375
      Left            =   9720
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumDirDepTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   9
      Top             =   8040
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":064C
      Caption         =   "frmEntryTS.frx":066C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":06E2
      Keys            =   "frmEntryTS.frx":0700
      Spin            =   "frmEntryTS.frx":074A
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
   Begin VB.CommandButton cmdEmpAdd 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmpEdit 
      Caption         =   "&EDIT"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   375
      Left            =   13920
      TabIndex        =   1
      Top             =   9360
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEMP 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5535
      _cx             =   9763
      _cy             =   11668
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumCheckTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   10
      Top             =   8520
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0772
      Caption         =   "frmEntryTS.frx":0792
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":07FE
      Keys            =   "frmEntryTS.frx":081C
      Spin            =   "frmEntryTS.frx":0866
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumNetPayTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   11
      Top             =   7560
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":088E
      Caption         =   "frmEntryTS.frx":08AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":091E
      Keys            =   "frmEntryTS.frx":093C
      Spin            =   "frmEntryTS.frx":0986
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumDedTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   12
      Top             =   7080
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":09AE
      Caption         =   "frmEntryTS.frx":09CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0A40
      Keys            =   "frmEntryTS.frx":0A5E
      Spin            =   "frmEntryTS.frx":0AA8
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumERNTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   13
      Top             =   6120
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0AD0
      Caption         =   "frmEntryTS.frx":0AF0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0B60
      Keys            =   "frmEntryTS.frx":0B7E
      Spin            =   "frmEntryTS.frx":0BC8
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
   Begin VSFlex8Ctl.VSFlexGrid fgDED 
      Height          =   4215
      Left            =   5760
      TabIndex        =   14
      Top             =   5400
      Width           =   5655
      _cx             =   9975
      _cy             =   7435
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
   Begin VSFlex8Ctl.VSFlexGrid fgERN 
      Height          =   2895
      Left            =   5760
      TabIndex        =   15
      Top             =   1920
      Width           =   9375
      _cx             =   16536
      _cy             =   5106
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumTaxTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   16
      Top             =   6600
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0BF0
      Caption         =   "frmEntryTS.frx":0C10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0C76
      Keys            =   "frmEntryTS.frx":0C94
      Spin            =   "frmEntryTS.frx":0CDE
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumHrTotal 
      Height          =   375
      Left            =   11760
      TabIndex        =   17
      Top             =   5640
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0D06
      Caption         =   "frmEntryTS.frx":0D26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0D90
      Keys            =   "frmEntryTS.frx":0DAE
      Spin            =   "frmEntryTS.frx":0DF8
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumBOHrs 
      Height          =   375
      Left            =   9360
      TabIndex        =   26
      Top             =   480
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0E20
      Caption         =   "frmEntryTS.frx":0E40
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0EAC
      Keys            =   "frmEntryTS.frx":0ECA
      Spin            =   "frmEntryTS.frx":0F14
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumBTlHrs 
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   840
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":0F3C
      Caption         =   "frmEntryTS.frx":0F5C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":0FC4
      Keys            =   "frmEntryTS.frx":0FE2
      Spin            =   "frmEntryTS.frx":102C
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumBRegErn 
      Height          =   375
      Left            =   11880
      TabIndex        =   29
      Top             =   120
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":1054
      Caption         =   "frmEntryTS.frx":1074
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":10DC
      Keys            =   "frmEntryTS.frx":10FA
      Spin            =   "frmEntryTS.frx":1144
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumBOEarng 
      Height          =   375
      Left            =   11880
      TabIndex        =   30
      Top             =   480
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":116C
      Caption         =   "frmEntryTS.frx":118C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":11FC
      Keys            =   "frmEntryTS.frx":121A
      Spin            =   "frmEntryTS.frx":1264
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumBTlEarng 
      Height          =   375
      Left            =   11880
      TabIndex        =   31
      Top             =   840
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      Calculator      =   "frmEntryTS.frx":128C
      Caption         =   "frmEntryTS.frx":12AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmEntryTS.frx":1318
      Keys            =   "frmEntryTS.frx":1336
      Spin            =   "frmEntryTS.frx":1380
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
   Begin VB.Label lblDept 
      Caption         =   "DeptInfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   37
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblEEName 
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   855
      Left            =   2760
      TabIndex        =   35
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "Employee Totals:"
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
      Left            =   11760
      TabIndex        =   28
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "D E D U C T I O N S"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "E A R N I N G S"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblCheckDate 
      Caption         =   "Check Date"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblPEDate 
      Caption         =   "PE Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2175
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
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmEntryTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EMP As New ADODB.Recordset
Dim DED As New ADODB.Recordset
Dim ERN As New ADODB.Recordset
Dim JC As New ADODB.Recordset
Dim ERItem As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim rsDedBasis As New ADODB.Recordset

Dim FirstFlag As Boolean
Dim EditFlag As Boolean

Dim ERNDrop As String
Dim DEDDrop As String
Dim CityDrop As String
Dim StateDrop As String
Dim StateAbbrev As String
Dim JobDrop As String

Dim SortCol, SortOrder, GridFocus As Byte

Dim SSPct, MedPct As Double
Dim SSMax, FUNMax, SUNMax As Currency
Dim FedAllow, OHAllow, OHSDAllow As Currency
Dim FWTAGI, SWTAGI, p1, P2, P3, P4 As Currency
Dim TaxYear, TaxMonth As Integer

Dim NextCheckNumber, HiCheckNumber As Long
Dim EECount As Long

Dim YTDSSWage, YTDFUNWage, YTDSUNWage As Currency
Dim SWTWageTL, SWTTaxTL, CWTWageTL, CWTTaxTL As Currency

Dim CourtRate, CourtTax As Currency
Dim CourtCityID As Long
Dim CourtCityName As String
Dim CourtAdd As Byte

Dim DedCount, ErnCount, i, j, k As Integer

Public HiCheckNum, StartCheckNumber, BatchID As Long
Dim NotInNetTotal As Currency

Dim JobDist As Boolean
Dim TimeSheet As Boolean
Dim SQLString2 As String

Private Sub Form_Load()

    ' define the recordset for ERN and DED items and amounts
    DefineRS
    
    ' use Job Dist???
    If frmSelTimeSheets.UseDist = True Then
        
        JobDist = True
        SQLString = "SELECT * FROM JCJob"
        
        ' *** only jobs w/ City Rate filled in
        SQLString = "SELECT * FROM JCJob WHERE CityID <> 0"
        
        If JCJob.GetBySQL(SQLString) Then
            Do
                JC.AddNew
                JC!JobID = JCJob.JobID
                JC!CityID = JCJob.CityID
                                
                ' ******************
                ' *** stuff it   ***
                ' JC!CityID = (JCJob.JobID Mod 10) + 1
                ' ******************
                
                If PRCity.GetByID(JC!CityID) Then
                    JC!CityRate = PRCity.CityRate
                Else
                    JC!CityRate = 0
                End If
                If Trim(JCJob.FullName) <> "" Then
                    JC!Name = Mid(JCJob.FullName, 1, 30)
                ElseIf Trim(JCJob.Name) <> "" Then
                    JC!Name = Mid(JCJob.Name, 1, 30)
                Else
                    X = Trim(JCJob.FirstName) & " " & Trim(JCJob.MidInit) & " " & Trim(JCJob.LastName)
                    If X = "" Then
                        JC!Name = "Job ID: " & JCJob.JobID
                    Else
                        JC!Name = Mid(X, 1, 30)
                    End If
                End If
                JC!Name = Trim(JC!Name)
                JC.Update
                
                If JCJob.GetNext = False Then Exit Do
            Loop
        End If
        
        JC.Sort = "Name"
    Else
        JobDist = False
        JobDrop = ""
    End If

    FirstFlag = True
    Me.lblCompanyName = Trim(PRCompany.Name)

    ' ##### stuff it #####
    ' BatchID = 999
    ' BatchID = 16
    ' BatchID = 47
    
    If Not PRBatch.GetByID(BatchID) Then
        MsgBox "Batch Error: " & BatchID, vbCritical
        End
    End If
    
    Me.lblPEDate = "PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy")
    Me.lblCheckDate = "Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
    NextCheckNumber = nNull(StartCheckNumber)
    
'    If Not PRBatch.GetByID(BatchID) Then
'        MsgBox "PRBatch NF: " & BatchID, vbCritical
'        End
'    End If
    
    ' ##### stuff it #####

    PRHist.OpenRS
    PRDist.OpenRS
    PRItemHist.OpenRS

    ' initialize the employee grid
    LoadEmpGrid
    
    ' load PRHist - assign check number fields to the EMP recordset
    ' LoadHistory

    ' resort the employee grid
    ' EMP.Sort = "EmployeeNumber"
    
    ' sort column
    ' check # column
     
    Select Case frmNewBatch.SortOrder
        Case PREquate.SortOrderName
            SortOrder = 1       ' ascending
            SortCol = 1         ' ee name col
            fgEMP.Cell(flexcpFontBold, 0, 1) = True
            SQLString = "EmployeeName"
        Case PREquate.SortOrderNumber
            SortOrder = 1       ' ascending
            SortCol = 0         ' ee # col
            fgEMP.Cell(flexcpFontBold, 0, 0) = True
            SQLString = "EmployeeNumber"
        Case PREquate.SortOrderDeptNumber
            SortCol = 11
            SortOrder = 1
            SQLString = "DptEE"
        Case PREquate.SortOrderDeptName
            SortCol = 11
            SortOrder = 1
            SQLString = "DptEE"
        Case Else
            MsgBox "Form Error?", vbExclamation, "PR Entry"
            GoBack
    End Select
    EMP.Sort = SQLString
    
    fgEMP.Select 1, 0, 1, 2
    fgEMP.ShowCell 1, 2

    ' assign the recordsets to the grids
    SetGrid ERN, fgERN
    SetGrid DED, fgDED

    fgERN.ColWidth(0) = 1400        ' title
    fgERN.TextMatrix(0, 0) = "Earning Type"
    
    fgERN.ColWidth(1) = 1000        ' hours
    fgERN.ColWidth(2) = 1000        ' rate
    fgERN.ColWidth(3) = 1500        ' amount
    
    fgERN.ColWidth(4) = 270         ' amount manual
    fgERN.TextMatrix(0, 4) = "M"
   
    If JobDist = True Then
        fgERN.ColWidth(0) = 1400        ' title
        fgERN.ColWidth(5) = 2500        ' job
        fgERN.ColWidth(6) = 1500        ' city
    Else
        fgERN.ColWidth(0) = 1800        ' title
        fgERN.ColWidth(5) = 0           ' job
        fgERN.ColWidth(6) = 2500        ' city
    End If
    
    fgERN.TextMatrix(0, 5) = "Job Name"
    fgERN.TextMatrix(0, 6) = "City Name"
    
    ' hide the other columns of the ERN grid
    For i = 7 To 35
        fgERN.ColWidth(i) = 0
    Next i
    
    If JobDist = False Then
        fgERN.TextMatrix(0, 5) = "City"
    Else
        fgERN.TextMatrix(0, 5) = "Job"
    End If
    
    ' Earnings Hours Column
    fgERN.ColFormat(1) = "##0.00"
    
    fgDED.ScrollBars = flexScrollBarVertical
    
    fgDED.ColWidth(0) = 1700        ' title
    fgDED.TextMatrix(0, 0) = "Deduction Type"
    
    fgDED.ColWidth(1) = 2000        ' desc
    fgDED.TextMatrix(0, 1) = "Deduction Basis"
    
    fgDED.ColWidth(2) = 1500        ' amount
    
    fgDED.ColWidth(3) = 270         ' amount manual
    fgDED.TextMatrix(0, 3) = "M"
    
    DfltJobID = 0
    DfltCityID = 999999
    DfltStateID = 36
    
    ' =========== ************************
    ' PRBatch.YearMonth = 200811
    
    TaxMonth = PRBatch.YearMonth Mod 100
    TaxYear = Int(PRBatch.YearMonth / 100)
    
    ' get tax parameters
    SSPct = PRGlobal.GetAmount(PREquate.GlobalTypeSSPct, TaxYear)
    SSMax = PRGlobal.GetAmount(PREquate.GlobalTypeSSMax, TaxYear)
    MedPct = PRGlobal.GetAmount(PREquate.GlobalTypeMEDPct, TaxYear)
    FedAllow = PRGlobal.GetAmount(PREquate.GlobalTypeFWTAllow, TaxYear)
    OHAllow = PRGlobal.GetAmount(PREquate.GLobalTypeOHAllow, TaxYear)
    OHSDAllow = PRGlobal.GetAmount(PREquate.GlobalTypeOHSDTaxAllow, TaxYear)
    FUNMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, TaxYear)
    ' =========== ************************
    
    tdbIntegerSet Me.tdbnumCheckNum
    Me.tdbnumCheckNum.Format = "#########0"
    Me.tdbnumCheckNum.DisplayFormat = ""
    
    ' select the first employee
    SetDataGrids
    CalcGrids
    BatchTotals
    
    ' disable these buttons for now
    Me.cmdEmpEdit.Enabled = False
    
    ' check numbering set
    If StartCheckNumber = 0 Then    ' not a brand new batch
        SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & BatchID & _
                    " ORDER BY CheckNumber DESC"
        If Not PRHist.GetBySQL(SQLString) Then
            SQLString = "SELECT * FROM PRHist ORDER BY CheckNumber DESC"
            If Not PRHist.GetBySQL(SQLString) Then
                HiCheckNumber = 100
            Else
                HiCheckNumber = PRHist.CheckNumber + 1
            End If
        Else
            HiCheckNumber = PRHist.CheckNumber + 1
        End If
    Else
        HiCheckNumber = nNull(StartCheckNumber)
    End If
    
    If EMP!CheckNumber = 0 Then
        Me.tdbnumCheckNum = HiCheckNumber
    Else
        Me.tdbnumCheckNum = EMP!CheckNumber
    End If
    
    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    
    Me.lblEEName = PREmployee.EmployeeNumber & " " & PREmployee.LFName
    
    If PRDepartment.GetByID(PREmployee.DepartmentID) Then
        Me.lblDept = PRDepartment.DepartmentNumber & " " & PRDepartment.Name
    Else
        Me.lblDept = ""
    End If
    
End Sub


Private Sub cmdExit_Click()
    
    PRCompany.LastCheckNum = NextCheckNumber - 1
    PRCompany.Save (Equate.RecPut)
    
    On Error Resume Next
    
    EMP.Close
    ERN.Close
    DED.Close
    JC.Close
    ERItem.Close
    rs.Close
    rsDedBasis.Close
    
    Set EMP = Nothing
    Set ERN = Nothing
    Set DED = Nothing
    Set JC = Nothing
    Set ERItem = Nothing
    Set rs = Nothing
    Set rsDedBasis = Nothing
    
    EditFlag = False
    
    On Error GoTo 0
    
    Unload Me

End Sub
Private Sub fgEMP_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    ' warn to save data if edits made
    If EditFlag Then
        If MsgBox("Save Entries?", vbQuestion + vbYesNo, "Payroll Data Entry") = vbYes Then
            cmdSave_Click
        End If
    End If

End Sub

Private Sub fgEMP_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If Not FirstFlag Then

        ' update and clear the ERN and DED record sets
        ERN.MoveFirst
        Do
            ERN.Delete
            ERN.MoveNext
            If ERN.EOF Then Exit Do
        Loop
                
        DED.MoveFirst
        Do
            DED.Delete
            DED.MoveNext
            If DED.EOF Then Exit Do
        Loop
        
        If rsDedBasis.RecordCount > 0 Then
            rsDedBasis.MoveFirst
            Do
                rsDedBasis.Delete
                rsDedBasis.MoveNext
            Loop Until rsDedBasis.EOF
        End If
        
        SetDataGrids
        CalcGrids
        BatchTotals
    
        EMP!CheckNumber = nNull(EMP!CheckNumber)
        EMP.Update
    
        If EMP!CheckNumber <> 0 Then
            Me.tdbnumCheckNum = EMP!CheckNumber
        Else
            Me.tdbnumCheckNum = NextCheckNumber
        End If
    
    End If

    ' position the cursor
    fgERN.ShowCell 1, 1
    fgERN.Select 1, 1
    fgERN.Refresh
    
    fgDED.ShowCell 1, 2
    fgDED.Select 1, 2
    fgDED.Refresh

    Me.lblEEName = PREmployee.EmployeeNumber & " " & PREmployee.LFName

    If PRDepartment.GetByID(PREmployee.DepartmentID) Then
        Me.lblDept = PRDepartment.DepartmentNumber & " " & PRDepartment.Name
    Else
        Me.lblDept = ""
    End If

End Sub
Private Sub fgERN_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    fgERN.CellBackColor = vbYellow

End Sub

Private Sub fgERN_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    GridFocus = 1
    CalcGrids
    fgERN.CellBackColor = vbDefault

    ' Job ID changed - set the CityID
    ' If JobDist = True And Col = 5 And ERN!AmountManual = False Then
    If JobDist = True And Col = 5 Then
        If ERN.EOF Then ERN.MoveLast        ' if on last line
        If ERN!AmountManual = False Then
            JC.Find "JobID = " & ERN!JobID, 0, adSearchForward, 1
            If JC.EOF = False Then
                ERN!CityID = JC!CityID
                ERN.Update
                CalcGrids
            End If
        End If
    End If

    EditFlag = True

End Sub

Private Sub fgDED_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' don't allow edit of description column
    If Col = 1 Then
        Cancel = True
        Exit Sub
    End If
    
    fgDED.CellBackColor = vbYellow

End Sub
Private Sub fgDED_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    GridFocus = 2
    CalcGrids
    fgDED.CellBackColor = vbDefault

    EditFlag = True

End Sub
Private Sub LoadEmpGrid()

Dim WkcID As Long

    EECount = 0

    ' set up temp record set for employees
    EMP.CursorLocation = adUseClient
   
    EMP.Fields.Append "EmployeeNumber", adDouble
    EMP.Fields.Append "EmployeeName", adVarChar, 60, adFldIsNullable
    EMP.Fields.Append "CheckNumber", adDouble
    EMP.Fields.Append "Salaried", adBoolean
    EMP.Fields.Append "Saved", adBoolean
    EMP.Fields.Append "HistID", adDouble
    EMP.Fields.Append "HistFlag", adBoolean
    EMP.Fields.Append "EmployeeID", adDouble
    EMP.Fields.Append "TempID", adDouble
    EMP.Fields.Append "WkcPct", adDouble
    EMP.Fields.Append "DptNumber", adDouble
    EMP.Fields.Append "DptEE", adVarChar, 60, adFldIsNullable
    
    EMP.Open , , adOpenDynamic, adLockOptimistic

    ' load employees from existing history records
    SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & BatchID
    If PRHist.GetBySQL(SQLString) Then
        Do
            
            If Not PREmployee.GetByID(PRHist.EmployeeID) Then
                MsgBox "Employee NF: " & PRHist.EmployeeID, vbCritical
                End
            End If
            
            EMP.AddNew
            EMP!EmployeeNumber = PREmployee.EmployeeNumber
            EMP!EmployeeID = PRHist.EmployeeID
            EMP!EmployeeName = Trim(PREmployee.LFName) & ""
            EMP!CheckNumber = PRHist.CheckNumber
            If PREmployee.Salaried Then EMP!Salaried = True
            EMP!Saved = True
            EMP!HistID = PRHist.HistID
            EMP!HistFlag = True
                
            EECount = EECount + 1
            EMP!TempID = EECount
            
            If PRDepartment.GetByID(PREmployee.DepartmentID) Then
                EMP!DptNumber = PRDepartment.DepartmentNumber
            Else
                EMP!DptNumber = 0
            End If
            
            If frmNewBatch.cmbSortOrder.ListIndex = PREquate.SortOrderDeptNumber Then
                EMP!DptEE = Mid(Format(EMP!DptNumber, "0000") & PREmployee.EmployeeNumber, 1, 60)
            ElseIf frmNewBatch.cmbSortOrder.ListIndex = PREquate.SortOrderDeptName Then
                EMP!DptEE = Mid(Format(EMP!DptNumber, "0000") & PREmployee.LFName, 1, 60)
            Else
                EMP!DptEE = ""
            End If
            EMP.Update
            
            If Not PRHist.GetNext Then Exit Do
        
        Loop
    
    End If

    ' load the employee temp ADO record set
    ' load active employees not already loaded from PRHist
    SQLString = "SELECT * FROM PREmployee WHERE Inactive = 0"
    If PREmployee.GetBySQL(SQLString) Then
    
        Do
        
            ' don't add if already exists
            SQLString = "EmployeeID = " & PREmployee.EmployeeID
            EMP.Find SQLString, 0, adSearchForward, 1
        
            If EMP.EOF Then
        
                EMP.AddNew
                EMP!EmployeeNumber = PREmployee.EmployeeNumber
                EMP.Fields("EmployeeID") = PREmployee.EmployeeID
                EMP.Fields("EmployeeName") = Trim(PREmployee.LFName)
                EMP.Fields("CheckNumber") = 0
                If PREmployee.Salaried Then EMP!Salaried = True
                EMP!Saved = False
                EMP.Fields("HistID") = 0
                EMP.Fields("HistFlag") = False
                
                ' wkc comp pct
                WkcID = 0
                If PREmployee.WkcUseDept = 1 Then
                    If PRDepartment.GetByID(PREmployee.DepartmentID) Then
                        WkcID = PRDepartment.WkcCat
                    End If
                Else
                    WkcID = PREmployee.WkcCat
                End If
                If WkcID = 0 Then
                    EMP!WkcPct = 0
                Else
                    If PRGlobal.GetByID(WkcID) Then
                        EMP!WkcPct = PRGlobal.Percent
                    End If
                End If
            
                If PRDepartment.GetByID(PREmployee.DepartmentID) Then
                    EMP!DptNumber = PRDepartment.DepartmentNumber
                Else
                    EMP!DptNumber = 0
                End If
                
                If frmNewBatch.cmbSortOrder.ListIndex = PREquate.SortOrderDeptNumber Then
                    EMP!DptEE = Mid(Format(EMP!DptNumber, "0000") & PREmployee.EmployeeNumber, 1, 60)
                ElseIf frmNewBatch.cmbSortOrder.ListIndex = PREquate.SortOrderDeptName Then
                    EMP!DptEE = Mid(Format(EMP!DptNumber, "0000") & PREmployee.LFName, 1, 60)
                Else
                    EMP!DptEE = ""
                End If
                EMP.Update
            
            End If
            
            If Not PREmployee.GetNext Then Exit Do
            
        Loop

    End If

    ' sort emp rs
    With frmNewBatch.cmbSortOrder
        If .ListIndex = PREquate.SortOrderNumber Then
            SQLString2 = "EmployeeNumber"
        ElseIf .ListIndex = PREquate.SortOrderName Then
            SQLString2 = "EmployeeName"
        ElseIf .ListIndex = PREquate.SortOrderDeptNumber Then
            SQLString2 = "DptNumber, EmployeeNumber"
        ElseIf .ListIndex = PREquate.SortOrderDeptName Then
            SQLString2 = "DptNumber, EmployeeName"
        Else
            MsgBox "Form Error? ", vbExclamation, "PR Entry"
            GoBack
        End If
    End With

    ' assign the temporary recordset to the grid
    SetGrid EMP, fgEMP
    
    EMP.Sort = SQLString2
    
    ' not editable
    fgEMP.Editable = flexEDNone
    
    ' select the entire row
    fgEMP.SelectionMode = flexSelectionByRow

    fgEMP.ScrollBars = flexScrollBarVertical

    ' format the flex grid
    fgEMP.ColWidth(0) = 900
    fgEMP.ColWidth(1) = 2000
    fgEMP.ColWidth(2) = 1000
    fgEMP.ColWidth(3) = 500
    fgEMP.ColWidth(4) = 600

    ' column titles
    fgEMP.TextMatrix(0, 0) = "EE#"
    fgEMP.TextMatrix(0, 1) = "N A M E"
    fgEMP.TextMatrix(0, 2) = "Check# +"
    fgEMP.TextMatrix(0, 3) = "SAL"
    fgEMP.TextMatrix(0, 4) = "Saved"

    ' sort column
    ' check # column
    With frmNewBatch.cmbSortOrder
        If .ListIndex = PREquate.SortOrderName Then
            SortOrder = 1       ' ascending
            SortCol = 1         ' ee name col
            fgEMP.Cell(flexcpFontBold, 0, 1) = True
        Else: .ListIndex = PREquate.SortOrderNumber
            SortOrder = 1       ' ascending
            SortCol = 0         ' ee # col
            fgEMP.Cell(flexcpFontBold, 0, 0) = True
        ' ???
        End If
    End With

    ' fixed fields setup
    tdbAmountSet Me.tdbnumCheckTotal
    tdbAmountSet Me.tdbnumDirDepTotal
    tdbAmountSet Me.tdbnumNetPayTotal
    tdbAmountSet Me.tdbnumERNTotal
    tdbAmountSet Me.tdbnumDedTotal
    tdbAmountSet Me.tdbnumTaxTotal
    tdbAmountSet Me.tdbnumHrTotal
    
    ' batch total fields
    tdbAmountSet Me.tdbnumBRegHrs
    tdbAmountSet Me.tdbnumBOHrs
    tdbAmountSet Me.tdbnumBTlHrs
    tdbAmountSet Me.tdbnumBRegErn
    tdbAmountSet Me.tdbnumBOEarng
    tdbAmountSet Me.tdbnumBTlEarng
    
    tdbAmountSet Me.tdbnumBChecks
    Me.tdbnumBChecks.Format = "########0"
    Me.tdbnumBChecks.DisplayFormat = ""
    
    ' not editable
    Me.tdbnumCheckTotal.ReadOnly = True
    Me.tdbnumDirDepTotal.ReadOnly = True
    Me.tdbnumNetPayTotal.ReadOnly = True
    Me.tdbnumERNTotal.ReadOnly = True
    Me.tdbnumDedTotal.ReadOnly = True
    Me.tdbnumTaxTotal.ReadOnly = True
    Me.tdbnumHrTotal.ReadOnly = True

    Me.tdbnumBRegHrs.ReadOnly = True
    Me.tdbnumBOHrs.ReadOnly = True
    Me.tdbnumBTlHrs.ReadOnly = True
    Me.tdbnumBRegErn.ReadOnly = True
    Me.tdbnumBOEarng.ReadOnly = True
    Me.tdbnumBTlEarng.ReadOnly = True

    Me.tdbnumBChecks.ReadOnly = True

End Sub

Private Sub DefineRS()

    ' ========================================================================
    
    ' set up records for ERN grid
    ERN.CursorLocation = adUseClient
    ERN.Fields.Append "Title", adDouble
    ERN.Fields.Append "Hours", adSingle
    ERN.Fields.Append "Rate", adCurrency
    ERN.Fields.Append "Amount", adCurrency
    ERN.Fields.Append "AmountManual", adBoolean
    ERN.Fields.Append "JobID", adDouble
    ERN.Fields.Append "CityID", adDouble

    ' tax flag info
    ERN.Fields.Append "NoSSTax", adBoolean
    ERN.Fields.Append "NoMEDTax", adBoolean
    ERN.Fields.Append "NoFWTTax", adBoolean
    ERN.Fields.Append "NoSWTTax", adBoolean
    ERN.Fields.Append "NoCWTTax", adBoolean
    ERN.Fields.Append "NoFUNTax", adBoolean
    ERN.Fields.Append "NoSUNTax", adBoolean
    ERN.Fields.Append "Basis", adInteger
    ERN.Fields.Append "AmtPct", adCurrency
    ERN.Fields.Append "Tips", adBoolean
    ERN.Fields.Append "NotInNet", adBoolean
    ERN.Fields.Append "EmployerItemID", adDouble
    ERN.Fields.Append "Salary", adBoolean

    ERN.Fields.Append "CityManual", adBoolean
    ERN.Fields.Append "CityWage", adCurrency
    ERN.Fields.Append "CityTax", adCurrency
    ERN.Fields.Append "CourtTax", adCurrency
    ERN.Fields.Append "StateWage", adCurrency
    ERN.Fields.Append "StateTax", adCurrency

    ERN.Fields.Append "DistID", adDouble

    ERN.Fields.Append "NewFlag", adBoolean

    ERN.Fields.Append "MaxAmount", adCurrency

    ' if EMPLOYEE marked as non-taxable
    ERN.Fields.Append "EENoSSTax", adBoolean
    ERN.Fields.Append "EENoMEDTax", adBoolean
    ERN.Fields.Append "EENoFWTTax", adBoolean
    ERN.Fields.Append "EENoSWTTax", adBoolean
    ERN.Fields.Append "EENoCWTTax", adBoolean
    ERN.Fields.Append "EENoFUNTax", adBoolean
    ERN.Fields.Append "EENoSUNTax", adBoolean

    ERN.Fields.Append "RateDifference", adInteger

    ERN.Open , , adOpenDynamic, adLockOptimistic

    ' ========================================================================
    
    ' set up records for DED grid
    DED.CursorLocation = adUseClient
    
    DED.Fields.Append "Title", adDouble
    DED.Fields.Append "Desc", adVarChar, 30, adFldIsNullable
    DED.Fields.Append "Amount", adCurrency
    DED.Fields.Append "AmountManual", adBoolean
    
    ' tax flag info
    DED.Fields.Append "NoSSTax", adBoolean
    DED.Fields.Append "NoMEDTax", adBoolean
    DED.Fields.Append "NoFWTTax", adBoolean
    DED.Fields.Append "NoSWTTax", adBoolean
    DED.Fields.Append "NoCWTTax", adBoolean
    DED.Fields.Append "NoFUNTax", adBoolean
    DED.Fields.Append "NoSUNTax", adBoolean
    DED.Fields.Append "Basis", adInteger
    DED.Fields.Append "AmtPct", adCurrency
    DED.Fields.Append "DirDepType", adInteger
    DED.Fields.Append "DirDepBasis", adInteger
    DED.Fields.Append "DirDepBank", adVarChar, 20, adFldIsNullable
    DED.Fields.Append "DirDepAmtPct", adCurrency
    DED.Fields.Append "EmployerItemID", adDouble
    DED.Fields.Append "ItemType", adInteger
    DED.Fields.Append "ItemID", adDouble
    DED.Fields.Append "CityID", adDouble
    DED.Fields.Append "CityRate", adCurrency
    DED.Fields.Append "CityWage", adCurrency
    DED.Fields.Append "DedSort", adInteger
    
    DED.Fields.Append "ItemHistID", adDouble
    
    DED.Fields.Append "MaxAmount", adCurrency
    DED.Fields.Append "SDTax", adBoolean
    DED.Fields.Append "OrigAmount", adCurrency
    
    DED.Fields.Append "NotInNet", adBoolean
    DED.Fields.Append "WageExcluded", adCurrency
    
    DED.Open , , adOpenDynamic, adLockOptimistic
    
    ' Job Cost
    JC.CursorLocation = adUseClient
    JC.Fields.Append "JobID", adDouble
    JC.Fields.Append "CityID", adDouble
    JC.Fields.Append "CityRate", adCurrency
    JC.Fields.Append "Name", adVarChar, 30, adFldIsNullable
    JC.Open , , adOpenDynamic, adLockOptimistic
    
    rsDedBasis.CursorLocation = adUseClient
    rsDedBasis.Fields.Append "DeductionID", adDouble
    rsDedBasis.Fields.Append "EarningID", adDouble
    rsDedBasis.Fields.Append "Amount", adCurrency
    rsDedBasis.Open , , adOpenDynamic, adLockOptimistic
    
End Sub

Private Sub SetDataGrids()

Dim ET As Integer
Dim ErnCount, DedCount As Integer

    FirstFlag = False
    
    ' get the employee to load
    PREmployee.GetByID (CLng(fgEMP.TextMatrix(fgEMP.Row, 7)))

    DfltCityID = PREmployee.DefaultCityID
    If DfltCityID = 0 Then DfltCityID = 999999

    DfltJobID = PREmployee.DefaultJobID

    ' load the earning and ded types for the employee
    ' for drop down selection
    
    ' string setup for Flex Grid drop down ColComboList
    ' |#nnn;xxxxxxx
    ' nnn = PRITem.ItemID
    ' xxx = PRItem.Title
    
    ' automatically include regular and overtime
    ERNDrop = "|#99991;RegPay|#99992;OvtPay"
    DEDDrop = ""
    
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = " & CStr(EMP!EmployeeID) & _
                " AND PRItem.Active = 1" & _
                " ORDER BY ItemType, EmployerItemID"
    
    If PRItem.GetBySQL(SQLString) Then
    
        Do
        
            ' get the employer item
            ' always used for the title
            ' use for tax flags if employee PRItem.UseEmployer = 0
            ' using a secondary recordset since both are from PRItem
            If PRItem.ItemType <> PREquate.ItemTypeDirDepDed Then
                
                SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & CStr(PRItem.EmployerItemID)
                rsInit SQLString, cn, ERItem
                
                If ERItem.BOF And ERItem.EOF Then
                    MsgBox "Employer Item NF: " & PRItem.ItemID, vbCritical
                    End
                End If
                ERItem.MoveFirst
            
                If PRItem.ItemType = PREquate.ItemTypeOE Then
                    ERNDrop = Trim(ERNDrop) & "|#" & CStr(PRItem.ItemID) & ";" & Trim(ERItem!Title)
                End If
            
                If PRItem.ItemType = PREquate.ItemTypeDED Or PRItem.ItemType = PREquate.ItemTypeSDTax Then
                    DEDDrop = Trim(DEDDrop) & "|#" & CStr(PRItem.ItemID) & ";" & Trim(ERItem!Title)
                End If
        
            Else
        
                If PRItem.ItemType = PREquate.ItemTypeDirDepDed Then
                    DEDDrop = Trim(DEDDrop) & "|#" & CStr(PRItem.ItemID) & ";" & Trim(PRItem.DirDepBank)
                End If
        
            End If
            
            If Not PRItem.GetNext Then Exit Do
        
        Loop
    
    End If
                        
    ' add standard taxes to the drop list
    DEDDrop = Trim(DEDDrop) & "|#99991;SS Tax"
    DEDDrop = Trim(DEDDrop) & "|#99992;MED Tax"
    DEDDrop = Trim(DEDDrop) & "|#99993;FWT Tax"
    DEDDrop = Trim(DEDDrop) & "|#99994;SWT Tax"
    DEDDrop = Trim(DEDDrop) & "|#99995;CWT Tax"
    DEDDrop = Trim(DEDDrop) & "|#99996;COURT Tax"
                        
    ' load the ERN & DED grids
    ' from blank - no history record for this employee
    
    If EMP!HistID = 0 Then
        
        LoadHistNew
        
    Else
    
        LoadHistExisting
        
    End If
        
    ' add some blank lines
    For i = 1 To 3
    
'        ERN.AddNew
'        ERN.Fields("Title") = 0
'        ERN.Fields("Hours") = 0
'        ERN.Fields("Rate") = 0
'        ERN.Fields("CityID") = DfltCityID
'        ERN.Fields("Amount") = 0
'        ERN.Update
        
'        DED.AddNew
'        DED.Fields("Title") = 0
'        DED.Fields("Desc") = ""
'        DED.Fields("Amount") = 0
'        DED.Fields("ItemType") = PREquate.ItemTypeDED
'        DED.Update
        
    Next i
    
    ' ===> assign the city drop down ColComboList
    ' *** use innerjoin to get state name
    CityDrop = "|#999999;NON TAX"
    SQLString = "SELECT * FROM PRCity ORDER BY CityName"
    If PRCity.GetBySQL(SQLString) Then
        Do
            CityDrop = Trim(CityDrop) & "|#" & CStr(PRCity.CityID) & ";" & Trim(PRCity.ShortName) & vbTab & Format(PRCity.CityRate, "##0.00")
            If Not PRCity.GetNext Then Exit Do
        Loop
    End If
    
    ' gather job info
    If JobDist = True Then
        JobDrop = "|#0;NONE"
        If JC.RecordCount > 0 Then
            JC.MoveFirst
            Do
                JobDrop = Trim(JobDrop) & "|#" & JC!JobID & ";" & JC!Name
                JC.MoveNext
            Loop Until JC.EOF
        End If
        
    End If
    
    ' assign dropdown control to the grids
    fgERN.ColComboList(0) = ERNDrop
    fgERN.ColComboList(5) = JobDrop
    fgERN.ColComboList(6) = CityDrop
    
    ' show the job column
    If JobDist = False Then fgERN.ColWidth(5) = 0
    
    fgDED.ColComboList(0) = DEDDrop

    DED.Sort = "DedSort"

    ' un-editable columns

    ' no earnings records? add one
    If ERN.RecordCount = 0 Then
        ERN.AddNew
        ERN.Update
        ERN.MoveFirst
    End If
    
    ' position the grids
    fgERN.ShowCell 1, 1
    fgERN.Select 1, 1
    fgERN.Refresh
    
    fgDED.ShowCell 1, 1
    fgDED.Select 1, 1
    fgDED.Refresh

End Sub

Private Sub LoadHistNew()
            
    StateAbbrev = "OH"
    If DfltStateID <> 0 Then
        If PRState.GetByID(DfltStateID) Then
            StateAbbrev = PRState.StateAbbrev
        Else
            StateAbbrev = "OH"
        End If
    End If

    If Not PREmployee.GetByID(EMP!EmployeeID) Then
        MsgBox "EE NF: " & EMP!EmployeeID, vbCritical
        End
    End If

    ' use time sheet entry
    TimeSheet = False
    If JobDist = True And frmSelTimeSheets.OK = True Then
        GetTimeSheetData frmSelTimeSheets.rsTimeSheet
    End If

    If TimeSheet = True Then        ' mark the rows as from TimeSheet
    
        If ERN.RecordCount > 0 Then
            ERN.MoveFirst
            Do
                fgERN.RowData(fgERN.Row) = ERN!JobID
                fgERN.Cell(flexcpFontItalic, fgERN.Row, 0, fgERN.Row, fgERN.Cols - 1) = True
                fgERN.Cell(flexcpForeColor, fgERN.Row, 0, fgERN.Row, fgERN.Cols - 1) = vbBlue
                ERN.MoveNext
            Loop Until ERN.EOF
        End If
    
    Else        ' no timesheet data - pop from items
        
        ' add reg and ovt pay by default
        ERN.AddNew
        ERN.Fields("Title") = 99991
        ERN.Fields("Hours") = 0
            
        If PREmployee.Salaried = 1 Then
            ERN.Fields("Rate") = PREmployee.SalaryAmount
            ERN.Fields("Amount") = PREmployee.SalaryAmount
            ERN.Fields("Salary") = True
        Else
            ERN.Fields("Rate") = PREmployee.HourlyAmount
            ERN.Fields("Salary") = False
            ERN.Fields("Amount") = 0
        End If
        
        ERN.Fields("CityID") = DfltCityID
        ERN.Fields("JobID") = DfltJobID
        If DfltJobID <> 0 And JobDist = True Then
            If JCJob.GetByID(DfltJobID) Then
                If PRCity.GetByID(JCJob.CityID) Then
                    ERN.Fields("CityID") = PRCity.CityID
                End If
            End If
        End If
        
        ERN.Fields("CityWage") = ERN!Amount
        
        ERN!NewFlag = False
        
        ERN.Update
    
        '  OverTime
        ERN.AddNew
        ERN.Fields("Title") = 99992
        ERN.Fields("Hours") = 0
        
        If PREmployee.Salaried = 0 Then
            If PRCompany.DfltOTRate <> 0 Then
                ERN.Fields("Rate") = PREmployee.HourlyAmount * PRCompany.DfltOTRate
            Else
                ERN.Fields("Rate") = PREmployee.HourlyAmount * 1.5
            End If
        Else
            ERN.Fields("Rate") = 0
        End If
        
        ERN.Fields("CityID") = DfltCityID
        ERN.Fields("JobID") = DfltJobID
        If DfltJobID <> 0 And JobDist = True Then
            If JCJob.GetByID(DfltJobID) Then
                If PRCity.GetByID(JCJob.CityID) Then
                    ERN.Fields("CityID") = PRCity.CityID
                End If
            End If
        End If
        
        ERN.Fields("Amount") = 0
        ERN!NewFlag = False
        
        ERN.Update
    
        ErnCount = 2
    
        fgERN.RowData(fgERN.Row) = 0
    
    End If
    
    If TimeSheet = False Then

        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & CStr(EMP!EmployeeID) & _
                    " AND Active = 1 " & _
                    " AND (ItemType = " & CStr(PREquate.ItemTypeOE) & _
                    " OR ItemType = " & CStr(PREquate.ItemTypeDED) & _
                    " OR ItemType = " & CStr(PREquate.ItemTypeSDTax) & ")" & _
                    " ORDER BY ItemType, EmployerItemID"
    
    Else
    
        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & CStr(EMP!EmployeeID) & _
                    " AND Active = 1 " & _
                    " AND (ItemType = " & CStr(PREquate.ItemTypeDED) & _
                    " OR ItemType = " & CStr(PREquate.ItemTypeSDTax) & ")" & _
                    " ORDER BY ItemType, EmployerItemID"
    
    End If
    
    If PRItem.GetBySQL(SQLString) Then

        Do
        
            SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & CStr(PRItem.EmployerItemID)
            rsInit SQLString, cn, ERItem
            If ERItem.BOF And ERItem.EOF Then
                MsgBox "Employer Item NF: " & PRItem.ItemID, vbCritical
                End
            End If
        
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                
                ERN.AddNew
                ERN.Fields("Title") = PRItem.ItemID
                ERN.Fields("Hours") = 0
                ERN.Fields("Rate") = 0
                ERN.Fields("AmtPct") = PRItem.AmtPct
                
                If PRItem.UseEmployer = 0 Then      ' use the employee OE info

                    ERN.Fields("NoSSTax") = PRItem.NoSSTax
                    ERN.Fields("NoMedTax") = PRItem.NoMedTax
                    ERN.Fields("NoFWTTax") = PRItem.NoFWTTax
                    ERN.Fields("NoSWTTax") = PRItem.NoSWTTax
                    ERN.Fields("NoCWTTax") = PRItem.NoCWTTax
                    ERN.Fields("NoFUNTax") = PRItem.NoFUNTax
                    ERN.Fields("NoSUNTax") = PRItem.NoSUNTax
                    ERN.Fields("Tips") = PRItem.Tips
                    ERN.Fields("NotInNet") = PRItem.NotInNet
                    ERN.Fields("RateDifference") = PRItem.RateDifference
                
                Else                                ' use the employer OE info
                    
                    ERN.Fields("NoSSTax") = ERItem!NoSSTax
                    ERN.Fields("NoMedTax") = ERItem!NoMedTax
                    ERN.Fields("NoFWTTax") = ERItem!NoFWTTax
                    ERN.Fields("NoSWTTax") = ERItem!NoSWTTax
                    ERN.Fields("NoCWTTax") = ERItem!NoCWTTax
                    ERN.Fields("NoFUNTax") = ERItem!NoFUNTax
                    ERN.Fields("NoSUNTax") = ERItem!NoSUNTax
                    
                    ERN.Fields("Tips") = ERItem!Tips
                    ERN.Fields("NotInNet") = ERItem!NotInNet
                    ERN.Fields("RateDifference") = nNull(ERItem!RateDifference)
                
                End If
                
                ' always use the EMPLOYEE item for the basis, rate and amount
                ERN.Fields("Basis") = PRItem.Basis
                
                If PRItem.Basis = PREquate.BasisHourly Then
                    ERN.Fields("Rate") = PRItem.AmtPct
                    ERN.Fields("Amount") = 0
                Else
                    ERN.Fields("Rate") = 0
                    ERN.Fields("Amount") = PRItem.AmtPct
                End If
                
                ERN.Fields("JobID") = DfltJobID
                                
                ' get the dflt city from the job?
                ERN.Fields("CityID") = DfltCityID
                If DfltJobID <> 0 And JobDist = True Then
                    If JCJob.GetByID(DfltJobID) Then
                        If PRCity.GetByID(JCJob.CityID) Then
                            ERN.Fields("CityID") = PRCity.CityID
                        End If
                    End If
                End If
                
                ERN.Fields("CityWage") = ERN!Amount
                
                ERN!EmployerItemID = PRItem.EmployerItemID
                ERN!NewFlag = False
                ERN!MaxAmount = PRItem.MaxAmount
                
                ' rate difference for hourly?
                If ERN!Basis = PREquate.BasisHourly And ERN!RateDifference <> 0 Then
                    If ERN!RateDifference = PREquate.BasisAmount Then
                        ERN!Rate = PREmployee.HourlyAmount + ERN!AmtPct
                    End If
                    If ERN!RateDifference = PREquate.BasisPercent Then
                        ERN!Rate = PREmployee.HourlyAmount + Round(ERN!AmtPct / 100 * PREmployee.HourlyAmount, 2)
                    End If
                End If
                
                ERN.Update
                ErnCount = ErnCount + 1
                
            Else
            
                ' see if this deduction was selected
                
                ' **** ??? GVille Patch 11/23/09 ?????
                If frmNewBatch.rsItem.RecordCount > 0 Then frmNewBatch.rsItem.MoveFirst
                
                frmNewBatch.rsItem.Find "ItemID = " & PRItem.EmployerItemID, 0, adSearchForward, 1
                If Not frmNewBatch.rsItem.EOF And frmNewBatch.rsItem!Select = True Then
            
                    DED.AddNew
                    DED.Fields("Title") = PRItem.ItemID
                    DED.Fields("Desc") = ""
                    DED.Fields("DedSort") = 10
                    
                    ' handle SD tax
                    DED.Fields("ItemType") = PREquate.ItemTypeDED
                    DED.Fields("ItemType") = PRItem.ItemType
                    
                    If PRItem.UseEmployer = 0 Then      ' use the employee OE info
    
                        DED.Fields("NoSSTax") = PRItem.NoSSTax
                        DED.Fields("NoMedTax") = PRItem.NoMedTax
                        DED.Fields("NoFWTTax") = PRItem.NoFWTTax
                        DED.Fields("NoSWTTax") = PRItem.NoSWTTax
                        DED.Fields("NoCWTTax") = PRItem.NoCWTTax
                        DED.Fields("NoFUNTax") = PRItem.NoFUNTax
                        DED.Fields("NoSUNTax") = PRItem.NoSUNTax
                        DED.Fields("NotInNet") = PRItem.NotInNet
                        
                    Else                                ' use the employer DED tax info
                        
                        DED.Fields("NoSSTax") = ERItem!NoSSTax
                        DED.Fields("NoMedTax") = ERItem!NoMedTax
                        DED.Fields("NoFWTTax") = ERItem!NoFWTTax
                        DED.Fields("NoSWTTax") = ERItem!NoSWTTax
                        DED.Fields("NoCWTTax") = ERItem!NoCWTTax
                        DED.Fields("NoFUNTax") = ERItem!NoFUNTax
                        DED.Fields("NoSUNTax") = ERItem!NoSUNTax
                        DED.Fields("NotInNet") = ERItem!NotInNet
                        
                    End If
                    
                    DED.Fields("ItemType") = PRItem.ItemType
                        
                    ' always use the EMPLOYEE for the basis and amount
                    DED.Fields("Basis") = PRItem.Basis
                    DED.Fields("AmtPct") = PRItem.AmtPct
                    If PRItem.Basis = PREquate.BasisPercent Then
                        DED.Fields("Desc") = Format(PRItem.AmtPct, "##0.00") & "%"
                        DED.Fields("Amount") = 0
                    Else
                        DED.Fields("Desc") = "$" & Format(PRItem.AmtPct, "##,##0.00")
                        DED.Fields("Amount") = PRItem.AmtPct
                    End If
                    
                    DED!EmployerItemID = PRItem.EmployerItemID
                    DED!ItemID = PRItem.ItemID
                    DED!MaxAmount = PRItem.MaxAmount
                    
                    If PRItem.ItemType = PREquate.ItemTypeSDTax Then
                        DED!SDTax = True
                    Else
                        DED!SDTax = False
                    End If
                    
                    DED.Update
                    DedCount = DedCount + 1
                
                End If
            
            End If
    
            If Not PRItem.GetNext Then Exit Do
            
        Loop
    
    End If
    
    ' create the deduct basis recordset
    DedBasisCreate
    
    ' add standard taxes
    For i = 1 To 4
        DED.AddNew
        DED.Fields("Title") = 99990 + i
        DED.Fields("Desc") = ""
        DED.Fields("Amount") = 0
        DED.Fields("ItemType") = PREquate.ItemTypeRegTax
        DED.Fields("DedSort") = 20 + i
        DED!SDTax = False
        DED.Update
    Next i
    
    ' add dir dep
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & CStr(EMP!EmployeeID) & _
                " AND Active = 1 " & _
                " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeDirDepDed)
                
    If PRItem.GetBySQL(SQLString) Then
    
        Do
    
            DED.AddNew
            DED.Fields("Title") = PRItem.ItemID
            If PRItem.DirDepType = PREquate.DirDepTypeChecking Then
                DED.Fields("Desc") = "DD CHK"
            Else
                DED.Fields("Desc") = "DD SVG"
            End If
            
            If PRItem.DirDepBasis = PREquate.DirDepBasisNet Then
                DED!Desc = Trim(DED!Desc) & " NET"
                DED.Fields("DedSort") = 31
            ElseIf PRItem.DirDepBasis = PREquate.DirDepBasisAmt Then
                DED.Fields("DedSort") = 30
            Else
                DED.Fields("DedSort") = 30
            End If
            
            DED.Fields("ItemType") = PREquate.ItemTypeDirDepDed
            
            ' always use the EMPLOYEE for the direct deposit info
            DED.Fields("DirDepType") = PRItem.DirDepType
            DED.Fields("DirDepBasis") = PRItem.DirDepBasis
            DED.Fields("DirDepBank") = Mid(PRItem.DirDepBank, 1, 20)
            DED.Fields("DirDepAmtPct") = PRItem.DirDepAmtPct
            
            DED.Update
    
            If Not PRItem.GetNext Then Exit Do
        
        Loop
    
    End If
    
    ' update EE tax flags to earnings
    If ERN.RecordCount > 0 Then
        ERN.MoveFirst
        Do
            If PREmployee.NoSSTax Then ERN!EENoSSTax = True
            If PREmployee.NoMedTax Then ERN!EENoMedTax = True
            If PREmployee.NoFedTax Then ERN!EENoFWTTax = True
            If PREmployee.NoStateTax Then ERN!EENoSWTTax = True
            If PREmployee.NoCityTax Then ERN!EENoCWTTax = True
            If PREmployee.NoFedUnemp Then ERN!EENoFUNTax = True
            If PREmployee.NoStateUnemp Then ERN!EENoSUNTax = True
            ERN!NewFlag = False
            ERN.Update
            ERN.MoveNext
        Loop Until ERN.EOF
    End If
End Sub

Private Sub LoadHistExisting()

Dim ItemAmt As Currency

    ' ********************************************
    ' recordset to store city distribution
    Dim rsCD As New ADODB.Recordset

    rsCD.CursorLocation = adUseClient
    rsCD.Fields.Append "CityID", adDouble
    rsCD.Fields.Append "Tax", adCurrency
    rsCD.Fields.Append "Wage", adCurrency
    rsCD.Fields.Append "Courtesy", adInteger
    rsCD.Open , , adOpenDynamic, adLockOptimistic

    ' ===> assign from existing history and dist
    If Not PRHist.GetByID(EMP!HistID) Then
        MsgBox "PRHist NF: " & ERN!HistID, vbCritical
        End
    End If
    
    ' loop thru prdist for history record
    SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
    If PRDist.GetBySQL(SQLString) Then
        
        Do
        
            ERN.AddNew
            
            If PRDist.DistType = PREquate.DistTypeReg Then
                ERN!Title = 99991
                ERN!salary = PREmployee.Salaried
            ElseIf PRDist.DistType = PREquate.DistTypeOT Then
                ERN!Title = 99992
            ElseIf PRDist.DistType = PREquate.DistTypeItem Then
                ERN!Title = PRDist.ItemID
            Else
                MsgBox "DistType NF: " & PRDist.DistID & " " & PRDist.DistType, vbCritical
                End
            End If
            
            ' .... get pritem
            If Not PRItem.GetByID(PRDist.ItemID) Then
                MsgBox "PRItem from PRDist NF: " & PRDist.DistID & " " & PRDist.ItemID, vbCritical
                End
            End If
            
            ' use employer
            If PRItem.UseEmployer Then
                If Not PRItem.GetByID(PRItem.EmployerItemID) Then
                    MsgBox "PRItem ER from PRDist NF: " & PRDist.DistID & " " & PRItem.EmployerItemID, vbCritical
                    End
                End If
            End If
            
            ERN!Basis = PRItem.Basis
                        
            ERN!Hours = PRDist.Hours
            ERN!Rate = PRDist.Rate
            ERN!Amount = PRDist.Amount
            ERN!AmountManual = True
            ERN!CityID = PRDist.CityID
            ERN!JobID = PRDist.JobID
            ERN!CityTax = PRDist.CityTax
            ERN!CityManual = True
            ERN!CityWage = PRDist.CityWage
            ERN!CourtTax = PRDist.CourtesyCityTax
            ERN!DistID = PRDist.DistID
            ERN!EmployerItemID = PRDist.EmployerItemID
            
            ' get tax flags for other earnings
            If PRDist.DistType = PREquate.DistTypeItem Then
                
                ' use the employer record
                If PRItem.UseEmployer = 1 Then
                    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & PRItem.EmployerItemID
                    If Not PRItem.GetBySQL(SQLString) Then
                        MsgBox "Employer Item NF: " & PRItem.EmployerItemID, vbCritical
                        End
                    End If
                End If
                
                ' assign tax flag to ern record set
                ERN!NoSSTax = PRItem.NoSSTax
                ERN!NoMedTax = PRItem.NoMedTax
                ERN!NoFWTTax = PRItem.NoFWTTax
                ERN!NoSWTTax = PRItem.NoSWTTax
                ERN!NoCWTTax = PRItem.NoCWTTax
            
            End If
            
            If PRDist.JobID > 0 Then
                ERN!JobID = PRDist.JobID
                fgERN.RowData(fgERN.Row) = ERN!JobID
                fgERN.Cell(flexcpFontItalic, fgERN.Row, 0, fgERN.Row, fgERN.Cols - 1) = True
                fgERN.Cell(flexcpForeColor, fgERN.Row, 0, fgERN.Row, fgERN.Cols - 1) = vbBlue
            End If
            
            ERN!NewFlag = False
            ERN.Update
            
            ' store city distribution
            SQLString = "CityID = " & PRDist.CityID
            rsCD.Find SQLString, 0, adSearchForward, 1
            If rsCD.EOF Then
                rsCD.AddNew
                rsCD!CityID = PRDist.CityID
                rsCD!Courtesy = 0
                rsCD.Update
            End If
            rsCD!Tax = rsCD!Tax + PRDist.CityTax
            rsCD!Wage = rsCD!Wage + PRDist.CityWage
            rsCD.Update
            
            ' courtesy WH?
            If PRDist.CourtesyCityTax <> 0 Then
                SQLString = "CityID = " & PRDist.CityID & " AND Courtesy = 1"
                rsCD.Filter = SQLString
                If rsCD.RecordCount = 0 Then
                    rsCD.Filter = adFilterNone
                    rsCD.AddNew
                    rsCD!CityID = PRDist.CourtesyCityID
                    rsCD!Courtesy = 1
                    rsCD.Update
                End If
                rsCD!Tax = rsCD!Tax + PRDist.CourtesyCityTax
                rsCD!Wage = rsCD!Wage + PRDist.CityWage
                rsCD.Update
                rsCD.Filter = adFilterNone
            End If
            
            ' update the tax totals
            SWTWageTL = SWTWageTL + PRDist.StateWage
            SWTTaxTL = SWTTaxTL + PRDist.StateTax
            CWTWageTL = CWTWageTL + PRDist.CityWage
            CWTTaxTL = CWTTaxTL + PRDist.CityTax
            
            If Not PRDist.GetNext Then Exit Do
        
        Loop
    
    End If
                
    ' get the deductions
    SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.HistID = " & PRHist.HistID
    
    If PRItemHist.GetBySQL(SQLString) Then
        Do
        
            ' .... get pritem
            If Not PRItem.GetByID(PRItemHist.ItemID) Then
                MsgBox "PRItem from PRItemHist NF: " & PRItemHist.ItemHistID & " " & PRItemHist.ItemID, vbCritical
                End
            End If
                
            DED.AddNew
            DED!Title = PRItemHist.ItemID
            DED!ItemID = PRItemHist.ItemID
            DED!EmployerItemID = PRItemHist.EmployerItemID
                        
            DED!Amount = PRItemHist.Amount
            DED!AmountManual = True
            
            If PRItem.ItemType = PREquate.ItemTypeDED Or PRItem.ItemType = PREquate.ItemTypeSDTax Then
                
                DED!dedSort = 10
                
                ' handle SD tax
                DED!ItemType = PREquate.ItemTypeDED
                DED!ItemType = PRItem.ItemType
            
                ' store the amt and percent from the employee item defn
                ItemAmt = PRItem.AmtPct
                
                ' use the employer ???
                If PRItem.UseEmployer = 1 Then
                    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & PRItem.EmployerItemID
                    If Not PRItem.GetBySQL(SQLString) Then
                        MsgBox "Employer Item NF: " & PRItem.ItemID & vbCr & PRItem.EmployerItemID, vbCritical
                        End
                    End If
                End If
                
                If PRItem.Basis = PREquate.BasisAmount Or PRItem.Basis = 0 Then
                    DED!Desc = "$" & Format(ItemAmt, "##,##0.00")
                ElseIf PRItem.Basis = PREquate.BasisPercent Then
                    DED!Desc = Format(ItemAmt, "##0.00") & "%"
                Else
                    DED!Desc = "..."
                End If
                
                DED!NoSSTax = PRItem.NoSSTax
                DED!NoMedTax = PRItem.NoMedTax
                DED!NoFWTTax = PRItem.NoFWTTax
                DED!NoSWTTax = PRItem.NoSWTTax
                DED!NoCWTTax = PRItem.NoCWTTax
                
                DED!ItemType = PREquate.ItemTypeDED
                ' handle sd tax
                DED!ItemType = PRItem.ItemType
                DED!NotInNet = PRItem.NotInNet
            
            Else        ' direct deposit
                
                DED!dedSort = 30    ' direct deposit
                DED!ItemType = PREquate.ItemTypeDirDepDed
            
                If PRItem.DirDepType = PREquate.DirDepTypeChecking Then
                    DED!Desc = "DD CHK"
                Else
                    DED!Desc = "DD SVG"
                End If
                
                DED!DirDepAmtPct = PRItem.DirDepAmtPct
                DED!DirDepBasis = PRItem.DirDepBasis
                
            End If
            
            DED!ItemHistID = PRItemHist.ItemHistID
            DED!EmployerItemID = PRItem.EmployerItemID
            
            If PRItem.ItemType = PREquate.ItemTypeSDTax Then
                DED!SDTax = True
            Else
                DED!SDTax = False
            End If
            
            DED.Update
                
            If Not PRItemHist.GetNext Then Exit Do
        
        Loop
    
    End If
            
    DedBasisCreate
            
    ' add the regular taxes
    DED.AddNew
    DED!Title = 99991
    DED!Desc = Format(PRHist.SSWage, "$##,###,##0.00")
    DED!Amount = PRHist.SSTax
    DED!AmountManual = True
    DED!ItemType = PREquate.ItemTypeRegTax
    DED!dedSort = 21
    DED!SDTax = False
    DED.Update
    
    DED.AddNew
    DED!Title = 99992
    DED!Desc = Format(PRHist.MEDWage, "$##,###,##0.00")
    DED!Amount = PRHist.MedTax
    DED!AmountManual = True
    DED!ItemType = PREquate.ItemTypeRegTax
    DED!dedSort = 22
    DED!SDTax = False
    DED.Update
    
    DED.AddNew
    DED!Title = 99993
    DED!Desc = Format(PRHist.FWTWage, "$##,###,##0.00")
    DED!Amount = PRHist.FWTTax
    DED!AmountManual = True
    DED!ItemType = PREquate.ItemTypeRegTax
    DED!dedSort = 23
    DED!SDTax = False
    DED.Update
    
    DED.AddNew
    DED!Title = 99994
    DED!Desc = Format(PRHist.SWTWage, "$##,###,##0.00")
    DED!Amount = PRHist.SWTTax
    DED!AmountManual = True
    DED!ItemType = PREquate.ItemTypeRegTax
    DED!dedSort = 24
    DED!SDTax = False
    DED.Update
    
'    DED.AddNew
'    DED!Title = 99995
'    DED!Desc = Format(CWTWageTL, "$##,###,##0.00")
'    DED!Amount = CWTTaxTL
'    DED!AmountManual = True
'    DED!ItemType = PREquate.ItemTypeRegTax
'    DED!DedSort = 25
'    DED!SDTax = False
'    DED.Update

    ' add city tax from rsCD
    If rsCD.RecordCount > 0 Then
        rsCD.Sort = "Courtesy, CityID"
        rsCD.MoveFirst
        Do
            
            DED.AddNew
            If rsCD!Courtesy = 0 Then
                DED!Title = 99995
            Else
                DED!Title = 99996
            End If
            
            X = ""
            If PRCity.GetByID(rsCD!CityID) Then
                X = Mid(PRCity.ShortName, 1, 7)
            End If
            
            DED!Desc = Trim(X) & " " & Format(rsCD!Wage, "$###,##0.00")
            DED!Amount = rsCD!Tax
            DED!AmountManual = True
            DED!ItemType = PREquate.ItemTypeRegTax
            DED!dedSort = 25
            DED!SDTax = False
            DED!CityID = rsCD!CityID
            DED.Update
            
            rsCD.MoveNext
        
        Loop Until rsCD.EOF
    End If

    ' update EE tax flags to earnings
    If ERN.RecordCount > 0 Then
        ERN.MoveFirst
        Do
            If PREmployee.NoSSTax Then ERN!EENoSSTax = True
            If PREmployee.NoMedTax Then ERN!EENoMedTax = True
            If PREmployee.NoFedTax Then ERN!EENoFWTTax = True
            If PREmployee.NoStateTax Then ERN!EENoSWTTax = True
            If PREmployee.NoCityTax Then ERN!EENoCWTTax = True
            If PREmployee.NoFedUnemp Then ERN!EENoFUNTax = True
            If PREmployee.NoStateUnemp Then ERN!EENoSUNTax = True
            ERN!NewFlag = False
            ERN.Update
            ERN.MoveNext
        Loop Until ERN.EOF
    End If

    rsCD.Close

End Sub
Private Sub CalcGrids()

Dim rsSS As New ADODB.Recordset
Dim ERNTotal As Currency
Dim DEDTotal As Currency
Dim TaxTotal As Currency
Dim DDTotal As Currency
Dim CHKTotal As Currency
Dim NETTotal As Currency
Dim HRTotal As Currency

Dim SSWage As Currency
Dim MEDWage As Currency
Dim FWTWage As Currency
Dim SWTWage As Currency
Dim CWTWage As Currency

Dim OrigNet, Net, CWTDed, CWTDist, TotalCityWage As Currency
Dim OrigAmt As Currency
Dim SetFlag As Boolean
Dim MarSng As String
Dim fgERNRow, fgERNCol, fgDEDRow, fgDEDCol As Integer

Dim CityManual As Boolean
Dim ErnCityID, ErnCount As Long
Dim SWTState As String
Dim YTDAmt As Currency
Dim CalcStateID As Long
Dim SDEXAmt, AnnSWTWage As Currency
Dim DedBasis As Currency
 
    p1 = 0
    P2 = 0
    CalcStateID = 0

    ' store the original amount for deductions
    ' used to get CWT back if from existing
    If DED.RecordCount > 0 Then
        DED.MoveFirst
        Do
            DED!OrigAmount = DED!Amount
            DED.Update
            DED.MoveNext
        Loop Until DED.EOF
    End If

    ' get the YTD SSWage
    SQLString = "SELECT HistID, CheckDate, SSWage " & _
                " FROM PRHist WHERE PRHist.EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7)
    
    If EMP!HistID <> 0 Then
        SQLString = "SELECT HistID, CheckDate, SSWage, FUNWage, SUNWage " & _
                    " FROM PRHist WHERE PRHist.EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7) & _
                    " AND PRHist.CheckDate <= " & CLng(Int(PRBatch.CheckDate)) & _
                    " AND PRHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
                    " AND PRHist.HistID < " & EMP!HistID & _
                    " ORDER BY PRHist.HistID"
    Else
        SQLString = "SELECT HistID, CheckDate, SSWage, FUNWage, SUNWage " & _
                    " FROM PRHist WHERE PRHist.EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7) & _
                    " AND PRHist.CheckDate <= " & CLng(Int(PRBatch.CheckDate)) & _
                    " AND PRHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
                    " ORDER BY PRHist.HistID"
        
'        ' ????? check date compare if two checks same pay ??????
'        SQLString = "SELECT HistID, CheckDate, SSWage " & _
'                    " FROM PRHist WHERE PRHist.EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7) & _
'                    " AND PRHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
'                    " ORDER BY PRHist.HistID"
    End If
                        
    YTDSSWage = 0
    YTDFUNWage = 0
    YTDSUNWage = 0
                        
    rsInit SQLString, cn, rsSS
    If rsSS.RecordCount > 0 Then
        Do
            YTDSSWage = YTDSSWage + rsSS!SSWage
            YTDFUNWage = YTDFUNWage + rsSS!FUNWage
            YTDSUNWage = YTDSUNWage + rsSS!SUNWage
            rsSS.MoveNext
        Loop Until rsSS.EOF
        rsSS.Close
    End If
    
    ' save the cursor positions
    fgERNRow = fgERN.Row
    fgERNCol = fgERN.Col
    fgDEDRow = fgDED.Row
    fgDEDCol = fgDED.Col

    ' ??? one earnings line - select after dropdown of first column ???
    If ERN.RecordCount = 1 And ERN.EOF Then ERN.MoveFirst

    ' actions if on Earnings Grid
    If GridFocus = 1 Then
            
        If fgERN.Col = 0 And PREmployee.Salaried = 0 Then
            
            ERN!CityID = DfltCityID
            
            If ERN!Title = 99992 Then
                ERN!Rate = SuperRound(PREmployee.HourlyAmount, 1.5)
            Else
                ERN!Rate = PREmployee.HourlyAmount
            End If
        
            ERN.Update
        
        End If
    
    End If

    ' ======================================================================================

    ' loop thru the ERN
    NotInNetTotal = 0
    TotalCityWage = 0
    ERN.MoveFirst
    
    ' include in basis for deduct by percent?
    ' clear the amounts
    If rsDedBasis.RecordCount > 0 Then
        rsDedBasis.MoveFirst
        Do
            rsDedBasis!Amount = 0
            rsDedBasis.Update
            rsDedBasis.MoveNext
        Loop Until rsDedBasis.EOF
    End If
    
    Do
        
        ' set the flags from the earning item
        ' ERN!Title has the employee item id
        If ERN!Title < 99990 And ERN!NewFlag = True Then
            
            If PRItem.GetByID(ERN!Title) = False Then
                MsgBox "PRItem Not Found: " & ERN!Title, vbExclamation
                Exit Sub
            End If
            
            ' basis and amount per employee
            ERN!Basis = PRItem.Basis
            If PRItem.Basis = PREquate.BasisHourly Then
                ERN!Rate = PRItem.AmtPct
            ElseIf PRItem.Basis = PREquate.BasisAmount Then
                ERN!Amount = PRItem.AmtPct
            End If
            ERN!EmployerItemID = PRItem.EmployerItemID
            
            ' use the employer item record?
            If PRItem.UseEmployer Then
                If PRItem.GetByID(PRItem.EmployerItemID) = False Then
                    MsgBox "Employer Item NF: " & PRItem.EmployerItemID, vbExclamation
                    Exit Sub
                End If
            End If
            
            ERN!NoSSTax = PRItem.NoSSTax
            ERN!NoMedTax = PRItem.NoMedTax
            ERN!NoFWTTax = PRItem.NoFWTTax
            ERN!NoSWTTax = PRItem.NoSWTTax
            ERN!NoCWTTax = PRItem.NoCWTTax
            ERN!NoFUNTax = PRItem.NoFUNTax
            ERN!NoSUNTax = PRItem.NoSUNTax
            ERN!Tips = PRItem.Tips
            ERN!NotInNet = PRItem.NotInNet
            ERN!NewFlag = False
            ERN.Update
        End If
        
        ' extend the hourly amounts
        SetFlag = True
        
        If ERN!Basis <> PREquate.BasisHourly Then SetFlag = False
        
        ' regular pay - not salaried
        If ERN!Title = 99991 And ERN!salary = False Then SetFlag = True
        
        ' OverTime
        If ERN!Title = 99992 Then SetFlag = True
        
        ' the manual flag trumps 'em all!!!
        If ERN!AmountManual = True Then SetFlag = False
        
        If SetFlag = True Then
            ERN.Fields("Amount") = Round(ERN!Hours * ERN!Rate, 2)
            ERN!Amount = SuperRound(ERN!Hours, ERN!Rate)
        End If
        
        ' include in basis for deduct by percent?
        If rsDedBasis.RecordCount > 0 Then
            rsDedBasis.MoveFirst
            Do
                If ERN!Title = "99991" Or ERN!Title = "99992" Then
                    If rsDedBasis!EarningID = ERN!Title Then
                        rsDedBasis!Amount = rsDedBasis!Amount + ERN!Amount
                        rsDedBasis.Update
                    End If
                Else
                    If rsDedBasis!EarningID = ERN!EmployerItemID Then
                        rsDedBasis!Amount = rsDedBasis!Amount + ERN!Amount
                        rsDedBasis.Update
                    End If
                End If
                rsDedBasis.MoveNext
            Loop Until rsDedBasis.EOF
        End If
        
        ERNTotal = ERNTotal + ERN!Amount
        HRTotal = HRTotal + ERN!Hours
        
        ' taxable wage totals
        If ERN!NoSSTax = False Then SSWage = SSWage + ERN!Amount
        If ERN!NoMedTax = False Then MEDWage = MEDWage + ERN!Amount
        If ERN!NoFWTTax = False Then FWTWage = FWTWage + ERN!Amount
        If ERN!NoSWTTax = False Then SWTWage = SWTWage + ERN!Amount
        If ERN!NoCWTTax = False Then CWTWage = CWTWage + ERN!Amount
        
        If ERN!NoCWTTax = False Then
            ERN!CityWage = ERN!Amount
            TotalCityWage = TotalCityWage + ERN!Amount
        End If
        
'        ' calculate city tax
'        If Not PRCity.GetByID(ERN!CityID) Then
'            MsgBox "City Error: " & ERN!CityID, vbCritical
'            End
'        End If
                
        If ERN!NotInNet = True Then
            NotInNetTotal = NotInNetTotal + ERN!Amount
        End If
        
        ' store a city id from earnings
        ' use to get state for SWT basis
        If ERN!Amount > 0 Then ErnCityID = ERN!CityID
        
        ERN.Update
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    
    Loop
    
    ' ======================================================================================
    
    ' loop thru DED
    CWTDed = 0
    DED.MoveFirst
    Do
        
        ' dont' do SD tax here!!!
        If DED!ItemType = PREquate.ItemTypeDED Then
        
            ' calculate deductions amounts
            If DED!AmountManual = False Then
                
                ' track hours also ......
                DedBasis = ERNTotal
                DED!WageExcluded = 0
                ' earning amounts to take from ded basis?
                If rsDedBasis.RecordCount > 0 Then
                    rsDedBasis.MoveFirst
                    Do
                        If rsDedBasis!DeductionID = DED!EmployerItemID Then
                            DedBasis = DedBasis - rsDedBasis!Amount
                            DED!WageExcluded = DED!WageExcluded + rsDedBasis!Amount
                        End If
                        rsDedBasis.MoveNext
                    Loop Until rsDedBasis.EOF
                End If
                
                If DED!Basis = PREquate.BasisPercent Then           ' percent of gross
                    DED.Fields("Amount") = Round(DedBasis * DED!AmtPct / 100, 2)
                ElseIf DED!Basis = PREquate.BasisHourly Then
                    DED!Amount = Round(DED!AmtPct * HRTotal, 2)
                Else
                    DED.Fields("Amount") = DED!AmtPct               ' by amount otherwise
                End If
            End If
            
            ' max amount?
            If DED!MaxAmount <> 0 Then
                YTDAmt = 0
                
                ' ItemHist ID ?????
'                SQLString = "SELECT * FROM PRItemHist WHERE EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7) & _
'                            " AND PRItemHist.CheckDate <= " & CLng(Int(PRBatch.CheckDate)) & _
'                            " AND PRItemHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
'                            " AND PRItemHist.ItemHistID < " & EMP!DistID & _
'                            " AND PRItemHist.ItemID = " & ERN!ItemID & _
'                            " ORDER BY PRItemHist.DistID"
                
                SQLString = "SELECT * FROM PRItemHist WHERE EmployeeID = " & fgEMP.TextMatrix(fgEMP.Row, 7) & _
                            " AND PRItemHist.CheckDate <= " & CLng(Int(PRBatch.CheckDate)) & _
                            " AND PRItemHist.YearMonth >= " & (Int(PRBatch.YearMonth / 100) * 100) + 1 & _
                            " AND PRItemHist.ItemID = " & DED!ItemID & _
                            " ORDER BY PRItemHist.ItemHistID"
                
                If PRItemHist.GetBySQL(SQLString) Then
                    Do
                        YTDAmt = YTDAmt + PRItemHist.Amount
                        If Not PRItemHist.GetNext Then Exit Do
                    Loop
                End If
                
                DED!Amount = AmtMax(DED!Amount, YTDAmt, DED!MaxAmount)
                
            End If
            
            ' taxable wage totals
            If DED!NoSSTax = True Then SSWage = SSWage - DED!Amount
            If DED!NoMedTax = True Then MEDWage = MEDWage - DED!Amount
            If DED!NoFWTTax = True Then FWTWage = FWTWage - DED!Amount
            If DED!NoSWTTax = True Then SWTWage = SWTWage - DED!Amount
            
            If DED!NoCWTTax = True Then
                CWTWage = CWTWage - DED!Amount
                CWTDed = CWTDed + DED!Amount
            End If
            
            DED.Update
        
            DEDTotal = DEDTotal + DED!Amount
        
        End If
        
        DED.MoveNext
        If DED.EOF Then Exit Do
        
    Loop
    
    ' ======================================================================================
    
    ' can't be negative
    If SSWage < 0 Then SSWage = 0
    If MEDWage < 0 Then MEDWage = 0
    If FWTWage < 0 Then FWTWage = 0
    If SWTWage < 0 Then SWTWage = 0
    If CWTWage < 0 Then CWTWage = 0
    
    ' remove then create the CWT lines
    DED.MoveFirst
    Do
        If (DED!Title = 99995 Or DED!Title = 99996) And DED!AmountManual = False Then
        ' If (DED!Title = 99995 Or DED!Title = 99996) Then
            DED.Delete
        End If
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop
    
    ' courtesy CWT?
    If PREmployee.CourtesyCityID <> 0 Then
        If Not PRCity.GetByID(PREmployee.CourtesyCityID) Then
            CourtRate = 0
        Else
            CourtRate = PRCity.CityRate
            CourtCityName = PRCity.CityName
            CourtCityID = PREmployee.CourtesyCityID
        End If
        CourtAdd = PREmployee.CourtesyAdd
    Else
        CourtRate = 0
    End If
    
    ' clear courtesy amounts / records
    CourtTax = 0
    ErnCount = 0
    CWTDist = CWTDed
    ERN.MoveFirst
    Do
        
        ErnCount = ErnCount + 1
        
        ' **** use city wage .....
        If ERN!CityWage <> 0 Then
            SQLString = "CityID = " & ERN!CityID
            DED.Find SQLString, 0, adSearchForward, 1
            If DED.EOF Then
                ' get the PRCity record
                If ERN!CityID = 999999 Then ' non tax
                    DED.AddNew
                    DED!Title = 99995
                    DED!CityID = 999999
                    DED!CityRate = 0
                    DED!Amount = 0
                    DED!dedSort = 25
                    DED!Desc = "NON TAX"
                Else
                    If Not PRCity.GetByID(ERN!CityID) Then
                        MsgBox "City Error: " & ERN!CityID, vbCritical
                        End
                    End If
                    DED.AddNew
                    DED!Title = 99995
                    DED!CityID = PRCity.CityID
                    DED!CityRate = PRCity.CityRate
                    DED!Amount = 0
                    DED!dedSort = 25
                    DED!Desc = PRCity.ShortName
                End If
                DED!SDTax = False
                DED.Update
            End If
                    
            ' **** use city wage
            If ERN!CityManual = False Then
                ' re-calc city wage if necessary
                If CWTDed <> 0 Then
                    If ErnCount = ERN.RecordCount Then  ' last record - use remaining
                        p1 = CWTDist
                        CWTDist = 0
                    Else
                        If TotalCityWage <> 0 Then
                            p1 = Round(ERN!CityWage / TotalCityWage * CWTDed, 2)
                        Else
                            p1 = 0
                        End If
                        If p1 > CWTDist Then
                            p1 = CWTDist
                            CWTDist = 0
                        Else
                            CWTDist = CWTDist - p1
                        End If
                    End If
                    ERN!CityWage = ERN!CityWage - p1
                End If
                    
                ERN!CityTax = Round(DED!CityRate / 100 * ERN!CityWage, 2)
                CityManual = False
                
                ' courtesy CWT ???
                ' differential
                If ERN!CityTax <> 0 And PRCity.CityRate < CourtRate And CourtAdd = 0 Then
                    ERN!CourtTax = Round(((CourtRate - DED!CityRate) * ERN!CityWage / 100), 2)
                    CourtTax = CourtTax + ERN!CourtTax
                ElseIf CourtRate <> 0 And CourtAdd = 1 Then     ' additional
                    ERN!CourtTax = Round(CourtRate * ERN!CityWage / 100, 2)
                    CourtTax = CourtTax + ERN!CourtTax
                End If
            
            Else
                DED!AmountManual = True
                CourtTax = CourtTax + ERN!CourtTax
                CityManual = True
            End If
            DED!CityWage = DED!CityWage + ERN!CityWage
                        
            If PREmployee.NoCityTax Then
                DED!Desc = "X " & Trim(DED!Desc)
                ERN!CityTax = 0
            End If
            
            If PREmployee.NoCityTax And ERN!CityManual = False Then p1 = 0
                        
            DED!Amount = DED!Amount + Round(ERN!CityTax, 2)
            DED.Update
            ERN.Update
    
        End If
        
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    
    Loop
    
    ' add a line for courtesy CWT?
    If CourtTax <> 0 Then
        ' does one already exist?
        p1 = 0
        DED.MoveFirst
        Do
            If DED!Title = 99996 Then
                If DED!AmountManual = True Then
                    p1 = 2
                    CourtTax = DED!Amount
                Else
                    p1 = 1
                End If
                Exit Do
            End If
            DED.MoveNext
        Loop Until DED.EOF
        
        If p1 = 2 Then
            ' manual - leave it alone
        Else
            If p1 = 0 Then DED.AddNew
            DED!Title = 99996
            DED!CityID = PREmployee.CourtesyCityID
            DED!CityRate = CourtRate
            DED!Amount = CourtTax
            DED!dedSort = 26
            DED!Desc = CourtCityName
            DED!SDTax = False
            If CityManual = True Then DED!AmountManual = True
            DED.Update
        End If
    
    End If
    
'''    ' loop again if cwtded <> 0
'''    ' split up the deduction from CWT proportionally
'''    If CWTDed <> 0 And CityManual = False Then
'''        CourtTax = 0
'''        CWTDist = CWTDed
'''        DED.MoveFirst
'''        Do
'''            If DED!Title = 99995 Then
'''                P1 = Round(DED!CityWage / TotalCityWage * CWTDed, 2)
'''                CWTDist = CWTDist - P1
'''                DED!CityWage = DED!CityWage - P1
'''                DED!Amount = Round(DED!CityRate / 100 * DED!CityWage, 2)
'''                DED!Desc = Trim(DED!Desc) & " " & Format(DED!CityWage, "###,##0.00")
'''                DED.Update
'''            End If
'''            DED.MoveNext
'''            If DED.EOF Then Exit Do
'''        Loop
'''
'''        ' rounding
'''        If CWTDist <> 0 Then
'''            DED.MoveFirst
'''            DED!CityWage = DED!CityWage + (CWTDist - CWTDed)
'''            DED!Amount = Round(DED!CityRate / 100 * DED!CityWage, 2)
'''            If DED!CityID <> 0 Then     ' ???
'''                If Not PRCity.GetByID(DED!CityID) Then
'''                    MsgBox "PRCity Error: " & DED!CityID, vbCritical
'''                    End
'''                End If
'''            End If
'''            DED!Desc = Trim(PRCity.ShortName) & " " & Format(DED!CityWage, "###,##0.00")
'''            DED.Update
'''        End If
'''
'''    End If
    
    ' ======================================================================================
    
    ' show the taxable wage on the standard tax lines and tax withheld
    DED.MoveFirst
    Do
    
        If DED!SDTax = True Then GoTo NxtDed
    
        ' SS Tax
        If DED!Title = 99991 Then
            
            ' SS Tax - ===> *** YTD MAX
            ' YTDSSWage does not include this pay
            If YTDSSWage >= SSMax Then
                SSWage = 0
            ElseIf YTDSSWage + SSWage >= SSMax Then
                SSWage = SSMax - YTDSSWage
            End If
            
            ' *** global percentage ???
            If DED!AmountManual = False Then
                If PREmployee.NoSSTax = 1 Then
                    DED!Amount = 0
                    DED.Fields("Desc") = "X " & Format(SSWage, "$###,##0.00")
                Else
                    DED.Fields("Amount") = Round(SSWage * SSPct / 100, 2)
                    DED.Fields("Desc") = Format(SSWage, "$###,##0.00")
                End If
            End If
            TaxTotal = TaxTotal + DED!Amount
        End If
        
        ' MED Tax
        If DED!Title = 99992 Then
            If DED!AmountManual = False Then
                If PREmployee.NoMedTax = 1 Then
                    DED!Amount = 0
                    DED.Fields("Desc") = "X " & Format(MEDWage, "$###,##0.00")
                Else
                    DED!Amount = Round(MEDWage * MedPct / 100, 2)
                    DED.Fields("Desc") = Format(MEDWage, "$###,##0.00")
                End If
            End If
            TaxTotal = TaxTotal + DED!Amount
        End If
            
        ' FWT Tax
        If DED!Title = 99993 Then
            If DED!AmountManual Then OrigAmt = DED!Amount
            If PREmployee.NoFedTax = 1 Then
                p1 = 0
                DED.Fields("Desc") = "X " & Format(FWTWage, "$###,##0.00")
            ElseIf PREmployee.FWTBasis = PREquate.BasisPercent Then
                p1 = Round(FWTWage * PREmployee.FWTAmount / 100, 2)
                DED.Fields("Desc") = Format(PREmployee.FWTAmount / 100, "##0.00%") & " " & Format(FWTWage, "$###,##0.00")
            Else
                If PREmployee.FWTMarried = 1 Then
                    MarSng = "M"
                Else
                    MarSng = "S"
                End If
                If DED!AmountManual = False Then
                    If PREmployee.PaysPerYear = 0 Then PREmployee.PaysPerYear = 52
                    FWTAGI = Round((FWTWage * PREmployee.PaysPerYear), 2) - Round((PREmployee.FWTAmount * FedAllow), 2)
                    If FWTAGI > 0 Then
                        p1 = PRFWTTable.GetFWT(0, MarSng, Int(PRBatch.YearMonth / 100), PRBatch.YearMonth Mod 100, FWTAGI)
                        p1 = p1 / PREmployee.PaysPerYear
                    Else
                        p1 = 0
                    End If
                    p1 = Round(p1, 2)
                End If
                DED.Fields("Desc") = Trim(MarSng) & PREmployee.FWTAmount & " " & Format(FWTWage, "$###,##0.00")
            End If
            If PREmployee.FWTExtraAmount <> 0 Then
                If PREmployee.FWTExtraBasis = PREquate.BasisAmount Then
                    DED!Amount = p1 + PREmployee.FWTExtraAmount
                    If DED!Amount < 0 Then DED!Amount = 0
                    DED!Desc = "* " & Trim(DED!Desc)
                ElseIf PREmployee.FWTExtraBasis = PREquate.BasisPercent Then
                    DED!Amount = p1 + Round((p1 * PREmployee.FWTExtraAmount / 100), 2)
                    If DED!Amount < 0 Then DED!Amount = 0
                    DED!Desc = "* " & Trim(DED!Desc)
                End If
            Else
                DED!Amount = p1
            End If
            If DED!AmountManual = True Then DED!Amount = OrigAmt
            TaxTotal = TaxTotal + DED!Amount
        End If
        
        ' SWT Tax
        If DED!Title = 99994 Then
            
            ' what state to use?
            ' save button enforces one state per check
            ' ErnCityID has last CityID from the earnings grid
            '
            ' if not assigned - use employer default city
            If ErnCityID = 0 Then ErnCityID = PRCompany.DfltCityID
            
            ' default to OHIO if all else fails ....
            SWTState = "OH"
            If ErnCityID > 0 And PRCity.GetByID(ErnCityID) = True Then
                If PRState.GetByID(PRCity.StateID) = True Then
                    SWTState = PRState.StateAbbrev
                End If
            End If
            
            If DED!AmountManual = True Then OrigAmt = DED!Amount
            
            ' Ohio state tax by default
            If PREmployee.NoStateTax = 1 Then
                p1 = 0
                DED.Fields("Desc") = "X " & Format(SWTWage, "$###,##0.00")
            ElseIf PREmployee.SWTBasis = PREquate.BasisPercent Then
                p1 = Round(SWTWage * PREmployee.SWTAmount / 100, 2)
                DED.Fields("Desc") = Format(PREmployee.SWTAmount / 100, "##0.00%") & " " & Format(SWTWage, "$###,##0.00")
            Else
                If PREmployee.PaysPerYear = 0 Then PREmployee.PaysPerYear = 52
                
                If SWTState = "MD" Then
                    
                    ' standard deduction - 15% of wage - min $1,500 / max $2,000
                    p1 = Round((SWTWage * PREmployee.PaysPerYear) * 0.15, 2)
                    If p1 < 1500 Then p1 = 1500
                    If p1 > 2000 Then p1 = 2000
                        
                    ' state exemption
                    P2 = PREmployee.SWTAmount * 3200
                    
                    SWTAGI = (SWTWage * PREmployee.PaysPerYear) - p1 - P2
                    If SWTAGI < 5000 Then SWTAGI = 0
                    
                Else        ' default to OH otherwise
                    SWTAGI = Round((SWTWage * PREmployee.PaysPerYear), 2) - Round((PREmployee.SWTAmount * OHAllow), 2)
                End If
                
                If PREmployee.FWTMarried = 1 Then
                    MarSng = "M"
                Else
                    MarSng = "S"
                End If
                
                If SWTAGI > 0 Then
                    If SWTState = "MD" Then
                        p1 = SWTAGI * 0.06
                    Else
                        p1 = PRFWTTable.GetFWT(36, "X", Int(PRBatch.YearMonth / 100), PRBatch.YearMonth Mod 100, SWTAGI)
                    End If
                    p1 = p1 / PREmployee.PaysPerYear
                Else
                    p1 = 0
                End If
                p1 = Round(p1, 2)
                DED.Fields("Desc") = Trim(MarSng) & PREmployee.SWTAmount & " " & Format(SWTWage, "$###,##0.00")
            End If
            If PREmployee.SWTExtraAmount <> 0 Then
                If PREmployee.SWTExtraBasis = PREquate.BasisAmount Then
                    DED!Amount = p1 + PREmployee.SWTExtraAmount
                    If DED!Amount < 0 Then DED!Amount = 0
                    DED!Desc = "* " & Trim(DED!Desc)
                ElseIf PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
                    DED!Amount = p1 + Round((p1 * PREmployee.SWTExtraAmount / 100), 2)
                    If DED!Amount < 0 Then DED!Amount = 0
                    DED!Desc = "* " & Trim(DED!Desc)
                End If
            Else
                DED!Amount = p1
            End If
            If DED!AmountManual = True Then DED!Amount = OrigAmt
            TaxTotal = TaxTotal + DED!Amount
        End If
        
        ' CWT Tax
        If DED!Title = 99995 Then
            
            ' *** this is already calculated ***
            ' DEDTotal = DEDTotal + DED!Amount
            X = ""
            If PRCity.GetByID(DED!CityID) Then
                X = Mid(PRCity.ShortName, 1, 7)
            End If
            
            ' put the orig amt back if manual flag set
            If DED!AmountManual = True Then
                DED!Amount = DED!OrigAmount
            End If
            
            DED.Fields("Desc") = Trim(X) & " " & Format(DED!CityWage, "$###,##0.00")
            TaxTotal = TaxTotal + DED!Amount
        
        End If
        
        ' Courtesy Tax
        If DED!Title = 99996 Then
            
            ' *** this is already calculated ***
            ' DEDTotal = DEDTotal + DED!Amount
            X = ""
            If PRCity.GetByID(DED!CityID) Then
                X = Mid(PRCity.ShortName, 1, 7)
            End If
            
            ' put the orig amt back if manual flag set
            If DED!AmountManual = True Then
                DED!Amount = DED!OrigAmount
            End If
            
            DED.Fields("Desc") = Trim(X) & " " & Format(TotalCityWage - CWTDed, "$###,##0.00")
            TaxTotal = TaxTotal + DED!Amount
        
        End If
        
        DED!Amount = Round(DED!Amount, 2)
        DED.Update
        
NxtDed:
        DED.MoveNext
        If DED.EOF Then Exit Do
    
    Loop

    ' OHIO school district tax
    ' SWT wage had to be calculated first
    DED.MoveFirst
    Do
        If DED!SDTax = True Then
            If DED!AmountManual = True Then
                ' leave it alone
            ElseIf SWTWage = 0 Then
                DED!Amount = 0
            Else
                ' get the PRItem record
                If Not PRItem.GetByID(DED!Title) Then
                    MsgBox "PRItem not found: " & DED!Title, vbExclamation
                    Exit Sub
                End If
                ' annualize the SWT wage
                AnnSWTWage = SuperRound(SWTWage, PREmployee.PaysPerYear)
                ' deduct the state exemptions
                If PREmployee.SWTBasis = PREquate.BasisExemptions Then
                    SDEXAmt = SuperRound(OHSDAllow, PREmployee.SWTAmount)
                Else
                    SDEXAmt = 0  ' ??
                End If
                AnnSWTWage = AnnSWTWage - SDEXAmt
                DED!Amount = AnnSWTWage * PRItem.AmtPct / 100
                DED!Amount = Round(DED!Amount / PREmployee.PaysPerYear, 2)
            End If
            DEDTotal = DEDTotal + DED!Amount
        End If
        DED.MoveNext
    Loop Until DED.EOF
    
    ' not in net
    If NotInNetTotal <> 0 Then
        DED.MoveFirst
        Do
            If DED!NotInNet = True And DED!AmountManual = False Then
                DED!Amount = NotInNetTotal
                DED.Update
                NotInNetTotal = 0
                DEDTotal = DEDTotal + DED!Amount
                Exit Do
            End If
            DED.MoveNext
        Loop Until DED.EOF
    End If
    
    ' direct deposit calcs
    ' loop thru twice
    ' once for non-net basis
    ' second for remaining net
    Net = ERNTotal - DEDTotal - TaxTotal - NotInNetTotal
    
    OrigNet = Net
    DDTotal = 0
    DED.MoveFirst
    Do
        If DED!ItemType = PREquate.ItemTypeDirDepDed And DED!DirDepBasis <> PREquate.DirDepBasisNet Then
            If DED!AmountManual = False Then
                
                If DED!DirDepBasis = PREquate.DirDepBasisAmt Then
                    p1 = DED!DirDepAmtPct
                Else
                    p1 = Round(DED!DirDepAmtPct * OrigNet / 100, 2)
                End If
                
                If p1 <= Net Then
                    DED!Amount = p1
                    Net = Net - p1
                Else
                    DED!Amount = Net
                    Net = 0
                End If
                
                DED.Update
            
            End If
            DDTotal = DDTotal + DED!Amount
        End If
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop
    
    ' loop again for the DD amount to net
    DED.MoveFirst
    Do
        If DED!ItemType = PREquate.ItemTypeDirDepDed And DED!DirDepBasis = PREquate.DirDepBasisNet Then
            If DED!AmountManual = False Then
                DED!Amount = Net
                DED.Update
            End If
            DDTotal = DDTotal + DED!Amount
        End If
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop
    
'========================
        
    ' sort the deductions
    DED.Sort = "DedSort"

    ' display totals
    Me.tdbnumHrTotal = HRTotal
    Me.tdbnumDedTotal = DEDTotal
    Me.tdbnumERNTotal = ERNTotal
    Me.tdbnumTaxTotal = TaxTotal
    Me.tdbnumNetPayTotal = ERNTotal - DEDTotal - TaxTotal - NotInNetTotal
    Me.tdbnumDirDepTotal = DDTotal
    Me.tdbnumCheckTotal = ERNTotal - DEDTotal - TaxTotal - DDTotal - NotInNetTotal

    fgERN.ShowCell fgERNRow, fgERNCol
    fgERN.Select fgERNRow, fgERNCol
    fgERN.Refresh
    
    If fgDEDRow > fgDED.Rows Then
        fgDEDRow = fgDED.Rows
    End If
    If fgDEDCol > fgDED.Cols Then
        fgDEDCol = fgDED.Cols
    End If
    
    If fgDEDRow < fgDED.Rows Then
        fgDED.ShowCell fgDEDRow, fgDEDCol
        fgDED.Select fgDEDRow, fgDEDCol
    End If
    
    fgDED.Refresh

End Sub

Private Sub cmdClearManual_Click()

    ERN.MoveFirst
    Do
        ERN!AmountManual = False
        ERN!CityManual = False
        ERN.Update
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    Loop

    DED.MoveFirst
    Do
        DED!AmountManual = False
        DED.Update
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop

End Sub

Private Sub cmdSetManual_Click()

    ERN.MoveFirst
    Do
        ERN!AmountManual = True
        ERN!CityManual = True
        ERN.Update
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    Loop

    DED.MoveFirst
    Do
        DED!AmountManual = True
        DED.Update
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop

End Sub

Private Sub cmdReCalc_Click()
    CalcGrids
End Sub

Private Sub fgemp_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
Dim SortString, ColHeader As String
Dim ERecs As Long
Dim TID As Long
    
    If Button <> 1 Then Exit Sub
    If Shift <> 0 Then Exit Sub
    If fgEMP.MouseCol > 2 Then Exit Sub
    If fgEMP.MouseRow <> 0 Then Exit Sub
    
    TID = EMP!TempID
    
    ' unbold the old sort column
    fgEMP.Cell(flexcpFontBold, 0, SortCol) = False
    
    ' what order - toggle or ascending for new sort col
    If fgEMP.MouseCol = SortCol Then
        If SortOrder = 1 Then
            SortOrder = 2
        Else
            SortOrder = 1
        End If
    Else
        SortCol = fgEMP.MouseCol
        SortOrder = 1
    End If
    
    ' col header modify
    fgEMP.Cell(flexcpFontBold, 0, SortCol) = True
        
    ' reset the headers
    fgEMP.TextMatrix(0, 0) = "EE#"
    fgEMP.TextMatrix(0, 1) = "N A M E"
    fgEMP.TextMatrix(0, 2) = "Check#"
    
    ' sort it
    If SortCol = 0 Then
        SortString = "EmployeeNumber"
    ElseIf SortCol = 1 Then
        SortString = "EmployeeName"
    ElseIf SortCol = 2 Then
        SortString = "CheckNumber"
    ElseIf SortCol = 11 Then
        SortString = "DptEE"
    End If

    ' signal the order
    If SortOrder = 1 Then
        fgEMP.TextMatrix(0, SortCol) = fgEMP.TextMatrix(0, SortCol) & "+"
    Else
        fgEMP.TextMatrix(0, SortCol) = fgEMP.TextMatrix(0, SortCol) & "-"
        SortString = Trim(SortString) & " DESC"
    End If
    
    TID = EMP!TempID
    
    EMP.Sort = SortString
    
    fgEMP.Select 1, 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF10: cmdSave_Click
        Case vbKeyF11: cmdSkip_Click
    End Select
    
End Sub

Private Sub cmdSave_Click()

Dim LastDistID, StateID As Long
Dim ErnCount As Long
Dim CourtTaxAmt, CourtTaxRmn As Currency
Dim TotalCityWage As Currency
Dim CourtManual As Boolean
Dim SaveFlag As Boolean
Dim Ovr As Boolean

    ' ************************************************
    ' record set to write CWT to PRDist
    ' store city wage / tax by city for all earnings
Dim rsCWT As New ADODB.Recordset
    
    rsCWT.CursorLocation = adUseClient
    rsCWT.Fields.Append "CityID", adDouble
    rsCWT.Fields.Append "CityWage", adCurrency
    rsCWT.Fields.Append "CityTax", adCurrency
    rsCWT.Fields.Append "CityTaxRmn", adCurrency
    rsCWT.Fields.Append "Count", adDouble
    rsCWT.Open , , adOpenDynamic, adLockOptimistic
    ' ************************************************

    ' negative net pay warning
    If IsNull(tdbnumNetPayTotal.Value) Then tdbnumNetPayTotal.Value = 0
    If tdbnumNetPayTotal.Value < 0 Then
        If MsgBox("Warning - Negative Net Pay Amount" & vbCr & _
                  "OK to save???", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    ' Only ONE State allowed
    ' also update earnings by CityID recordset
    StateID = 0
    TotalCityWage = 0
    ERN.MoveFirst
    Do
        
        ' get the StateID from PRCity
        If PRCity.GetByID(ERN!CityID) Then
            If StateID <> 0 And PRCity.StateID <> StateID Then
                MsgBox "Only ONE state allowed per check!", vbExclamation
                Exit Sub
            End If
        End If
        StateID = PRCity.StateID

        ' update record set by city
        SQLString = "CityID = " & ERN!CityID
        rsCWT.Find SQLString, 0, adSearchForward, 1
        If rsCWT.EOF Then
            rsCWT.AddNew
            rsCWT!CityID = ERN!CityID
            rsCWT!CityWage = 0
            rsCWT!CityTax = 0
            rsCWT!CityTaxRmn = 0
            rsCWT!Count = 0
        End If
        rsCWT!CityWage = rsCWT!CityWage + ERN!CityWage
        rsCWT!Count = rsCWT!Count + 1
        rsCWT.Update

        TotalCityWage = TotalCityWage + ERN!CityWage

        ERN.MoveNext
    Loop Until ERN.EOF
        
    ' loop thru the city tax ded lines
    CourtTaxAmt = 0
    DED.MoveFirst
    Do
        If DED!Title = 99995 Then
            SQLString = "CityID = " & DED!CityID
            rsCWT.Find SQLString, 0, adSearchForward, 1
            If rsCWT.EOF Then
                MsgBox "CWT data error!", vbExclamation
                GoBack
            End If
            rsCWT!CityTax = DED!Amount
            rsCWT!CityTaxRmn = DED!Amount
            rsCWT.Update
        End If
        If DED!Title = 99996 Then
            CourtManual = DED!AmountManual
            CourtTaxAmt = DED!Amount
            CourtTaxRmn = DED!Amount
        End If
        DED.MoveNext
    Loop Until DED.EOF
        
    ' get the employee record
    If Not PREmployee.GetByID(EMP!EmployeeID) Then
        MsgBox "Employee Error!!!", vbCritical
        End
    End If

    ' ask if to overwrite existing history and related
    If EMP!HistID <> 0 Then
        If MsgBox("OK to overwrite existing history???", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
        Ovr = True
            
        ' delete related PRDist and PRHist
        SQLString = "DELETE * FROM PRDist WHERE PRDist.HistID = " & EMP!HistID
        rsInit SQLString, cn, rs
        PRDist.OpenRS
        
        SQLString = "DELETE * FROM PRItemHist WHERE PRItemHist.HistID = " & EMP!HistID
        rsInit SQLString, cn, rs
        PRItemHist.OpenRS
    
        If Not PRHist.GetByID(EMP!HistID) Then
            MsgBox "PRHist NF: " & EMP!HistID, vbExclamation
            End
        End If
    
    Else    ' create new PRHist record
        
        Ovr = False
        
        PRHist.Clear
        PRHist.EmployeeID = EMP!EmployeeID
        PRHist.BatchID = PRBatch.BatchID
        PRHist.YearMonth = PRBatch.YearMonth
        PRHist.CheckDate = PRBatch.CheckDate
        PRHist.PEDate = PRBatch.PEDate
        PRHist.DepartmentID = PREmployee.DepartmentID
        PRHist.NotInNetAmount = NotInNetTotal
        PRHist.Save (Equate.RecAdd)
        
        EMP!HistID = PRHist.HistID
        
    End If
        
    EMP!Saved = True
    EMP!CheckNumber = nNull(Me.tdbnumCheckNum)
    EMP.Update
    
    If PRBatch.PEDate > PREmployee.DateLastPaid Then
        PREmployee.DateLastPaid = PRBatch.PEDate
        PREmployee.Save (Equate.RecPut)
    End If
    
    ' update the history fields
    PRHist.CheckNumber = EMP!CheckNumber
    
    PRHist.RegHours = 0
    PRHist.RegAmount = 0
    PRHist.OTHours = 0
    PRHist.OTAmount = 0
    PRHist.OEHours = 0
    PRHist.OERate = 0
    PRHist.OEAmount = 0
    PRHist.SSWage = 0
    PRHist.SSTax = 0
    PRHist.MEDWage = 0
    PRHist.MedTax = 0
    PRHist.FWTWage = 0
    PRHist.FWTTax = 0
    PRHist.SWTWage = 0
    PRHist.SWTTax = 0
    PRHist.CWTWage = 0
    PRHist.CWTTax = 0
    PRHist.Deductions = 0
    PRHist.DirectDeposit = 0
    PRHist.Gross = 0
    PRHist.Net = 0
    PRHist.FUNWage = 0
    PRHist.SUNWage = 0
        
    ' loop the earnings
    StateID = 0
    ErnCount = 0
    ERN.MoveFirst
    Do
        
        ErnCount = ErnCount + 1
        
        ' ******************************
        ' If ERN!Hours = 0 And ERN!Rate = 0 And ERN!Amount = 0 Then GoTo NextErn
        ' ******************************
                
        If ERN!Title = 99991 Then
            PRHist.RegHours = PRHist.RegHours + ERN!Hours
            PRHist.RegRate = ERN!Rate
            PRHist.RegAmount = PRHist.RegAmount + ERN!Amount
        ElseIf ERN!Title = 99992 Then
            PRHist.OTHours = PRHist.OTHours + ERN!Hours
            PRHist.OTRate = ERN!Rate
            PRHist.OTAmount = PRHist.OTAmount + ERN!Amount
        Else
            PRHist.OEHours = PRHist.OEHours + ERN!Hours
            PRHist.OERate = ERN!Rate    ' ???
            PRHist.OEAmount = PRHist.OEAmount + ERN!Amount
        End If
        
        If Not ERN!NoSSTax Then PRHist.SSWage = PRHist.SSWage + ERN!Amount
        If Not ERN!NoMedTax Then PRHist.MEDWage = PRHist.MEDWage + ERN!Amount
        If Not ERN!NoFWTTax Then PRHist.FWTWage = PRHist.FWTWage + ERN!Amount
        If Not ERN!NoSWTTax Then PRHist.SWTWage = PRHist.SWTWage + ERN!Amount
        If Not ERN!NoCWTTax Then PRHist.CWTWage = PRHist.CWTWage + ERN!Amount
        If Not ERN!NoFUNTax Then PRHist.FUNWage = PRHist.FUNWage + ERN!Amount
        If Not ERN!NoSUNTax Then PRHist.SUNWage = PRHist.SUNWage + ERN!Amount
        
        PRHist.Gross = PRHist.Gross + ERN!Amount
        
        ' multiple states???
        PRHist.SWTTax = PRHist.SWTTax + ERN!StateTax
        ' city tax totaled during ded loop
        
        ' get the City record
        If ERN!CityID <> 999999 Then
            If Not PRCity.GetByID(ERN!CityID) Then
                MsgBox "PRCity NF: " & ERN!CityID, vbCritical
                End
            End If
        End If
        
        ' write to PRDist
        PRDist.Clear
        PRDist.EmployeeID = PREmployee.EmployeeID
        PRDist.BatchID = PRBatch.BatchID
        PRDist.HistID = PRHist.HistID
        PRDist.StateID = PRCity.StateID
        If ERN!CityID = 999999 Then
            PRDist.CityID = 0
        Else
            PRDist.CityID = ERN!CityID
        End If
        PRDist.DepartmentID = PREmployee.DepartmentID
        PRDist.YearMonth = PRBatch.YearMonth
        PRDist.PEDate = PRBatch.PEDate
        PRDist.CheckDate = PRBatch.CheckDate
        
        If ERN!Title = 99991 Then
            PRDist.DistType = PREquate.DistTypeReg
            PRDist.ItemID = 1
            PRDist.ItemType = PREquate.ItemTypeRegPay
        ElseIf ERN!Title = 99992 Then
            PRDist.DistType = PREquate.DistTypeOT
            PRDist.ItemID = 1
            PRDist.ItemType = PREquate.ItemTypeOvtPay
        Else
            PRDist.DistType = PREquate.DistTypeItem
            PRDist.ItemID = ERN!Title
            PRDist.ItemType = PREquate.ItemTypeOE
            PRDist.EmployerItemID = ERN!EmployerItemID
        End If
            
        PRDist.Hours = ERN!Hours
        PRDist.Rate = ERN!Rate
        PRDist.Amount = ERN!Amount
        PRDist.ManualAmount = ERN!AmountManual
                
        PRDist.GrossWage = ERN!Amount
        
        If Not ERN!NoSWTTax Then PRDist.StateWage = ERN!Amount
        PRDist.StateTax = ERN!StateTax
        ' PRDist.ManualStateTax = ERN!ManualStateTax
        
        If ERN!NoCWTTax = False Then PRDist.CityWage = ERN!CityWage
        
        ' write CityTax proportionally per city
        ' necessary in case CWT is manual
        SQLString = "CityID = " & ERN!CityID
        rsCWT.Find SQLString, 0, adSearchForward, 1
        If rsCWT.EOF Then
            MsgBox "CWT Recordset Err!", vbExclamation
            GoBack
        End If
        
        If rsCWT!Count = 1 Then         ' take the rest
            PRDist.CityTax = rsCWT!CityTaxRmn
        ElseIf rsCWT!CityWage = 0 Then
            PRDist.CityTax = 0
        Else
            p1 = Round(ERN!CityWage / rsCWT!CityWage * rsCWT!CityTax, 2)
            PRDist.CityTax = p1
            rsCWT!CityTaxRmn = rsCWT!CityTaxRmn - p1
        End If
        rsCWT!Count = rsCWT!Count - 1
        rsCWT.Update
        
        PRDist.ManualCityTax = ERN!CityManual
        
        ' if courtesy tax not manual - take from earnings record
        If CourtManual = False Then
            PRDist.CourtesyCityTax = ERN!CourtTax
        Else        ' split by city wage
            If TotalCityWage = 0 Then
                p1 = 0
            Else
                p1 = Round(PRDist.CityWage / TotalCityWage * CourtTaxAmt, 2)
            End If
            CourtTaxRmn = CourtTaxRmn - p1
            PRDist.CourtesyCityTax = p1
            LastDistID = PRDist.DistID
        End If
        
        PRDist.CourtesyCityID = CourtCityID
        
        ' other fields ....
            
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        
        PRDist.BillingRate = 0
        PRDist.JobID = ERN!JobID
        PRDist.CustomerID = 0
        
        If ERN!NotInNet = True Then
            PRDist.NotInNet = 1
        Else
            PRDist.NotInNet = 0
        End If
        
        SaveFlag = False
        If PRDist.Amount <> 0 Then SaveFlag = True
        If PRDist.Hours <> 0 Then SaveFlag = True
        If PRDist.CityWage <> 0 Then SaveFlag = True
        If PRDist.CityTax <> 0 Then SaveFlag = True
        If PRDist.CourtesyCityTax <> 0 Then SaveFlag = True
        If PRDist.Amount <> 0 Then SaveFlag = True
        If SaveFlag = True Then
            PRDist.Save (Equate.RecAdd)
        End If

NextErn:
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    Loop

    ' update TimeSheet records?
    If JobDist = True Then
    
        With frmSelTimeSheets.rsTimeSheet
    
            If .RecordCount > 0 Then
        
                .MoveFirst
                Do
                    
                    If !Selected = True Then
                    
                        SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(!WEDate) & _
                                    " AND EmployeeID = " & EMP!EmployeeID & _
                                    " AND TotalHours <> 0" & _
                                    " AND BatchID = 0"
                        If PRTimeSheet.GetBySQL(SQLString) = True Then
                            Do
                                PRTimeSheet.BatchID = PRDist.BatchID
                                PRTimeSheet.HistID = PRDist.HistID
                                PRTimeSheet.Save (Equate.RecPut)
                                If PRTimeSheet.GetNext = False Then Exit Do
                            Loop
                        End If
                    End If
                    
                    .MoveNext
                Loop Until .EOF
            End If
    
        End With
    
    End If

    ' round off court tax if manual
    If CourtManual = True And CourtTaxRmn <> 0 And LastDistID <> 0 Then
        If PRDist.GetByID(LastDistID) = False Then
            MsgBox "Manual Courtesy error!", vbExclamation
            GoBack
        End If
        PRDist.CourtesyCityTax = PRDist.CourtesyCityTax + p1
        PRDist.Save (Equate.RecPut)
    End If

    ' ??? multiple states ???
    PRHist.CWTTax = 0

    ' loop thru the deductions
    DED.MoveFirst
    Do
        
        If DED!Amount = 0 Then GoTo NextDed
        
        If DED!Title = 99991 Then
            PRHist.SSTax = DED!Amount
        ElseIf DED!Title = 99992 Then
            PRHist.MedTax = DED!Amount
        ElseIf DED!Title = 99993 Then
            PRHist.FWTTax = DED!Amount
        ElseIf DED!Title = 99994 Then
            PRHist.SWTTax = DED!Amount
        ElseIf DED!Title = 99995 Then
            PRHist.CWTTax = PRHist.CWTTax + DED!Amount
        ElseIf DED!Title = 99996 Then
            PRHist.CWTTax = PRHist.CWTTax + DED!Amount
        Else
            If DED!ItemType = PREquate.ItemTypeDED Or DED!ItemType = PREquate.ItemTypeSDTax Then
                PRHist.Deductions = PRHist.Deductions + DED!Amount
            End If
            If DED!ItemType = PREquate.ItemTypeDirDepDed Then
                PRHist.DirectDeposit = PRHist.DirectDeposit + DED!Amount
            End If
        End If
        
        If DED!Title >= 99991 Then GoTo NextDed
        
        PRItemHist.Clear
        PRItemHist.EmployeeID = PREmployee.EmployeeID
        PRItemHist.HistID = PRHist.HistID
        PRItemHist.BatchID = PRBatch.BatchID
        PRItemHist.ItemID = DED!Title
        PRItemHist.ItemType = DED!ItemType
        PRItemHist.YearMonth = PRBatch.YearMonth
        PRItemHist.PEDate = PRBatch.PEDate
        PRItemHist.CheckDate = PRBatch.CheckDate
        PRItemHist.Amount = DED!Amount
        PRItemHist.ManualAmount = DED!AmountManual
        PRItemHist.EmployerItemID = DED!EmployerItemID
        PRItemHist.DepartmentID = PREmployee.DepartmentID
        
        If DED!NoSSTax Then PRHist.SSWage = PRHist.SSWage - DED!Amount
        If DED!NoMedTax Then PRHist.MEDWage = PRHist.MEDWage - DED!Amount
        If DED!NoFWTTax Then PRHist.FWTWage = PRHist.FWTWage - DED!Amount
        If DED!NoSWTTax Then PRHist.SWTWage = PRHist.SWTWage - DED!Amount
        
        ' %%% court cwt also %%%
        If DED!NoCWTTax Then PRHist.CWTWage = PRHist.CWTWage - DED!Amount
        
        If DED!NoFUNTax Then PRHist.FUNWage = PRHist.FUNWage - DED!Amount
        If DED!NoSUNTax Then PRHist.SUNWage = PRHist.SUNWage - DED!Amount
        
        PRItemHist.Percent = DED!AmtPct
                
        ' wage excluded from basis for deduct by percent (401k match purposes)
        PRItemHist.WageExcluded = nNull(DED!WageExcluded)
        
        ' PRItemHist.WageBase = DED!Basis
        PRItemHist.Save (Equate.RecAdd)

NextDed:
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop
    
    ' state info
    PRHist.StateID = PRDist.StateID
    ' SUN max wage
    If Not PRState.GetByID(PRDist.StateID) Then
        SUNMax = 99999999.99
    Else
        SUNMax = PRState.UnEmpMax
        If SUNMax = 0 Then SUNMax = 99999999.99
    End If

    ' %%% set wage base amounts equal to taxable wage
    PRHist.SSWageBase = PRHist.SSWage
    PRHist.FUNWageBase = PRHist.FUNWage
    PRHist.SUNWageBase = PRHist.SUNWage

    ' %%% top off ss / fun / sun to max
    If YTDSSWage + PRHist.SSWageBase > SSMax Then
        PRHist.SSWage = SSMax - YTDSSWage
    End If
    If YTDFUNWage + PRHist.FUNWageBase > FUNMax Then
        PRHist.FUNWage = FUNMax - YTDFUNWage
    End If
    If YTDSUNWage + PRHist.SUNWageBase > SUNMax Then
        PRHist.SUNWage = SUNMax - YTDSUNWage
    End If

    ' %%% don't store EMPLOYER expense tagable wage
    '     if employee flag is set
    If PREmployee.NoFedUnemp = 1 Then
        PRHist.FUNWage = 0
        PRHist.FUNWageBase = 0
    End If
    If PREmployee.NoStateUnemp = 1 Then
        PRHist.SUNWage = 0
        PRHist.SUNWageBase = 0
    End If

    PRHist.WkcAmount = Round(EMP!WkcPct / 100 * PRHist.Gross, 2)
 
    PRHist.Net = Me.tdbnumNetPayTotal - PRHist.DirectDeposit
    PRHist.NotInNetAmount = NotInNetTotal
    PRHist.Save (Equate.RecPut)

    EditFlag = False

    If fgEMP.Row < fgEMP.Rows - 1 Then
        fgEMP.Row = fgEMP.Row + 1
    End If

    ' increment the next check number if not history override
    If Ovr = False Then
        NextCheckNumber = NextCheckNumber + 1
    End If
    
    ' assign the check number if not already set
    EMP!CheckNumber = nNull(EMP!CheckNumber)
    EMP.Update
    If EMP!CheckNumber = 0 Then
        Me.tdbnumCheckNum = NextCheckNumber
    Else
        Me.tdbnumCheckNum = EMP!CheckNumber
    End If

    BatchTotals

End Sub

Private Sub cmdSkip_Click()

    If fgEMP.Row < fgEMP.Rows - 1 Then
        fgEMP.Row = fgEMP.Row + 1
    End If

End Sub

Private Sub cmdAddEarn_Click()

    ERN.AddNew
    ERN!CityID = DfltCityID
    ERN!NewFlag = True
    ERN!Title = 99991       ' dflt to reg pay
    ERN.Update

End Sub

Private Sub BatchTotals()

Dim RegHrs, OHrs, TlHrs As Currency
Dim RegErn, OErn, TlErn As Currency
Dim Checks As Long

    ' clear totals
    RegHrs = 0
    OHrs = 0
    TlHrs = 0
    RegErn = 0
    OErn = 0
    TlErn = 0
    Checks = 0
    HiCheckNum = 0

    ' update to PRBatch File - in case last one is deleted
    If Not PRBatch.GetByID(BatchID) Then
        MsgBox "Batch Error: " & BatchID, vbCritical
        End
    End If
    PRBatch.RecCount = Checks
    PRBatch.Save (Equate.RecPut)
    
    ' update to screen controls - in case last one is deleted
    Me.tdbnumBChecks = Checks
    Me.tdbnumBRegHrs = RegHrs
    Me.tdbnumBOHrs = OHrs
    Me.tdbnumBTlHrs = TlHrs
    Me.tdbnumBRegErn = RegErn
    Me.tdbnumBOEarng = OErn
    Me.tdbnumBTlEarng = TlErn

    SQLString = "SELECT * FROM PRHist WHERE BatchID = " & BatchID
    If Not PRHist.GetBySQL(SQLString) Then Exit Sub
    
    Do
    
        Checks = Checks + 1
        SQLString = "SELECT * FROM PRDist WHERE PRDist.HistID = " & PRHist.HistID
        If PRDist.GetBySQL(SQLString) Then
            Do
                If PRDist.DistType = PREquate.DistTypeReg Then
                    RegHrs = RegHrs + PRDist.Hours
                    RegErn = RegErn + PRDist.Amount
                Else
                    OHrs = OHrs + PRDist.Hours
                    OErn = OErn + PRDist.Amount
                End If
                TlHrs = TlHrs + PRDist.Hours
                TlErn = TlErn + PRDist.Amount
                If Not PRDist.GetNext Then Exit Do
            Loop
        End If
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop
    
    ' update to screen controls
    Me.tdbnumBChecks = Checks
    Me.tdbnumBRegHrs = RegHrs
    Me.tdbnumBOHrs = OHrs
    Me.tdbnumBTlHrs = TlHrs
    Me.tdbnumBRegErn = RegErn
    Me.tdbnumBOEarng = OErn
    Me.tdbnumBTlEarng = TlErn
    
    ' update to PRBatch File
    If Not PRBatch.GetByID(BatchID) Then
        MsgBox "Batch Error: " & BatchID, vbCritical
        End
    End If
    
    PRBatch.RecCount = Checks
    
    PRBatch.Save (Equate.RecPut)

End Sub

Private Sub cmdEmpAdd_Click()
    
Dim EmpRow As Long
    
    frmAddEmployee.Init
    frmAddEmployee.Show vbModal
    If frmAddEmployee.EmpID = -1 Then Exit Sub      ' canceled

    If Not PREmployee.GetByID(frmAddEmployee.EmpID) Then
        MsgBox "Employee Not Found! " & frmAddEmployee.EmpID, vbCritical
        End
    End If

    EmpRow = fgEMP.Row
    FirstFlag = True
    EECount = EECount + 1
    EMP.AddNew
    EMP!EmployeeNumber = PREmployee.EmployeeNumber
    EMP.Fields("EmployeeID") = PREmployee.EmployeeID
    EMP.Fields("EmployeeName") = Trim(PREmployee.LFName)
    EMP.Fields("CheckNumber") = 0
    If PREmployee.Salaried Then EMP!Salaried = True
    EMP!Saved = False
    EMP.Fields("HistID") = 0
    EMP.Fields("HistFlag") = False
    EMP!TempID = EECount
    EMP.Update
    FirstFlag = False
    fgEMP.Refresh
    Me.tdbnumCheckNum = NextCheckNumber
    Unload frmAddEmployee

    ' goto the new line - force a change if new line is same as old line
    SQLString = "TempID = " & EECount
    EMP.Find SQLString, 0, adSearchForward, 1
    If fgEMP.Row = EmpRow Then
        fgEMP.Row = fgEMP.Row + 1
        fgEMP.Row = fgEMP.Row - 1
    End If
    fgEMP.Select fgEMP.Row, 0

End Sub

Private Sub cmdDelete_Click()

Dim ChkNum As Long

    If EMP!Saved = False Then Exit Sub

    If MsgBox("OK to delete history record for: " & vbCr & _
              Trim(EMP!EmployeeNumber) & " " & Trim(EMP!EmployeeName) & "?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub
              
    ChkNum = nNull(EMP!CheckNumber)
              
    SQLString = "DELETE * FROM PRDist WHERE PRDist.HistID = " & EMP!HistID
    rsInit SQLString, cn, rs
    PRDist.OpenRS
    
    SQLString = "DELETE * FROM PRItemHist WHERE PRItemHist.HistID = " & EMP!HistID
    rsInit SQLString, cn, rs
    PRItemHist.OpenRS
    
    SQLString = "DELETE * FROM PRHist WHERE PRHist.HistID = " & EMP!HistID
    rsInit SQLString, cn, rs
    PRHist.OpenRS

    EMP!CheckNumber = 0
    EMP!HistID = 0
    EMP!Saved = False
    EMP.Update

    ' update and clear the ERN and DED record sets
    ERN.MoveFirst
    Do
        ERN.Delete
        ERN.MoveNext
        If ERN.EOF Then Exit Do
    Loop
            
    DED.MoveFirst
    Do
        DED.Delete
        DED.MoveNext
        If DED.EOF Then Exit Do
    Loop
    
    SetDataGrids
    CalcGrids
    BatchTotals

    ' deleting the last ck# - re-use it
    If ChkNum = NextCheckNumber - 1 Then
        NextCheckNumber = NextCheckNumber - 1
    End If
    Me.tdbnumCheckNum = NextCheckNumber
    
    ' position the cursor
    fgERN.ShowCell 1, 1
    fgERN.Select 1, 1
    fgERN.Refresh
    
    fgDED.ShowCell 1, 2
    fgDED.Select 1, 2
    fgDED.Refresh


End Sub


Private Sub cmdSave2_Click()
    cmdSave_Click
End Sub

Private Sub GetTimeSheetData(ByRef rsTS As ADODB.Recordset)

Dim TSFlag As Boolean

    ErnCount = 0
    rsTS.MoveFirst
    Do
        If rsTS!Selected = True Then
            SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(rsTS!WEDate) & _
                        " AND EmployeeID = " & EMP!EmployeeID & _
                        " AND TotalHours <> 0" & _
                        " AND BatchID = 0"
            
            If PRTimeSheet.GetBySQL(SQLString) Then
                Do
                    TSFlag = False
                    If ERN.RecordCount > 0 Then
                        ERN.MoveFirst
                        Do
                            If ERN!EmployerItemID = PRTimeSheet.ItemID Then
                                If ERN!JobID = PRTimeSheet.JobID Then
                                    TSFlag = True
                                    Exit Do
                                End If
                            End If
                            ERN.MoveNext
                        Loop Until ERN.EOF
                    End If
                    
                    If TSFlag = False Then
                        ERN.AddNew
                        ERN!JobID = PRTimeSheet.JobID
                                                                    
                        ' suzy
                        ' get the city id
                        If JCJob.GetByID(PRTimeSheet.JobID) = False Then
                            'MsgBox "Job Not Found: " & PRTimeSheet.JobID, vbExclamation
                            'End
                            ERN!CityID = PREmployee.DefaultCityID
                        Else
                            ERN!CityID = JCJob.CityID
                        End If
                        
                        ' regular / overtime
                        If PRTimeSheet.ItemID = 99991 Or PRTimeSheet.ItemID = 99992 Then
                            ERN!Title = PRTimeSheet.ItemID
                            ERN!EmployerItemID = PRTimeSheet.ItemID
                            If PREmployee.Salaried = 1 Then
                                ERN!Rate = PREmployee.SalaryAmount
                                ERN!salary = True
                            Else
                                ERN!Rate = PREmployee.HourlyAmount
                                ERN!salary = False
                            End If
                        Else
                            ' get the employer item id
                            If PRItem.GetByID(PRTimeSheet.ItemID) = False Then
                                MsgBox "Employer Item NF: " & PRTimeSheet.ItemID, vbExclamation
                                End
                            End If
                        
                            ERN.Fields("NoSSTax") = PRItem.NoSSTax
                            ERN.Fields("NoMedTax") = PRItem.NoMedTax
                            ERN.Fields("NoFWTTax") = PRItem.NoFWTTax
                            ERN.Fields("NoSWTTax") = PRItem.NoSWTTax
                            ERN.Fields("NoCWTTax") = PRItem.NoCWTTax
                            ERN.Fields("NoFUNTax") = PRItem.NoFUNTax
                            ERN.Fields("NoSUNTax") = PRItem.NoSUNTax
                            ERN.Fields("Tips") = PRItem.Tips
                            ERN.Fields("NotInNet") = PRItem.NotInNet
                        
                            ' get the EMPLOYEE item
                            SQLString = "SELECT * FROM PRItem WHERE " & _
                                        "EmployeeID = " & EMP!EmployeeID & " AND " & _
                                        "EmployerItemID = " & PRTimeSheet.ItemID
                            If PRItem.GetBySQL(SQLString) = False Then
                                MsgBox "Employee Item NF: " & PRTimeSheet.ItemID, vbExclamation
                                End
                            End If
                        
                            ERN!Title = PRItem.ItemID
                        
                            ' always use the EMPLOYEE item for the basis, rate and amount
                            ERN.Fields("Basis") = PRItem.Basis
                        
                            If PRItem.Basis = PREquate.BasisHourly Then
                                ERN.Fields("Rate") = PRItem.AmtPct
                                ERN.Fields("Amount") = 0
                            Else
                                ERN.Fields("Rate") = 0
                                ' ERN.Fields("Amount") = PRItem.AmtPct
                            End If
                                                    
                            ' use the employee defn of tax flags???
                            If PRItem.UseEmployer = 0 Then
                                ERN.Fields("NoSSTax") = PRItem.NoSSTax
                                ERN.Fields("NoMedTax") = PRItem.NoMedTax
                                ERN.Fields("NoFWTTax") = PRItem.NoFWTTax
                                ERN.Fields("NoSWTTax") = PRItem.NoSWTTax
                                ERN.Fields("NoCWTTax") = PRItem.NoCWTTax
                                ERN.Fields("NoFUNTax") = PRItem.NoFUNTax
                                ERN.Fields("NoSUNTax") = PRItem.NoSUNTax
                                ERN.Fields("Tips") = PRItem.Tips
                                ERN.Fields("NotInNet") = PRItem.NotInNet
                            End If
                        
                        End If
                                                                
                    End If
                                            
                    ERN!Hours = ERN!Hours + PRTimeSheet.TotalHours
                    ERN!Amount = ERN!Hours * ERN!Rate
                    ERN.Update
                    TimeSheet = True
                    ErnCount = ErnCount + 1
                                
                    If PRTimeSheet.GetNext = False Then Exit Do
                
                Loop
            End If
        End If
        rsTS.MoveNext
    
    Loop Until rsTS.EOF

    If ErnCount = 0 Then TimeSheet = False

End Sub

Private Sub DedBasisCreate()
    
Dim rsDB As New ADODB.Recordset
    
    ' if the deduction is by percent
    ' make a recordset of earning items to be excluded
    If DED.RecordCount = 0 Then Exit Sub
            
    DED.MoveFirst
    Do
        
        If PRItem.GetByID(DED!ItemID) = False Then
        End If
        
        If PRItem.UseEmployer Then
            If PRItem.GetByID(DED!EmployerItemID) = False Then
            End If
            SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDeductBasis & _
                        " AND UserID = " & PRCompany.CompanyID & _
                        " AND Description = '" & PRItem.ItemID & "'" & _
                        " AND Var1 = '0'"
        Else
            SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDeductBasis & _
                        " AND UserID = " & PRCompany.CompanyID & _
                        " AND Description = '" & PRItem.EmployerItemID & "'" & _
                        " AND Var1 = '" & PREmployee.EmployeeID & "'"
        End If
        
        If PRGlobal.GetBySQL(SQLString) Then
        
            ' check the string to parse
            If PRGlobal.Var2 = "" Then
            ElseIf IsNull(PRGlobal.Var2) Then
            Else
                Set rsDB = ParseString(PRGlobal.Var2, "/")
                If rsDB.RecordCount > 0 Then
                    rsDB.MoveFirst
                    Do
                        rsDedBasis.AddNew
                        rsDedBasis!DeductionID = DED!EmployerItemID
                        rsDedBasis!EarningID = rsDB!listvalue
                        rsDedBasis.Update
                        rsDB.MoveNext
                    Loop Until rsDB.EOF
                End If
            End If
        End If
            
        DED.MoveNext
    
    Loop Until DED.EOF
    
End Sub

Private Sub tdbnumCheckNum_LostFocus()
    ' check number was change manually
    NextCheckNumber = Me.tdbnumCheckNum
End Sub
Private Sub fgERN_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If fgERN.Cell(flexcpForeColor, Row, 0) = vbBlue Then
        MsgBox "Edit of TimeSheet data not allowed!", vbExclamation
        Cancel = True
    End If
    
    
' MsgBox fgERN.Row & vbCr & fgERN.RowData(Row)
    
'    If IsNull(fgERN.RowData(Row)) Then
'        Cancel = False
'        Exit Sub
'    End If
'
'    If fgERN.RowData(Row) <> 0 Then
'        MsgBox "Edit of TimeSheet data not allowed!", vbExclamation
'        Cancel = True
'    End If

End Sub


