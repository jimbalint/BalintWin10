VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmInvProcess 
   Caption         =   "Invoice Processing"
   ClientHeight    =   11565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvProcess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11565
   ScaleWidth      =   14610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cndNew2 
      Caption         =   "&NEW"
      Height          =   495
      Left            =   10080
      Picture         =   "frmInvProcess.frx":030A
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtApptTime 
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Text            =   "Appt Time"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdQBJobRefresh 
      Caption         =   "REFRESH CUSTOMERS FROM QB"
      Height          =   495
      Left            =   11400
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   720
      Width           =   2295
   End
   Begin TDBDate6Ctl.TDBDate tdbApptDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   4080
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   661
      Calendar        =   "frmInvProcess.frx":0614
      Caption         =   "frmInvProcess.frx":0714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":0778
      Keys            =   "frmInvProcess.frx":0796
      Spin            =   "frmInvProcess.frx":07F4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
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
      Text            =   "09/11/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40432
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdInvNow 
      Caption         =   "IN&V NOW"
      Height          =   735
      Left            =   12000
      Picture         =   "frmInvProcess.frx":081C
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin TDBText6Ctl.TDBText tdbSoldCity 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":0B26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":0B8A
      Key             =   "frmInvProcess.frx":0BA8
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
      Text            =   "TDBText2"
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
   Begin TDBText6Ctl.TDBText tdbSoldAddr1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":0BEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":0C50
      Key             =   "frmInvProcess.frx":0C6E
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
      Text            =   "TDBText1"
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
   Begin VB.CommandButton cmdFind 
      Caption         =   "&FIND"
      Height          =   735
      Left            =   6000
      Picture         =   "frmInvProcess.frx":0CB2
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&NEXT"
      Height          =   735
      Left            =   4800
      Picture         =   "frmInvProcess.frx":0FBC
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "PRE&V"
      Height          =   735
      Left            =   3600
      Picture         =   "frmInvProcess.frx":12C6
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&SEARCH"
      Height          =   735
      Left            =   7200
      Picture         =   "frmInvProcess.frx":15D0
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin TDBNumber6Ctl.TDBNumber tdbPkgCount 
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   4080
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":18DA
      Caption         =   "frmInvProcess.frx":18FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":196A
      Keys            =   "frmInvProcess.frx":1988
      Spin            =   "frmInvProcess.frx":19D2
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
   Begin VB.CommandButton cmdNew 
      Caption         =   "&NEW"
      Height          =   735
      Left            =   10800
      Picture         =   "frmInvProcess.frx":19FA
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   735
      Left            =   8400
      Picture         =   "frmInvProcess.frx":1D04
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelLine 
      Caption         =   "&DEL"
      Height          =   735
      Left            =   13560
      Picture         =   "frmInvProcess.frx":200E
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton cmdAddLine 
      Caption         =   "&ADD"
      Height          =   855
      Left            =   13560
      Picture         =   "frmInvProcess.frx":2318
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&CLEAR"
      Height          =   735
      Left            =   13560
      Picture         =   "frmInvProcess.frx":2622
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton cmdPriceLookup 
      Caption         =   "LOOK UP$"
      Height          =   975
      Left            =   13560
      Picture         =   "frmInvProcess.frx":292C
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7680
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid fgTrans 
      Height          =   1215
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   11175
      _cx             =   19711
      _cy             =   2143
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4575
      Left            =   120
      TabIndex        =   23
      Top             =   6000
      Width           =   13215
      _cx             =   23310
      _cy             =   8070
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin VB.ComboBox cmbTerms 
      Height          =   345
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin TDBText6Ctl.TDBText tdbtxtPO1 
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":2C36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":2C96
      Key             =   "frmInvProcess.frx":2CB4
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
      Text            =   "TDBText1"
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
   Begin TDBDate6Ctl.TDBDate tdbOrderDate 
      Height          =   375
      Left            =   11640
      TabIndex        =   24
      Top             =   2760
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calendar        =   "frmInvProcess.frx":2CF8
      Caption         =   "frmInvProcess.frx":2DF8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":2E62
      Keys            =   "frmInvProcess.frx":2E80
      Spin            =   "frmInvProcess.frx":2EDE
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
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
      Text            =   "08/09/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40399
      CenturyMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumInvNum 
      Height          =   375
      Left            =   11400
      TabIndex        =   29
      Top             =   1680
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":2F06
      Caption         =   "frmInvProcess.frx":2F26
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":2F94
      Keys            =   "frmInvProcess.frx":2FB2
      Spin            =   "frmInvProcess.frx":2FFC
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   735
      Left            =   9600
      Picture         =   "frmInvProcess.frx":3024
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   13200
      Picture         =   "frmInvProcess.frx":332E
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   10680
      Width           =   1095
   End
   Begin TrueOleDBList80.TDBCombo tdbcmbSoldTo 
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   635
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      _DropdownWidth  =   0
      _EDITHEIGHT     =   635
      _GAPHEIGHT      =   53
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AllowRowSizing=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      Appearance      =   1
      BorderStyle     =   1
      ComboStyle      =   0
      AutoCompletion  =   0   'False
      LimitToList     =   0   'False
      ColumnHeaders   =   -1  'True
      ColumnFooters   =   0   'False
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      Caption         =   ""
      EditFont        =   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      LayoutName      =   ""
      LayoutFileName  =   ""
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   0
      AutoSize        =   -1  'True
      ListField       =   ""
      BoundColumn     =   ""
      IntegralHeight  =   0   'False
      CellTipsWidth   =   0
      CellTipsDelay   =   1000
      AutoDropdown    =   0   'False
      RowTracking     =   -1  'True
      RightToLeft     =   0   'False
      RowMember       =   ""
      MouseIcon       =   0
      MouseIcon.vt    =   3
      MousePointer    =   0
      MatchEntryTimeout=   2000
      OLEDragMode     =   0
      OLEDropMode     =   0
      AnimateWindow   =   0
      AnimateWindowDirection=   0
      AnimateWindowTime=   200
      AnimateWindowClose=   0
      DropdownPosition=   0
      Locked          =   0   'False
      ScrollTrack     =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      AddItemSeparator=   ";"
      _PropDict       =   $"frmInvProcess.frx":3638
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Arial"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText tdbtxtPO2 
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   3600
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":36E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3742
      Key             =   "frmInvProcess.frx":3760
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
      Text            =   "TDBText1"
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
   Begin TDBNumber6Ctl.TDBNumber tdbFreight 
      Height          =   375
      Left            =   11400
      TabIndex        =   26
      Top             =   3720
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":37A4
      Caption         =   "frmInvProcess.frx":37C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3836
      Keys            =   "frmInvProcess.frx":3854
      Spin            =   "frmInvProcess.frx":389E
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
   Begin TDBText6Ctl.TDBText tdbSoldAddr2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":38C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":392A
      Key             =   "frmInvProcess.frx":3948
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbSoldAddr3 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":398C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":39F0
      Key             =   "frmInvProcess.frx":3A0E
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbSoldAddr4 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3A52
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3AB6
      Key             =   "frmInvProcess.frx":3AD4
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbSoldState 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3000
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3B18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3B7C
      Key             =   "frmInvProcess.frx":3B9A
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
      Text            =   "TDBText2"
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
   Begin TDBText6Ctl.TDBText tdbSoldZip 
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3BDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3C42
      Key             =   "frmInvProcess.frx":3C60
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
      Text            =   "TDBText2"
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
   Begin TDBText6Ctl.TDBText tdbShipAddr1 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   1560
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3CA4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3D08
      Key             =   "frmInvProcess.frx":3D26
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbShipAddr2 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1920
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3D6A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3DCE
      Key             =   "frmInvProcess.frx":3DEC
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbShipAddr3 
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2280
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3E30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3E94
      Key             =   "frmInvProcess.frx":3EB2
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbShipAddr4 
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   2640
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3EF6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":3F5A
      Key             =   "frmInvProcess.frx":3F78
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
      Text            =   "TDBText1"
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
   Begin TDBText6Ctl.TDBText tdbShipCity 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   3000
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":3FBC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":4020
      Key             =   "frmInvProcess.frx":403E
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
      Text            =   "TDBText2"
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
   Begin TDBText6Ctl.TDBText tdbShipState 
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   3000
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":4082
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":40E6
      Key             =   "frmInvProcess.frx":4104
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
      Text            =   "TDBText2"
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
   Begin TDBText6Ctl.TDBText tdbShipZip 
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   3000
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "frmInvProcess.frx":4148
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":41AC
      Key             =   "frmInvProcess.frx":41CA
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
      Text            =   "TDBText2"
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
   Begin TDBNumber6Ctl.TDBNumber tdbSalesTax 
      Height          =   375
      Left            =   11400
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":420E
      Caption         =   "frmInvProcess.frx":422E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":4296
      Keys            =   "frmInvProcess.frx":42B4
      Spin            =   "frmInvProcess.frx":42FE
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
   Begin TDBNumber6Ctl.TDBNumber tdbInvTotal 
      Height          =   375
      Left            =   11400
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4680
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":4326
      Caption         =   "frmInvProcess.frx":4346
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":43B6
      Keys            =   "frmInvProcess.frx":43D4
      Spin            =   "frmInvProcess.frx":441E
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
   Begin TDBNumber6Ctl.TDBNumber tdbItemTotal 
      Height          =   375
      Left            =   11400
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":4446
      Caption         =   "frmInvProcess.frx":4466
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":44D0
      Keys            =   "frmInvProcess.frx":44EE
      Spin            =   "frmInvProcess.frx":4538
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
   Begin TDBNumber6Ctl.TDBNumber tdbPalletCount 
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   4080
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmInvProcess.frx":4560
      Caption         =   "frmInvProcess.frx":4580
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvProcess.frx":45EE
      Keys            =   "frmInvProcess.frx":460C
      Spin            =   "frmInvProcess.frx":4656
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
   Begin VB.Label lblQBUpd 
      Caption         =   "lblQBUpd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   11640
      TabIndex        =   55
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label lblRev 
      Caption         =   "1.7"
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
      Left            =   13920
      TabIndex        =   54
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCustMsg2 
      Caption         =   "Cust Msg2"
      Height          =   255
      Left            =   120
      TabIndex        =   51
      Top             =   11040
      Width           =   3375
   End
   Begin VB.Label lblCustMsg1 
      Caption         =   "Cust Msg1"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   10680
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Appt Date/Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lblInvDate 
      Caption         =   "Invoice Date:"
      Height          =   255
      Left            =   11520
      TabIndex        =   48
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Terms:"
      Height          =   255
      Left            =   1320
      TabIndex        =   47
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Select Customer:"
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "SHIP TO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   45
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "SOLD TO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   1200
      Width           =   975
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
      Height          =   375
      Left            =   240
      TabIndex        =   43
      Top             =   120
      Width           =   13815
   End
End
Attribute VB_Name = "frmInvProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 2012-06-09
' create CalcSalesTaxPct as a sub
' call with tdbcmbSoldTo.LostFocus & FindInvoice
' have to get JCCustomer record first
' --> fixed where did not get JCCustomer info on invoice find

Dim xdbJob As New XArrayDB
Dim I, J, K As Long
Dim X, Y, Z As String
Dim txtString(6) As String
Dim LoadFlag As Boolean
Dim Flg As Boolean
Dim Price, Amount As Double
Dim dbl As Double
Dim boo As Boolean

Dim rsCol As New ADODB.Recordset
Dim rsTrans As New ADODB.Recordset
Dim rsStock As New ADODB.Recordset
Dim rsInvH As New ADODB.Recordset

Dim CommDrop, StockDrop As String
Dim TruckDrop, TrailerDrop, DriverDrop As String
Dim ColNum As Long
Dim NewInvoice As Boolean

' General QB variables
Dim requestMsgSet As IMsgSetRequest
Dim responseMsgSet As IMsgSetResponse
Dim ResponseList As IResponseList
Dim Response As IResponse
Dim ResponseType As Integer
Dim orItemRetList As IORItemRetList

' QB Item variables
Dim ItemQuery As IItemQuery
Dim orItemRet As IORItemRet
Dim itemServiceAdd As IItemServiceAdd
Dim itemServiceRet As IItemServiceRet

' QB Invoice Variables
Dim invoiceAdd As IInvoiceAdd
Dim orInvoiceLineAdd1 As IORInvoiceLineAdd
Dim orInvoiceLineAddORElement2 As String
Dim orRateORElement3 As String
Dim orRatePriceLevelORElement4 As String
Dim dataExt5 As IDataExt
Dim dataExt6 As IDataExt
Dim orDiscountLineAddORElement7 As String
Dim orSalesTaxLineAddORElement8 As String
Dim invoiceRet As IInvoiceRet

Dim SalesTaxPct As Double

Public InvDate As Date
Public OK As Boolean

' variables for screen comparison
Dim ScreenVals(29) As String
Dim fgTransVals(3, 5) As String
Dim fgVals(100, 10) As String

' variables for update to QB
Dim QBIDAR, QBIDTpl, QBIDFreight, QBIDMisc As String

Private Sub Form_Load()

    ' new fields ...
    If AddField("InvHeader", "ApptDate", "String", cn) Then
    End If
    If AddField("InvHeader", "ApptTime", "String", cn) Then
    End If
    
    ' ***********************************************
    ' get the QB update items
    
    ' get the QB AR account and template to use
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeQBSetup
    If InvGlobal.GetBySQL(SQLString) = False Then
        MsgBox "QB setup is not complete!" & vbCr & _
               "Go To Global Maintenance to complete the QB Setup", vbExclamation
        GoBack
    End If

'    If InvHeader.QBInvoiceID <> "" Then
'        MsgBox "This invoice has already been updated to QB!", vbExclamation
'        Exit Sub
'    End If
    
    If QBAccount.GetByID(NumValue(InvGlobal.Var1)) = False Then
        MsgBox "QB setup is not complete!  A/R Account" & vbCr & _
               "Go To Global Maintenance to complete the QB Setup", vbExclamation
        GoBack
    End If
    QBIDAR = QBAccount.QBID
    
    If QBAccount.GetByID(NumValue(InvGlobal.Var2)) = False Then
        MsgBox "QB setup is not complete!  Invoice Templage" & vbCr & _
               "Go To Global Maintenance to complete the QB Setup", vbExclamation
        GoBack
    End If
    QBIDTpl = QBAccount.QBID

    If InvStock.GetByID(NumValue(InvGlobal.Var3)) = False Then
        MsgBox "QB setup is not complete!  Freight Item" & vbCr & _
               "Go To Global Maintenance to complete the QB Setup", vbExclamation
        GoBack
    End If
    
    QBIDFreight = InvStock.QBID
    
    If InvStock.GetByID(NumValue(InvGlobal.Var4)) = False Then
        MsgBox "QB setup is not complete!  Misc Item" & vbCr & _
               "Go To Global Maintenance to complete the QB Setup", vbExclamation
        GoBack
    End If
    
    QBIDMisc = InvStock.QBID
    
    ' use sales tax?
    UseSalesTax = False
    If InvGlobal.Var5 = "1" Then UseSalesTax = True
    
'    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeSalesTax & _
'                " AND CompanyID = " & PRCompany.CompanyID
'    If InvGlobal.GetBySQL(SQLString) = True Then
'        If InvGlobal.Byte1 = 1 Then UseSalesTax = True
'    End If

    Init
    
    Me.KeyPreview = True
    
    Me.lblCompanyName = PRCompany.Name
    
    FormEnabled False
    
    'boo = FindInvoice(12)
    'cmdPrint_Click

End Sub

Private Sub cmdInvNow_Click()
    
    If Me.tdbcmbSoldTo.SelectedItem = 0 Then
        MsgBox "You must select a Customer:Job!", vbExclamation
        Exit Sub
    End If
    
    cmdSave_Click
    frmInvNow.Show vbModal
    If OK = False Then Exit Sub
    
    OK = False
    
    Me.lblInvDate = "Invoice Date: " & Format(InvDate, "mm/dd/yyyy")
    
    QBUpdate
    
    ' update to QB not complete
    If InvHeader.QBInvoiceID = "" Then
        MsgBox "Update of invoice to QB not complete!", vbExclamation
        Exit Sub
    End If
    
    LoadScreenVals          ' disable discard changes message
    cmdClear_Click          ' clear the screen
    FormEnabled False       ' disable screen controls
    LoadScreenVals          ' disable discard changes message

End Sub
Private Sub cndNew2_Click()
    cmdNew_Click
End Sub

Private Sub InvCalc()

Dim fgRWS, fgRW, fgCOL As Long
Dim InvTotal As Currency
Dim ThisQS, QS As Long

    ' calc subtotal of items
    QS = 0
    
    With fg
    
        If .Rows = 1 Then Exit Sub
    
        fgRW = .Row
        fgCOL = .Col
        InvTotal = 0
        
        For fgRWS = 1 To .Rows - 1
            
            ThisQS = NumValue(.TextMatrix(fgRWS, GetCol("QtyShipped")))
            QS = QS + ThisQS
            
            ' calc the extended amt
            If ThisQS > 0 Then
                .TextMatrix(fgRWS, GetCol("Amount")) = ThisQS * NumValue(.TextMatrix(fgRWS, GetCol("Price")))
            End If
            
            InvTotal = InvTotal + NumValue(.TextMatrix(fgRWS, GetCol("Amount")))
        
        Next fgRWS
        
        Me.tdbItemTotal.Value = InvTotal
        
        If SalesTaxPct = 0 Then
            Me.tdbSalesTax.Value = 0
        Else
            Me.tdbSalesTax.Value = Round(InvTotal * SalesTaxPct / 100, 2)
            InvTotal = InvTotal + Me.tdbSalesTax.Value
        End If
        
        InvTotal = InvTotal + Me.tdbFreight.Value
        
        Me.tdbInvTotal.Value = InvTotal
    
    End With

    Me.tdbPkgCount.Value = QS
    
    Me.Refresh

End Sub

Private Sub QBUpdate()

Dim Transpo(3) As Long
Dim TestMode As Boolean
Dim QBItemID As String
    
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    'MsgBox "QB Update ..."
    'Exit Sub
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    If UCase(User.Logon) = "JIM" Then
        TestMode = True
    Else
        TestMode = False
    End If
    
    Me.MousePointer = vbHourglass
    
    Me.lblCompanyName.Caption = "QB Open ..."
    Me.Refresh
    
    If QBOpen(Me, Me.lblCompanyName) = False Then
        Me.lblCompanyName = PRCompany.Name
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    Me.lblCompanyName.Caption = "Start cmdSave "
    Me.Refresh
    
    cmdSave_Click   ' save the current invoice
    
    Me.lblCompanyName.Caption = "Finish cmdSave "
    Me.Refresh
    
    ' start the new invoice
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    Set invoiceAdd = requestMsgSet.AppendInvoiceAddRq
    
    Me.lblCompanyName.Caption = "QB Update A ..."
    Me.Refresh
    
    If JCJob.GetByID(xdbJob.Value(Me.tdbcmbSoldTo.SelectedItem, 2)) = False Then
        Me.MousePointer = vbArrow
        MsgBox "QB Customer Find Error!", vbExclamation
        Me.lblCompanyName = PRCompany.Name
        Exit Sub
    End If
    
    Me.lblCompanyName.Caption = "QB Update B ..."
    Me.Refresh
    
    If JCJob.QBID = "ORIG" Then
        invoiceAdd.CustomerRef.ListID.SetValue JCJob.QBParentID
    Else
        invoiceAdd.CustomerRef.ListID.SetValue JCJob.QBID
    End If
    
    Me.lblCompanyName.Caption = "QB Update C ..."
    Me.Refresh
    
    invoiceAdd.ARAccountRef.ListID.SetValue Trim(QBIDAR)
    invoiceAdd.TemplateRef.ListID.SetValue Trim(QBIDTpl)
    
    Me.lblCompanyName.Caption = "QB Update D ..."
    Me.Refresh
    
    invoiceAdd.TxnDate.SetValue InvDate
    invoiceAdd.IsPending.SetValue False
    invoiceAdd.IsToBePrinted.SetValue False
    
    Me.lblCompanyName.Caption = "QB Update E ..."
    Me.Refresh
    
    invoiceAdd.BillAddress.Addr1.SetValue Trim(Me.tdbSoldAddr1)
    invoiceAdd.BillAddress.Addr2.SetValue Trim(Me.tdbSoldAddr2)
    invoiceAdd.BillAddress.Addr3.SetValue Trim(Me.tdbSoldAddr3)
    invoiceAdd.BillAddress.Addr4.SetValue Trim(Me.tdbSoldAddr4)
    invoiceAdd.BillAddress.City.SetValue Trim(Me.tdbSoldCity)
    invoiceAdd.BillAddress.State.SetValue Trim(Me.tdbSoldState)
    invoiceAdd.BillAddress.PostalCode.SetValue Trim(Me.tdbSoldZip)
    
    Me.lblCompanyName.Caption = "QB Update F ..."
    Me.Refresh
    
    invoiceAdd.ShipAddress.Addr1.SetValue Trim(Me.tdbShipAddr1)
    invoiceAdd.ShipAddress.Addr2.SetValue Trim(Me.tdbShipAddr2)
    invoiceAdd.ShipAddress.Addr3.SetValue Trim(Me.tdbShipAddr3)
    invoiceAdd.ShipAddress.Addr4.SetValue Trim(Me.tdbShipAddr4)
    invoiceAdd.ShipAddress.City.SetValue Trim(Me.tdbShipCity)
    invoiceAdd.ShipAddress.State.SetValue Trim(Me.tdbShipState)
    invoiceAdd.ShipAddress.PostalCode.SetValue Trim(Me.tdbShipZip)
    
    Me.lblCompanyName.Caption = "QB Update G ..."
    Me.Refresh
    
    invoiceAdd.PONumber.SetValue Trim(Me.tdbtxtPO1.Text)
    
    Me.lblCompanyName.Caption = "QB Update G2 ..."
    Me.Refresh
    
    invoiceAdd.RefNumber.SetValue Me.tdbnumInvNum.Value
    
    Me.lblCompanyName.Caption = "QB Update H ..."
    Me.Refresh
    
    With Me.cmbTerms
        If .ListIndex <> -1 Then
            If InvGlobal.GetByID(.ItemData(.ListIndex)) = True Then
                invoiceAdd.TermsRef.ListID.SetValue Trim(InvGlobal.Var1)
            End If
        End If
    End With
    
    Me.lblCompanyName.Caption = "QB Update I ..."
    Me.Refresh
    
    ' get that tax code / item from the customer record
    If UseSalesTax = True Then
    
        If JCCustomer.GetByID(JCJob.ParentID) = False Then
            MsgBox "Customer record not found for Job #: " & JCJob.JobID, vbExclamation
            GoBack
        End If
            
        If JCCustomer.QBTaxItem = "" Then
            MsgBox "QB Tax Item Not Set: " & JCCustomer.Name, vbExclamation
            GoBack
        End If
            
        If JCCustomer.QBTaxCode = "" Then
            MsgBox "QB Tax Code Not Set: " & JCCustomer.Name, vbExclamation
            GoBack
        End If
        
        Me.lblCompanyName.Caption = "QB Update J ..."
        Me.Refresh
    
        invoiceAdd.ItemSalesTaxRef.ListID.SetValue Trim(JCCustomer.QBTaxItem)
        invoiceAdd.CustomerSalesTaxCodeRef.ListID.SetValue Trim(JCCustomer.QBTaxCode)
    
    End If

    ' -------------------------------------------------------------------------------
    ' add QB invoice body lines for KP extra info
    
    Me.lblCompanyName.Caption = "QB Update K ..."
    Me.Refresh
    
    ' second PO ?
    If Trim(InvHeader.PO2) <> "" Then AddDescLine "PO 2:" & InvHeader.PO2

    ' transportation info
    For I = 1 To 3
        
        If I = 1 Then
            Transpo(1) = InvHeader.TruckID1
            Transpo(2) = InvHeader.TrailerID1
            Transpo(3) = InvHeader.DriverID1
        End If
        If I = 2 Then
            Transpo(1) = InvHeader.TruckID2
            Transpo(2) = InvHeader.TrailerID2
            Transpo(3) = InvHeader.DriverID2
        End If
        If I = 3 Then
            Transpo(1) = InvHeader.TruckID3
            Transpo(2) = InvHeader.TrailerID3
            Transpo(3) = InvHeader.DriverID3
        End If
    
        X = ""
        For J = 1 To 3
            
            If Transpo(J) <> 0 Then
                If InvGlobal.GetByID(Transpo(J)) = True Then
                    If X = "" Then
                        X = InvGlobal.Description
                    Else
                        X = X & "/" & InvGlobal.Description
                    End If
                End If
            End If
        
        Next J
    
        Me.lblCompanyName.Caption = "QB Update Trans ... " & I
        Me.Refresh
        
        If X <> "" Then AddDescLine X
    
    Next I
    
    ' -------------------------------------------------------------------------------
        
    AddDescLine ""
        
    ' add the lines for the invoice body
    SQLString = "SELECT * FROM InvBody WHERE HeaderID = " & InvHeader.HeaderID & _
                " ORDER BY LineNum"
    If InvBody.GetBySQL(SQLString) = True Then
        
        Do
            
            Set orInvoiceLineAdd1 = invoiceAdd.ORInvoiceLineAddList.Append
            
            X = ""
            QBItemID = QBIDMisc
            If InvBody.StockID <> 0 Then
                
                SQLString = "SELECT * FROM InvStock WHERE JobID = " & JCJob.JobID & _
                            " AND StockID = " & InvBody.StockID
                
                ' 2/26/2011
                ' prices are per CUSTOMER not JOB
                SQLString = "SELECT * FROM InvStock WHERE JobID = " & JCJob.ParentID & _
                            " AND StockID = " & InvBody.StockID
                
                If InvStock.GetBySQL(SQLString) = True Then
                    QBItemID = InvStock.QBID
                End If
                
                ' update stock file fields
                InvStock.CustomerPrice = InvBody.Price
                InvStock.LastDate = InvHeader.InvoiceDate
                InvStock.rsPut
            
                
            Else        ' use misc item id?
            
                If InvBody.QtyOrdered <> 0 Or _
                   InvBody.QtyShipped <> 0 Or _
                   InvBody.Amount <> 0 Then
                    
                    QBItemID = QBIDMisc
                    If InvBody.QtyShipped = 0 And InvBody.Amount <> 0 Then
                        InvBody.QtyShipped = 1
                        InvBody.Price = InvBody.Amount
                    End If
                   
                End If
            
            End If
            
            If TestMode Then
                MsgBox InvBody.Description & "/" & InvBody.StockID & "/" & InvStock.QBID & "/" & X
            End If
            
            Me.lblCompanyName.Caption = "QB Update j ..."
            Me.Refresh
    
            orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.ListID.SetValue Trim(QBItemID)
            orInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue InvBody.QtyShipped
            orInvoiceLineAdd1.InvoiceLineAdd.Desc.SetValue Trim(InvBody.Description)
            orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue InvBody.Price
            orInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue Round(InvBody.Amount, 2)
            
            Me.lblCompanyName.Caption = "QB Update k ..."
            Me.Refresh
    
            If UseSalesTax = True Then
                orInvoiceLineAdd1.InvoiceLineAdd.SalesTaxCodeRef.ListID.SetValue Trim(JCCustomer.QBTaxCode)
                orInvoiceLineAdd1.InvoiceLineAdd.IsTaxable.SetValue True
            End If
            
            If InvBody.GetNext = False Then Exit Do
        
        Loop
    
    End If

    Me.lblCompanyName.Caption = "QB Update l ..."
    Me.Refresh
    
    ' add the freight amount
    If Me.tdbFreight.Value <> 0 Then
        AddDescLine ""
        Set orInvoiceLineAdd1 = invoiceAdd.ORInvoiceLineAddList.Append
        orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.ListID.SetValue Trim(QBIDFreight)
        orInvoiceLineAdd1.InvoiceLineAdd.Desc.SetValue "Freight"
        orInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue Me.tdbFreight.Value
    End If

    Me.lblCompanyName.Caption = "QB Update m ..."
    Me.Refresh
    
    ' customer auto comments
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeInvMessage & _
                " AND CompanyID = " & PRCompany.CompanyID & _
                " AND UserID = " & InvHeader.SoldJobID
    If InvGlobal.GetBySQL(SQLString) = True Then
        For I = 1 To 5
            
            If I = 1 Then X = InvGlobal.Var1
            If I = 2 Then X = InvGlobal.Var2
            If I = 3 Then X = InvGlobal.Var3
            If I = 4 Then X = InvGlobal.Var4
            If I = 5 Then X = InvGlobal.Var5
            If X <> "" Then
                
                Me.lblCompanyName.Caption = "QB Update n ..." & X
                Me.Refresh
    
                AddDescLine X
            End If
        Next I
    End If
    
    Me.lblCompanyName.Caption = "QB Update o ..."
    Me.Refresh
    
    ' appt date / time
    If InvHeader.TruckID1 <> 0 And InvHeader.ApptDate <> 0 Then
        
        X = "APPOINTMENT SCHEDULED FOR:"
        AddDescLine X
        
        X = InvHeader.ApptTime
        X = X & Format(InvHeader.ApptDate, " dddd mm/dd/yyyy")
        AddDescLine X
    
    End If

    Me.lblCompanyName.Caption = "QB Update p ..."
    Me.Refresh
    
    ' package / pallet count
    If InvHeader.PackageCount <> 0 Then
        AddDescLine "TOTAL NUMBER OF PACKAGES: " & InvHeader.PackageCount
    End If
    If InvHeader.PalletCount <> 0 Then
        AddDescLine "TOTAL NUMBER OF PALLETS:  " & InvHeader.PalletCount
    End If
    
    Me.lblCompanyName.Caption = "QB Update q ..."
    Me.Refresh
    
    ' process the QB request
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    Set ResponseList = responseMsgSet.ResponseList
    orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.ListID.SetValue Trim(QBIDFreight)

    If (ResponseList Is Nothing) Then
        AddNote ("Invoice Add - response is nothing!")
        MsgBox "Invoice Add - response is nothing!", vbExclamation
        Me.MousePointer = vbArrow
        GoBack
    End If
    
    Me.lblCompanyName.Caption = "QB Update r ..."
    Me.Refresh
    
    For I = 0 To ResponseList.Count - 1
        
        Set Response = ResponseList.GetAt(I)
        
        ' Check the status returned for the response.
        If Response.StatusCode >= 1000 Then
            MsgBox Response.StatusCode & vbCr & _
                   Response.StatusMessage, vbExclamation
            ' GoTo InvParseNxtI
            AddNote (Response.StatusCode & " " & Response.StatusMessage)
            GoBack
        End If
        
        If TestMode Then MsgBox "response list b " & I
        
        If (Response.Detail Is Nothing) Then GoTo InvParseNxtI
        
        If TestMode Then MsgBox "response list c " & I
        
        ResponseType = Response.Type.GetValue
        
        If TestMode Then MsgBox "response list d " & I
        
        If ResponseType <> rtInvoiceAddRs Then GoTo InvParseNxtI
        
        If TestMode Then MsgBox "response list e " & I
        
        Set invoiceRet = Response.Detail
        
        If TestMode Then MsgBox "response list f " & I
        
        If invoiceRet Is Nothing Then
            MsgBox "invoiceRet is nothing ", vbExclamation
            Me.MousePointer = vbArrow
            AddNote ("invoiceRet is nothing ")
            GoBack
        End If
        
InvParseNxtI:
    Next I
    
    ' QB updates are done ... OK to set the invoice record
    frmInvProcess.OK = True
    
    ' only update if save to QB is OK
    InvHeader.QBInvoiceID = invoiceRet.TxnID.GetValue
    InvHeader.rsPut
        
    Me.MousePointer = vbArrow
    Me.lblCompanyName = PRCompany.Name
    
    SessMgr.EndSession
    SessMgr.CloseConnection

End Sub

Private Sub AddNote(ByVal ErrMessage)
    
    Notes.OpenRS
    Notes.Clear
    Notes.NoteType = 99
    Notes.NoteCat = 99
    Notes.RelatedID = InvHeader.HeaderID
    Notes.Subject = "Error"
    Notes.User = "SYSTEM"
    Notes.DateTm = Now()
    Notes.Notation = ErrMessage
    Notes.Save (Equate.RecAdd)

End Sub

Private Sub AddDescLine(ByVal Str As String)

    Set orInvoiceLineAdd1 = invoiceAdd.ORInvoiceLineAddList.Append
    orInvoiceLineAdd1.InvoiceLineAdd.Desc.SetValue Trim(Str)
    'orInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue 0
     'orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue 0
    'orInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue 0
    'orInvoiceLineAdd1.InvoiceLineAdd.SalesTaxCodeRef.ListID.SetValue JCCustomer.QBTaxCode
    'orInvoiceLineAdd1.InvoiceLineAdd.IsTaxable.SetValue True

End Sub

Private Sub cmdSearch_Click()
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    frmInvFind.Show vbModal
    If frmInvFind.InvNum = 0 Then
        ' reget the data .......
        If NumValue(CStr(Me.tdbnumInvNum.Value & "")) = 0 Then Exit Sub
        SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber = " & Me.tdbnumInvNum.Value
        FindInvoice Me.tdbnumInvNum.Value
        Exit Sub
    End If
    FindInvoice frmInvFind.InvNum

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    GoBack

End Sub

Private Sub cmdFind_Click()
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    X = InputBox("Enter Invoice Number to find:", "Search by Invoice #")
    If X = "" Then Exit Sub
    If NumValue(X) = 0 Then Exit Sub
    If FindInvoice(CLng(X)) = False Then Exit Sub

End Sub

Private Sub cmdNew_Click()
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    cmdClear_Click
    
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeInvNumber & _
                " AND CompanyID = " & PRCompany.CompanyID
    If InvGlobal.GetBySQL(SQLString) = False Then
        InvGlobal.Clear
        InvGlobal.TypeCode = InvEquate.GlobalTypeInvNumber
        InvGlobal.CompanyID = PRCompany.CompanyID
        InvGlobal.UserID = 0
        InvGlobal.rsAdd
    End If
    
    Me.lblInvDate = "Invoice Date: "
    Me.lblQBUpd = ""
    InvDate = 0
    
    InvGlobal.UserID = InvGlobal.UserID + 1
    InvGlobal.rsPut
    Me.tdbnumInvNum.Value = InvGlobal.UserID
    
    InvHeader.InvoiceNumber = Me.tdbnumInvNum.Text
    
    FormEnabled True

    LoadScreenVals

    Me.tdbcmbSoldTo.SetFocus

End Sub

Private Function FindInvoice(ByVal InvNum As Long) As Boolean

    ' *******************************************************
    
    SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber = " & InvNum
    If InvHeader.GetBySQL(SQLString) = False Then
        MsgBox "Invoice # " & InvNum & " not found!", vbExclamation
        FindInvoice = False
        Exit Function
    End If
    
    FindInvoice = True
        
    With Me
        
        I = xdbJob.Find(0, 2, InvHeader.SoldJobID, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
        If I < 0 Then
            MsgBox "Customer Job not found: " & InvHeader.SoldJobID, vbExclamation
            GoBack
        End If
        
        boo = JCJob.GetByID(InvHeader.SoldJobID)
        
        ' **** added 2012-06-09 ******************************
        boo = JCCustomer.GetByID(JCJob.ParentID)
        CalcSalesTaxPct
        ' **** added 2012-06-09 ******************************
        
        .tdbcmbSoldTo.SelectedItem = I
        
        .tdbSoldAddr1 = InvHeader.SoldAddr1
        .tdbSoldAddr2 = InvHeader.SoldAddr2
        .tdbSoldAddr3 = InvHeader.SoldAddr3
        .tdbSoldAddr4 = InvHeader.SoldAddr4
        .tdbSoldCity = InvHeader.SoldCity
        .tdbSoldState = InvHeader.SoldState
        .tdbSoldZip = InvHeader.SoldZip
        
        .tdbShipAddr1 = InvHeader.ShipAddr1
        .tdbShipAddr2 = InvHeader.ShipAddr2
        .tdbShipAddr3 = InvHeader.ShipAddr3
        .tdbShipAddr4 = InvHeader.ShipAddr4
        .tdbShipCity = InvHeader.ShipCity
        .tdbShipState = InvHeader.ShipState
        .tdbShipZip = InvHeader.ShipZip
        
        .tdbnumInvNum = InvHeader.InvoiceNumber
        .tdbOrderDate.Value = InvHeader.OrderDate
        
        If InvHeader.InvoiceDate <> 0 Then
            .lblInvDate.Caption = "Invoice Date: " & Format(InvHeader.InvoiceDate, "mm/dd/yyyy")
            InvDate = InvHeader.InvoiceDate
        Else
            .lblInvDate.Caption = "Invoice Date: "
        End If
        
        If InvHeader.QBInvoiceID = "" Then
            Me.lblQBUpd = ""
        Else
            Me.lblQBUpd = "Updated to QB"
        End If
        
        .tdbtxtPO1 = InvHeader.PO1
        .tdbtxtPO2 = InvHeader.PO2
        .tdbFreight = InvHeader.Freight
        .tdbPkgCount = InvHeader.PackageCount
        .tdbPalletCount = InvHeader.PalletCount
        
        .tdbApptDate = InvHeader.ApptDate
        .txtApptTime = InvHeader.ApptTime
        
        SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeTerms & _
                    " AND Var1 = '" & InvHeader.Terms & "'"
        If InvGlobal.GetBySQL(SQLString) = True Then
            cmbPoint .cmbTerms, InvGlobal.GlobalID
        Else
            .cmbTerms.ListIndex = -1
        End If
        
        TransFill 1, InvHeader.TruckID1, InvHeader.TrailerID1, _
                  InvHeader.DriverID1
        TransFill 2, InvHeader.TruckID2, InvHeader.TrailerID2, _
                  InvHeader.DriverID2
        TransFill 3, InvHeader.TruckID3, InvHeader.TrailerID3, _
                  InvHeader.DriverID3
        rsTrans.MoveFirst
                  
        LoadStock

        ' remove all existing rows
        Do
            If fg.Rows = 1 Then Exit Do
            fg.RemoveItem 1
        Loop
        
        SQLString = "SELECT * FROM InvBody WHERE HeaderID = " & InvHeader.HeaderID & _
                    " ORDER BY LineNum"
        If InvBody.GetBySQL(SQLString) = True Then
            I = 0
            Do
                I = I + 1
                If fg.Rows < I + 1 Then
                    X = ""
                    fg.AddItem X, I
                End If
                fg.TextMatrix(I, GetCol("StockID")) = InvBody.StockID
                fg.TextMatrix(I, GetCol("QtyOrdered")) = InvBody.QtyOrdered
                fg.TextMatrix(I, GetCol("QtyShipped")) = InvBody.QtyShipped
                fg.TextMatrix(I, GetCol("Description")) = InvBody.Description
                fg.TextMatrix(I, GetCol("Price")) = InvBody.Price
                fg.TextMatrix(I, GetCol("Amount")) = InvBody.Amount
                If InvBody.GetNext = False Then Exit Do
            Loop
        
        End If
    
        ' add a few extra rows
        For I = 1 To 5
            X = ""
            For J = 1 To rsCol.RecordCount
                X = X & ""
                If J <> rsCol.RecordCount Then X = X & vbTab
            Next J
            fg.AddItem X
        Next I
    
        ' already invoiced ? restrict edits
        ' *** use QBInvoiceID ***
        ' If InvHeader.InvoiceDate <> 0 Then
        If InvHeader.QBInvoiceID <> "" Then
            FormEnabled False
        Else
            FormEnabled True
        End If
            
        CustomerComments
    
    End With
    
    InvCalc
    LoadScreenVals
    
End Function

Private Sub CustomerComments()
        
    ' customer special message display
    With Me
        SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeInvMessage & _
                    " AND CompanyID = " & PRCompany.CompanyID & _
                    " AND UserID = " & JCJob.JobID
        If InvGlobal.GetBySQL(SQLString) = True Then
            J = 0
            .lblCustMsg1.Caption = ""
            .lblCustMsg2.Caption = ""
            For I = 1 To 5
                If I = 1 Then X = InvGlobal.Var1
                If I = 2 Then X = InvGlobal.Var2
                If I = 3 Then X = InvGlobal.Var3
                If I = 4 Then X = InvGlobal.Var4
                If I = 5 Then X = InvGlobal.Var5
                If X <> "" Then
                    J = J + 1
                    If J = 1 Then .lblCustMsg1.Caption = X
                    If J = 2 Then .lblCustMsg2.Caption = X
                    If J = 2 Then Exit For
                End If
            Next I
        End If
    End With

End Sub

Private Sub FormEnabled(ByVal fBoo As Boolean)

    With Me
        
        .tdbnumInvNum.Enabled = fBoo
        
        .tdbcmbSoldTo.Enabled = fBoo
        
        .tdbSoldAddr1.Enabled = fBoo
        .tdbSoldAddr2.Enabled = fBoo
        .tdbSoldAddr3.Enabled = fBoo
        .tdbSoldAddr4.Enabled = fBoo
        .tdbSoldCity.Enabled = fBoo
        .tdbSoldState.Enabled = fBoo
        .tdbSoldZip.Enabled = fBoo
        
        .tdbShipAddr1.Enabled = fBoo
        .tdbShipAddr2.Enabled = fBoo
        .tdbShipAddr3.Enabled = fBoo
        .tdbShipAddr4.Enabled = fBoo
        .tdbShipCity.Enabled = fBoo
        .tdbShipState.Enabled = fBoo
        .tdbShipZip.Enabled = fBoo

        .cmbTerms.Enabled = fBoo
        .tdbOrderDate.Enabled = fBoo
        .tdbFreight.Enabled = fBoo
        
        .tdbtxtPO1.Enabled = fBoo
        .tdbtxtPO2.Enabled = fBoo
        
        If fBoo = False Then
            .fgTrans.Editable = flexEDNone
            .fg.Editable = flexEDNone
        Else
            .fgTrans.Editable = flexEDKbdMouse
            .fg.Editable = flexEDKbdMouse
        End If
        
        .tdbPkgCount.Enabled = fBoo
        
        .cmdSave.Enabled = fBoo
        .cmdInvNow.Enabled = fBoo
        .cmdAddLine.Enabled = fBoo
        .cmdDelLine.Enabled = fBoo
        .cmdPriceLookup.Enabled = fBoo

        .tdbApptDate.Enabled = fBoo
        .txtApptTime.Enabled = fBoo
    
    End With

End Sub

Private Sub TransFill(ByVal Count As Byte, _
                        ByVal TruckID As Long, _
                        ByVal TrailerID As Long, _
                        ByVal DriverID As Long)

    rsTrans.Find "Count = " & Count, 0, adSearchForward, 1
    If rsTrans.EOF Then
        rsTrans.AddNew
        rsTrans!Count = Count
    End If
    rsTrans!TruckID = TruckID
    rsTrans!TrailerID = TrailerID
    rsTrans!DriverID = DriverID
    rsTrans.Update

End Sub

Private Sub cmdPrev_Click()
    
Dim InvNo As Long
    
    If NumValue(CStr(Me.tdbnumInvNum.Value & "")) = 0 Then
        InvNo = 999999999
    Else
        InvNo = NumValue(CStr(Me.tdbnumInvNum.Value & ""))
    End If
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    I = Me.tdbnumInvNum.Value
    SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber < " & InvNo & _
                " ORDER BY InvoiceNumber DESC"
    If InvHeader.GetBySQL(SQLString) = False Then
        MsgBox "No Previous Invoice Number Exists", vbInformation
    Else
        I = InvHeader.InvoiceNumber
    End If
    FindInvoice I
End Sub
Private Sub cmdNext_Click()
    
Dim InvNo As Long
    
    If NumValue(CStr(Me.tdbnumInvNum.Value & "")) = 0 Then
        InvNo = 1
    Else
        InvNo = NumValue(CStr(Me.tdbnumInvNum.Value & ""))
    End If
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    I = Me.tdbnumInvNum.Value
    SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber > " & InvNo & _
                " ORDER BY InvoiceNumber"
    If InvHeader.GetBySQL(SQLString) = False Then
        MsgBox "No Next Invoice Number Exists", vbInformation
    Else
        I = InvHeader.InvoiceNumber
    End If
    FindInvoice I
End Sub

Private Sub cmdSave_Click()

    ' ***
    NewInvoice = True

    If IsNull(Me.tdbcmbSoldTo.SelectedItem) Then Exit Sub
    
    If Me.tdbcmbSoldTo.SelectedItem = 0 Then
        MsgBox "You must select a Customer:Job!", vbExclamation
        Exit Sub
    End If
    
    If NumValue(Me.tdbnumInvNum & "") = 0 Then
        MsgBox "Invoice Number not assigned!", vbExclamation
        Exit Sub
    End If
    
    InvCalc
    
    ' re-set the variables tracking changes made
    LoadScreenVals
    
    ' save header info
        
    ' see if inv# already exists
    SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber = " & Me.tdbnumInvNum.Value
    If InvHeader.GetBySQL(SQLString) = True Then
    Else
        InvHeader.Clear
        InvHeader.InvoiceNumber = Me.tdbnumInvNum.Value
        InvHeader.rsAdd
    
        ' update for the next invoice number to prompt for
        SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeInvNumber & _
                    " AND CompanyID = " & PRCompany.CompanyID
        If InvGlobal.GetBySQL(SQLString) = False Then
            InvGlobal.Clear
            InvGlobal.TypeCode = InvEquate.GlobalTypeInvNumber
            InvGlobal.CompanyID = PRCompany.CompanyID
            InvGlobal.rsAdd
        End If
        InvGlobal.UserID = Me.tdbnumInvNum.Value
        InvGlobal.rsPut
    
    End If
        
    ' create new InvHeader record
    With Me
        
        InvHeader.InvoiceNumber = .tdbnumInvNum.Value
        
        InvHeader.SoldAddr1 = .tdbSoldAddr1
        InvHeader.SoldAddr2 = .tdbSoldAddr2
        InvHeader.SoldAddr3 = .tdbSoldAddr3
        InvHeader.SoldAddr4 = .tdbSoldAddr4
        InvHeader.SoldCity = .tdbSoldCity
        InvHeader.SoldState = .tdbSoldState
        InvHeader.SoldZip = .tdbSoldZip
        
        InvHeader.ShipAddr1 = .tdbShipAddr1
        InvHeader.ShipAddr2 = .tdbShipAddr2
        InvHeader.ShipAddr3 = .tdbShipAddr3
        InvHeader.ShipAddr4 = .tdbShipAddr4
        InvHeader.ShipCity = .tdbShipCity
        InvHeader.ShipState = .tdbShipState
        InvHeader.ShipZip = .tdbShipZip
        
        InvHeader.SoldJobID = JCJob.JobID
        InvHeader.InvoiceDate = InvDate
        
        If IsNull(.tdbOrderDate.Value) = False Then
            InvHeader.OrderDate = .tdbOrderDate.Value
        Else
            InvHeader.OrderDate = 0
        End If
        
        InvHeader.PO1 = .tdbtxtPO1.Text
        InvHeader.PO2 = .tdbtxtPO2.Text
        
        ' trans info
        I = 0
        rsTrans.MoveFirst
        Do
            I = I + 1
            If rsTrans!TruckID <> 0 Then
                If I = 1 Then
                    InvHeader.TruckID1 = rsTrans!TruckID
                    InvHeader.TrailerID1 = rsTrans!TrailerID
                    InvHeader.DriverID1 = rsTrans!DriverID
                ElseIf I = 2 Then
                    InvHeader.TruckID2 = rsTrans!TruckID
                    InvHeader.TrailerID2 = rsTrans!TrailerID
                    InvHeader.DriverID2 = rsTrans!DriverID
                Else
                    InvHeader.TruckID3 = rsTrans!TruckID
                    InvHeader.TrailerID3 = rsTrans!TrailerID
                    InvHeader.DriverID3 = rsTrans!DriverID
                End If
            Else
                If I = 1 Then
                    InvHeader.TruckID1 = 0
                    InvHeader.TrailerID1 = 0
                    InvHeader.DriverID1 = 0
                ElseIf I = 2 Then
                    InvHeader.TruckID2 = 0
                    InvHeader.TrailerID2 = 0
                    InvHeader.DriverID2 = 0
                Else
                    InvHeader.TruckID3 = 0
                    InvHeader.TrailerID3 = 0
                    InvHeader.DriverID3 = 0
                End If
            End If
            rsTrans.MoveNext
        Loop Until rsTrans.EOF
        rsTrans.MoveFirst
    
        ' terms
        If .cmbTerms.ListIndex <> -1 Then
            boo = InvGlobal.GetByID(.cmbTerms.ItemData(.cmbTerms.ListIndex))
            InvHeader.Terms = InvGlobal.Var1
        End If
        
        InvHeader.ItemTotal = 0
        InvHeader.SalesTax = 0
        InvHeader.Freight = 0
        InvHeader.PackageCount = 0
        InvHeader.TotalAmount = 0
        
        ' time as a string
        InvHeader.ApptDate = .tdbApptDate.Value
        InvHeader.ApptTime = Mid(.txtApptTime, 1, 10)
        
        If IsNull(.tdbItemTotal) = False Then InvHeader.ItemTotal = .tdbItemTotal.Value
        If IsNull(.tdbSalesTax) = False Then InvHeader.SalesTax = .tdbSalesTax.Value
        If IsNull(.tdbFreight) = False Then InvHeader.Freight = .tdbFreight.Value
        If IsNull(.tdbInvTotal) = False Then InvHeader.TotalAmount = .tdbInvTotal
        If IsNull(.tdbPkgCount) = False Then InvHeader.PackageCount = .tdbPkgCount.Value
        If IsNull(.tdbPalletCount) = False Then InvHeader.PalletCount = .tdbPalletCount.Value
         
        InvHeader.SaveFlag = 1
        
        InvHeader.rsPut
    
    End With

    ' save the body info
    Dim LastRow As Long
    With fg
        
        ' find the last row used
        For I = .Rows - 1 To 0 Step -1
            Flg = False
            For J = 0 To .Cols - 1
                If .TextMatrix(I, J) <> "" Then
                    Flg = True
                    Exit For
                End If
            Next J
            If Flg Then
                LastRow = I
                Exit For
            End If
        Next I
        If Flg = True Then
            For I = 1 To LastRow
                SQLString = "SELECT * FROM InvBody WHERE HeaderID = " & InvHeader.HeaderID & _
                            " AND LineNum = " & I
                If InvBody.GetBySQL(SQLString) = False Then
                    InvBody.Clear
                    InvBody.HeaderID = InvHeader.HeaderID
                    InvBody.LineNum = I
                    InvBody.rsAdd
                End If
                
                InvBody.QtyOrdered = NumValue(.TextMatrix(I, GetCol("QtyOrdered")))
                InvBody.QtyShipped = NumValue(.TextMatrix(I, GetCol("QtyShipped")))
                InvBody.Description = .Cell(flexcpTextDisplay, I, GetCol("Description"))
                InvBody.StockID = NumValue(.TextMatrix(I, GetCol("StockID")))
                InvBody.Price = NumValue(.TextMatrix(I, GetCol("Price")))
                InvBody.Amount = NumValue(.TextMatrix(I, GetCol("Amount")))
                InvBody.rsPut
            
            Next I
        
        End If
    
        ' delete trailing records if necessary
        SQLString = "DELETE * FROM InvBody WHERE HeaderID = " & InvHeader.HeaderID & _
                    " AND LineNum > " & LastRow
        cn.Execute SQLString
    
    End With

End Sub

Private Sub cmdPriceLookup_Click()

    If IsNull(Me.tdbcmbSoldTo.SelectedItem) Then Exit Sub
    
    With frmInvPriceLookup
        
        ' prices are per CUSTOMER
        ' .JobID = JCJob.JobID
        .JobID = JCJob.ParentID
        
        .Init
        .Show vbModal
        If .OK = False Then Exit Sub
        
        I = Me.fg.Row
        If I = 0 Then I = 1
        With .rs
            .MoveFirst
            Do
                If !StockID1 <> 0 And !Quantity1 > 0 Then
                    X = " " & vbTab & !StockID1 & vbTab & !Quantity1 & vbTab & !Quantity1 & vbTab & _
                        !Description1 & vbTab & !Price1
                    fg.AddItem X, I
                    fg.TextMatrix(I, GetCol("Amount")) = Round(!Quantity1 * !Price1, 2)
                    I = I + 1
                End If
                If !StockID2 <> 0 And !Quantity2 > 0 Then
                    X = " " & vbTab & !StockID2 & vbTab & !Quantity2 & vbTab & !Quantity2 & vbTab & _
                        !Description2 & vbTab & !Price2
                    fg.AddItem X, I
                    fg.TextMatrix(I, GetCol("Amount")) = Round(!Quantity2 * !Price2, 2)
                    I = I + 1
                End If
                .MoveNext
            Loop Until .EOF
        End With
    
    End With

    InvCalc

End Sub

Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If LoadFlag = True Then Exit Sub
    If OldRow = 0 Then Exit Sub
    If OldRow > fg.Rows - 1 Then Exit Sub
    
    ' last cell....
    If OldRow = 3 And OldCol = 3 Then
        Me.fg.SetFocus
        Exit Sub
    End If
    
    With fg
        
        ' *** add a new comment ***
        If OldCol = GetCol("Description") And .TextMatrix(OldRow, OldCol) = "999999" Then
            X = InputBox("Enter New Invoice Comment")
            If X = "" Then
                .TextMatrix(OldRow, OldCol) = ""
                Exit Sub
            End If
            InvGlobal.Clear
            InvGlobal.TypeCode = InvEquate.GlobalTypeComment
            InvGlobal.Description = X
            InvGlobal.rsAdd
            CommDrop = CommDrop & "|#" & InvGlobal.GlobalID & ";" & X
            .TextMatrix(OldRow, OldCol) = InvGlobal.GlobalID
            .ColComboList(GetCol("Description")) = CommDrop
        End If
    
        ' fill in stock info
        If OldCol = GetCol("StockID") And .TextMatrix(OldRow, OldCol) <> "" Then
            
            If InvStock.GetByID(CLng(NumValue(.TextMatrix(OldRow, OldCol)))) = False Then Exit Sub
            .TextMatrix(OldRow, GetCol("Description")) = InvStock.Description
            .TextMatrix(OldRow, GetCol("Price")) = InvStock.CustomerPrice
        
        End If
    
        ' copy QtyOrdered to QtyShipped
        ' If OldCol = GetCol("QtyOrdered") And .TextMatrix(OldRow, GetCol("QtyShipped")) <> "" Then
        If OldCol = GetCol("QtyOrdered") Then
            .TextMatrix(OldRow, GetCol("QtyShipped")) = .TextMatrix(OldRow, GetCol("QtyOrdered"))
        End If
    
        ' calc total amount
        ' ??????
        If .TextMatrix(OldRow, GetCol("StockID")) <> "" Then
            .TextMatrix(OldRow, GetCol("Amount")) = NumValue(.TextMatrix(OldRow, GetCol("QtyShipped"))) * _
                                    NumValue(.TextMatrix(OldRow, GetCol("Price")))
        End If
                                            
    End With

    InvCalc

End Sub

Private Sub InitSoldTo()

    JCJob.OpenRS
    SQLString = "SELECT * FROM JCJob WHERE Active = 1 ORDER BY FullName"
    If JCJob.GetBySQL(SQLString) = False Then
        MsgBox "No Customer information exists!", vbExclamation
        GoBack
    End If
    
    xdbJob.ReDim 0, JCJob.Records + 1, 1, 2
    xdbJob.Value(0, 1) = "<Select a Customer:Job>"
    xdbJob.Value(0, 2) = 0
    
    I = 1
    Do
        xdbJob.Value(I, 1) = JCJob.FullName
        xdbJob.Value(I, 2) = JCJob.JobID
        I = I + 1
        If JCJob.GetNext = False Then Exit Do
    Loop
    
    tdbcmbSet Me.tdbcmbSoldTo

End Sub

Private Sub Init()

    LoadFlag = True
    
    InitSoldTo
    
    Me.lblCompanyName = ""
    
    ' if sat - add 2 days / else add 1 day
    If Int(Now()) Mod 7 = 7 Then
        tdbDateSet Me.tdbApptDate, Now() + 2
    Else
        tdbDateSet Me.tdbApptDate, Now() + 1
    End If
        
    ' populate the terms drop-down list
    ' use the InvGlobalID for the itemdata - must be a LONG value
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeTerms & _
                " ORDER BY Description"
    If InvGlobal.GetBySQL(SQLString) = True Then
        Do
            With Me.cmbTerms
                .AddItem InvGlobal.Description
                .ItemData(.NewIndex) = InvGlobal.GlobalID
            End With
            If InvGlobal.GetNext = False Then Exit Do
        Loop
    End If

    With Me
    
        .tdbSoldAddr1.Text = ""
        .tdbSoldAddr2.Text = ""
        .tdbSoldAddr3.Text = ""
        .tdbSoldAddr4.Text = ""
        .tdbSoldCity.Text = ""
        .tdbSoldState.Text = ""
        .tdbSoldZip.Text = ""
        
        .tdbShipAddr1.Text = ""
        .tdbShipAddr2.Text = ""
        .tdbShipAddr3.Text = ""
        .tdbShipAddr4.Text = ""
        .tdbShipCity.Text = ""
        .tdbShipState.Text = ""
        .tdbShipZip.Text = ""
        
        .tdbtxtPO1.Text = ""
        .tdbtxtPO2.Text = ""
        
        tdbAmountSet .tdbItemTotal
        tdbAmountSet .tdbFreight
        tdbAmountSet .tdbSalesTax
        tdbAmountSet .tdbInvTotal
        
        .tdbItemTotal.ReadOnly = True
        .tdbSalesTax.ReadOnly = True
        .tdbInvTotal.ReadOnly = True
        
        tdbIntegerSet .tdbPkgCount
        tdbDateSet .tdbOrderDate, Now()
        .lblInvDate.Caption = "Invoice Date:"
        
        With .tdbnumInvNum
            .MinValue = 0
            .MaxValue = 999999999
            .Format = "########0;(########0)"
            .DisplayFormat = "########0;########0; ; "
        End With
    
        ' force all caps for text fields
        CapSet .tdbtxtPO1, 25
        CapSet .tdbtxtPO2, 25
        
        CapSet .tdbSoldAddr1, 40
        CapSet .tdbSoldAddr2, 40
        CapSet .tdbSoldAddr3, 40
        CapSet .tdbSoldAddr4, 40
        CapSet .tdbSoldCity, 30
        CapSet .tdbSoldState, 2
        CapSet .tdbSoldZip, 10
    
        CapSet .tdbShipAddr1, 40
        CapSet .tdbShipAddr2, 40
        CapSet .tdbShipAddr3, 40
        CapSet .tdbShipAddr4, 40
        CapSet .tdbShipCity, 30
        CapSet .tdbShipState, 2
        CapSet .tdbShipZip, 10
    
        .txtApptTime.MaxLength = 10
    
        .fg.TabBehavior = flexTabCells
        .fgTrans.TabBehavior = flexTabCells
    
    End With
    
    ' description drop down - comments from InvGlobal
    CommDrop = "|#999999;<Add New>"
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeComment & _
                " ORDER BY Description"
    If InvGlobal.GetBySQL(SQLString) = True Then
        Do
            CommDrop = CommDrop & "|#" & InvGlobal.GlobalID & ";" & InvGlobal.Description
            If InvGlobal.GetNext = False Then Exit Do
        Loop
    End If

'    rs.CursorLocation = adUseClient
'    rs.Fields.Append "StockID", adDouble
'    rs.Fields.Append "QtyOrdered", adDouble
'    rs.Fields.Append "QtyShipped", adDouble
'    rs.Fields.Append "Description", adVarChar, 40, adFldIsNullable
'    rs.Fields.Append "UnitPrice", adDouble
'    rs.Fields.Append "Amount", adDouble
'    rs.Fields.Append "LineType", adInteger
'    rs.Fields.Append "RelatedID", adDouble
'    rs.Open , , adOpenDynamic, adLockOptimistic
'
'    SetGrid rs, fg

    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    
    rsCol.Open , , adOpenDynamic, adLockOptimistic

    ' set the column data
    AddCol "", "Col0", 300
    AddCol "Stock Item", "StockID", 2200
    AddCol "Qty Ordered", "QtyOrdered", 1200, , "###,##0"
    AddCol "Qty Shipped", "QtyShipped", 1200, , "###,##0"
    AddCol "Description", "Description", 5500
    AddCol "Unit Price", "Price", 1200, , "###,##0.0000"
    AddCol "Amount", "Amount", 1200, , "###,##0.00"
    AddCol "LineType", "LineType", 0
    AddCol "RelatedID", "RelatedID", 0

    With fg
        
        ' *** disconnected flex grid ***
        .Rows = 1
        .Cols = rsCol.RecordCount
        
        .FixedRows = 1
        .FixedCols = 1
        
        .ExplorerBar = flexExMoveRows
        .AllowBigSelection = False
        .Editable = flexEDKbdMouse
    
        I = 0
        rsCol.MoveFirst
        Do
            .TextMatrix(0, I) = rsCol!Title
            .ColWidth(I) = rsCol!Width
            .ColData(I) = rsCol!Abbrev
            If rsCol!dataType <> 0 Then
                .ColDataType(I) = rsCol!dataType
            End If
            If rsCol!Format <> 0 Then
                .ColFormat(I) = rsCol!Format
            End If
            I = I + 1
            rsCol.MoveNext
        Loop Until rsCol.EOF
    
        StockDrop = "|#0; "
        .ColComboList(GetCol("StockID")) = StockDrop
        .ColComboList(GetCol("Description")) = CommDrop
        
        .TabBehavior = flexTabCells
    
    End With

    ' fgTrans - Truck / Trailer / Driver grid init
    rsTrans.CursorLocation = adUseClient
    rsTrans.Fields.Append "Count", adInteger
    rsTrans.Fields.Append "TruckID", adDouble
    rsTrans.Fields.Append "TrailerID", adDouble
    rsTrans.Fields.Append "DriverID", adDouble
    rsTrans.Open , , adOpenDynamic, adLockOptimistic
    
    ' three entries available
    For I = 1 To 3
        rsTrans.AddNew
        rsTrans!Count = I
        rsTrans!TruckID = 0
        rsTrans!TrailerID = 0
        rsTrans!DriverID = 0
        rsTrans.Update
    Next I
    
    SetGrid rsTrans, fgTrans
    fgTrans.BackColorAlternate = 0
    
    With fgTrans
        
        For I = 0 To .Cols - 1
            .ColKey(I) = .TextMatrix(0, I)
        Next I
    
        .TextMatrix(0, .ColIndex("TruckID")) = "No - Truck - Lic"
        .TextMatrix(0, .ColIndex("TrailerID")) = "No - Trailer - Lic"
        .TextMatrix(0, .ColIndex("DriverID")) = "Driver"
    
        .ColWidth(.ColIndex("Count")) = 700
        J = 3360
        .ColWidth(.ColIndex("TruckID")) = J
        .ColWidth(.ColIndex("TrailerID")) = J
        .ColWidth(.ColIndex("DriverID")) = J
    
        .ColComboList(.ColIndex("TruckID")) = TransDropInit(InvEquate.GlobalTypeTruck)
        .ColComboList(.ColIndex("TrailerID")) = TransDropInit(InvEquate.GlobalTypeTrailer)
        .ColComboList(.ColIndex("DriverID")) = TransDropInit(InvEquate.GlobalTypeDriver)
        
        .Select 0, 1
    
    End With
    
    ' *** add test records ***
    For I = 1 To 15
        X = ""
        For J = 1 To rsCol.RecordCount
            X = X & ""
            If J <> rsCol.RecordCount Then X = X & vbTab
        Next J
        fg.AddItem X
    Next I
    ' *************************

    LoadFlag = False

End Sub

Private Sub CapSet(ByRef tdbTXT As TDBText, Optional lng As Integer)

    With tdbTXT
        .Format = "A9#@"
        .FormatMode = dbiIncludeFormat
        
        If lng > 0 Then
            .MaxLength = lng
        End If
        
    End With

End Sub

Private Function TransDropInit(ByVal bytTypeCode As Byte) As String

    TransDropInit = "|#0;NONE"
    
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & bytTypeCode & _
                " ORDER BY Description"
    If InvGlobal.GetBySQL(SQLString) = True Then
        Do
            TransDropInit = TransDropInit & "|#" & InvGlobal.GlobalID & ";" & InvGlobal.Description
            If InvGlobal.GetNext = False Then Exit Do
        Loop
    End If
 
End Function

Private Sub tdbcmbSet(ByRef tdbcmb As TDBCombo)

    With tdbcmb
        .Array = xdbJob
        .ScrollBars = dblVertical
        .Caption = "Select Customer:Job"
        .Columns(0).Width = 1000
        .Columns(1).Visible = False
        .AutoCompletion = True
        .AutoDropdown = True
        .LimitToList = True
    End With

End Sub

Private Sub tdbcmbSoldTo_LostFocus()

    ' fill the sold/ship text boxes
    If Me.tdbcmbSoldTo.Text = "" Then Exit Sub
    If Me.tdbcmbSoldTo.SelectedItem = 0 Then Exit Sub
    
    boo = JCJob.GetByID(xdbJob.Value(Me.tdbcmbSoldTo.SelectedItem, 2))
    
    ' update the stock list for the QB job
    Me.MousePointer = vbHourglass
    ' ItemUpd JCJob.JobID
    ItemUpd JCJob.ParentID
    Me.MousePointer = vbArrow
    
    ' get the sales tax info from the JCCustomer record
    boo = JCCustomer.GetByID(JCJob.ParentID)
    
    CalcSalesTaxPct
    
    With Me
        
        .tdbSoldAddr1 = JCJob.BillAddr1
        .tdbSoldAddr2 = JCJob.BillAddr2
        .tdbSoldAddr3 = JCJob.BillAddr3
        .tdbSoldAddr4 = JCJob.BillAddr4
        .tdbSoldCity = JCJob.BillCity
        .tdbSoldState = JCJob.BillState
        .tdbSoldZip = JCJob.BillZip
        
        .tdbShipAddr1 = JCJob.ShipAddr1
        .tdbShipAddr2 = JCJob.ShipAddr2
        .tdbShipAddr3 = JCJob.ShipAddr3
        .tdbShipAddr4 = JCJob.ShipAddr4
        .tdbShipCity = JCJob.ShipCity
        .tdbShipState = JCJob.ShipState
        .tdbShipZip = JCJob.ShipZip
    
    End With
    
    ' terms
    ' find using the QBID
    With Me.cmbTerms
        .ListIndex = -1
        If JCJob.Terms <> "" Then
            SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeTerms & _
                        " AND Var1 = '" & JCJob.Terms & "'"
            If InvGlobal.GetBySQL(SQLString) = True Then
                cmbPoint Me.cmbTerms, InvGlobal.GlobalID
            End If
        End If
    End With

    CustomerComments

    LoadStock
    
    InvCalc

End Sub

Private Sub LoadStock()
    
    ' create stock records for this client if none exist
    ' use master pricing
    '
    ' if job is marked inactive (deletes stock items) and then active again
    ' this will need to be done
    
    ' inventory item dropdown
    StockDrop = "|#0; "
    SQLString = "SELECT * FROM InvStock WHERE JobID = " & JCJob.JobID & _
                " ORDER BY Description"
    
    ' prices per CUSTOMER not JOB
    SQLString = "SELECT * FROM InvStock WHERE JobID = " & JCJob.ParentID & _
                " ORDER BY Description"
    
    If InvStock.GetBySQL(SQLString) = True Then
        Do
            If InvStock.StockSelect = True Then
                StockDrop = StockDrop & "|#" & InvStock.StockID & ";" & InvStock.Description
            End If
            If InvStock.GetNext = False Then Exit Do
        Loop
    Else        ' none exist - create then from the master
        
        ' !!!! should not be necessary !!!!
        
        SQLString = "SELECT * FROM InvStock WHERE JobID = 0 ORDER BY Description"
        rsInit SQLString, cn, rsStock
        If rsStock.RecordCount > 0 Then
            rsStock.MoveFirst
            Do
                InvStock.Clear
                InvStock.Cost = rsStock!Cost
                InvStock.CustomerPrice = rsStock!MasterPrice
                InvStock.Description = Trim(rsStock!Description)
                InvStock.InventoryItem = rsStock!InventoryItem
                InvStock.JobID = JCJob.JobID
                InvStock.QBID = Trim(rsStock!QBID)
                InvStock.QBName = Trim(rsStock!QBName)
                InvStock.StockSelect = True
                InvStock.rsAdd
                
                StockDrop = StockDrop & "|#" & InvStock.StockID & ";" & InvStock.Description
    
                rsStock.MoveNext
            Loop Until rsStock.EOF
        End If
    End If
    
    With fg
        .ColComboList(GetCol("StockID")) = StockDrop
    End With

End Sub

Private Sub cmdAddLine_Click()
    With fg
        If .Row = 0 Then Exit Sub
        K = .Row
        X = ""
        For J = 1 To rsCol.RecordCount
            X = X & ""
            If J <> rsCol.RecordCount Then X = X & vbTab
        Next J
        fg.AddItem X, K
    End With
End Sub
Private Sub cmdDelLine_Click()
    With fg
        If .Rows = 2 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub

Private Sub cmdClear_Click()
    
    ' ok to discard changes?
    If CheckForChange = False Then Exit Sub
    
    With Me
        
        ' added 2012-05-29
        ' UseSalesTax = False
        
        .tdbPalletCount = 0
        
        .tdbcmbSoldTo.SelectedItem = 0
        .cmbTerms.ListIndex = -1
        .tdbnumInvNum.Value = 0
        .tdbOrderDate.Value = Null
        
        InvDate = 0
        .lblInvDate = "Invoice Date: "
        .lblQBUpd = ""
        .tdbOrderDate.Value = Now()
        
        .tdbSoldAddr1.Text = ""
        .tdbSoldAddr2.Text = ""
        .tdbSoldAddr3.Text = ""
        .tdbSoldAddr4.Text = ""
        .tdbSoldCity.Text = ""
        .tdbSoldState.Text = ""
        .tdbSoldZip.Text = ""
        
        .tdbShipAddr1.Text = ""
        .tdbShipAddr2.Text = ""
        .tdbShipAddr3.Text = ""
        .tdbShipAddr4.Text = ""
        .tdbShipCity.Text = ""
        .tdbShipState.Text = ""
        .tdbShipZip.Text = ""
        
        .tdbtxtPO1.Text = ""
        .tdbtxtPO2.Text = ""
        
        rsTrans.MoveFirst
        Do
            rsTrans!TruckID = 0
            rsTrans!TrailerID = 0
            rsTrans!DriverID = 0
            rsTrans.Update
            rsTrans.MoveNext
        Loop Until rsTrans.EOF
        rsTrans.MoveFirst
        
        ' delete the lines from the fg
        Do
            If fg.Rows = 1 Then Exit Do
            fg.Row = fg.Rows - 1
            fg.RemoveItem
        Loop
        X = ""
        For I = 1 To 10
            fg.AddItem X, I
        Next I
    
        .lblCustMsg1.Caption = ""
        .lblCustMsg2.Caption = ""
        
        .tdbItemTotal.Value = 0
        .tdbFreight.Value = 0
        .tdbSalesTax.Value = 0
        .tdbInvTotal.Value = 0
        
        ' if sat - add 2 days / else add 1 day
        If Int(Now()) Mod 7 = 7 Then
            tdbDateSet Me.tdbApptDate, Now() + 2
        Else
            tdbDateSet Me.tdbApptDate, Now() + 1
        End If
        
        .txtApptTime = "03:30 AM"
    
    End With

    InvHeader.Clear
    LoadScreenVals

End Sub


Private Function SetText(ByVal CompName As String, _
                         ByVal Addr1 As String, _
                         ByVal Addr2 As String, _
                         ByVal Addr3 As String, _
                         ByVal Addr4 As String, _
                         ByVal City As String, _
                         ByVal State As String, _
                         ByVal Zip As String) As String

    J = 0
    For I = 1 To 6
        If I = 1 Then X = CompName
        If I = 2 Then X = Addr1
        If I = 3 Then X = Addr2
        If I = 4 Then X = Addr3
        If I = 5 Then X = Addr4
        If I = 6 Then
            X = City
            If X <> "" Then X = X & ", "
            X = X & State & "  " & Zip
        End If
        If X <> "" Then
            J = J + 1
            If J = 1 Then
                SetText = X
            Else
                SetText = SetText & vbCrLf & X
            End If
        End If
    Next I

End Function

Private Sub tdbcmbSoldTo_NotInList(NewEntry As String, Retry As Integer)
    MsgBox "Please enter a valid Customer:Job", vbExclamation
    Me.tdbcmbSoldTo.SetFocus
End Sub

Private Sub AddCol(ByVal Title As String, _
                   ByVal Abbrev As String, _
                   ByVal Width As Long, _
                   Optional DType As Byte, _
                   Optional Fmt As String)

    rsCol.AddNew
    rsCol!Title = Mid(Title, 1, 30)
    rsCol!Abbrev = Mid(Abbrev, 1, 30)
    rsCol!Width = Width
    rsCol!Number = ColNum
    rsCol!dataType = DType
    rsCol!Format = Fmt
    rsCol.Update
    
    ColNum = ColNum + 1

End Sub

Private Function GetCol(ByVal ColData As String) As Long

    SQLString = "Abbrev = '" & ColData & "'"
    rsCol.Find SQLString, 0, adSearchForward, 1
    If rsCol.EOF Then
        GetCol = -1
    Else
        GetCol = rsCol!Number
    End If

End Function

Private Sub cmdPrint_Click()

    cmdSave_Click   ' save the current invoice
    
    VertAdj = 0
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeVAdj
    If InvGlobal.GetBySQL(SQLString) = True Then
        VertAdj = InvGlobal.Byte1
    End If
    
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeInvPrinter
    If InvGlobal.GetBySQL(SQLString) = False Then
        MsgBox "Use Global Maintenance to select the invoice printer!", vbExclamation
        Exit Sub
    End If
    
    KP_PrintInvoice Me.tdbnumInvNum, InvGlobal.Var1

End Sub

Private Sub fgTrans_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = fgTrans.ColIndex("Count") Then Cancel = True
End Sub

Private Sub tdbFreight_Change()
    If LoadFlag = True Then Exit Sub
    InvCalc
End Sub

Private Function DataChanged() As Boolean

End Function

Private Sub LoadScreenVals()

    ' load the screen values into variables
    ' use to compare to screen values later
    ' give warning to save

    'Dim ScreenVals(26) As String
    'Dim fgTransVals(3, 5) As String
    'Dim fgVals(100, 6) As String

    With Me
            
        ScreenVals(1) = .tdbcmbSoldTo.SelectedItem
        ScreenVals(2) = .tdbnumInvNum.Text
        ScreenVals(3) = .tdbOrderDate.Text
        ScreenVals(4) = .tdbOrderDate.Text
        ScreenVals(5) = .lblInvDate.Caption
        ScreenVals(6) = .tdbItemTotal.Text
        ScreenVals(7) = .tdbFreight.Text
        ScreenVals(8) = .tdbSalesTax.Text
        ScreenVals(9) = .tdbInvTotal.Text
        
        ScreenVals(10) = .tdbSoldAddr1.Text
        ScreenVals(11) = .tdbSoldAddr2.Text
        ScreenVals(12) = .tdbSoldAddr3.Text
        ScreenVals(13) = .tdbSoldAddr4.Text
            
        ScreenVals(14) = .tdbSoldCity.Text
        ScreenVals(15) = .tdbSoldState.Text
        ScreenVals(16) = .tdbSoldZip.Text
        
        ScreenVals(17) = .tdbShipAddr1.Text
        ScreenVals(18) = .tdbShipAddr2.Text
        ScreenVals(19) = .tdbShipAddr3.Text
        ScreenVals(20) = .tdbShipAddr4.Text
            
        ScreenVals(21) = .tdbShipCity.Text
        ScreenVals(22) = .tdbShipState.Text
        ScreenVals(23) = .tdbShipZip.Text
            
        ScreenVals(24) = .tdbtxtPO1.Text
        ScreenVals(25) = .tdbtxtPO2.Text
            
        ScreenVals(26) = .tdbApptDate.Text
        ScreenVals(27) = .txtApptTime.Text
        ScreenVals(28) = .tdbPkgCount.Text
        ScreenVals(29) = .tdbPalletCount.Text
        
        With .fgTrans
            For I = 1 To .Rows - 1
                For J = 0 To .Cols - 1
                    fgTransVals(I, J) = .TextMatrix(I, J)
                Next J
            Next I
        End With
    
        With .fg
            .Rows = 25
            For I = 1 To .Rows - 1
                For J = 0 To .Cols - 1
                    fgVals(I, J) = .TextMatrix(I, J)
                Next J
            Next I
        End With
    
    End With

End Sub

Private Function CheckForChange() As Boolean

Dim ChangeFlag As Boolean

    ' blank screen - don't check
    If Me.tdbnumInvNum.Text = "" Then
        CheckForChange = True
        Exit Function
    End If
    
    ChangeFlag = False
    
    With Me
            
        For I = 1 To 29

            If I = 1 Then X = .tdbcmbSoldTo.SelectedItem
            
            If I = 2 Then X = .tdbnumInvNum.Text
            If I = 3 Then X = .tdbOrderDate.Text
            If I = 4 Then X = .tdbOrderDate.Text
            If I = 5 Then X = .lblInvDate.Caption
            If I = 6 Then X = .tdbItemTotal.Text
            If I = 7 Then X = .tdbFreight.Text
            If I = 8 Then X = .tdbSalesTax.Text
            If I = 9 Then X = .tdbInvTotal.Text
            
            If I = 10 Then X = .tdbSoldAddr1.Text
            If I = 11 Then X = .tdbSoldAddr2.Text
            If I = 12 Then X = .tdbSoldAddr3.Text
            If I = 13 Then X = .tdbSoldAddr4.Text
                
            If I = 14 Then X = .tdbSoldCity.Text
            If I = 15 Then X = .tdbSoldState.Text
            If I = 16 Then X = .tdbSoldZip.Text
            
            If I = 17 Then X = .tdbShipAddr1.Text
            If I = 18 Then X = .tdbShipAddr2.Text
            If I = 19 Then X = .tdbShipAddr3.Text
            If I = 20 Then X = .tdbShipAddr4.Text
                
            If I = 21 Then X = .tdbShipCity.Text
            If I = 22 Then X = .tdbShipState.Text
            If I = 23 Then X = .tdbShipZip.Text
                
            If I = 24 Then X = .tdbtxtPO1.Text
            If I = 25 Then X = .tdbtxtPO2.Text
                
            If I = 26 Then X = .tdbApptDate.Text
            If I = 27 Then X = .txtApptTime.Text
            If I = 28 Then X = .tdbPkgCount.Text
            If I = 29 Then X = .tdbPalletCount.Text
            
            If X <> ScreenVals(I) Then
                'MsgBox I & vbCr & X & vbCr & ScreenVals(I)
                ChangeFlag = True
            End If
            
        Next I

        With .fgTrans
            For I = 1 To .Rows - 1
                For J = 0 To .Cols - 1
                    If fgTransVals(I, J) <> .TextMatrix(I, J) Then ChangeFlag = True
                Next J
            Next I
        End With

        With .fg
            For I = 1 To .Rows - 1
                For J = 0 To .Cols - 1
                    If fgVals(I, J) <> .TextMatrix(I, J) Then ChangeFlag = True
                Next J
            Next I
        End With
    
    End With

    If ChangeFlag = False Then
        ' nothing changed
        CheckForChange = True
    Else
        ' ask if OK to discard changes
        X = "OK to discard changes to this invoice?"
        If MsgBox(X, vbYesNo + vbQuestion) = vbNo Then
            CheckForChange = False      ' not ok to go ahead
        Else
            CheckForChange = True       ' OK to go ahead
        End If
    End If

End Function

Private Sub cmdQBJobRefresh_Click()
    
Dim JobID As Long

Dim preferencesQuery As IPreferencesQuery
Dim preferencesRet As IPreferencesRet
    
    ' save the selected job is there is one
    If Me.tdbcmbSoldTo.Text <> "" Then
        JobID = xdbJob.Value(Me.tdbcmbSoldTo.SelectedItem, 2)
    Else
        JobID = 0
    End If
    
    frmJCGetQBData.Show vbModal
    
    ' *************************************************************************************
    ' get whether the QB company has sales tax
    SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeSalesTax & _
                " AND CompanyID = " & PRCompany.CompanyID
    If InvGlobal.GetBySQL(SQLString) = False Then
        InvGlobal.Clear
        InvGlobal.TypeCode = InvEquate.GlobalTypeSalesTax
        InvGlobal.CompanyID = PRCompany.CompanyID
        InvGlobal.rsAdd
    End If
    
    If QBOpen(Me, Me.lblCompanyName) = False Then
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    ' make the QB request
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    Set preferencesQuery = requestMsgSet.AppendPreferencesQueryRq
    preferencesQuery.IncludeRetElementList.Add "SalesTaxPreferences"
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    SessMgr.EndSession
    SessMgr.CloseConnection
    
    ' read the response
    Set ResponseList = responseMsgSet.ResponseList
    If Not (ResponseList Is Nothing) Then
        For I = 0 To ResponseList.Count - 1
            Set Response = ResponseList.GetAt(I)
            If (Response.StatusCode = 0) Then
                If (Not Response.Detail Is Nothing) Then
                    ResponseType = Response.Type.GetValue
                    If (ResponseType = rtPreferencesQueryRs) Then
                        Set preferencesRet = Response.Detail
                        If (Not preferencesRet.SalesTaxPreferences Is Nothing) Then
                            InvGlobal.Byte1 = 1     ' sales tax
                            UseSalesTax = True
                        Else
                            InvGlobal.Byte1 = 0     ' no sales tax used
                            UseSalesTax = False
                        End If
                        InvGlobal.rsPut
                    End If
                End If
            End If
        Next I
    End If
    
    ' *************************************************************************************
    
    InitSoldTo
    Me.tdbcmbSoldTo.ReBind  ' ***
    
    ' point to the job again
    If JobID <> 0 Then
        I = xdbJob.Find(0, 2, JobID, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
        If I < 0 Then
            MsgBox "Customer Job not found: " & InvHeader.SoldJobID, vbExclamation
            GoBack
        End If
        boo = JCJob.GetByID(InvHeader.SoldJobID)
        Me.tdbcmbSoldTo.SelectedItem = I
    End If

End Sub


Private Sub tdbnumInvNum_GotFocus()
    
    With Me.tdbnumInvNum
        If InvHeader.SaveFlag = 0 Then
            .Enabled = True
        Else
            .Enabled = False
        End If
    End With
        
End Sub

Private Sub tdbnumInvNum_LostFocus()

    With Me.tdbnumInvNum
        If InvHeader.SaveFlag = 0 Then
            SQLString = "SELECT * FROM InvHeader WHERE InvoiceNumber = " & .Value
            rsInit SQLString, cn, rsInvH
            If rsInvH.RecordCount > 0 Then
                MsgBox "This inv/order number already exists!", vbExclamation
                .Value = InvHeader.InvoiceNumber
                .Text = InvHeader.InvoiceNumber
            End If
        End If
    End With
    
End Sub

Private Sub CalcSalesTaxPct()

    ' 2012-06-09 moved to sub - call on sold to lose focus & invcalc
    
    SalesTaxPct = 0
    If UseSalesTax = True Then
        If JCCustomer.QBTaxCode = "" Then
            MsgBox "QB Tax CODE not set for this CUSTOMER record", vbExclamation
            cmdClear_Click
            Exit Sub
        End If
        
        If JCCustomer.QBTaxItem = "" Then
            MsgBox "QB Tax ITEM not set for this CUSTOMER record", vbExclamation
            cmdClear_Click
            Exit Sub
        End If
        
        SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'SALESTAXCODE' AND " & _
                    "QBID = '" & JCCustomer.QBTaxCode & "'"
        If QBAccount.GetBySQL(SQLString) = True Then
            If QBAccount.Description = "True" Then
                SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'SALESTAX' AND " & _
                            "QBID = '" & JCCustomer.QBTaxItem & "'"
                If QBAccount.GetBySQL(SQLString) = True Then
                    SalesTaxPct = CDbl(QBAccount.AccountNumber) / 10000
                End If
            End If
        End If
    End If

End Sub
