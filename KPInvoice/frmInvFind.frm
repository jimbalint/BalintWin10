VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmInvFind 
   Caption         =   "Search for Order/Invoice"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMark 
      Caption         =   "&MARK"
      Height          =   615
      Left            =   8760
      TabIndex        =   19
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   615
      Left            =   8760
      TabIndex        =   18
      Top             =   9360
      Width           =   1575
   End
   Begin TrueOleDBList80.TDBCombo tdbCustomer 
      Height          =   390
      Left            =   3840
      TabIndex        =   17
      Top             =   2280
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   688
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      _DropdownWidth  =   0
      _EDITHEIGHT     =   688
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
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
      EditFont        =   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
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
      _PropDict       =   $"frmInvFind.frx":030A
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=3,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Arial"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
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
   Begin VB.CheckBox chkAllInvoiceDates 
      Caption         =   "ALL &INVOICE DATES"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CheckBox chkAllOrderDates 
      Caption         =   "ALL &ORDER DATES"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   855
      Left            =   9120
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Default         =   -1  'True
      Height          =   975
      Left            =   5280
      TabIndex        =   8
      Top             =   8880
      Width           =   2055
   End
   Begin TDBDate6Ctl.TDBDate tdbStartOrderDate 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calendar        =   "frmInvFind.frx":03B4
      Caption         =   "frmInvFind.frx":04B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvFind.frx":0514
      Keys            =   "frmInvFind.frx":0532
      Spin            =   "frmInvFind.frx":0590
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
      Text            =   "08/12/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40402
      CenturyMode     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Select Invoice Type  "
      Height          =   975
      Left            =   3533
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      Begin VB.OptionButton optClosed 
         Caption         =   "&CLOSED"
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optPending 
         Caption         =   "&PENDING"
         Height          =   375
         Left            =   2220
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&ALL"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   12495
      _cx             =   22040
      _cy             =   7858
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   11040
      TabIndex        =   0
      Top             =   9360
      Width           =   1575
   End
   Begin TDBDate6Ctl.TDBDate tdbEndOrderDate 
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   2880
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calendar        =   "frmInvFind.frx":05B8
      Caption         =   "frmInvFind.frx":06B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvFind.frx":0714
      Keys            =   "frmInvFind.frx":0732
      Spin            =   "frmInvFind.frx":0790
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
      Text            =   "08/12/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40402
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate tdbStartInvoiceDate 
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calendar        =   "frmInvFind.frx":07B8
      Caption         =   "frmInvFind.frx":08B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvFind.frx":0918
      Keys            =   "frmInvFind.frx":0936
      Spin            =   "frmInvFind.frx":0994
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
      Text            =   "08/12/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40402
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate tdbEndInvoiceDate 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   3360
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calendar        =   "frmInvFind.frx":09BC
      Caption         =   "frmInvFind.frx":0ABC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvFind.frx":0B18
      Keys            =   "frmInvFind.frx":0B36
      Spin            =   "frmInvFind.frx":0B94
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
      Text            =   "08/12/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40402
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Select Customer:"
      Height          =   255
      Left            =   2093
      TabIndex        =   12
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   9000
      Width           =   4695
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
Attribute VB_Name = "frmInvFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Public InvNum As Long
Dim xdbJob As New XArrayDB
Dim rs As New ADODB.Recordset

Private Sub cmdDelete_Click()
    
    Dim InvNum As Long
    
    If rs!InvoiceDate <> "" Then
        MsgBox "Invoiced Orders can not bet deleted!", vbExclamation
        Exit Sub
    End If
    
    InvNum = rs!InvoiceNumber
    
    X = "OK to delete Invoice#: " & InvNum
    If MsgBox(X, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    X = "DELETE * FROM InvHeader WHERE HeaderID = " & rs!HeaderID
    cn.Execute X
    
    X = "DELETE * FROM InvBody WHERE HeaderID = " & rs!HeaderID
    cn.Execute X
    
    MsgBox "Invoice #: " & InvNum & " has been deleted", vbInformation
    cmdLoad_Click
    
End Sub

Private Sub cmdMark_Click()
    If MsgBox("OK to mark as invoiced?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    SQLString = "UPDATE InvHeader SET InvoiceDate = OrderDate, " & _
                "QBInvoiceID = '---' " & _
                "WHERE HeaderID = " & rs!HeaderID
    cn.Execute SQLString
    cmdLoad_Click

End Sub

Private Sub cmdSelect_Click()
    InvNum = 0
    On Error Resume Next
    InvNum = rs!InvoiceNumber
    On Error GoTo 0
    Me.Hide
End Sub

Private Sub Form_Load()

    Init
    
    If UCase(User.Logon) <> "JIM" Then Me.cmdMark.Visible = False
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    InvNum = 0
    Me.Hide
End Sub

Public Sub Init()

    SQLString = "SELECT * FROM JCJob ORDER BY FullName"
    If JCJob.GetBySQL(SQLString) = False Then
        MsgBox "No Customer information exists!", vbExclamation
        GoBack
    End If
    
    Set xdbJob = New XArrayDB
    xdbJob.ReDim 0, JCJob.Records, 1, 2
    
    xdbJob.Value(0, 1) = "<ALL>"
    xdbJob.Value(0, 2) = 0
    
    I = 1
    Do
        xdbJob.Value(I, 1) = JCJob.FullName
        xdbJob.Value(I, 2) = JCJob.JobID
        I = I + 1
        If JCJob.GetNext = False Then Exit Do
    Loop
    
    With Me.tdbCustomer
        .ScrollBars = dblVertical
        .Caption = "Select Customer:Job"
        .Columns(0).Width = 1000
        .Columns(1).Visible = False
        .AutoCompletion = True
        .AutoDropdown = True
        .LimitToList = True
        .Array = xdbJob
        .SelectedItem = 0
    End With

    Me.optAll = False
    Me.optPending = False
    Me.optClosed = True
    
    Me.chkAllOrderDates = 1
    Me.tdbStartOrderDate.Visible = False
    Me.tdbEndOrderDate.Visible = False
    
    Me.chkAllInvoiceDates = 1
    Me.tdbStartInvoiceDate.Visible = False
    Me.tdbEndInvoiceDate.Visible = False
    
    tdbDateSet Me.tdbStartOrderDate, Now
    tdbDateSet Me.tdbEndOrderDate, Now
    tdbDateSet Me.tdbStartInvoiceDate, Now
    tdbDateSet Me.tdbEndInvoiceDate, Now

End Sub
Private Sub cmdLoad_Click()

    SQLString = "SELECT * FROM InvHeader "
    With Me
        
        If .optAll = True Then
        ElseIf .optPending = True Then
            SQLString = SQLString & " WHERE QBInvoiceID = ''"
        Else
            SQLString = SQLString & " WHERE QBInvoiceID <> ''"
        End If
        
        If .chkAllOrderDates = 0 Then
            If Mid(SQLString, Len(SQLString), 1) <> " " Then
                SQLString = SQLString & " AND "
            Else
                SQLString = SQLString & " WHERE "
            End If
            SQLString = SQLString & "Int(CDbl(OrderDate)) >= " & Int(CDbl(.tdbStartOrderDate.Value)) & _
                        " AND Int(CDbl(OrderDate)) <= " & Int(CDbl(.tdbEndOrderDate.Value))
        End If
        
        If .chkAllInvoiceDates = 0 Then
            If Mid(SQLString, Len(SQLString), 1) <> " " Then
                SQLString = SQLString & " AND "
            Else
                SQLString = SQLString & " WHERE "
            End If
            SQLString = SQLString & "Int(CDbl(InvoiceDate)) >= " & Int(CDbl(.tdbStartInvoiceDate.Value)) & _
                        " AND Int(CDbl(InvoiceDate)) <= " & Int(CDbl(.tdbEndInvoiceDate.Value))
        End If
        
        If .tdbCustomer.SelectedItem <> 0 Then
            If Mid(SQLString, Len(SQLString), 1) <> " " Then
                SQLString = SQLString & " AND "
            Else
                SQLString = SQLString & " WHERE "
            End If
            SQLString = SQLString & " SoldJobID = " & xdbJob.Value(.tdbCustomer.SelectedItem, 2)
        End If

    End With
    
    SQLString = SQLString & " ORDER BY InvoiceNumber"
    
    If InvHeader.GetBySQL(SQLString) = False Then
        MsgBox "No Invoices found for the ranges selected", vbInformation
        Exit Sub
    End If

    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "InvoiceNumber", adDouble
    rs.Fields.Append "CustomerName", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "InvoiceAmount", adCurrency
    rs.Fields.Append "OrderDate", adVarChar, 10, adFldIsNullable
    rs.Fields.Append "InvoiceDate", adVarChar, 10, adFldIsNullable
    rs.Fields.Append "HeaderID", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic

    Do
        rs.AddNew
        rs!InvoiceNumber = InvHeader.InvoiceNumber
        
        If InvHeader.InvoiceDate <> 0 Then
            rs!InvoiceDate = Format(InvHeader.InvoiceDate, "mm/dd/yyyy")
        Else
            rs!InvoiceDate = ""
        End If
        
        If InvHeader.OrderDate <> 0 Then
            rs!OrderDate = Format(InvHeader.OrderDate, "mm/dd/yyyy")
        Else
            rs!OrderDate = ""
        End If
        
        rs!HeaderID = InvHeader.HeaderID
        
        rs!InvoiceAmount = InvHeader.TotalAmount
        boo = JCJob.GetByID(InvHeader.SoldJobID)
        rs!CustomerName = Mid(JCJob.FullName, 1, 30)
        
        rs.Update
        If InvHeader.GetNext = False Then Exit Do
    Loop
    
    SetGrid rs, fg

    fg.ColWidth(5) = 0

    rs.MoveFirst

End Sub

Private Sub chkAllOrderDates_Click()
    
    If Me.chkAllOrderDates = 0 Then
        Me.tdbStartOrderDate.Visible = True
        Me.tdbEndOrderDate.Visible = True
    Else
        Me.tdbStartOrderDate.Visible = False
        Me.tdbEndOrderDate.Visible = False
    End If

End Sub
Private Sub chkAllInvoiceDates_Click()
    
    If Me.chkAllInvoiceDates = 0 Then
        Me.tdbStartInvoiceDate.Visible = True
        Me.tdbEndInvoiceDate.Visible = True
    Else
        Me.tdbStartInvoiceDate.Visible = False
        Me.tdbEndInvoiceDate.Visible = False
    End If

End Sub



