VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmW2 
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9360
   ScaleWidth      =   12570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   9000
      TabIndex        =   87
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE CHANGES"
      Height          =   375
      Left            =   9960
      TabIndex        =   86
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CheckBox chkSkip 
      Caption         =   "Skip this Employee"
      Height          =   375
      Left            =   7560
      TabIndex        =   85
      Top             =   7080
      Width           =   2055
   End
   Begin VB.ComboBox cmbW2Box14d 
      Height          =   345
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbW2Box14c 
      Height          =   345
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   78
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ComboBox cmbW2Box14b 
      Height          =   345
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   77
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox cmbW2Box14a 
      Height          =   345
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   76
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Height          =   375
      Left            =   11160
      Picture         =   "frmW2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   375
      Left            =   7560
      Picture         =   "frmW2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdStateDel 
      Caption         =   "DEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   73
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdStateAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   72
      Top             =   6000
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid fgState 
      Height          =   735
      Left            =   240
      TabIndex        =   70
      Top             =   6480
      Width           =   6495
      _cx             =   11456
      _cy             =   1296
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin VB.ComboBox cmbEEList 
      Height          =   345
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   6720
      Width           =   4575
   End
   Begin VB.CommandButton cmdCityDel 
      Caption         =   "DEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   67
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmdCityAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   65
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   375
      Left            =   8760
      Picture         =   "frmW2.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Height          =   375
      Left            =   9960
      Picture         =   "frmW2.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame fraOrder 
      Caption         =   "   SORT  ORDER  "
      Height          =   735
      Left            =   7560
      TabIndex        =   60
      Top             =   7680
      Width           =   4695
      Begin VB.OptionButton optOrderNumber 
         Caption         =   "EMPLOYEE NUMBER"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   62
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optOrderName 
         Caption         =   "EMPLOYEE NAME"
         Height          =   375
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&CREATE"
      Height          =   495
      Left            =   9960
      TabIndex        =   59
      Top             =   8760
      Width           =   855
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   345
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   10920
      TabIndex        =   56
      Top             =   8760
      Width           =   735
   End
   Begin TDBNumber6Ctl.TDBNumber TDBLine1 
      Height          =   495
      Left            =   4800
      TabIndex        =   37
      Top             =   600
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":0C28
      Caption         =   "frmW2.frx":0C48
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":0CD0
      Keys            =   "frmW2.frx":0CEE
      Spin            =   "frmW2.frx":0D38
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
   Begin VB.ComboBox cmbW2Box12d 
      Height          =   345
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5280
      Width           =   4215
   End
   Begin VB.ComboBox cmbW2Box12c 
      Height          =   345
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4560
      Width           =   4215
   End
   Begin VB.ComboBox cmbW2Box12b 
      Height          =   345
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3840
      Width           =   4215
   End
   Begin VB.ComboBox cmbW2Box12a 
      Height          =   345
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3120
      Width           =   4215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   11760
      TabIndex        =   29
      Top             =   8760
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid fgCity 
      Height          =   1335
      Left            =   240
      TabIndex        =   28
      Top             =   7800
      Width           =   6495
      _cx             =   11456
      _cy             =   2355
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin VB.CheckBox chkLine13c 
      Caption         =   "Check2"
      Height          =   225
      Left            =   6720
      TabIndex        =   21
      Top             =   4155
      Width           =   255
   End
   Begin VB.CheckBox chkLine13b 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5640
      TabIndex        =   20
      Top             =   4155
      Width           =   255
   End
   Begin VB.CheckBox chkLine13a 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   5040
      TabIndex        =   19
      Top             =   4155
      Width           =   255
   End
   Begin TDBText6Ctl.TDBText TDBLinea 
      Height          =   500
      Left            =   3360
      TabIndex        =   2
      Top             =   40
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   882
      Caption         =   "frmW2.frx":0D60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":0DFC
      Key             =   "frmW2.frx":0E1A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   11
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
   Begin VB.CheckBox chkVoid 
      Alignment       =   1  'Right Justify
      Caption         =   "Void"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   285
      Width           =   800
   End
   Begin TDBNumber6Ctl.TDBNumber TDBEmpNumber 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":0E5E
      Caption         =   "frmW2.frx":0E7E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":0EEA
      Keys            =   "frmW2.frx":0F08
      Spin            =   "frmW2.frx":0F52
      AlignHorizontal =   1
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineb 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   882
      Caption         =   "frmW2.frx":0F7A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":101C
      Key             =   "frmW2.frx":103A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC5 
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "frmW2.frx":107E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":10E2
      Key             =   "frmW2.frx":1100
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC2 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1740
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":1144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":11A8
      Key             =   "frmW2.frx":11C6
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC3 
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   2085
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":120A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":126E
      Key             =   "frmW2.frx":128C
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC4 
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   2415
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":12D0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1334
      Key             =   "frmW2.frx":1352
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC1 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   882
      Caption         =   "frmW2.frx":1396
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":143A
      Key             =   "frmW2.frx":1458
      BackColor       =   -2147483643
      EditMode        =   1
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLined 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   882
      Caption         =   "frmW2.frx":149C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1512
      Key             =   "frmW2.frx":1530
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   2
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   11
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
   Begin TDBText6Ctl.TDBText TDBLineE2 
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":1574
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":15D8
      Key             =   "frmW2.frx":15F6
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineE3 
      Height          =   300
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":163A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":169E
      Key             =   "frmW2.frx":16BC
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
   Begin TDBText6Ctl.TDBText TDBLineE4 
      Height          =   300
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   529
      Caption         =   "frmW2.frx":1700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1764
      Key             =   "frmW2.frx":1782
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineE5 
      Height          =   300
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   529
      Caption         =   "frmW2.frx":17C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":182A
      Key             =   "frmW2.frx":1848
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineE6 
      Height          =   300
      Left            =   1080
      TabIndex        =   14
      Top             =   5640
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "frmW2.frx":188C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":18F0
      Key             =   "frmW2.frx":190E
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText TDBLineC6 
      Height          =   300
      Left            =   960
      TabIndex        =   36
      Top             =   2760
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "frmW2.frx":1952
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":19B6
      Key             =   "frmW2.frx":19D4
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
      AlignVertical   =   2
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine2 
      Height          =   495
      Left            =   8160
      TabIndex        =   38
      Top             =   600
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":1A18
      Caption         =   "frmW2.frx":1A38
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1AC8
      Keys            =   "frmW2.frx":1AE6
      Spin            =   "frmW2.frx":1B30
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine3 
      Height          =   495
      Left            =   4800
      TabIndex        =   39
      Top             =   1080
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":1B58
      Caption         =   "frmW2.frx":1B78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1BFC
      Keys            =   "frmW2.frx":1C1A
      Spin            =   "frmW2.frx":1C64
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine4 
      Height          =   495
      Left            =   8160
      TabIndex        =   40
      Top             =   1080
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":1C8C
      Caption         =   "frmW2.frx":1CAC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1D24
      Keys            =   "frmW2.frx":1D42
      Spin            =   "frmW2.frx":1D8C
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine5 
      Height          =   495
      Left            =   4800
      TabIndex        =   41
      Top             =   1560
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":1DB4
      Caption         =   "frmW2.frx":1DD4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1E52
      Keys            =   "frmW2.frx":1E70
      Spin            =   "frmW2.frx":1EBA
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine6 
      Height          =   495
      Left            =   8160
      TabIndex        =   42
      Top             =   1560
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":1EE2
      Caption         =   "frmW2.frx":1F02
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":1F7C
      Keys            =   "frmW2.frx":1F9A
      Spin            =   "frmW2.frx":1FE4
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine7 
      Height          =   495
      Left            =   4800
      TabIndex        =   43
      Top             =   2040
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":200C
      Caption         =   "frmW2.frx":202C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":20AE
      Keys            =   "frmW2.frx":20CC
      Spin            =   "frmW2.frx":2116
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine8 
      Height          =   495
      Left            =   8160
      TabIndex        =   44
      Top             =   2040
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":213E
      Caption         =   "frmW2.frx":215E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":21D4
      Keys            =   "frmW2.frx":21F2
      Spin            =   "frmW2.frx":223C
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine9 
      Height          =   495
      Left            =   4800
      TabIndex        =   45
      Top             =   2520
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":2264
      Caption         =   "frmW2.frx":2284
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2304
      Keys            =   "frmW2.frx":2322
      Spin            =   "frmW2.frx":236C
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine10 
      Height          =   495
      Left            =   8160
      TabIndex        =   46
      Top             =   2520
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":2394
      Caption         =   "frmW2.frx":23B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":243E
      Keys            =   "frmW2.frx":245C
      Spin            =   "frmW2.frx":24A6
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine11 
      Height          =   495
      Left            =   4800
      TabIndex        =   47
      Top             =   3000
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   873
      Calculator      =   "frmW2.frx":24CE
      Caption         =   "frmW2.frx":24EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":256E
      Keys            =   "frmW2.frx":258C
      Spin            =   "frmW2.frx":25D6
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine12a 
      Height          =   300
      Left            =   8160
      TabIndex        =   48
      Top             =   3480
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":25FE
      Caption         =   "frmW2.frx":261E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2672
      Keys            =   "frmW2.frx":2690
      Spin            =   "frmW2.frx":26DA
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine12b 
      Height          =   300
      Left            =   8160
      TabIndex        =   49
      Top             =   4200
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2702
      Caption         =   "frmW2.frx":2722
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2776
      Keys            =   "frmW2.frx":2794
      Spin            =   "frmW2.frx":27DE
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine12c 
      Height          =   300
      Left            =   8160
      TabIndex        =   50
      Top             =   4920
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2806
      Caption         =   "frmW2.frx":2826
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":287A
      Keys            =   "frmW2.frx":2898
      Spin            =   "frmW2.frx":28E2
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
   Begin TDBNumber6Ctl.TDBNumber TDBLine12d 
      Height          =   300
      Left            =   8160
      TabIndex        =   51
      Top             =   5640
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":290A
      Caption         =   "frmW2.frx":292A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":297E
      Keys            =   "frmW2.frx":299C
      Spin            =   "frmW2.frx":29E6
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
   Begin TDBNumber6Ctl.TDBNumber tdbLine14a 
      Height          =   300
      Left            =   6120
      TabIndex        =   52
      Top             =   4680
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2A0E
      Caption         =   "frmW2.frx":2A2E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2A82
      Keys            =   "frmW2.frx":2AA0
      Spin            =   "frmW2.frx":2AEA
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
   Begin TDBNumber6Ctl.TDBNumber tdbLine14b 
      Height          =   300
      Left            =   6120
      TabIndex        =   53
      Top             =   5010
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2B12
      Caption         =   "frmW2.frx":2B32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2B86
      Keys            =   "frmW2.frx":2BA4
      Spin            =   "frmW2.frx":2BEE
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
   Begin TDBNumber6Ctl.TDBNumber tdbLine14c 
      Height          =   300
      Left            =   6120
      TabIndex        =   54
      Top             =   5400
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2C16
      Caption         =   "frmW2.frx":2C36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2C8A
      Keys            =   "frmW2.frx":2CA8
      Spin            =   "frmW2.frx":2CF2
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
   Begin TDBNumber6Ctl.TDBNumber tdbLine14d 
      Height          =   300
      Left            =   6120
      TabIndex        =   55
      Top             =   5760
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   529
      Calculator      =   "frmW2.frx":2D1A
      Caption         =   "frmW2.frx":2D3A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2D8E
      Keys            =   "frmW2.frx":2DAC
      Spin            =   "frmW2.frx":2DF6
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
   Begin TDBText6Ctl.TDBText tdbLineE_FirstName 
      Height          =   300
      Left            =   240
      TabIndex        =   81
      Top             =   4200
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   529
      Caption         =   "frmW2.frx":2E1E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2E82
      Key             =   "frmW2.frx":2EA0
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText tdbLineE_LastName 
      Height          =   300
      Left            =   2640
      TabIndex        =   82
      Top             =   4200
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   529
      Caption         =   "frmW2.frx":2EE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":2F48
      Key             =   "frmW2.frx":2F66
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
      AlignVertical   =   2
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
   Begin TDBText6Ctl.TDBText tdbLineE_MidInit 
      Height          =   300
      Left            =   2040
      TabIndex        =   83
      Top             =   4200
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   529
      Caption         =   "frmW2.frx":2FAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW2.frx":300E
      Key             =   "frmW2.frx":302C
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
      AlignVertical   =   2
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
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label17"
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
      Height          =   495
      Left            =   7080
      TabIndex        =   84
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label16 
      Caption         =   "e Employee Name, Address, City, State and Zip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   80
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   12600
      X2              =   7080
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7080
      X2              =   7080
      Y1              =   6120
      Y2              =   9360
   End
   Begin VB.Label Label15 
      Caption         =   "State Wages/Tax"
      Height          =   255
      Left            =   240
      TabIndex        =   71
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Employee Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "City Wages/Tax"
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label lblTaxYear 
      Caption         =   "Tax Year:"
      Height          =   255
      Left            =   7560
      TabIndex        =   58
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "12a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   35
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "12d"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   34
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "12c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   33
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "12b"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   32
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   5640
      TabIndex        =   31
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Emp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   30
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "14  Other"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   27
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   26
      Top             =   3750
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Sick Pay"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   25
      Top             =   3915
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Third Party "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   24
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Retirement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5640
      TabIndex        =   23
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Stat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   22
      Top             =   3720
      Width           =   495
   End
End
Attribute VB_Name = "frmW2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sqlstring1 As String
Dim SQLString2 As String

Dim rs As New ADODB.Recordset

Dim rsCity As New ADODB.Recordset
Dim rsC As New ADODB.Recordset

Dim rsState As New ADODB.Recordset
Dim rsS As New ADODB.Recordset

Dim i, j, k As Long

Dim Box12ID(4) As Long
Dim Box14ID(4) As Long

Dim StartYM, EndYM As Long
Dim LoadFlag As Boolean
Dim ID As Long
Dim StateDrop As String

' ---------------------------------------------------------
' School District tax
'
' added SDTax field to rsCity and PRW2City
'   = 1 indicates it is a school district tax
'   the CityID is the ItemID used for the title
'   the state is stuffed to 36 for Ohio - will work like
'     all other city taxes
' set if this client has any SD tax items
' the city wage is stuffed from PRHist.CWTWage
Dim SDTax As Boolean
Dim CityWages As Currency
Dim SDItemID As Long
Dim SDFlag As Boolean
' ---------------------------------------------------------


Private Sub Form_Load()
        
    SQLString = "SELECT * FROM GLCompany WHERE ID = " & PRCompany.GLCompanyID
    If GLCompany.GetBySQL(SQLString) Then
    End If
    
    LoadFlag = True
    
    Me.lblCompanyName = PRCompany.Name
    
    ' disable fg add/del buttons for now
    Me.cmdCityAdd.Enabled = False
    Me.cmdCityDel.Enabled = False
    Me.cmdStateAdd.Enabled = False
    Me.cmdStateDel.Enabled = False
    
' If TableExists("PRW2", cn) = True Then
'    SQLString = "DROP TABLE PRW2"
'    cn.Execute SQLString
' End If
' If TableExists("PRW2City", cn) = True Then
'    SQLString = "DROP TABLE PRW2City"
'    cn.Execute SQLString
' End If
' If TableExists("PRW2State", cn) = True Then
'    SQLString = "DROP TABLE PRW2State"
'    cn.Execute SQLString
' End If
    
    If TableExists("PRW2", cn) = False Then
        PRW2Create
    End If
    If TableExists("PRW2City", cn) = False Then
        PRW2CityCreate
    End If
    If TableExists("PRW2State", cn) = False Then
        PRW2StateCreate
    End If
    If AddField("PRW2City", "SDTax", "Byte", cn) Then
    End If

    ' does this client have any SD tax?
    SQLString = "SELECT * FROM PRItem WHERE ItemType = " & PREquate.ItemTypeSDTax
    SDTax = PRItem.GetBySQL(SQLString)

    FormInit
    Me.optOrderName = True

    With Me.cmbTaxYear
        GetData .ItemData(.ListIndex)
    End With
    
    DisplayData
    
    Me.KeyPreview = True

    LoadFlag = False

End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub FormInit()
    
Dim yr As Double
Dim SelYear As Long
    
    
    ' recordsets for accum of data to:
    '   PRW2State / PRW2City
    rsCity.CursorLocation = adUseClient
    rsCity.Fields.Append "CityID", adDouble
    rsCity.Fields.Append "CityWage", adCurrency
    rsCity.Fields.Append "CityTax", adCurrency
    rsCity.Fields.Append "SDTax", adInteger
    rsCity.Fields.Append "Courtesy", adInteger
    rsCity.Open , , adOpenDynamic, adLockOptimistic
    
    rsState.CursorLocation = adUseClient
    rsState.Fields.Append "StateID", adDouble
    rsState.Fields.Append "StateWage", adCurrency
    rsState.Fields.Append "StateTax", adCurrency
    rsState.Open , , adOpenDynamic, adLockOptimistic
    
    ' init the state dropdown
    StateDrop = ""
    SQLString = "SELECT * FROM PRState ORDER BY StateAbbrev"
    If PRState.GetBySQL(SQLString) Then
        Do
            StateDrop = Trim(StateDrop) & "|#" & PRState.StateID & ";" & PRState.StateAbbrev
            If PRState.GetNext = False Then Exit Do
        Loop
    End If
        
    ' ***********************************************************
    ' init the Tax Year dropdown
    rs.CursorLocation = adUseClient
    rs.Fields.Append "TaxYear", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRHist"
    If PRHist.GetBySQL(SQLString) = False Then
        ' does PRW2 records exist?
        SQLString = "SELECT * FROM PRW2"
        If PRW2.GetBySQL(SQLString) = False Then
            
            rs.AddNew
            rs!TaxYear = Year(Now) + 1
            rs.Update
            
            rs.AddNew
            rs!TaxYear = Year(Now)
            rs.Update
        
            rs.AddNew
            rs!TaxYear = Year(Now) - 1
            rs.Update
        
        Else
            Do
                rs.Find "TaxYear = " & PRW2.TaxYear, 0, adSearchForward, 1
                If rs.EOF Then
                    rs.AddNew
                    rs!TaxYear = Int(PRW2.TaxYear / 100)
                    rs.Update
                End If
                If PRW2.GetNext = False Then Exit Do
            Loop
        End If
    Else
        Do
            rs.Find "TaxYear = " & Int(PRHist.YearMonth / 100), 0, adSearchForward, 1
            If rs.EOF Then
                rs.AddNew
                rs!TaxYear = Int(PRHist.YearMonth / 100)
                rs.Update
            End If
            If PRHist.GetNext = False Then Exit Do
        Loop
    End If
    
    SelYear = 0
    rs.Sort = "TaxYear DESC"
    rs.MoveFirst
    i = 0
    Do
        With Me.cmbTaxYear
            .AddItem rs!TaxYear
            .ItemData(.NewIndex) = rs!TaxYear
            If Month(Now) >= 10 And rs!TaxYear = Year(Now) Then
                SelYear = i
            ElseIf Month(Now) < 10 And rs!TaxYear = Year(Now) - 1 Then
                SelYear = i
            End If
        End With
        i = i + 1
        rs.MoveNext
    Loop Until rs.EOF
    Me.cmbTaxYear.ListIndex = SelYear
    
    ' ***********************************************************
    ' Box 12 Init
    Box12Init Me.cmbW2Box12a
    Box12Init Me.cmbW2Box12b
    Box12Init Me.cmbW2Box12c
    Box12Init Me.cmbW2Box12d
    
    ' ***********************************************************
    ' Box 14 Init
    Box14Init Me.cmbW2Box14a
    Box14Init Me.cmbW2Box14b
    Box14Init Me.cmbW2Box14c
    Box14Init Me.cmbW2Box14d
    
    tdbIntegerSet Me.TDBEmpNumber
    Me.TDBEmpNumber.Enabled = False
    
    tdbAmountSet Me.TDBLine1
    tdbAmountSet Me.TDBLine2
    tdbAmountSet Me.TDBLine3
    tdbAmountSet Me.TDBLine4
    tdbAmountSet Me.TDBLine5
    tdbAmountSet Me.TDBLine6
    tdbAmountSet Me.TDBLine7
    tdbAmountSet Me.TDBLine8
    tdbAmountSet Me.TDBLine9
    tdbAmountSet Me.TDBLine10
    tdbAmountSet Me.TDBLine11
    tdbAmountSet Me.tdbLine14a
    tdbAmountSet Me.tdbLine14b
    tdbAmountSet Me.tdbLine14c
    tdbAmountSet Me.tdbLine14d
    
    tdbAmountSet Me.TDBLine12a
    tdbAmountSet Me.TDBLine12b
    tdbAmountSet Me.TDBLine12c
    tdbAmountSet Me.TDBLine12d
    
    tdbTextSet Me.TDBLineC1, 50
    tdbTextSet Me.TDBLineC2, 50
    tdbTextSet Me.TDBLineC3, 50
    tdbTextSet Me.TDBLineC4, 50
    tdbTextSet Me.TDBLineC5, 2
    tdbTextSet Me.TDBLineC6, 10
    
    tdbTextSet Me.tdbLineE_FirstName, 50
    tdbTextSet Me.tdbLineE_LastName, 50
    tdbTextSet Me.tdbLineE_MidInit, 2
    tdbTextSet Me.TDBLineE2, 50
    tdbTextSet Me.TDBLineE3, 50
    tdbTextSet Me.TDBLineE4, 50
    tdbTextSet Me.TDBLineE5, 2
    tdbTextSet Me.TDBLineE6, 10
    
    tdbAmountSet Me.tdbLine14a
    tdbAmountSet Me.tdbLine14b
    tdbAmountSet Me.tdbLine14c
    tdbAmountSet Me.tdbLine14d
    
    Me.TDBLineb = Mid(PRCompany.FederalID, 1, 15)
    Me.TDBLineC1 = Mid(PRCompany.Name, 1, 50)
    Me.TDBLineC2 = Mid(PRCompany.Address1, 1, 50)
    Me.TDBLineC3 = Mid(PRCompany.Address2, 1, 50)
    Me.TDBLineC4 = Mid(PRCompany.City, 1, 50)
    
    ' Me.TDBLineC5 = PRState.StateAbbrev        '  Get State Abbrev.
    Me.TDBLineC6 = Mid(PRCompany.ZipCode, 1, 10)
    
    Me.Caption = "TAX YEAR: " & Me.cmbTaxYear.Text & "  -  " & PRCompany.Name

    Me.optOrderName = True

End Sub

Private Sub Box12Init(ByRef cmb As ComboBox)
        
    With cmb
        .AddItem ""
        .ItemData(.NewIndex) = 0
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW2Box12 & _
                    " ORDER BY Description"
        If PRGlobal.GetBySQL(SQLString) Then
            Do
                .AddItem PRGlobal.Description
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If Not PRGlobal.GetNext Then Exit Do
            Loop
        End If
    End With

End Sub

Private Sub Box14Init(ByRef cmb As ComboBox)
        
    With cmb
        .AddItem ""
        .ItemData(.NewIndex) = 0
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW2Box14 & _
                    " ORDER BY Description"
        If PRGlobal.GetBySQL(SQLString) Then
            Do
                .AddItem PRGlobal.Description
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If Not PRGlobal.GetNext Then Exit Do
            Loop
        End If
    End With

End Sub

Private Sub CreateData(ByVal TaxYear As Long)
    
Dim W2ID As Long
    
    ' data already exists warning
    SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear
    If PRW2.GetBySQL(SQLString) = True Then
        
        If MsgBox("Data already exists for tax year: " & TaxYear & vbCr & _
               "OK to delete and re-create it?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        SQLString = "DELETE * FROM PRW2 WHERE TaxYear = " & TaxYear
        cn.Execute SQLString
    
        SQLString = "DELETE * FROM PRW2City WHERE TaxYear = " & TaxYear
        cn.Execute SQLString
    
        SQLString = "DELETE * FROM PRW2State WHERE TaxYear = " & TaxYear
        cn.Execute SQLString
    
    End If
    
    frmProgress.Show
    frmProgress.lblMsg1 = "Loading data for tax year: " & TaxYear
    
    ' init variables
    For ct = 1 To 4
        Box12ID(ct) = 0
        Box14ID(ct) = 0
    Next ct
    StartYM = TaxYear * 100 + 1
    EndYM = TaxYear * 100 + 12
    
    PRW2State.OpenRS
    PRW2City.OpenRS
    
    If Me.optOrderName = True Then
        SQLString = "SELECT * FROM PREmployee ORDER BY LastName, FirstName"
    Else
        SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
    End If
    
    If PREmployee.GetBySQL(SQLString) = False Then
        MsgBox "No Employees Found!", vbExclamation
        GoBack
    End If
        
    Do
    
        ct = ct + 1
        frmProgress.lblMsg2 = PREmployee.LFName & " " & _
                              PREmployee.EmployeeNumber & " " & _
                              Format(ct, "#,##0")
        frmProgress.Refresh
        
        
        SQLString = "SELECT * FROM PRHist WHERE EmployeeID = " & PREmployee.EmployeeID & " " & _
                    "AND YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM
        If PRHist.GetBySQL(SQLString) = False Then GoTo NextEmployee
        
        CityWages = 0   ' used for wage for SD tax
        
        PRW2.Clear
        
        PRW2.TaxYear = TaxYear
        PRW2.EmployeeID = PREmployee.EmployeeID
        PRW2.EmployeeNumber = PREmployee.EmployeeNumber
        PRW2.BoxA_SSNumber = PREmployee.SSN
        
        PRW2.BoxB_FedID = Mid(PRCompany.FederalID, 1, 15)
        
        PRW2.BoxC_ERName = Mid(PRCompany.Name, 1, 50)
        PRW2.BoxC_ERAddr1 = Mid(PRCompany.Address1, 1, 50)
        PRW2.BoxC_ERAddr2 = Mid(PRCompany.Address2, 1, 50)
        PRW2.BoxC_ERCity = Mid(PRCompany.City, 1, 50)
        If PRState.GetByID(PRCompany.AddrStateID) = False Then
            PRW2.BoxC_ERState = ""
        Else
            PRW2.BoxC_ERState = Mid(PRState.StateAbbrev, 1, 2)
        End If
        PRW2.BoxC_ERZip = Mid(PRCompany.ZipCode, 1, 10)
        
        PRW2.BoxE_EEFirstName = Mid(PREmployee.FirstName, 1, 50)
        PRW2.BoxE_EELastName = Mid(PREmployee.LastName, 1, 50)
        PRW2.BoxE_EEMidInit = Mid(PREmployee.MidInit, 1, 2)
        PRW2.BoxE_EEAddr1 = Mid(PREmployee.Address1, 1, 50)
        PRW2.BoxE_EEAddr2 = Mid(PREmployee.Address2, 1, 50)
        PRW2.BoxE_EECity = Mid(PREmployee.City, 1, 50)
        PRW2.BoxE_EEState = Mid(PREmployee.State, 1, 2)
        PRW2.BoxE_EEZip = Mid(PREmployee.ZipCode, 1, 10)
        
        PRW2.Box13_StatEmp = PREmployee.Statutory
        
        Dim GrossPay As Currency
        GrossPay = 0
        
        Do
            
            GrossPay = GrossPay + PRHist.Gross
            PRW2.Box1_Wages = PRW2.Box1_Wages + PRHist.FWTWage
            PRW2.Box2_FedTax = PRW2.Box2_FedTax + PRHist.FWTTax
            PRW2.Box3_SSWages = PRW2.Box3_SSWages + PRHist.SSWage
            PRW2.Box4_SSTax = PRW2.Box4_SSTax + PRHist.SSTax
            PRW2.Box5_MedWages = PRW2.Box5_MedWages + PRHist.MEDWage
            PRW2.Box6_MedTax = PRW2.Box6_MedTax + PRHist.MedTax
            ' prw2.Box7_SSTips=prw2.Box7_SSTips+prhist.tip
            ' prw2.Box8_AllocTips
            ' prw2.Box9_EIC
            ' prw2.Box10_DCBen
            ' prw2.Box11_NQPlans
                       
            rsState.Find "StateID = " & PRHist.StateID, 0, adSearchForward, 1
            If rsState.EOF Then
                rsState.AddNew
                rsState!StateID = PRHist.StateID
                rsState.Update
            End If
            rsState!StateWage = rsState!StateWage + PRHist.SWTWage
            rsState!StateTax = rsState!StateTax + PRHist.SWTTax
            rsState.Update
                       
            CityWages = CityWages + PRHist.CWTWage
                       
            If PRHist.GetNext = False Then Exit Do
        
        Loop
        
        ' don't process zero gross employees
        ' 2018-12-15 - change to total gross
'MsgBox (PREmployee.EmployeeNumber & vbTab & PREmployee.LastName & vbTab & PREmployee.FirstName & vbTab & PRW2.Box1_Wages)
'        If PRW2.Box1_Wages <= 0 Then
'            GoTo NextEmployee
'        End If
        If GrossPay <= 0 Then
            GoTo NextEmployee
        End If
        
          ' 2020-01-15 - use 1099 flag
        ' - above code commented out for ET???? idk
        If PREmployee.x1099Employee <> 0 Then GoTo NextEmployee
        
        SQLString = "SELECT * FROM PRDist WHERE EmployeeID = " & PREmployee.EmployeeID & " " & _
                    "AND YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM
        
        If PRDist.GetBySQL(SQLString) Then
            Do
                
                '>>>>>> find PRItem <<<<<<<
                ' get the item record
                If PRDist.ItemType <> PREquate.ItemTypeRegPay And PRDist.ItemType <> PREquate.ItemTypeOvtPay Then
                    
                    Box_12_14_Pop PRDist.ItemID, PRDist.Amount
                    
                    If PRItem.GetByID(PRDist.ItemID) = False Then
                        MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRDist.ItemID, vbExclamation
                        GoBack
                    End If
                    
                    If PRItem.UseEmployer = 1 Then
                        If PRItem.GetByID(PRItem.EmployerItemID) = False Then
                            MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRItem.EmployerItemID, vbExclamation
                            GoBack
                        End If
                    End If
                    
                    ' PRItem record found in above sub
                    ' update to tips?
                    If PRItem.Tips = 1 Then
                        PRW2.Box7_SSTips = PRW2.Box7_SSTips + PRDist.Amount
                        PRW2.Box3_SSWages = PRW2.Box3_SSWages - PRDist.Amount
                    End If
                
                End If
                
                ' update the state and city temp tables
                If PRDist.CityID <> 0 Then
                    rsCity.Find "CityID = " & PRDist.CityID, 0, adSearchForward, 1
                    If rsCity.EOF Then
                        rsCity.AddNew
                        rsCity!CityID = PRDist.CityID
                        rsCity!Courtesy = 0
                        rsCity.Update
                    End If
                    rsCity!CityWage = rsCity!CityWage + PRDist.CityWage
                    rsCity!CityTax = rsCity!CityTax + PRDist.CityTax
                    rsCity!SDTax = 0
                    rsCity.Update
                End If
                
                ' courtesy WH
                If PRDist.CourtesyCityTax <> 0 Then
                    ' update the state and city temp tables
                    rsCity.Find "CityID = " & PRDist.CourtesyCityID, 0, adSearchForward, 1
                    If rsCity.EOF Then
                        rsCity.AddNew
                        rsCity!CityID = PRDist.CourtesyCityID
                        rsCity!Courtesy = 1
                        rsCity.Update
                    End If
                    rsCity!CityWage = rsCity!CityWage + PRDist.CityWage
                    rsCity!CityTax = rsCity!CityTax + PRDist.CourtesyCityTax
                    rsCity!SDTax = 0
                    rsCity.Update
                End If
                
                If PRDist.GetNext = False Then Exit Do
            Loop
        End If
            
        SQLString = "SELECT * FROM PRItemHist WHERE EmployeeID = " & PREmployee.EmployeeID & " " & _
                    "AND YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                    " AND ItemType <> " & PREquate.ItemTypeDirDepDed
        If PRItemHist.GetBySQL(SQLString) Then
            Do
                
                Box_12_14_Pop PRItemHist.ItemID, PRItemHist.Amount
                
                
                ' ============================================================================
                ' process school district tax as a local
                If SDTax Then
                
                    If PRItemHist.ItemType = PREquate.ItemTypeSDTax Then
                    
                        If PRItem.GetByID(PRItemHist.ItemID) = False Then
                            MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRItemHist.ItemID, vbExclamation
                            GoBack
                        End If
                        
                        If PRItem.UseEmployer = 1 Then
                            If PRItem.GetByID(PRItem.EmployerItemID) = False Then
                                MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRItem.EmployerItemID, vbExclamation
                                GoBack
                            End If
                        End If
                        
                        SDItemID = PRItem.ItemID
                        
                        If PRItem.ItemType = PREquate.ItemTypeSDTax Then
                            
'                            rsCity.Find "CityID = " & SDItemID, 0, adSearchForward, 1
'                            If rsCity.EOF Then
'                                rsCity.AddNew
'                                rsCity!CityID = SDItemID
'                                rsCity.Update
'                            End If
'                            rsCity!CityWage = CityWages     ' from PRHist above
'                            rsCity!CityTax = rsCity!CityTax + PRItemHist.Amount
'                            rsCity!SDTax = 1
'                            rsCity.Update
                        
                            ' corrected search
                            ' 140108
                            SDFlag = False
                            If rsCity.RecordCount > 0 Then
                                rsCity.MoveFirst
                                Do
                                    If rsCity!SDTax = 1 And rsCity!CityID = SDItemID Then
                                        SDFlag = True
                                        Exit Do
                                    End If
                                    rsCity.MoveNext
                                    If rsCity.EOF Then
                                        Exit Do
                                    End If
                                Loop
                            End If

                            If SDFlag = False Then
                                rsCity.AddNew
                                rsCity!CityID = SDItemID
                                rsCity!SDTax = 1
                                rsCity.Update
                            End If
                            rsCity!CityWage = CityWages     ' from PRHist above
                            rsCity!CityTax = rsCity!CityTax + PRItemHist.Amount
                            rsCity.Update
                        
                        End If
                
                    End If
                
                End If
                ' ============================================================================
                
                If PRItemHist.GetNext = False Then Exit Do
            Loop
        End If

        PRW2.Save (Equate.RecAdd)

        ' update the state and city temp rs data
        If rsCity.RecordCount > 0 Then
            rsCity.MoveFirst
            Do
                
                PRW2City.Clear
                PRW2City.W2ID = PRW2.W2ID
                PRW2City.CityID = rsCity!CityID
                PRW2City.CityWage = rsCity!CityWage
                PRW2City.CityTax = rsCity!CityTax
                
                If rsCity!SDTax = 0 Then
                    If PRCity.GetByID(rsCity!CityID) = False Then
                        MsgBox "CityID not found: " & rsCity!CityID, vbExclamation
                        GoBack
                    End If
                    PRW2City.CityName = Left(PRCity.ShortName, 20)
                    PRW2City.StateID = PRCity.StateID
                    PRW2City.Courtesy = rsCity!Courtesy
                Else        ' handle SD tax entries - get name from PRItem
                    PRW2City.SDTax = 1
                    PRW2City.CityID = rsCity!CityID
                    PRW2City.StateID = 36   ' SD tax for OHIO only !!!
                    If PRItem.GetByID(rsCity!CityID) = True Then
                        PRW2City.CityName = PRItem.Abbreviation
                    End If
                End If
                
                PRW2City.TaxYear = TaxYear
                PRW2City.Save (Equate.RecAdd)
                
                rsCity.Delete
                rsCity.MoveNext
            Loop Until rsCity.EOF
        End If
        
        If rsState.RecordCount > 0 Then
            rsState.MoveFirst
            Do
                PRW2State.Clear
                PRW2State.W2ID = PRW2.W2ID
                PRW2State.StateID = rsState!StateID
                PRW2State.StateWage = rsState!StateWage
                PRW2State.StateTax = rsState!StateTax
                
                ' !!!! handle multi state employers !!!!
                PRW2State.ERStateID = PRCompany.StateID
                If rsState!StateID <> 36 Then
                    If PRState.GetByID(rsState!StateID) Then
                        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOtherStateID & _
                                    " AND Description = '" & PRState.StateAbbrev & "'" & _
                                    " AND Var2 = '" & GLCompany.ID & "'"
                        If PRGlobal.GetBySQL(SQLString) = True Then
                            PRW2State.ERStateID = PRGlobal.Var1
                        End If
                    End If
                End If
                
                PRW2State.TaxYear = TaxYear
                PRW2State.Save (Equate.RecAdd)
                rsState.Delete
                rsState.MoveNext
            Loop Until rsState.EOF
        End If
        
NextEmployee:
        If PREmployee.GetNext = False Then Exit Do
    
    Loop
               
    ' create W3 data
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
            PRGlobal.TypeCode = j
            PRGlobal.Year = TaxYear
            PRGlobal.UserID = PRCompany.CompanyID
            PRGlobal.Save (Equate.RecAdd)
        End If
    
    Next i
    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3E & " AND " & _
                "UserID = " & User.ID

    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeW3E
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
               
    ' clear out zero city records
    SQLString = "DELETE * FROM PRW2City WHERE CityWage = 0 and CityTax = 0"
    cn.Execute SQLString
    
    Unload frmProgress
               
    ' zero out wages by employee tax flag - NOT Box1
               
End Sub

Private Sub Box_12_14_Pop(ByVal ItemID As Long, ByVal Amount As Currency)

    If ItemID = 0 Then Exit Sub

    ' skip direct deposit
    If PRItemHist.ItemType = PREquate.ItemTypeDirDepDed Then Exit Sub
    
    ' get the item record
    If PRItem.GetByID(ItemID) = False Then
        MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRItemHist.ItemID, vbExclamation
        GoBack
    End If
    
    If PRItem.UseEmployer = 1 Then
        If PRItem.GetByID(PRItem.EmployerItemID) = False Then
            MsgBox "Item ID not found: " & PREmployee.EmployeeID & " " & PRItemHist.EmployerItemID, vbExclamation
            GoBack
        End If
    End If
    
    If Amount > 0 Then
        If PRItem.Pension = 1 Then PRW2.Box13_RetirePlan = 1
        If PRItem.SickPay = 1 Then PRW2.Box13_3rdParty = 1
    End If
    
    j = 0
    If PRItem.W2Box12Code <> 0 Then
        For i = 1 To 4
            If Trim(Box12ID(i)) = Trim(PRItem.W2Box12Code) Then
                j = 1
                Exit For
            End If
            If Box12ID(i) = 0 Then
                Box12ID(i) = PRItem.W2Box12Code
                j = 1
                Exit For
            End If
        Next i
        If j = 0 Then
            MsgBox "Box 12 Codes Error: " & PRItem.W2Box12Code, vbExclamation
            GoBack
        End If
        If i = 1 Then
            PRW2.Box12A_ID = PRItem.W2Box12Code
            PRW2.Box12A_Amount = PRW2.Box12A_Amount + Amount
        ElseIf i = 2 Then
            PRW2.Box12B_ID = PRItem.W2Box12Code
            PRW2.Box12B_Amount = PRW2.Box12B_Amount + Amount
        ElseIf i = 3 Then
            PRW2.Box12C_ID = PRItem.W2Box12Code
            PRW2.Box12C_Amount = PRW2.Box12C_Amount + Amount
        ElseIf i = 4 Then
            PRW2.Box12D_ID = PRItem.W2Box12Code
            PRW2.Box12D_Amount = PRW2.Box12D_Amount + Amount
        End If
    End If
    
    j = 0
    If PRItem.W2Box14Code <> 0 Then
        For i = 1 To 4
            If Box14ID(i) = PRItem.W2Box14Code Then
                j = 1
                Exit For
            End If
            If Box14ID(i) = 0 Then
                Box14ID(i) = PRItem.W2Box14Code
                j = 1
                Exit For
            End If
        Next i
        If j = 0 Then
            MsgBox "Box 14 Codes Error: " & PRItem.W2Box14Code & " EE#: " & PRW2.EmployeeNumber, vbExclamation
            GoBack
        End If
        If i = 1 Then
            PRW2.Box14A_ID = PRItem.W2Box14Code
            PRW2.Box14A_Amount = PRW2.Box14A_Amount + Amount
        ElseIf i = 2 Then
            PRW2.Box14B_ID = PRItem.W2Box14Code
            PRW2.Box14B_Amount = PRW2.Box14B_Amount + Amount
        ElseIf i = 3 Then
            PRW2.Box14C_ID = PRItem.W2Box14Code
            PRW2.Box14C_Amount = PRW2.Box14C_Amount + Amount
        ElseIf i = 4 Then
            PRW2.Box14D_ID = PRItem.W2Box14Code
            PRW2.Box14D_Amount = PRW2.Box14D_Amount + Amount
        End If
    End If

End Sub

Private Sub DisplayData()

    ' get the employee record
    If PREmployee.GetByID(PRW2.EmployeeID) = False Then
        MsgBox "EmployeeID not found: " & PRW2.EmployeeID, vbExclamation
        
        ' 2018-12-15 WTF
        Exit Sub
        
        GoBack
    End If
    
    Me.TDBEmpNumber = PRW2.EmployeeNumber
    Me.tdbLineE_FirstName = PRW2.BoxE_EEFirstName
    Me.tdbLineE_LastName = PRW2.BoxE_EELastName
    Me.tdbLineE_MidInit = PRW2.BoxE_EEMidInit
    Me.TDBLinea = PRW2.BoxA_SSNumber
    Me.TDBLineb = PRW2.BoxB_FedID
    
    Me.TDBLineC1 = PRW2.BoxC_ERName
    Me.TDBLineC2 = PRW2.BoxC_ERAddr1
    Me.TDBLineC3 = PRW2.BoxC_ERAddr2
    Me.TDBLineC4 = PRW2.BoxC_ERCity
    Me.TDBLineC5 = PRW2.BoxC_ERState
    Me.TDBLineC6 = PRW2.BoxC_ERZip
    
    'Me.TDBLineE1 = PRW2.BoxE_EEName
    Me.TDBLineE2 = PRW2.BoxE_EEAddr1
    Me.TDBLineE3 = PRW2.BoxE_EEAddr2
    Me.TDBLineE4 = PRW2.BoxE_EECity
    Me.TDBLineE5 = PRW2.BoxE_EEState
    Me.TDBLineE6 = PRW2.BoxE_EEZip
    
    Me.TDBLine1 = PRW2.Box1_Wages
    Me.TDBLine2 = PRW2.Box2_FedTax
    Me.TDBLine3 = PRW2.Box3_SSWages
    Me.TDBLine4 = PRW2.Box4_SSTax
    Me.TDBLine5 = PRW2.Box5_MedWages
    Me.TDBLine6 = PRW2.Box6_MedTax
    Me.TDBLine7 = PRW2.Box7_SSTips
    Me.TDBLine8 = PRW2.Box8_AllocTips
    Me.TDBLine9 = PRW2.Box9_EIC
    Me.TDBLine10 = PRW2.Box10_DCBen
    Me.TDBLine11 = PRW2.Box11_NQPlans
    
    cmbPoint Me.cmbW2Box12a, PRW2.Box12A_ID
    Me.TDBLine12a = PRW2.Box12A_Amount
    
    cmbPoint Me.cmbW2Box12b, PRW2.Box12B_ID
    Me.TDBLine12b = PRW2.Box12B_Amount
    
    cmbPoint Me.cmbW2Box12c, PRW2.Box12C_ID
    Me.TDBLine12c = PRW2.Box12C_Amount
    
    cmbPoint Me.cmbW2Box12d, PRW2.Box12D_ID
    Me.TDBLine12d = PRW2.Box12D_Amount
    
    Me.chkLine13a = PRW2.Box13_StatEmp
    Me.chkLine13b = PRW2.Box13_RetirePlan
    Me.chkLine13c = PRW2.Box13_3rdParty

    cmbPoint Me.cmbW2Box14a, PRW2.Box14A_ID
    Me.tdbLine14a = PRW2.Box14A_Amount
    
    cmbPoint Me.cmbW2Box14b, PRW2.Box14B_ID
    Me.tdbLine14b = PRW2.Box14B_Amount

    cmbPoint Me.cmbW2Box14c, PRW2.Box14C_ID
    Me.tdbLine14c = PRW2.Box14C_Amount

    cmbPoint Me.cmbW2Box14d, PRW2.Box14D_ID
    Me.tdbLine14d = PRW2.Box14D_Amount

    ' state/city grid
    SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & " " & _
                "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
    
    rsInit SQLString, cn, rsS
    SetGrid rsS, Me.fgState
    
    fgState.ColWidth(0) = 0
    fgState.ColWidth(1) = 0
    fgState.ColWidth(2) = 0
    fgState.ColComboList(3) = StateDrop
    fgState.TextMatrix(0, 3) = "State"
    fgState.ColWidth(5) = 1800
    fgState.ColWidth(6) = 1300
    
    SQLString = "SELECT * FROM PRW2City WHERE W2ID = " & PRW2.W2ID & " " & _
                "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
 
    rsInit SQLString, cn, rsC
    SetGrid rsC, fgCity
    
    fgCity.ColWidth(0) = 0
    fgCity.ColWidth(1) = 0
    fgCity.ColWidth(2) = 0
    fgCity.ColWidth(3) = 0
    fgCity.ColWidth(5) = 1800
    fgCity.ColWidth(6) = 1300
    fgCity.ColWidth(7) = 0
    fgCity.ColWidth(8) = 0
    fgCity.ColWidth(9) = 0

    Me.chkVoid = PRW2.Void
    Me.chkSkip = PRW2.Skip

End Sub

Private Sub SaveData()
        
    If LoadFlag = True Then Exit Sub
        
    PRW2.BoxA_SSNumber = Me.TDBLinea
    PRW2.BoxB_FedID = Me.TDBLineb
    
    PRW2.BoxC_ERName = Me.TDBLineC1
    PRW2.BoxC_ERAddr1 = Me.TDBLineC2
    PRW2.BoxC_ERAddr2 = Me.TDBLineC3
    PRW2.BoxC_ERCity = Me.TDBLineC4
    PRW2.BoxC_ERState = Me.TDBLineC5
    PRW2.BoxC_ERZip = Me.TDBLineC6
                                       
    PRW2.BoxE_EEFirstName = Me.tdbLineE_FirstName
    PRW2.BoxE_EELastName = Me.tdbLineE_LastName
    PRW2.BoxE_EEMidInit = Me.tdbLineE_MidInit
    PRW2.BoxE_EEAddr1 = Me.TDBLineE2
    PRW2.BoxE_EEAddr2 = Me.TDBLineE3
    PRW2.BoxE_EECity = Me.TDBLineE4
    PRW2.BoxE_EEState = Me.TDBLineE5
    PRW2.BoxE_EEZip = Me.TDBLineE6
                                      
    PRW2.Box1_Wages = Me.TDBLine1
    PRW2.Box2_FedTax = Me.TDBLine2
    PRW2.Box3_SSWages = Me.TDBLine3
    PRW2.Box4_SSTax = Me.TDBLine4
    PRW2.Box5_MedWages = Me.TDBLine5
    PRW2.Box6_MedTax = Me.TDBLine6
    PRW2.Box7_SSTips = Me.TDBLine7
    PRW2.Box8_AllocTips = Me.TDBLine8
    PRW2.Box9_EIC = Me.TDBLine9
    PRW2.Box10_DCBen = Me.TDBLine10
    PRW2.Box11_NQPlans = Me.TDBLine11

    PRW2.Box12A_Amount = Me.TDBLine12a
    PRW2.Box12A_ID = Me.cmbW2Box12a.ItemData(Me.cmbW2Box12a.ListIndex)

    PRW2.Box12B_Amount = Me.TDBLine12b
    PRW2.Box12B_ID = Me.cmbW2Box12b.ItemData(Me.cmbW2Box12b.ListIndex)

    PRW2.Box12C_Amount = Me.TDBLine12c
    PRW2.Box12C_ID = Me.cmbW2Box12c.ItemData(Me.cmbW2Box12c.ListIndex)

    PRW2.Box12D_Amount = Me.TDBLine12d
    PRW2.Box12D_ID = Me.cmbW2Box12d.ItemData(Me.cmbW2Box12d.ListIndex)

    PRW2.Box13_StatEmp = Me.chkLine13a
    PRW2.Box13_RetirePlan = Me.chkLine13b
    PRW2.Box13_3rdParty = Me.chkLine13c

    PRW2.Box14A_Amount = Me.tdbLine14a
    PRW2.Box14A_ID = Me.cmbW2Box14a.ItemData(Me.cmbW2Box14a.ListIndex)

    PRW2.Box14B_Amount = Me.tdbLine14b
    PRW2.Box14B_ID = Me.cmbW2Box14b.ItemData(Me.cmbW2Box14b.ListIndex)

    PRW2.Box14C_Amount = Me.tdbLine14c
    PRW2.Box14C_ID = Me.cmbW2Box14c.ItemData(Me.cmbW2Box14c.ListIndex)

    PRW2.Box14D_Amount = Me.tdbLine14d
    PRW2.Box14D_ID = Me.cmbW2Box14d.ItemData(Me.cmbW2Box14d.ListIndex)

    If rsS.RecordCount > 0 Then
        rsS.MoveFirst
        Do
            
            SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & " " & _
                        "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
            
            ' 140115 PM - add state id
            SQLString = "SELECT * FROM PRW2State WHERE W2ID = " & PRW2.W2ID & " " & _
                        "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex) & _
                        "AND StateID = " & rsS!StateID
            
            If PRW2State.GetBySQL(SQLString) = False Then
                PRW2State.Clear
                PRW2State.W2ID = PRW2.W2ID
                PRW2State.StateID = rsS!StateID
                PRW2State.TaxYear = Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
                PRW2State.Save (Equate.RecAdd)
            End If
            
            ' ??? 12/29/2010 ???
            ' PRW2State.ERStateID = rss!ERStateID & ""
            
            PRW2State.StateWage = rsS!StateWage
            PRW2State.StateTax = rsS!StateTax
            PRW2State.Save (Equate.RecPut)
            
            rsS.MoveNext
        Loop Until rsS.EOF
    End If
    
    If rsC.RecordCount > 0 Then
        rsC.MoveFirst
        Do
            
            ' 140108 - add SDTax & Courtesy parameters
            SQLString = "SELECT * FROM PRW2City WHERE W2ID = " & PRW2.W2ID & " " & _
                        "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex) & _
                        "AND SDTax = " & rsC!SDTax & _
                        "AND Courtesy = " & rsC!Courtesy
            
            ' 140115 PM - corrected - add CityID
            SQLString = "SELECT * FROM PRW2City WHERE W2ID = " & PRW2.W2ID & " " & _
                        "AND TaxYear = " & Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex) & _
                        "AND CityID = " & rsC!CityID & _
                        "AND SDTax = " & rsC!SDTax & _
                        "AND Courtesy = " & rsC!Courtesy
            
            If PRW2City.GetBySQL(SQLString) = False Then
                PRW2City.Clear
                PRW2City.W2ID = PRW2.W2ID
                PRW2City.TaxYear = Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
                PRW2City.CityID = rsC!CityID
                PRW2City.Courtesy = rsC!Courtesy
                PRW2City.SDTax = rsC!SDTax
                PRW2City.Save (Equate.RecAdd)
            End If
            
            ' wtf 140115 AM - this was commented out before
            ' 141118 - assign state id on prcity
            '          hard code to OH for SD tax
            If rsC!SDTax = 0 Then
                If PRCity.GetByID(PRW2City.CityID) = False Then
                    MsgBox "City NF: " & PRW2City.CityID, vbExclamation
                    GoBack
                End If
                PRW2City.StateID = PRCity.StateID
            Else
                PRW2City.StateID = 36
            End If
            
            PRW2City.CityWage = rsC!CityWage
            PRW2City.CityTax = rsC!CityTax
            PRW2City.Save (Equate.RecPut)
            
            rsC.MoveNext
        Loop Until rsC.EOF
    End If

    PRW2.Void = Me.chkVoid
    PRW2.Skip = Me.chkSkip

    PRW2.Save (Equate.RecPut)

End Sub
Private Sub cmdCreate_Click()
    LoadFlag = True
    CreateData Me.cmbTaxYear
    GetData Me.cmbTaxYear
    DisplayData
    LoadFlag = False
End Sub
Private Sub cmdLoad_Click()
    LoadFlag = True
    GetData Me.cmbTaxYear
    DisplayData
    LoadFlag = False
    Me.Caption = "TAX YEAR: " & Me.cmbTaxYear.Text & "  -  " & PRCompany.Name
End Sub

Private Sub GetData(ByVal TaxYear As Long)
    
    SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear
    
    If PRW2.GetBySQL(SQLString) = False Then
        CreateData Me.cmbTaxYear.ItemData(Me.cmbTaxYear.ListIndex)
    End If

    If Me.optOrderNumber = True Then
        SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear & " " & _
                    "ORDER BY EmployeeNumber"
    Else
        SQLString = "SELECT * FROM PRW2 WHERE TaxYear = " & TaxYear & " " & _
                    "ORDER BY BoxE_EELastName, BoxE_EEFirstName, BoxE_EEMidInit"
    End If
    
    If PRW2.GetBySQL(SQLString) = False Then
        MsgBox "No W2 data?", vbExclamation
        GoBack
    End If
    
    Me.cmbEEList.Clear
    
    Do
        With Me.cmbEEList
            .AddItem PRW2.EmployeeNumber & " " & Trim(PRW2.BoxE_EELastName) & ", " & Trim(PRW2.BoxE_EEFirstName) & _
                 " " & Trim(PRW2.BoxE_EEMidInit)
            .ItemData(.NewIndex) = PRW2.EmployeeNumber
        End With
        If PRW2.GetNext = False Then Exit Do
    Loop

    Me.cmbEEList.ListIndex = 0
    
    PRW2.GetFirst

End Sub

Private Sub cmdNext_Click()
    DisplayChange "NEXT"
End Sub
Private Sub cmdPrev_Click()
    DisplayChange "PREV"
End Sub
Private Sub cmdFirst_Click()
    DisplayChange "FIRST"
End Sub
Private Sub cmdLast_Click()
    DisplayChange "LAST"
End Sub

Private Sub cmbEEList_Click()
    If LoadFlag = True Then Exit Sub
    DisplayChange "CLICK"
End Sub
Private Sub cmdSave_Click()
    SaveData
End Sub

Private Sub DisplayChange(ByVal Action As String)
    
    SaveData
    
    LoadFlag = True
    
    With Me.cmbEEList
        
        If Action = "NEXT" Then
            If .ListIndex = .ListCount - 1 Then Exit Sub
            .ListIndex = .ListIndex + 1
        End If
        
        If Action = "PREV" Then
            If .ListIndex = 0 Then Exit Sub
            .ListIndex = .ListIndex - 1
        End If
        
        If Action = "FIRST" Then .ListIndex = 0
        If Action = "LAST" Then .ListIndex = .ListCount - 1
        
        SQLString = "SELECT * FROM PRW2 Where EmployeeNumber = " & .ItemData(.ListIndex) & _
                    " AND TaxYear = " & Me.cmbTaxYear.Text
        If PRW2.GetBySQL(SQLString) = False Then
            MsgBox "PRW2 File Error: " & .ItemData(.ListIndex), vbExclamation
            GoBack
        End If
    
    End With
    
    DisplayData
    LoadFlag = False

End Sub

Private Sub optOrderName_Click()
    Me.cmbEEList.Clear
    GetData Me.cmbTaxYear
End Sub

Private Sub optOrderNumber_Click()
    Me.cmbEEList.Clear
    GetData Me.cmbTaxYear
End Sub

Private Sub cmdPrint_Click()
    SaveData
    frmW2Print.TaxYear = Me.cmbTaxYear
    frmW2Print.Show vbModal
    LoadFlag = True
    rsS.Close
    rsC.Close
    PRW2.GetFirst
    Me.cmbEEList.ListIndex = 0
    LoadFlag = False
    DisplayData
End Sub


