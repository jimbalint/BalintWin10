VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm941_2011A 
   Caption         =   "Form 941 Rev Jan 2011"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&CALC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   9075
      TabIndex        =   80
      Top             =   40
      Width           =   735
   End
   Begin VB.ComboBox cmbChkDate12 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7440
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox EmployerName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.ComboBox cmbQtr 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.ComboBox cmbYear 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   3960
      TabIndex        =   42
      Top             =   40
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10920
      TabIndex        =   46
      Top             =   30
      Width           =   645
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   9930
      TabIndex        =   44
      Top             =   40
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   120
      TabIndex        =   40
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Form 941"
      TabPicture(0)   =   "frm941_Jan2011.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label38"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line15Check2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line15Check1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fg"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkManualFractions"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Form 941   Page 2"
      TabPicture(1)   =   "frm941_Jan2011.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label48"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label36"
      Tab(1).Control(4)=   "Label42"
      Tab(1).Control(5)=   "Label43"
      Tab(1).Control(6)=   "Label44"
      Tab(1).Control(7)=   "Label39"
      Tab(1).Control(8)=   "Label40"
      Tab(1).Control(9)=   "Label45"
      Tab(1).Control(10)=   "Label46"
      Tab(1).Control(11)=   "Label47"
      Tab(1).Control(12)=   "Label49"
      Tab(1).Control(13)=   "Label50"
      Tab(1).Control(14)=   "Label51"
      Tab(1).Control(15)=   "Label52"
      Tab(1).Control(16)=   "Label53"
      Tab(1).Control(17)=   "Label54"
      Tab(1).Control(18)=   "Label1"
      Tab(1).Control(19)=   "Line17Mo1"
      Tab(1).Control(20)=   "Line17Total"
      Tab(1).Control(21)=   "Line17Mo3"
      Tab(1).Control(22)=   "Line17Mo2"
      Tab(1).Control(23)=   "txtEIN"
      Tab(1).Control(24)=   "Line16"
      Tab(1).Control(25)=   "Line17Check1"
      Tab(1).Control(26)=   "Line17Check2"
      Tab(1).Control(27)=   "Line17Check3"
      Tab(1).Control(28)=   "Part4CheckNo"
      Tab(1).Control(29)=   "Line18Check"
      Tab(1).Control(30)=   "Line18Date"
      Tab(1).Control(31)=   "Line19"
      Tab(1).Control(32)=   "txtName"
      Tab(1).Control(33)=   "Line10Show"
      Tab(1).Control(34)=   "Line17Diff"
      Tab(1).Control(35)=   "Part4CheckYes"
      Tab(1).Control(36)=   "Part4Name"
      Tab(1).Control(37)=   "Part4Pin"
      Tab(1).Control(38)=   "Part4Phone"
      Tab(1).Control(39)=   "txtTradeName"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "Form 941   Pg 2  (Cont'd)"
      TabPicture(2)   =   "frm941_Jan2011.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Part5Phone"
      Tab(2).Control(1)=   "PrepAddr2"
      Tab(2).Control(2)=   "cmbPrepName"
      Tab(2).Control(3)=   "PrepCheck"
      Tab(2).Control(4)=   "Part5NameTitle"
      Tab(2).Control(5)=   "PrepFirm"
      Tab(2).Control(6)=   "PrepAddr1"
      Tab(2).Control(7)=   "PrepEIN"
      Tab(2).Control(8)=   "PrepZip"
      Tab(2).Control(9)=   "PrepDate"
      Tab(2).Control(10)=   "PrepSSN"
      Tab(2).Control(11)=   "Part5Date"
      Tab(2).Control(12)=   "PrepPhone"
      Tab(2).Control(13)=   "Label14"
      Tab(2).Control(14)=   "Label56"
      Tab(2).Control(15)=   "Label55"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Schedule B (Form 941)"
      TabPicture(3)   =   "frm941_Jan2011.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tdbNumHorzNudge"
      Tab(3).Control(1)=   "fgMo1"
      Tab(3).Control(2)=   "BMo1Tax"
      Tab(3).Control(3)=   "BMo2Tax"
      Tab(3).Control(4)=   "BMo3Tax"
      Tab(3).Control(5)=   "BTotalTax"
      Tab(3).Control(6)=   "fgMo2"
      Tab(3).Control(7)=   "fgMo3"
      Tab(3).Control(8)=   "tdbNumVertNudge"
      Tab(3).Control(9)=   "BLine10Show"
      Tab(3).Control(10)=   "BDifference"
      Tab(3).Control(11)=   "lblBMonth3"
      Tab(3).Control(12)=   "lblBMonth2"
      Tab(3).Control(13)=   "lblBMonth1"
      Tab(3).Control(14)=   "Label61"
      Tab(3).Control(15)=   "Label60"
      Tab(3).Control(16)=   "Label59"
      Tab(3).Control(17)=   "Label58"
      Tab(3).ControlCount=   18
      Begin VB.TextBox txtTradeName 
         Height          =   360
         Left            =   -70800
         TabIndex        =   91
         Top             =   735
         Width           =   3255
      End
      Begin VB.CheckBox chkManualFractions 
         Caption         =   "Override Fractions of cents - Line 7"
         Height          =   495
         Left            =   600
         TabIndex        =   90
         Top             =   8640
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   7335
         Left            =   360
         TabIndex        =   89
         Top             =   1080
         Width           =   10695
         _cx             =   18865
         _cy             =   12938
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Timer Timer1 
         Left            =   11040
         Top             =   360
      End
      Begin TDBText6Ctl.TDBText Part4Phone 
         Height          =   375
         Left            =   -73560
         TabIndex        =   24
         Top             =   7080
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "frm941_Jan2011.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":00CE
         Key             =   "frm941_Jan2011.frx":00EC
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
      Begin TDBText6Ctl.TDBText Part4Pin 
         Height          =   375
         Left            =   -69960
         TabIndex        =   25
         Top             =   7080
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   661
         Caption         =   "frm941_Jan2011.frx":0130
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":01CC
         Key             =   "frm941_Jan2011.frx":01EA
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
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   35
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
      Begin TDBText6Ctl.TDBText Part5Phone 
         Height          =   345
         Left            =   -71040
         TabIndex        =   29
         Top             =   1740
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":022E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":028C
         Key             =   "frm941_Jan2011.frx":02AA
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
      Begin TDBText6Ctl.TDBText PrepAddr2 
         Height          =   345
         Left            =   -72920
         TabIndex        =   33
         Top             =   5000
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":02EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":035A
         Key             =   "frm941_Jan2011.frx":0378
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
      Begin VB.ComboBox cmbPrepName 
         Height          =   315
         Left            =   -72920
         TabIndex        =   30
         Top             =   3600
         Width           =   5750
      End
      Begin TDBText6Ctl.TDBText Part4Name 
         Height          =   375
         Left            =   -73560
         TabIndex        =   23
         Top             =   6600
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   661
         Caption         =   "frm941_Jan2011.frx":03BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":042E
         Key             =   "frm941_Jan2011.frx":044C
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
      Begin VB.CheckBox Part4CheckYes 
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   22
         Top             =   6660
         Width           =   735
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Diff 
         Height          =   300
         Left            =   -66240
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":0490
         Caption         =   "frm941_Jan2011.frx":04B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0516
         Keys            =   "frm941_Jan2011.frx":0534
         Spin            =   "frm941_Jan2011.frx":057E
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
      Begin TDBNumber6Ctl.TDBNumber Line10Show 
         Height          =   300
         Left            =   -66240
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":05A6
         Caption         =   "frm941_Jan2011.frx":05C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0630
         Keys            =   "frm941_Jan2011.frx":064E
         Spin            =   "frm941_Jan2011.frx":0698
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
      Begin TDBNumber6Ctl.TDBNumber tdbNumHorzNudge 
         Height          =   615
         Left            =   -65640
         TabIndex        =   78
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":06C0
         Caption         =   "frm941_Jan2011.frx":06E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0760
         Keys            =   "frm941_Jan2011.frx":077E
         Spin            =   "frm941_Jan2011.frx":07C8
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
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74880
         TabIndex        =   9
         Top             =   735
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   77
         Top             =   -480
         Width           =   3855
      End
      Begin VB.CheckBox Line15Check1 
         Caption         =   "Apply to next return."
         Height          =   255
         Left            =   9720
         TabIndex        =   4
         Top             =   8640
         Width           =   1750
      End
      Begin VSFlex8Ctl.VSFlexGrid fgMo1 
         Height          =   2325
         Left            =   -73800
         TabIndex        =   6
         Top             =   660
         Width           =   8055
         _cx             =   14208
         _cy             =   4101
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm941_Jan2011.frx":07F0
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.CheckBox PrepCheck 
         Caption         =   "Check if you are self-employed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -72960
         TabIndex        =   39
         Top             =   5460
         Width           =   3495
      End
      Begin VB.CheckBox Line19 
         Caption         =   "Check here."
         Height          =   255
         Left            =   -65280
         TabIndex        =   21
         Top             =   5620
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate Line18Date 
         Height          =   285
         Left            =   -74400
         TabIndex        =   20
         Top             =   5220
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   503
         Calendar        =   "frm941_Jan2011.frx":08CA
         Caption         =   "frm941_Jan2011.frx":09E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0A8C
         Keys            =   "frm941_Jan2011.frx":0AAA
         Spin            =   "frm941_Jan2011.frx":0B08
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.CheckBox Line18Check 
         Caption         =   "Check here"
         Height          =   255
         Left            =   -65280
         TabIndex        =   19
         Top             =   4950
         Width           =   1575
      End
      Begin VB.CheckBox Part4CheckNo 
         Caption         =   "No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   26
         Top             =   7440
         Width           =   615
      End
      Begin VB.CheckBox Line17Check3 
         Caption         =   " You were a semiweekly schedule depositor for any part of this quarter.  Fill out Schedule B (Form 941):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   18
         Top             =   4080
         Width           =   9855
      End
      Begin VB.CheckBox Line17Check2 
         Caption         =   " You were a monthly schedule depositor for the entire quarter.  Fill out your tax liability for each month."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73320
         TabIndex        =   13
         Top             =   2040
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox Line17Check1 
         Caption         =   " Line 10 is less than $2,500.  Go to Part 3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73320
         TabIndex        =   12
         Top             =   1755
         Width           =   7215
      End
      Begin TDBText6Ctl.TDBText Line16 
         Height          =   375
         Left            =   -74520
         TabIndex        =   11
         Top             =   1185
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "frm941_Jan2011.frx":0B30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0B9C
         Key             =   "frm941_Jan2011.frx":0BBA
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
      Begin VB.TextBox txtEIN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -67440
         TabIndex        =   10
         Top             =   735
         Width           =   3855
      End
      Begin VB.CheckBox Line15Check2 
         Caption         =   "Send refund check."
         Height          =   255
         Left            =   9720
         TabIndex        =   5
         Top             =   8880
         Width           =   1750
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo2 
         Height          =   300
         Left            =   -71760
         TabIndex        =   15
         Top             =   2940
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":0BFE
         Caption         =   "frm941_Jan2011.frx":0C1E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0C88
         Keys            =   "frm941_Jan2011.frx":0CA6
         Spin            =   "frm941_Jan2011.frx":0CF0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   71499777
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo3 
         Height          =   300
         Left            =   -71760
         TabIndex        =   16
         Top             =   3300
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":0D18
         Caption         =   "frm941_Jan2011.frx":0D38
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0DA2
         Keys            =   "frm941_Jan2011.frx":0DC0
         Spin            =   "frm941_Jan2011.frx":0E0A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   25034753
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Total 
         Height          =   300
         Left            =   -73245
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3660
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8184
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":0E32
         Caption         =   "frm941_Jan2011.frx":0E52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0EE4
         Keys            =   "frm941_Jan2011.frx":0F02
         Spin            =   "frm941_Jan2011.frx":0F4C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   27590657
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo1 
         Height          =   300
         Left            =   -71760
         TabIndex        =   14
         Top             =   2580
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":0F74
         Caption         =   "frm941_Jan2011.frx":0F94
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":0FFE
         Keys            =   "frm941_Jan2011.frx":101C
         Spin            =   "frm941_Jan2011.frx":1066
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   28377089
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText Part5NameTitle 
         Height          =   345
         Left            =   -74880
         TabIndex        =   27
         Top             =   1260
         Width           =   11055
         _Version        =   65536
         _ExtentX        =   19500
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":108E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1112
         Key             =   "frm941_Jan2011.frx":1130
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
      Begin TDBText6Ctl.TDBText PrepFirm 
         Height          =   345
         Left            =   -74760
         TabIndex        =   31
         Top             =   4020
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":1174
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":11E6
         Key             =   "frm941_Jan2011.frx":1204
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
      Begin TDBText6Ctl.TDBText PrepAddr1 
         Height          =   345
         Left            =   -74760
         TabIndex        =   32
         Top             =   4500
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":1248
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":12B2
         Key             =   "frm941_Jan2011.frx":12D0
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
      Begin TDBText6Ctl.TDBText PrepEIN 
         Height          =   345
         Left            =   -66360
         TabIndex        =   35
         Top             =   4020
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":1314
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1376
         Key             =   "frm941_Jan2011.frx":1394
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
      Begin TDBText6Ctl.TDBText PrepZip 
         Height          =   345
         Left            =   -66840
         TabIndex        =   36
         Top             =   4500
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":13D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1444
         Key             =   "frm941_Jan2011.frx":1462
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
      Begin TDBDate6Ctl.TDBDate PrepDate 
         Height          =   345
         Left            =   -66480
         TabIndex        =   38
         Top             =   5460
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   609
         Calendar        =   "frm941_Jan2011.frx":14A6
         Caption         =   "frm941_Jan2011.frx":15BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1622
         Keys            =   "frm941_Jan2011.frx":1640
         Spin            =   "frm941_Jan2011.frx":169E
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
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "  /  /    "
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText PrepSSN 
         Height          =   345
         Left            =   -66960
         TabIndex        =   37
         Top             =   4980
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":16C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1732
         Key             =   "frm941_Jan2011.frx":1750
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
      Begin TDBDate6Ctl.TDBDate Part5Date 
         Height          =   345
         Left            =   -74880
         TabIndex        =   28
         Top             =   1740
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   609
         Calendar        =   "frm941_Jan2011.frx":1794
         Caption         =   "frm941_Jan2011.frx":18AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1910
         Keys            =   "frm941_Jan2011.frx":192E
         Spin            =   "frm941_Jan2011.frx":198C
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber BMo1Tax 
         Height          =   615
         Left            =   -65640
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":19B4
         Caption         =   "frm941_Jan2011.frx":19D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1A5E
         Keys            =   "frm941_Jan2011.frx":1A7C
         Spin            =   "frm941_Jan2011.frx":1AC6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   52428801
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber BMo2Tax 
         Height          =   615
         Left            =   -65640
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":1AEE
         Caption         =   "frm941_Jan2011.frx":1B0E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1B98
         Keys            =   "frm941_Jan2011.frx":1BB6
         Spin            =   "frm941_Jan2011.frx":1C00
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;;"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   9999999999
         MinValue        =   -99999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   1
         ValueVT         =   52428801
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber BMo3Tax 
         Height          =   615
         Left            =   -65640
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6840
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":1C28
         Caption         =   "frm941_Jan2011.frx":1C48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1CD2
         Keys            =   "frm941_Jan2011.frx":1CF0
         Spin            =   "frm941_Jan2011.frx":1D3A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;"
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
         ShowContextMenu =   1
         ValueVT         =   52428801
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber BTotalTax 
         Height          =   615
         Left            =   -65640
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   7560
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":1D62
         Caption         =   "frm941_Jan2011.frx":1D82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":1E10
         Keys            =   "frm941_Jan2011.frx":1E2E
         Spin            =   "frm941_Jan2011.frx":1E78
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;"
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
         ShowContextMenu =   1
         ValueVT         =   61472769
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VSFlex8Ctl.VSFlexGrid fgMo2 
         Height          =   2325
         Left            =   -73800
         TabIndex        =   7
         Top             =   3240
         Width           =   8055
         _cx             =   14208
         _cy             =   4101
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm941_Jan2011.frx":1EA0
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin VSFlex8Ctl.VSFlexGrid fgMo3 
         Height          =   2340
         Left            =   -73800
         TabIndex        =   8
         Top             =   5895
         Width           =   8055
         _cx             =   14208
         _cy             =   4128
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm941_Jan2011.frx":1F7A
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin TDBNumber6Ctl.TDBNumber tdbNumVertNudge 
         Height          =   615
         Left            =   -65640
         TabIndex        =   79
         Top             =   1320
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941_Jan2011.frx":2054
         Caption         =   "frm941_Jan2011.frx":2074
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":20F4
         Keys            =   "frm941_Jan2011.frx":2112
         Spin            =   "frm941_Jan2011.frx":215C
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
      Begin TDBNumber6Ctl.TDBNumber BLine10Show 
         Height          =   300
         Left            =   -66600
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   8205
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":2184
         Caption         =   "frm941_Jan2011.frx":21A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":220E
         Keys            =   "frm941_Jan2011.frx":222C
         Spin            =   "frm941_Jan2011.frx":2276
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;"
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
         ShowContextMenu =   1
         ValueVT         =   61472769
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber BDifference 
         Height          =   300
         Left            =   -70560
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   8520
         Width           =   6945
         _Version        =   65536
         _ExtentX        =   12259
         _ExtentY        =   529
         Calculator      =   "frm941_Jan2011.frx":229E
         Caption         =   "frm941_Jan2011.frx":22BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":2388
         Keys            =   "frm941_Jan2011.frx":23A6
         Spin            =   "frm941_Jan2011.frx":23F0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##;"
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
         ShowContextMenu =   1
         ValueVT         =   61472769
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText PrepPhone 
         Height          =   345
         Left            =   -66600
         TabIndex        =   34
         Top             =   3600
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   609
         Caption         =   "frm941_Jan2011.frx":2418
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jan2011.frx":2476
         Key             =   "frm941_Jan2011.frx":2494
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
      Begin VB.Label Label1 
         Caption         =   "Trade Name"
         Height          =   375
         Left            =   -70800
         TabIndex        =   92
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Preparer's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74760
         TabIndex        =   88
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label lblBMonth3 
         Caption         =   "month3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   87
         Top             =   6180
         Width           =   950
      End
      Begin VB.Label lblBMonth2 
         Caption         =   "month2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   86
         Top             =   3530
         Width           =   950
      End
      Begin VB.Label lblBMonth1 
         Caption         =   "month1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
         Top             =   945
         Width           =   950
      End
      Begin VB.Label Label5 
         Caption         =   "One"
         Height          =   225
         Left            =   9120
         TabIndex        =   76
         Top             =   8820
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "Month 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   5870
         Width           =   735
      End
      Begin VB.Label Label60 
         Caption         =   "Month 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   3230
         Width           =   735
      End
      Begin VB.Label Label59 
         Caption         =   "Month 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   73
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label58 
         Caption         =   "Report of Tax Liability for Semiweekly Schedule Depositors"
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
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label Label56 
         Caption         =   " PAID Preparers use Only"
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
         Left            =   -74880
         TabIndex        =   71
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label Label55 
         Caption         =   "Part 5:  Signature.  You MUST complete all pages of Form 941 and SIGN it."
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
         Left            =   -74880
         TabIndex        =   70
         Top             =   780
         Width           =   8175
      End
      Begin VB.Label Label54 
         Caption         =   "Do you want to allow an employee, a paid tax preparer, or another person to discuss this return with the IRS?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   69
         Top             =   6300
         Width           =   9615
      End
      Begin VB.Label Label53 
         Caption         =   "Part 4:  Third-Party Designee?"
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
         Left            =   -74880
         TabIndex        =   68
         Top             =   5940
         Width           =   4455
      End
      Begin VB.Label Label52 
         Caption         =   "Part 3:  This Section applies to your business.  If a question does NOT apply to your business, leave it blank."
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
         Left            =   -74880
         TabIndex        =   67
         Top             =   4620
         Width           =   11295
      End
      Begin VB.Label Label51 
         Caption         =   "If you are a seasonal employer and you do not have to file a return for every quarter of the year  . . . . . . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   66
         Top             =   5620
         Width           =   9375
      End
      Begin VB.Label Label50 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   5645
         Width           =   255
      End
      Begin VB.Label Label49 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   4980
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   $"frm941_Jan2011.frx":24D8
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   63
         Top             =   4930
         Width           =   9375
      End
      Begin VB.Label Label46 
         Caption         =   "Report of Tax Liability for Semiweekly Schedule Depositors, and attach it to this form."
         Height          =   255
         Left            =   -73005
         TabIndex        =   62
         Top             =   4335
         Width           =   6255
      End
      Begin VB.Label Label45 
         Caption         =   "Total must equal line 10."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   61
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label40 
         Caption         =   "  Tax Liability:      "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73080
         TabIndex        =   60
         Top             =   2580
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   1770
         Width           =   255
      End
      Begin VB.Label Label44 
         Caption         =   "Then go to Part 3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73005
         TabIndex        =   58
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Label43 
         Caption         =   "Check one:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   48
         Top             =   1740
         Width           =   1335
      End
      Begin VB.Label Label42 
         Caption         =   "in multiple states."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73800
         TabIndex        =   57
         Top             =   1390
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "Enter the state abbreviation for the state where you made your deposits OR enter ""MU"" if you made your deposits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73800
         TabIndex        =   56
         Top             =   1185
         Width           =   9975
      End
      Begin VB.Label Label17 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Employer Identification number (EIN)"
         Height          =   255
         Left            =   -67320
         TabIndex        =   54
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label48 
         Caption         =   "Name (not your trade name)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label38 
         Caption         =   "Check"
         Height          =   225
         Left            =   9120
         TabIndex        =   52
         Top             =   8640
         Width           =   495
      End
      Begin VB.Label Label33 
         Height          =   180
         Left            =   1395
         TabIndex        =   51
         Top             =   7230
         Width           =   7035
      End
   End
   Begin VB.Label Label10 
      Caption         =   "QTR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   50
      Top             =   80
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4800
      TabIndex        =   49
      Top             =   80
      Width           =   495
   End
End
Attribute VB_Name = "frm941_2011A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AlphaCheckLine4 As String
Public AlphaCheckLine15a As String
Public AlphaCheckLine15b As String
Public AlphaCheckLine17a As String
Public AlphaCheckLine17b As String
Public AlphaCheckLine17c As String
Public AlphaCheckLine18 As String
Public AlphaCheckLine19 As String
Public AlphaCheckPart4Yes As String
Public AlphaCheckPart4No As String
Public AlphaCheckPart5 As String
Public TotTaxLiability As Currency
Public Part4ID, Part5ID, PaidPrepID As Long

Dim StartYM, EndYM As Long
Dim LoadFlag As Boolean
Dim SSTax, MedTax As Currency
    
Dim rsTips As New ADODB.Recordset
Dim rsERTips As New ADODB.Recordset
Dim i, j, k As Long

Dim rsCol As New ADODB.Recordset
Dim RowNum, ColNum As Integer
Dim ERSSTax, MatchSS, MatchSSTotal As Currency

Private Sub Form_Load()
    
    LoadFlag = True
    
    ' setup the processing grid
    Grid941
    
    ' ************************************************************
    ' gather a list of RItems that are tips
    frmProgress.Show
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.lblMsg2 = "Gathering Employee Info ..."
    frmProgress.Refresh
    
    rsTips.CursorLocation = adUseClient
    rsTips.Fields.Append "EmployeeID", adDouble
    rsTips.Fields.Append "ItemID", adDouble
    rsTips.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID <> 0 " & _
                " AND PRItem.ItemType = " & PREquate.ItemTypeOE
    If PRItem.GetBySQL(SQLString) Then
        Do
            ' use the employer defn?
            If PRItem.UseEmployer = 1 Then
                SQLString = "SELECT * FROM PRItem WHERE ItemID = " & PRItem.EmployerItemID
                rsInit SQLString, cn, rsERTips
                If rsERTips.RecordCount > 0 Then
                    If rsERTips!Tips = 1 Then
                        rsTips.AddNew
                        rsTips!EmployeeID = PRItem.EmployeeID
                        rsTips!ItemID = PRItem.ItemID
                        rsTips.Update
                    End If
                End If
            Else
                If PRItem.Tips = 1 Then
                    rsTips.AddNew
                    rsTips!EmployeeID = PRItem.EmployeeID
                    rsTips!ItemID = PRItem.ItemID
                    rsTips.Update
                End If
            End If
            If PRItem.GetNext = False Then Exit Do
        Loop
    End If
    
    frmProgress.Hide
    
    tdbAmountSet Me.Line17Mo1
    tdbAmountSet Me.Line17Mo2
    tdbAmountSet Me.Line17Mo3
    tdbAmountSet Me.Line17Total
    tdbAmountSet Me.Line10Show
    tdbAmountSet Me.Line17Diff
    tdbAmountSet Me.BMo1Tax
    tdbAmountSet Me.BMo2Tax
    tdbAmountSet Me.BMo3Tax
    tdbAmountSet Me.BTotalTax
    tdbAmountSet Me.BLine10Show
    tdbAmountSet Me.BDifference
    
    tdbTextSet Me.Part4Name
    tdbTextSet Me.Part4Phone
    tdbTextSet Me.Part4Pin
    tdbTextSet Me.Part5NameTitle
    tdbTextSet Me.Part5Phone
    tdbTextSet Me.PrepFirm
    tdbTextSet Me.PrepAddr1
    tdbTextSet Me.PrepAddr2
    tdbTextSet Me.PrepPhone
    tdbTextSet Me.PrepEIN
    tdbTextSet Me.PrepSSN
    tdbTextSet Me.PrepZip
    
    Line17Total.ReadOnly = True
    Line17Diff.ReadOnly = True
    
    Me.Part4CheckNo = 1
    Me.Part4CheckYes = 0
    
    tdbDateSet Me.Part5Date, Int(Now())
    
    Me.cmbChkDate12.ToolTipText = "Check Date for EE Count - Line1"
    
    ' init the year qtr combo
    ' If cmbYrQtrSet(Me.cmbYear, Me.cmbQtr) = False Then GoBack
        
    ' init year / qtr combo even if no data exists
    With Me.cmbQtr
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    With Me.cmbYear
        j = Year(Now()) + 1
        For i = 1 To 5
            .AddItem j
            j = j - 1
        Next i
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
    
    ' *** stuff 1st qtr
     Me.cmbQtr.ListIndex = 0                                     ''''''''''''  TAKE OUT  '''''''''''''
    
    LoadFlag = False
    
    ' pop ChkDate12 combo
    PopChkDate12
    
    ' load the data
    Get941Data
    
    Me.AlphaCheckLine4 = " "
    Me.AlphaCheckLine15a = " "
    Me.AlphaCheckLine15b = " "

    EmployerName = UCase(PRCompany.Name)
    Me.txtEIN = PRCompany.FederalID
    Me.txtName = UCase(PRCompany.Name)
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
       MsgBox "WARNING: Employer State Not Filled In", vbExclamation, "Form 941 Entry"
       Me.Line16 = ""
    Else
        Me.Line16 = PRState.StateAbbrev
    End If
    CurrYear = Year(Now())

    SetNudge Me.tdbNumHorzNudge
    SetNudge Me.tdbNumVertNudge
    GetNudge User.ID, "941_2010Apr"
    Me.tdbNumHorzNudge = nNull(HorzNudge)
    Me.tdbNumVertNudge = nNull(VertNudge)

    Me.KeyPreview = True
    Me.SSTab1.Tab = 0
    Me.Timer1.Interval = 1
    
End Sub

Private Sub Grid941()

    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    rsCol.Open , , adOpenDynamic, adLockOptimistic
    
    ' columns for the matrix
    AddCol "Descr", "Descr", 5000
    
    i = 1800
    AddCol "Amt1", "Amt1", i
    AddCol "Amt2", "Amt2", i
    AddCol "Amt3", "Amt3", i
    
    i = 0
    AddCol "Edit1", "Edit1", i, adBoolean
    AddCol "Edit2", "Edit2", i, adBoolean
    AddCol "Edit3", "Edit3", i, adBoolean
    AddCol "Show1", "Show1", i, adBoolean
    AddCol "Show2", "Show2", i, adBoolean
    AddCol "Show3", "Show3", i, adBoolean

    RowNum = 0
    
    With Me.fg
    
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 10
        .Rows = 99
        .Editable = flexEDKbdMouse
        
        i = 0
        rsCol.MoveFirst
        Do
            .ColWidth(i) = rsCol!Width
            .ColData(i) = rsCol!Abbrev
            If rsCol!dataType <> 0 Then
                .ColDataType(i) = rsCol!dataType
            End If
            If rsCol!Format <> 0 Then
                .ColFormat(i) = rsCol!Format
            End If
            i = i + 1
            rsCol.MoveNext
        Loop Until rsCol.EOF
    
        AddRow " 1 ) Number of employees", False, False, True, False, False, True
        AddRow " 2 ) Wages, tips and other compensation", False, False, True, False, False, True
        AddRow " 3 ) Income tax withheld", False, False, True, False, False, True
        
        AddRow " 5a) Taxable social security wages", True, True, False, True, True, False
        AddRow " 5b) Taxable social security tips", True, True, False, True, True, False
        AddRow " 5c) Taxable Medicare wages & tips", True, True, False, True, True, False
        
        AddRow " 5d) Add Col 2 5a,Col 2 5b, Col 2 5c", False, False, False, False, False, True
        AddRow " 5e) Sec 3121(q) Notice and Demand-Tax due on unreported tips", False, False, True, False, False, True
        
        AddRow " 6e) Total taxes before adjustments", False, False, False, False, False, True
        AddRow " 7 ) Current qtr adj for fractions of cents", False, True, False, False, True, False
        AddRow " 8 ) Current qtr adj for sick pay", False, True, False, False, True, False
        AddRow " 9 ) Current qtr adj for tips and group-term life insurance", False, True, False, False, True, False
        
        AddRow "10 ) Total taxes after adjustments", False, False, False, False, False, True
        AddRow "11 ) Total deposits, incl prior qtr overpay", False, False, True, False, False, True
        
        AddRow "12a) COBRA premium asst payments", True, False, False, True, False, False
        AddRow "12b) Number of COBRA provided", True, False, False, True, False, False
        
        AddRow "13 ) Add lines 11 and 12a", False, False, False, False, False, True
        AddRow "14 ) Balance Due", False, False, False, False, False, True
        AddRow "15 ) Overpayment", False, False, False, False, True, False
        
        .Rows = RowNum
    
        .ColFormat(GetCol("Amt1")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt2")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt3")) = "##,###,##0.00-"
    
        ' color the grid
        For i = 1 To RowNum
            For j = 1 To 3
                If j = 1 Then k = GetCol("Show1")
                If j = 2 Then k = GetCol("Show2")
                If j = 3 Then k = GetCol("Show3")
                If .TextMatrix(i - 1, k) = "False" Then
                    .Select i - 1, k - 6
                    .CellBackColor = RGB(192, 192, 192)
                    .CellBackColor = RGB(100, 100, 100)
                End If
            Next j
        Next i
    
        ' set the amounts to zero
        Set941Val " 1 )", 3, 0
        Set941Val " 2 )", 3, 0
        Set941Val " 3 )", 3, 0
        Set941Val " 5a)", 1, 0
        Set941Val " 5a)", 2, 0
        Set941Val " 5b)", 1, 0
        Set941Val " 5b)", 2, 0
        Set941Val " 5c)", 1, 0
        Set941Val " 5c)", 2, 0
        Set941Val " 5d)", 3, 0
        Set941Val " 5e)", 3, 0
        Set941Val " 6e)", 3, 0
        Set941Val " 7 )", 2, 0
        Set941Val " 8 )", 2, 0
        Set941Val " 9 )", 2, 0
        Set941Val "10 )", 3, 0
        Set941Val "11 )", 3, 0
        Set941Val "12a)", 1, 0
        Set941Val "12b)", 1, 0
        Set941Val "13 )", 3, 0
        Set941Val "14 )", 3, 0
        Set941Val "15 )", 2, 0
    
    End With

End Sub

Private Sub BInitGrid(ByRef fg As VSFlexGrid)
        
Dim i, j As Integer
Dim k, m As Integer
        
    With fg
        
        fg.FixedCols = 0                   ' see all cols selected by SQL
        fg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
        fg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
    
        fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
        fg.TabBehavior = flexTabCells                       ' tab moves between cells
        fg.AllowSelection = False                          ' don't allow selection of ranges of cells
                
        .Cols = 8
        .Rows = 9

        .ColFormat(-1) = "#'"
        
        For i = 1 To 8
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = 0
            .TextMatrix(i, 2) = i + 8
            .TextMatrix(i, 3) = 0
            .TextMatrix(i, 4) = i + 16
            .TextMatrix(i, 5) = 0
            .TextMatrix(i, 6) = i + 24
            .TextMatrix(i, 7) = 0
        Next i
    
        .ColFormat(1) = "$###,###,##0.00"
        .ColFormat(3) = "$###,###,##0.00"
        .ColFormat(5) = "$###,###,##0.00"
        .ColFormat(7) = "$###,###,##0.00"


    For k = 0 To 7 Step 2
        .ColWidth(k) = 400
        .TextMatrix(0, k) = "Day"
    Next k
            
    For m = 1 To 8 Step 2
        .TextMatrix(0, m) = "Tax Amt"
    Next m
    
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

'Private Sub Line4_click()
'
'    If Me.Line4 = 1 Then
'        Me.AlphaCheckLine4 = "X"
'        Me.Line5a.TabStop = False
'        Me.Line5b.TabStop = False
'        Me.Line5c.TabStop = False
'        Me.Line5d.TabStop = False
'
'        If IsNull(Me.Line3) Then
'            Me.Line3 = 0
'        End If
'
'        Me.Line6e = Me.Line3 + Me.Line5d - Me.Line6d
'    Else
'        Me.AlphaCheckLine4 = ""
'    End If
'
'End Sub

Private Sub Line15Check1_Click()
    If Line15Check1 = 1 And Line15Check2 = 1 Then
        MsgBox "Please check EITHER Apply to Next Return or Send Refund", vbCritical, "Form 941"
    ElseIf Line15Check1 = 1 Then
        Me.AlphaCheckLine15a = "X"
    Else
        Me.AlphaCheckLine15a = ""
    End If
End Sub

Private Sub Line15Check2_Click()
    If Line15Check1 = 1 And Line15Check2 = 1 Then
        MsgBox "Please check EITHER Apply to Next Return or Send Refund", vbCritical, "Form 941"
    ElseIf Line15Check2 = 1 Then
        Me.AlphaCheckLine15b = "X"
    Else
        Me.AlphaCheckLine15b = ""
    End If
End Sub

Private Sub Line17Check1_Click()
    If Line17Check1 = 1 Then
        Me.AlphaCheckLine17a = "X"
        Me.Line17Check2 = 0
        Me.Line17Check3 = 0
    Else
        Me.AlphaCheckLine17a = ""
    End If
    
    Me.Line17Check2.TabStop = False
    Me.Line17Mo1.TabStop = False
    Me.Line17Mo2.TabStop = False
    Me.Line17Mo3.TabStop = False
    Me.Line17Check3.TabStop = False
    
    Me.Line17Mo1.Visible = False
    Me.Line17Mo2.Visible = False
    Me.Line17Mo3.Visible = False
    Me.Line17Total.Visible = False
    Me.Line10Show.Visible = False
    Me.Line17Diff.Visible = False
    Me.Label40.Visible = False
    Me.Label45.Visible = False

End Sub

Private Sub Line17Check2_Click()
    
    If Line17Check2 = 1 Then
        Me.AlphaCheckLine17b = "X"
        Me.Line17Check1 = 0
        Me.Line17Check3 = 0
    Else
        Me.AlphaCheckLine17b = ""
        Me.Line17Mo1.TabStop = False
        Me.Line17Mo2.TabStop = False
        Me.Line17Mo3.TabStop = False
    End If
    Me.Line17Mo1.Visible = True
    Me.Line17Mo2.Visible = True
    Me.Line17Mo3.Visible = True
    Me.Line17Total.Visible = True
    Me.Line10Show.Visible = True
    Me.Line17Diff.Visible = True
    Me.Label40.Visible = True
    Me.Label45.Visible = True


End Sub

Private Sub Line17Check3_Click()
    If Line17Check3 = 1 Then
        Me.AlphaCheckLine17c = "X"
        Me.Line17Check1 = 0
        Me.Line17Check2 = 0
    Else
        Me.AlphaCheckLine17c = ""
    End If
    Me.Line17Mo1.Visible = False
    Me.Line17Mo2.Visible = False
    Me.Line17Mo3.Visible = False
    Me.Line17Total.Visible = False
    Me.Line10Show.Visible = False
    Me.Line17Diff.Visible = False
    Me.Label40.Visible = False
    Me.Label45.Visible = False

End Sub

Private Sub Line18Check_Click()
    If Line18Check = 1 Then
        Me.AlphaCheckLine18 = "X"
        Line18Date.Visible = True
        Line18Date = Int(Now())
    Else
        Me.AlphaCheckLine18 = ""
        Line18Date.Visible = False
    End If
End Sub

Private Sub Line19_Click()
    If Line19 = 1 Then
        Me.AlphaCheckLine19 = "X"
    Else
        Me.AlphaCheckLine19 = ""
    End If
End Sub

Private Sub Part4CheckYes_Click()
    If Me.Part4CheckYes = 1 Then Me.Part4CheckNo = 0
End Sub

Private Sub Part4Name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Part4CheckNo_Click()
    If Me.Part4CheckNo = 1 Then Me.Part4CheckYes = 0
End Sub

Private Sub prepCheck_Click()
    If PrepCheck = 1 Then
        Me.AlphaCheckPart5 = "X"
    Else
        Me.AlphaCheckPart5 = ""
    End If
End Sub

Private Sub cmdExit_Click()
   GoBack
End Sub

Private Sub cmdPrint_Click()
    PrtInit ("Port")
    SetFont 10, Equate.Portrait

    HorzNudge = Me.tdbNumHorzNudge.Value
    VertNudge = Me.tdbNumVertNudge.Value
    
    SaveNudge User.ID, "941_2010Apr"
    
    Me.KeyPreview = True
    Form941A2011Jan
        
    If Me.Line17Check3 = 1 Then
        FormFeed
        
        VertNudge = VertNudge + 2
        HorzNudge = HorzNudge + 0
        
        Form941BHdr Me, Me.cmbYear.Text
        
'        Form941BPrint 2300, Me.fgMo1, BMo1Tax
'        Form941BPrint 6400, Me.fgMo2, BMo2Tax
'        Form941BPrint 10500, Me.fgMo3, BMo3Tax
    
        ' twk for eagl 07/02/10
        Form941BPrint 2270, Me.fgMo1, BMo1Tax
        Form941BPrint 6370, Me.fgMo2, BMo2Tax
        Form941BPrint 10470, Me.fgMo3, BMo3Tax
    
    End If

    Prvw.vsp.EndDoc
    Prvw.Show vbModal
End Sub

Private Sub fgMo1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    BGridUpdate fgMo1, Me.BMo1Tax
    BLine10Show = BLine10Show
End Sub
Private Sub fgMo1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If fgMo1.Col = 0 Or fgMo1.Col = 2 Or fgMo1.Col = 4 Or fgMo1.Col = 6 Then
        Cancel = True
    End If
End Sub

Private Sub fgMo2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    BGridUpdate fgMo2, Me.BMo2Tax
End Sub

Private Sub fgMo2_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If fgMo2.Col = 0 Or fgMo2.Col = 2 Or fgMo2.Col = 4 Or fgMo2.Col = 6 Then
        Cancel = True
    End If
End Sub

Private Sub fgMo3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    BGridUpdate fgMo3, Me.BMo3Tax
End Sub

Private Sub fgMo3_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If fgMo3.Col = 0 Or fgMo3.Col = 2 Or fgMo3.Col = 4 Or fgMo3.Col = 6 Then
        Cancel = True
    End If
End Sub

Private Sub Get941Data()

Dim Month1, Month2 As Byte
Dim TaxLiab As Currency

    If LoadFlag = True Then Exit Sub
                
    ' clear the grid values
    For i = 1 To fg.Rows - 1
        For j = 1 To 3
            If j = 1 Then k = GetCol("Show1")
            If j = 2 Then k = GetCol("Show2")
            If j = 3 Then k = GetCol("Show3")
            If fg.TextMatrix(i, k) = "True" Then
                fg.TextMatrix(i, j) = "0.00"
            End If
        Next j
    Next i
                
    SSTax = 0
    MatchSS = 0
    MedTax = 0
    MatchSSTotal = 0
                
    Part4Pin = " "
    PrepDate = Int(Now())
                
    Me.Line17Mo1 = 0
    Me.Line17Mo2 = 0
    Me.Line17Mo3 = 0
    
    BInitGrid fgMo1
    BInitGrid fgMo2
    BInitGrid fgMo3
                
    ' get the first and second month number of the quarter
    Month1 = (Me.cmbQtr.ListIndex * 3) + 1
    Month2 = Month1 + 1
    lblBMonth1 = MonthName(Month1)
    lblBMonth2 = MonthName(Month2)
    lblBMonth3 = MonthName(Month2 + 1)
                
    ' get start/end date for the quarter
    StartYM = CLng(Me.cmbYear.Text) * 100 + (Me.cmbQtr.ListIndex * 3) + 1
    EndYM = StartYM + 2
    
    frmProgress.Show
    
    frmProgress.lblMsg1 = PRCompany.Name
    frmProgress.Refresh
    
    ' get the PRHist data
    SQLString = "SELECT * FROM PRHist " & _
                " WHERE YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                " ORDER BY CheckDate, EmployeeID"
                    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Payroll History Found!!" & vbCr & StartYM & vbCr & EndYM, vbInformation
        frmProgress.Hide
        Exit Sub
    End If
    
    Recs = PRHist.Records
    Ct = 0
    
    Do
    
        If Ct Mod 10 = 1 Then
            frmProgress.lblMsg2 = Me.cmbQtr & Me.cmbYear & " " & _
                                 Format(Ct, CountFormat) & " of: " & Format(Recs, CountFormat)
            frmProgress.Refresh
        End If
    
        SSTax = SSTax + PRHist.SSTax
        MedTax = MedTax + PRHist.MedTax
        
        ' match SS# logic
        ' 2011 ER is still 6.2% / EE is 4.2%
        ERSSTax = (Round(PRHist.SSWage * 0.062, 2))
        MatchSS = MatchSS + ERSSTax
        ' MatchSS = MatchSS + PRHist.SSTax
        
        Add941Val " 2 )", 3, PRHist.FWTWage
        Add941Val " 3 )", 3, PRHist.FWTTax
        Add941Val " 5a)", 1, PRHist.SSWage
        
        ' *** Line5b - Tips ***
        If rsTips.RecordCount > 0 Then
            SQLString = "EmployeeID = " & PRHist.EmployeeID
            rsTips.Find SQLString, 0, adSearchForward, 1
            If rsTips.EOF = False Then
                SQLString = "SELECT * FROM PRDist WHERE HistID = " & PRHist.HistID & _
                            " AND ItemID = " & rsTips!ItemID
                If PRDist.GetBySQL(SQLString) Then
                    Add941Val " 5a)", 1, -PRDist.Amount
                    Add941Val " 5b)", 1, PRDist.Amount
                End If
            End If
        End If
        
        Add941Val " 5c)", 1, PRHist.MEDWage
        
        ' *** Line7b - sick pay ***
        ' *** Line7c - tips and group ins ***
        ' *** Line9 EIC payments ***
    
        ' tax liability per month
        ' use the 2011 match logic
        TaxLiab = PRHist.FWTTax + ((PRHist.SSTax + PRHist.MedTax) * 2)
        TaxLiab = PRHist.FWTTax + PRHist.SSTax + ERSSTax + PRHist.MedTax * 2
        
        If PRHist.YearMonth Mod 100 = Month1 Then
            Line17Mo1 = Line17Mo1 + TaxLiab
            BGridPop Me.fgMo1, TaxLiab, Day(PRHist.CheckDate)
        ElseIf PRHist.YearMonth Mod 100 = Month2 Then
            Line17Mo2 = Line17Mo2 + TaxLiab
            BGridPop Me.fgMo2, TaxLiab, Day(PRHist.CheckDate)
        Else
            Line17Mo3 = Line17Mo3 + TaxLiab
            BGridPop Me.fgMo3, TaxLiab, Day(PRHist.CheckDate)
        End If
        If Line18Check = 1 Then
            Line18Date.Visible = True
            Line18Date = Int(Now())
        Else
            Line18Date.Visible = False
        End If
        
        If Not PRHist.GetNext Then Exit Do
    Loop
    
    Calc941Data
    PopChkDate12
    PopPart4Part5
    frmProgress.Hide

End Sub

Private Sub BGridPop(ByRef fg As VSFlexGrid, ByVal TaxAmt As Currency, ByVal nDay As Byte)

Dim fgRow, fgCol As Integer
Dim CellValue As Currency

    fgCol = (Int((nDay - 1) / 8) * 2) + 1
    If nDay Mod 8 = 0 Then
        fgRow = 8
    Else
        fgRow = nDay Mod 8
    End If
    CellValue = fg.TextMatrix(fgRow, fgCol)
    CellValue = CellValue + TaxAmt
    fg.TextMatrix(fgRow, fgCol) = CellValue

End Sub

Private Sub Calc941Data()
    
Dim Cur As Currency
    
    ' calculated lines
    ' *** Line17Mo3 = Line10 - Line17Mo1 - Line17Mo2 ***

    Set941Val " 5a)", 2, Round(Get941Val(" 5a)", 1) * 0.104, 2)
    Set941Val " 5b)", 2, Round(Get941Val(" 5b)", 1) * 0.104, 2)
    Set941Val " 5c)", 2, Round(Get941Val(" 5c)", 1) * 0.029, 2)
            
    Cur = Get941Val(" 5a)", 2) + Get941Val(" 5b)", 2) + Get941Val(" 5c)", 2)
    Set941Val " 5d)", 3, Cur
    
    Cur = Get941Val(" 3 )", 3) + Get941Val(" 5d)", 3) + Get941Val(" 5e)", 3)
    Set941Val " 6e)", 3, Cur
    
    ' 7) fraction of cents
    If Me.chkManualFractions = 0 Then
        Cur = SSTax + MatchSS - Get941Val(" 5a)", 2) - Get941Val(" 5b)", 2) + MedTax * 2 - Get941Val(" 5c)", 2)
        Set941Val " 7 )", 2, Round(Cur, 2)
    End If
    
'    If Me.chkCents = 0 Then
'        Me.Line7a = Round(SSTax * 2 - Me.Line5aa - Me.Line5bb + MedTax * 2 - Me.Line5cc, 2)
'    End If
    
    ' 10) total of taxes after adjustments
    Cur = Get941Val(" 6e)", 3) + Get941Val(" 7 )", 2) + Get941Val(" 8 )", 2) + Get941Val(" 9 )", 2)
    Set941Val "10 )", 3, Cur
    
    ' 13)
    Set941Val "13 )", 3, Get941Val("11 )", 3) + Get941Val("12a)", 1)
    
    ' 14) Balance Due
    If Get941Val("10 )", 3) > Get941Val("13 )", 3) Then
        Cur = Get941Val("10 )", 3) - Get941Val("13", 3)
        Set941Val "14 )", 3, Cur
        Set941Val "15 )", 2, 0
        Me.Line15Check1.Enabled = False
        Me.Line15Check2.Enabled = False
    Else        ' overpayment
        Cur = Get941Val("13 )", 3) - Get941Val("10", 3)
        Set941Val "14 )", 3, 0
        Set941Val "15 )", 2, Cur
        Me.Line15Check1.Enabled = True
        Me.Line15Check2.Enabled = True
    End If
    
'    Me.Line8 = Me.Line6e + Me.Line7a + Me.Line7b + Me.Line7c
'    Me.Line10 = Me.Line8 - Me.Line9
'    Me.Line10Show = Me.Line10
'    BLine10Show = Line10
'    BDifference = Line10 - BLine10Show
'
'    BLine10Show = Get941Val("10 )", 3)
'    BDifference = BLine10Show
'
'    Me.Line12e = Round(Me.Line12d * 0.062, 2)
'
'    Me.Line13 = Me.Line11 + Me.Line12a + Me.Line12e
'
'    If Me.Line10 >= Me.Line13 Then   ' balance due
'        Me.Line14 = Me.Line10 - Me.Line13
'        Me.Line15 = 0
'        Me.Line15Check1.Enabled = False
'        Me.Line15Check2.Enabled = False
'    Else                            ' overpayment
'        Me.Line14 = 0
'        Me.Line15 = Me.Line13 - Me.Line10
'        Me.Line15Check1.Enabled = True
'        Me.Line15Check2.Enabled = True
'    End If
'

    Me.Line10Show = Get941Val("10 )", 3)
    Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Line17Diff = Get941Val("10 )", 3) - Line17Total

    BLine10Show = Get941Val("10 )", 3)
    BDifference = BLine10Show - BTotalTax

    BGridUpdate Me.fgMo1, BMo1Tax
    BGridUpdate Me.fgMo2, BMo2Tax
    BGridUpdate Me.fgMo3, BMo3Tax
    
    BTotalTax = BMo1Tax + BMo2Tax + BMo3Tax
    
    If Get941Val("10 )", 3) < 2500 Then
        Me.Line17Check1 = 1
        ' Me.Line17Check2.Enabled = False
    End If

End Sub

Private Sub BGridUpdate(ByRef fg As VSFlexGrid, ByRef MonthTotal As TDBNumber)
    
Dim CellValue As Currency
Dim i, j As Integer
    
    MonthTotal = 0
    For i = 1 To 8
        For j = 1 To 7 Step 2
            If fg.TextMatrix(i, j) <> "" Then
                CellValue = 0
                On Error Resume Next        ' turn of error handling
                CellValue = CCur(fg.TextMatrix(i, j))
                On Error GoTo 0             ' turn error handling back on
                MonthTotal = MonthTotal + CellValue
            End If
        Next j
    Next i

'    If fg.Col = 7 Then
'        If fg.Row = 8 Then
'        Else
'            fg.Row = fg.Row + 1
'            fg.Col = 0
'        End If
'    Else
'        fg.Col = fg.Col + 1
'    End If

    Me.BTotalTax = Me.BMo1Tax + Me.BMo2Tax + Me.BMo3Tax
    BDifference = Get941Val("10 )", 3) - Me.BTotalTax

End Sub

Private Sub PopChkDate12()
    
Dim DateDiff As Long
Dim Date12, ChkDate12 As Date
Dim rsEE As New ADODB.Recordset
Dim EECount, LastEE, LastDate As Long
Dim Pointer, DateCount As Long
    
    Me.cmbChkDate12.Clear
    
    ' determine 12th of the month date based on quarter selected
    Date12 = DateSerial(Me.cmbYear.Text, (Me.cmbQtr.ListIndex * 3) + 3, 12)
    DateDiff = 99999
    
    ' get start/end date for the quarter
    StartYM = CLng(Me.cmbYear.Text) * 100 + (Me.cmbQtr.ListIndex * 3) + 1
    EndYM = StartYM + 2
    
    ' get the PRHist data
    SQLString = "SELECT CheckDate, YearMonth, EmployeeID FROM PRHist " & _
                " WHERE YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                " ORDER BY CheckDate, EmployeeID"
                
    rsInit SQLString, cn, rsEE
                
    If rsEE.RecordCount = 0 Then
        MsgBox "No Payroll History Found!!" & vbCr & StartYM & vbCr & EndYM, vbInformation
        Exit Sub
    End If
    
    rsEE.MoveFirst
    
    LastEE = 0
    LastDate = 0
    EECount = 0
    Pointer = 0
    DateCount = 0
    
    Do
        ' break in check date
        If LastDate <> 0 And Int(rsEE!CheckDate) <> LastDate Then
            ' is this the closest to ChkDate12???
            If Abs(LastDate - Date12) < DateDiff Then
                DateCount = DateCount + 1
                DateDiff = Abs(LastDate - Date12)
                Pointer = DateCount - 1
            End If
            Me.cmbChkDate12.AddItem Format(LastDate, "mm/dd/yy") & " " & EECount
            EECount = 0
        End If
        LastDate = Int(rsEE!CheckDate)
        
        ' break in EmpID
        If rsEE!EmployeeID <> LastEE Then
            EECount = EECount + 1
        End If
        LastEE = rsEE!EmployeeID
        
        rsEE.MoveNext
    
    Loop Until rsEE.EOF
        
    If EECount <> 0 Then
        ' is this the closest to ChkDate12???
        If Abs(LastDate - Date12) < DateDiff Then
            DateCount = DateCount + 1
            DateDiff = Abs(LastDate - Date12)
            Pointer = DateCount - 1
        End If
        Me.cmbChkDate12.AddItem Format(LastDate, "mm/dd/yy") & " " & EECount
    End If
    
    Me.cmbChkDate12.ListIndex = Pointer

    rsEE.Close
    Set rsEE = Nothing

End Sub

Private Sub PopPart4Part5()
Dim i As Long
    
    Part4ID = 0
    Part5ID = 0
    PaidPrepID = 0
    
    ' Part 4 - Third Party Designee - Per User
    SQLString = "SELECT * FROM PRGlobal WHERE Typecode = " & _
                PREquate.GlobalType941Part4 & " AND UserID = " & User.ID
    
    If PRGlobal.GetBySQL(SQLString) Then
        Part4ID = PRGlobal.GlobalID
        Me.Part4Name = PRGlobal.Var1
        Me.Part4Phone = PRGlobal.Var2
        Me.Part4Pin = PRGlobal.Var3
        If Me.Part4Name <> "" Then
            Me.Part4CheckYes = 1
            Me.Part4CheckNo = 0
        End If
    End If
    
    ' Part 5 - Company Signature - Per Company
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & _
                PREquate.GlobalType941Part5 & " AND Userid = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) Then
        Part5ID = PRGlobal.GlobalID
        Me.Part5NameTitle = PRGlobal.Var1
        Me.Part5Phone = PRCompany.PhoneNumber
    End If

    ' populate the User Combo Box
    cmbPrepName.Clear
    SQLString = "SELECT * FROM Users ORDER BY NAME"
    If Not User.GetBySQL(SQLString) Then
       MsgBox "Users not found: " & UserID, vbCritical, "Form941 Entry"
       End
    End If

    Do
        cmbPrepName.AddItem UCase(User.Name)
        If Not User.GetNext Then Exit Do
    Loop
    
    ' re-get the User from the command prompt
    If User.GetByID(UserID) Then
    End If

'    ' Paid Preparer - Per Company
    
    SQLString = "SELECT * FROM PRglobal WHERE TypeCode = " & _
                PREquate.GlobalType941PaidPrep & " AND Userid = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        PaidPrepID = PRGlobal.GlobalID
        Me.cmbPrepName.AddItem User.Name
        Me.PrepFirm = PRGlobal.Var1
        Me.PrepAddr1 = PRGlobal.Var2
        Me.PrepAddr2 = PRGlobal.Var3
        Me.PrepPhone = PRGlobal.Var4
        Me.PrepEIN = PRGlobal.Var5
        Me.PrepZip = PRGlobal.Var6
        Me.PrepSSN = PRGlobal.Var7
        If PRGlobal.Var8 = "1" Then
            Me.PrepCheck = 1
        Else
            Me.PrepCheck = 0
        End If
        Me.cmbPrepName.Text = PRGlobal.Var9
    End If


End Sub

Private Sub cmbChkDate12_Click()
    ' Me.Line1 = Mid(Me.cmbChkDate12, 10, 10)
    Set941Val " 1 )", 3, Mid(Me.cmbChkDate12, 10, 10)
End Sub

Private Sub Timer1_Timer()
    If LoadFlag = True Then Exit Sub
    Calc941Data
End Sub
Private Sub cmbYear_Click()
    Get941Data
End Sub
Private Sub cmbQtr_Click()
    Get941Data
End Sub
Private Sub cmdCalc_Click()
    Calc941Data
End Sub
Private Sub Line11_LostFocus()
    Calc941Data
End Sub
Private Sub Line17Mo1_lostfocus()
    Me.Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Calc941Data
End Sub
Private Sub Line17Mo2_Change()
    Me.Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Calc941Data
End Sub

Private Sub Line17Mo3_lostfocus()
    Me.Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Calc941Data
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

Private Sub AddRow(ByVal Title As String, _
                   ByVal Edit1 As Boolean, _
                   ByVal Edit2 As Boolean, _
                   ByVal Edit3 As Boolean, _
                   ByVal Show1 As Boolean, _
                   ByVal Show2 As Boolean, _
                   ByVal Show3 As Boolean)
                   
    With Me.fg
        
        .TextMatrix(RowNum, GetCol("Descr")) = Title
        
        .TextMatrix(RowNum, GetCol("Edit1")) = Edit1
        .TextMatrix(RowNum, GetCol("Edit2")) = Edit2
        .TextMatrix(RowNum, GetCol("Edit3")) = Edit3
        
        .TextMatrix(RowNum, GetCol("Show1")) = Show1
        .TextMatrix(RowNum, GetCol("Show2")) = Show2
        .TextMatrix(RowNum, GetCol("Show3")) = Show3
    
    End With
    
    RowNum = RowNum + 1
                   
End Sub

Private Sub Set941Val(ByVal RowKey As String, ByVal ColNum As Byte, ByVal Amt As Currency)

    Dim fgRow, fgCol, fgI As Integer

    With Me.fg
        
        fgRow = .Row
        fgCol = .Col
        
        If ColNum < 1 Or ColNum > 3 Then
            MsgBox "Invalid ColNum: " + ColNum, vbExclamation
            End
        End If
        
        For fgI = 1 To .Rows
            If InStr(1, .TextMatrix(fgI - 1, GetCol("Descr")), RowKey, vbTextCompare) Then
                Exit For
            End If
        Next fgI
        
        If fgI = .Rows + 1 Then
            MsgBox "Row Key NF: " + RowKey, vbExclamation
            End
        End If
        
        .TextMatrix(fgI - 1, GetCol("Amt" & ColNum)) = Amt
    
    End With

End Sub

Private Sub Add941Val(ByVal RowKey As String, ByVal ColNum As Byte, ByVal Amt As Currency)

    Dim fgRow, fgCol, fgI As Integer
    Dim CellVal As Currency

    With Me.fg
        
        fgRow = .Row
        fgCol = .Col
        
        If ColNum < 1 Or ColNum > 3 Then
            MsgBox "Invalid ColNum: " + ColNum, vbExclamation
            End
        End If
        
        For fgI = 1 To .Rows
            If InStr(1, .TextMatrix(fgI - 1, GetCol("Descr")), RowKey, vbTextCompare) Then
                Exit For
            End If
        Next fgI
        
        If fgI = .Rows + 1 Then
            MsgBox "Row Key NF: " + RowKey, vbExclamation
            End
        End If
        
        CellVal = .TextMatrix(fgI - 1, GetCol("Amt" & ColNum))
        CellVal = CellVal + Amt
        .TextMatrix(fgI - 1, GetCol("Amt" & ColNum)) = CellVal
    
    End With

End Sub

Private Function Get941Val(ByVal RowKey As String, ByVal ColNum As Byte) As Currency

    Dim fgRow, fgCol, fgI As Integer

    With Me.fg
        
        fgRow = .Row
        fgCol = .Col
        
        If ColNum < 1 Or ColNum > 3 Then
            MsgBox "Invalid ColNum: " + ColNum, vbExclamation
            End
        End If
        
        For fgI = 1 To .Rows
            If InStr(1, .TextMatrix(fgI - 1, GetCol("Descr")), RowKey, vbTextCompare) Then
                Exit For
            End If
        Next fgI
        
        If fgI = .Rows + 1 Then
            MsgBox "Row Key NF: " + RowKey, vbExclamation
            End
        End If
        
        Get941Val = .TextMatrix(fgI - 1, GetCol("Amt" & ColNum))
    
    End With

End Function

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' can't edit control columns
    If Col = GetCol("Descr") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit1") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit2") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit3") Then Cancel = True: Exit Sub
    If Col = GetCol("Show1") Then Cancel = True: Exit Sub
    If Col = GetCol("Show2") Then Cancel = True: Exit Sub
    If Col = GetCol("Show3") Then Cancel = True: Exit Sub
    
    ' flagged as not editable
    If fg.TextMatrix(Row, Col + 3) = "False" Then
        Cancel = True
        Exit Sub
    End If

End Sub

'=======================================   FORM 941     ======================================
' Rev Jan 2011
'
Public Sub Form941A2011Jan()

Dim VertSpace, VertPosn, HorzPosn As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString, TelFmtString, ReportTitle As String
 
    CurrYear = Year(Now())
    Ln = 0
    SetEquates
    PrtInit ("Port")
    ReportTitle = "labels "
    SetFont 10, Equate.Portrait
    
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

    VertSpace = 492
    FmtString = "##,###,##0.00"
    TelFmtString = "###-###-####"
    
    With frm941_2011A
    
        PosPrint 3200, 1020, PRCompany.FederalID
        
        ' PosPrint 2500, 1490, PRCompany.Name
        PosPrint 2500, 1490, Me.txtName
        
        PosPrint 2200, 1720, Me.txtTradeName
        PosPrint 1500, 2200, PRCompany.Address1
        
        If PRCompany.Address2 <> "" Then
            PosPrint 1500, 2400, PRCompany.Address2
        End If
        If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
            PosPrint 1500, 2675, PRCompany.City & ", " & PRState.StateAbbrev & "  " & PRCompany.ZipCode
        End If
    
        If .cmbQtr = 1 Then
            PosPrint 8400, 1020, "X"
        End If   'col  'line
    
        If .cmbQtr = 2 Then
            PosPrint 8400, 1500, "X"
        End If
        If .cmbQtr = 3 Then
            PosPrint 8400, 1970, "X"
        End If
        If .cmbQtr = 4 Then
            PosPrint 8400, 2490, "X"     '  over   down
        End If
    
        
        PosPrint 10200, 3410, PadRight(Format(Get941Val(" 1 )", 3), "##,##0"), 6)
        
        PosPrint 9400, 3890, PadRight(Format(Get941Val(" 2 )", 3), FmtString), 13)
        PosPrint 9400, 4370, PadRight(Format(Get941Val(" 3 )", 3), FmtString), 13)
        
        PosPrint 8800, 4900, .AlphaCheckLine4
    
        PosPrint 4000, 5570, PadRight(Format(Get941Val(" 5a)", 1), FmtString), 13)
        PosPrint 6900, 5570, PadRight(Format(Get941Val(" 5a)", 2), FmtString), 13)
        
        PosPrint 4000, 6040, PadRight(Format(Get941Val(" 5b)", 1), FmtString), 13)
        PosPrint 6900, 6040, PadRight(Format(Get941Val(" 5b)", 2), FmtString), 13)
        
        PosPrint 4000, 6540, PadRight(Format(Get941Val(" 5c)", 1), FmtString), 13)
        PosPrint 6900, 6540, PadRight(Format(Get941Val(" 5c)", 2), FmtString), 13)
        PosPrint 9400, 7000, PadRight(Format(Get941Val(" 5d)", 3), FmtString), 13)
        
        PosPrint 9400, 7470, PadRight(Format(Get941Val(" 5e)", 3), FmtString), 13)
        
        VertPosn = 8880 + 360
        VertSpace = 470
        
        PosPrint 9400, VertPosn, PadRight(Format(Get941Val(" 6e)", 3), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 6900, VertPosn, PadRight(Format(Get941Val(" 7 )", 2), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 6900, VertPosn, PadRight(Format(Get941Val(" 8 )", 2), FmtString), 13)
                
        VertPosn = VertPosn + VertSpace
        PosPrint 6900, VertPosn, PadRight(Format(Get941Val(" 9 )", 2), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 9400, VertPosn, PadRight(Format(Get941Val("10 )", 3), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 9400, VertPosn, PadRight(Format(Get941Val("11 )", 3), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 4400, VertPosn, PadRight(Format(Get941Val("12a)", 1), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 4400, VertPosn, PadRight(Format(Get941Val("12b)", 1), "##,##0"), 6)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 9400, VertPosn, PadRight(Format(Get941Val("13 )", 3), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 9400, VertPosn, PadRight(Format(Get941Val("14 )", 3), FmtString), 13)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 6900, VertPosn, PadRight(Format(Get941Val("15 )", 2), FmtString), 13)
        
        PosPrint 9730, 13900, .AlphaCheckLine15a
        PosPrint 9730, 14170, .AlphaCheckLine15b     '
        
        FormFeed

    '   #######################  FORM 941 - PAGE 2  ####################################
    '
        
        ' *** FIX ***
        VertNudge = VertNudge + 6
        HorzNudge = HorzNudge + 2
        
        PosPrint 900, 1120, PRCompany.Name
        PosPrint 8490, 1120, PRCompany.FederalID
        PosPrint 950, 2100, .Line16
        PosPrint 1830, 2550, .AlphaCheckLine17a
        PosPrint 1830, 3100, .AlphaCheckLine17b

        If .Line17Check2 = 1 Then
            PosPrint 5000, 3800, PadRight(Format(.Line17Mo1, FmtString), 13)
            PosPrint 5000, 4250, PadRight(Format(.Line17Mo2, FmtString), 13)
            PosPrint 5000, 4710, PadRight(Format(.Line17Mo3, FmtString), 13)
            PosPrint 5000, 5230, PadRight(Format(.Line17Total, FmtString), 13)
        End If

        PosPrint 1820, 5540, .AlphaCheckLine17c
        PosPrint 9200, 6500, .AlphaCheckLine18
        If .Line18Check = 1 Then
            PosPrint 3900, 6950, .Line18Date
        End If
        PosPrint 9200, 7220, .AlphaCheckLine19
        
        Form941Pt4Pt5 frm941_2011A

        If .Part4CheckYes = 1 Then
            PosPrint 880, 8200, "X"
        Else
            PosPrint 900, 8900, "X"
        End If

        PosPrint 2600, 14470, .AlphaCheckPart5
        If IsNull(.PrepDate) = False Then
            PosPrint 8180, 14450, .PrepDate
        End If
    
        ' *** put it back (for sched B) ***
        'VertNudge = VertNudge - 6
        'HorzNudge = HorzNudge - 2
    
    End With

End Sub


