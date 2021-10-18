VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm941_2021_June 
   Caption         =   "Form 941 Rev June 2021"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   12390
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
      TabIndex        =   76
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
      TabIndex        =   41
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
      TabIndex        =   45
      Top             =   30
      Width           =   645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   120
      TabIndex        =   39
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
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
      TabCaption(0)   =   "Form 941 Part 1/3"
      TabPicture(0)   =   "frm941_Jun2021.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(1)=   "Label38"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Line15Check2"
      Tab(0).Control(4)=   "Line15Check1"
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(6)=   "fg"
      Tab(0).Control(7)=   "chkManualFractions"
      Tab(0).Control(8)=   "cmdPmt"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Form 941   Parts 2/3/4"
      TabPicture(1)   =   "frm941_Jun2021.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label48"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label43"
      Tab(1).Control(4)=   "Label44"
      Tab(1).Control(5)=   "Label39"
      Tab(1).Control(6)=   "Label40"
      Tab(1).Control(7)=   "Label45"
      Tab(1).Control(8)=   "Label46"
      Tab(1).Control(9)=   "Label47"
      Tab(1).Control(10)=   "Label49"
      Tab(1).Control(11)=   "Label51"
      Tab(1).Control(12)=   "Label52"
      Tab(1).Control(13)=   "Label53"
      Tab(1).Control(14)=   "Label54"
      Tab(1).Control(15)=   "Label1"
      Tab(1).Control(16)=   "Label3"
      Tab(1).Control(17)=   "Label7"
      Tab(1).Control(18)=   "Label8"
      Tab(1).Control(19)=   "Line16Mo1"
      Tab(1).Control(20)=   "Line16Total"
      Tab(1).Control(21)=   "Line16Mo3"
      Tab(1).Control(22)=   "Line16Mo2"
      Tab(1).Control(23)=   "txtEIN"
      Tab(1).Control(24)=   "Line16Check1"
      Tab(1).Control(25)=   "Line16Check2"
      Tab(1).Control(26)=   "Line16Check3"
      Tab(1).Control(27)=   "Part4CheckNo"
      Tab(1).Control(28)=   "Line17Check"
      Tab(1).Control(29)=   "Line17Date"
      Tab(1).Control(30)=   "Line18a"
      Tab(1).Control(31)=   "txtName"
      Tab(1).Control(32)=   "Line10Show"
      Tab(1).Control(33)=   "Line16Diff"
      Tab(1).Control(34)=   "Part4CheckYes"
      Tab(1).Control(35)=   "Part4Name"
      Tab(1).Control(36)=   "Part4Pin"
      Tab(1).Control(37)=   "Part4Phone"
      Tab(1).Control(38)=   "txtTradeName"
      Tab(1).Control(39)=   "Line18b"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "Form 941   Part 5"
      TabPicture(2)   =   "frm941_Jun2021.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label56"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "PrepPhone"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Part5Date"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "PrepSSN"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "PrepDate"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "PrepZip"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "PrepEIN"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "PrepAddr1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "PrepFirm"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Part5NameTitle"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "PrepCheck"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cmbPrepName"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "PrepAddr2"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Part5Phone"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Schedule B (Form 941)"
      TabPicture(3)   =   "frm941_Jun2021.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label58"
      Tab(3).Control(1)=   "Label59"
      Tab(3).Control(2)=   "Label60"
      Tab(3).Control(3)=   "Label61"
      Tab(3).Control(4)=   "lblBMonth1"
      Tab(3).Control(5)=   "lblBMonth2"
      Tab(3).Control(6)=   "lblBMonth3"
      Tab(3).Control(7)=   "BDifference"
      Tab(3).Control(8)=   "BLine10Show"
      Tab(3).Control(9)=   "tdbNumVertNudge"
      Tab(3).Control(10)=   "fgMo3"
      Tab(3).Control(11)=   "fgMo2"
      Tab(3).Control(12)=   "BTotalTax"
      Tab(3).Control(13)=   "BMo3Tax"
      Tab(3).Control(14)=   "BMo2Tax"
      Tab(3).Control(15)=   "BMo1Tax"
      Tab(3).Control(16)=   "fgMo1"
      Tab(3).Control(17)=   "tdbNumHorzNudge"
      Tab(3).ControlCount=   18
      Begin VB.CheckBox Line18b 
         Caption         =   "Check here"
         Height          =   255
         Left            =   -65280
         TabIndex        =   95
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdPmt 
         Caption         =   "$$"
         Height          =   375
         Left            =   -63360
         TabIndex        =   90
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox txtTradeName 
         Height          =   360
         Left            =   -70800
         TabIndex        =   87
         Top             =   735
         Width           =   3255
      End
      Begin VB.CheckBox chkManualFractions 
         Caption         =   "Override Fractions of cents - Line 7"
         Height          =   495
         Left            =   -74400
         TabIndex        =   86
         Top             =   8640
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid fg 
         Height          =   7335
         Left            =   -74640
         TabIndex        =   85
         Top             =   1080
         Width           =   11055
         _cx             =   19500
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
         Left            =   -63960
         Top             =   360
      End
      Begin TDBText6Ctl.TDBText Part4Phone 
         Height          =   375
         Left            =   -73680
         TabIndex        =   23
         Top             =   8280
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "frm941_Jun2021.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":00CE
         Key             =   "frm941_Jun2021.frx":00EC
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
         Left            =   -69840
         TabIndex        =   24
         Top             =   8280
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   661
         Caption         =   "frm941_Jun2021.frx":0130
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":01CC
         Key             =   "frm941_Jun2021.frx":01EA
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
         Left            =   3960
         TabIndex        =   28
         Top             =   1740
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":022E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":028C
         Key             =   "frm941_Jun2021.frx":02AA
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
         Left            =   240
         TabIndex        =   32
         Top             =   5000
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":02EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":035A
         Key             =   "frm941_Jun2021.frx":0378
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
         Left            =   2080
         TabIndex        =   29
         Top             =   3600
         Width           =   5750
      End
      Begin TDBText6Ctl.TDBText Part4Name 
         Height          =   375
         Left            =   -73320
         TabIndex        =   22
         Top             =   7680
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   661
         Caption         =   "frm941_Jun2021.frx":03BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":042E
         Key             =   "frm941_Jun2021.frx":044C
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
         Left            =   -74520
         TabIndex        =   21
         Top             =   7680
         Width           =   735
      End
      Begin TDBNumber6Ctl.TDBNumber Line16Diff 
         Height          =   300
         Left            =   -66240
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":0490
         Caption         =   "frm941_Jun2021.frx":04B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0516
         Keys            =   "frm941_Jun2021.frx":0534
         Spin            =   "frm941_Jun2021.frx":057E
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
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":05A6
         Caption         =   "frm941_Jun2021.frx":05C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0630
         Keys            =   "frm941_Jun2021.frx":064E
         Spin            =   "frm941_Jun2021.frx":0698
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
         TabIndex        =   74
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":06C0
         Caption         =   "frm941_Jun2021.frx":06E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0760
         Keys            =   "frm941_Jun2021.frx":077E
         Spin            =   "frm941_Jun2021.frx":07C8
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
         TabIndex        =   73
         Top             =   -480
         Width           =   3855
      End
      Begin VB.CheckBox Line15Check1 
         Caption         =   "Apply to next return."
         Height          =   255
         Left            =   -65280
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
         FormatString    =   $"frm941_Jun2021.frx":07F0
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
         Left            =   2040
         TabIndex        =   38
         Top             =   5460
         Width           =   3495
      End
      Begin VB.CheckBox Line18a 
         Caption         =   "Check here."
         Height          =   255
         Left            =   -65280
         TabIndex        =   20
         Top             =   5620
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate Line17Date 
         Height          =   285
         Left            =   -74400
         TabIndex        =   19
         Top             =   5220
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   503
         Calendar        =   "frm941_Jun2021.frx":08CA
         Caption         =   "frm941_Jun2021.frx":09E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0A8C
         Keys            =   "frm941_Jun2021.frx":0AAA
         Spin            =   "frm941_Jun2021.frx":0B08
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
      Begin VB.CheckBox Line17Check 
         Caption         =   "Check here"
         Height          =   255
         Left            =   -65280
         TabIndex        =   18
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
         Left            =   -74520
         TabIndex        =   25
         Top             =   8280
         Width           =   615
      End
      Begin VB.CheckBox Line16Check3 
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
         TabIndex        =   17
         Top             =   4080
         Width           =   9855
      End
      Begin VB.CheckBox Line16Check2 
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
         TabIndex        =   12
         Top             =   2040
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox Line16Check1 
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
         TabIndex        =   11
         Top             =   1755
         Width           =   7215
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
         Left            =   -65280
         TabIndex        =   5
         Top             =   8880
         Width           =   1750
      End
      Begin TDBNumber6Ctl.TDBNumber Line16Mo2 
         Height          =   300
         Left            =   -71760
         TabIndex        =   14
         Top             =   2940
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":0B30
         Caption         =   "frm941_Jun2021.frx":0B50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0BBA
         Keys            =   "frm941_Jun2021.frx":0BD8
         Spin            =   "frm941_Jun2021.frx":0C22
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
         ValueVT         =   34537473
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line16Mo3 
         Height          =   300
         Left            =   -71760
         TabIndex        =   15
         Top             =   3300
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":0C4A
         Caption         =   "frm941_Jun2021.frx":0C6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0CD4
         Keys            =   "frm941_Jun2021.frx":0CF2
         Spin            =   "frm941_Jun2021.frx":0D3C
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
         ValueVT         =   34537473
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line16Total 
         Height          =   300
         Left            =   -73245
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3660
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8184
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":0D64
         Caption         =   "frm941_Jun2021.frx":0D84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0E16
         Keys            =   "frm941_Jun2021.frx":0E34
         Spin            =   "frm941_Jun2021.frx":0E7E
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
         ValueVT         =   27262977
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line16Mo1 
         Height          =   300
         Left            =   -71760
         TabIndex        =   13
         Top             =   2580
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":0EA6
         Caption         =   "frm941_Jun2021.frx":0EC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":0F30
         Keys            =   "frm941_Jun2021.frx":0F4E
         Spin            =   "frm941_Jun2021.frx":0F98
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
         ValueVT         =   34537473
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText Part5NameTitle 
         Height          =   345
         Left            =   120
         TabIndex        =   26
         Top             =   1260
         Width           =   11055
         _Version        =   65536
         _ExtentX        =   19500
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":0FC0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1040
         Key             =   "frm941_Jun2021.frx":105E
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
         Left            =   240
         TabIndex        =   30
         Top             =   4020
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":10A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1114
         Key             =   "frm941_Jun2021.frx":1132
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
         Left            =   240
         TabIndex        =   31
         Top             =   4510
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":1176
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":11E0
         Key             =   "frm941_Jun2021.frx":11FE
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
         Left            =   8640
         TabIndex        =   34
         Top             =   4080
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":1242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":12A4
         Key             =   "frm941_Jun2021.frx":12C2
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
         Left            =   8160
         TabIndex        =   35
         Top             =   4560
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":1306
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1372
         Key             =   "frm941_Jun2021.frx":1390
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
         Left            =   8520
         TabIndex        =   37
         Top             =   5520
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   609
         Calendar        =   "frm941_Jun2021.frx":13D4
         Caption         =   "frm941_Jun2021.frx":14EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1550
         Keys            =   "frm941_Jun2021.frx":156E
         Spin            =   "frm941_Jun2021.frx":15CC
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
         Left            =   8040
         TabIndex        =   36
         Top             =   5040
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":15F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1660
         Key             =   "frm941_Jun2021.frx":167E
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
         Left            =   120
         TabIndex        =   27
         Top             =   1740
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   609
         Calendar        =   "frm941_Jun2021.frx":16C2
         Caption         =   "frm941_Jun2021.frx":17DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":183E
         Keys            =   "frm941_Jun2021.frx":185C
         Spin            =   "frm941_Jun2021.frx":18BA
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":18E2
         Caption         =   "frm941_Jun2021.frx":1902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":198C
         Keys            =   "frm941_Jun2021.frx":19AA
         Spin            =   "frm941_Jun2021.frx":19F4
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":1A1C
         Caption         =   "frm941_Jun2021.frx":1A3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1AC6
         Keys            =   "frm941_Jun2021.frx":1AE4
         Spin            =   "frm941_Jun2021.frx":1B2E
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6840
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":1B56
         Caption         =   "frm941_Jun2021.frx":1B76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1C00
         Keys            =   "frm941_Jun2021.frx":1C1E
         Spin            =   "frm941_Jun2021.frx":1C68
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   7560
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":1C90
         Caption         =   "frm941_Jun2021.frx":1CB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":1D3E
         Keys            =   "frm941_Jun2021.frx":1D5C
         Spin            =   "frm941_Jun2021.frx":1DA6
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
         FormatString    =   $"frm941_Jun2021.frx":1DCE
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
         FormatString    =   $"frm941_Jun2021.frx":1EA8
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
         TabIndex        =   75
         Top             =   1320
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941_Jun2021.frx":1F82
         Caption         =   "frm941_Jun2021.frx":1FA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":2022
         Keys            =   "frm941_Jun2021.frx":2040
         Spin            =   "frm941_Jun2021.frx":208A
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
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   8205
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":20B2
         Caption         =   "frm941_Jun2021.frx":20D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":213C
         Keys            =   "frm941_Jun2021.frx":215A
         Spin            =   "frm941_Jun2021.frx":21A4
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
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   8520
         Width           =   6945
         _Version        =   65536
         _ExtentX        =   12259
         _ExtentY        =   529
         Calculator      =   "frm941_Jun2021.frx":21CC
         Caption         =   "frm941_Jun2021.frx":21EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":22B6
         Keys            =   "frm941_Jun2021.frx":22D4
         Spin            =   "frm941_Jun2021.frx":231E
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
         Left            =   8400
         TabIndex        =   33
         Top             =   3600
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   609
         Caption         =   "frm941_Jun2021.frx":2346
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941_Jun2021.frx":23A4
         Key             =   "frm941_Jun2021.frx":23C2
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
      Begin VB.Label Label8 
         Caption         =   "If you're eligible for the emp. retentrion credit because your business is a recovery startup . . . . . . . . . . . . ."
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
         TabIndex        =   94
         Top             =   6000
         Width           =   9135
      End
      Begin VB.Label Label7 
         Caption         =   "18b"
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
         TabIndex        =   93
         Top             =   6000
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "*** Lines 19 to 28 are on the first tab ***"
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
         Left            =   -74160
         TabIndex        =   91
         Top             =   6360
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "Part 5: Sign here.  You MUST complete both pages of Form 941 and SIGN it."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label Label1 
         Caption         =   "Trade Name"
         Height          =   375
         Left            =   -70800
         TabIndex        =   88
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
         Left            =   240
         TabIndex        =   84
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   945
         Width           =   950
      End
      Begin VB.Label Label5 
         Caption         =   "One"
         Height          =   225
         Left            =   -65880
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         Left            =   120
         TabIndex        =   67
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
         Left            =   0
         TabIndex        =   66
         Top             =   -2640
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
         Left            =   -74520
         TabIndex        =   65
         Top             =   7200
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
         TabIndex        =   64
         Top             =   6840
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   5625
         Width           =   9015
      End
      Begin VB.Label Label49 
         Caption         =   "18a"
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
         TabIndex        =   61
         Top             =   5640
         Width           =   375
      End
      Begin VB.Label Label47 
         Caption         =   $"frm941_Jun2021.frx":2406
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
         TabIndex        =   60
         Top             =   4930
         Width           =   9375
      End
      Begin VB.Label Label46 
         Caption         =   "Report of Tax Liability for Semiweekly Schedule Depositors, and attach it to this form."
         Height          =   255
         Left            =   -73005
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   4920
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
         TabIndex        =   55
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
         TabIndex        =   47
         Top             =   1740
         Width           =   1335
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
         TabIndex        =   54
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Employer Identification number (EIN)"
         Height          =   255
         Left            =   -67320
         TabIndex        =   53
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label48 
         Caption         =   "Name (not your trade name)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label38 
         Caption         =   "Check"
         Height          =   225
         Left            =   -65880
         TabIndex        =   51
         Top             =   8640
         Width           =   495
      End
      Begin VB.Label Label33 
         Height          =   180
         Left            =   -73605
         TabIndex        =   50
         Top             =   7230
         Width           =   7035
      End
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
      TabIndex        =   43
      Top             =   40
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "18a"
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
      Left            =   600
      TabIndex        =   92
      Top             =   4320
      Width           =   375
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
      TabIndex        =   49
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
      TabIndex        =   48
      Top             =   80
      Width           =   495
   End
End
Attribute VB_Name = "frm941_2021_June"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public AlphaCheckLine4 As String
Public AlphaCheckLine15a As String
Public AlphaCheckLine15b As String
Public AlphaCheckLine16a As String
Public AlphaCheckLine16b As String
Public AlphaCheckLine16c As String
Public AlphaCheckLine17 As String
Public AlphaCheckLine18a As String
Public AlphaCheckLine18b As String
Public AlphaCheckPart4Yes As String
Public AlphaCheckPart4No As String
Public AlphaCheckPart5 As String
Public TotTaxLiability As Currency
Public Part4ID, Part5ID, PaidPrepID As Long

Dim StartYM, EndYM As Long
Dim LoadFlag As Boolean
Dim SSTax, MedTax, MedAddTl As Currency
    
Dim rsTips As New ADODB.Recordset
Dim rsERTips As New ADODB.Recordset
Dim I, J, K As Long

Dim rsCol As New ADODB.Recordset
Dim RowNum, ColNum As Integer
Dim ERSSTax, MatchSS, MatchSSTotal As Currency
        
Dim PanelVert As Integer
Dim PrtTest As Boolean


Private Sub cmdPmt_Click()
' Get941Val(" 5a)", 1)
    Dim d14 As Double
    d14 = Get941Val("12 ) Total taxes after adjustments and credits ", 3)
    Set941Val "13a)", 3, d14
End Sub

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
    
    tdbAmountSet Me.Line16Mo1
    tdbAmountSet Me.Line16Mo2
    tdbAmountSet Me.Line16Mo3
    tdbAmountSet Me.Line16Total
    tdbAmountSet Me.Line10Show
    tdbAmountSet Me.Line16Diff
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
    
    Line16Total.ReadOnly = True
    Line16Diff.ReadOnly = True
    
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
    
    ' *** stuff 1st qtr
    ' Me.cmbQtr.ListIndex = 0                                     ''''''''''''  TAKE OUT  '''''''''''''
    
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
    
'    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
'       MsgBox "WARNING: Employer State Not Filled In", vbExclamation, "Form 941 Entry"
'       Me.Line16 = ""
'    Else
'        Me.Line16 = PRState.StateAbbrev
'    End If
    
    CurrYear = Year(Now())

    SetNudge Me.tdbNumHorzNudge
    SetNudge Me.tdbNumVertNudge
    GetNudge User.ID, "941_2021June"
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
    
    I = 1800
    AddCol "Amt1", "Amt1", I
    AddCol "Amt2", "Amt2", I
    AddCol "Amt3", "Amt3", I
    
    I = 0
    AddCol "Edit1", "Edit1", I, adBoolean
    AddCol "Edit2", "Edit2", I, adBoolean
    AddCol "Edit3", "Edit3", I, adBoolean
    AddCol "Show1", "Show1", I, adBoolean
    AddCol "Show2", "Show2", I, adBoolean
    AddCol "Show3", "Show3", I, adBoolean

    RowNum = 0
    
    With Me.fg
    
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 10
        .Rows = 99
        .Editable = flexEDKbdMouse
        
        I = 0
        rsCol.MoveFirst
        Do
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
    
        ' Title, Edit1, Edit2, Edit3, Show1, Show2, Show3
    
        AddRow " 1 ) Number of employees", False, False, True, False, False, True
        AddRow " 2 ) Wages, tips and other compensation", False, False, True, False, False, True
        AddRow " 3 ) Income tax withheld", False, False, True, False, False, True
        
        AddRow " 5a)   Taxable social security wages", True, True, False, True, True, False
        AddRow " 5ai)  Qualified sick leave wages", True, True, False, True, True, False
        AddRow " 5aii) Qualified family leave wages", True, True, False, True, True, False
        AddRow " 5b)   Taxable social security tips", True, True, False, True, True, False
        AddRow " 5c)   Taxable Medicare wages & tips", True, True, False, True, True, False
        AddRow " 5d)   Taxable wages & tips Addl", True, True, False, True, True, False
        
        AddRow " 5e) Add Col 2 5a,Col 2 5b, Col 2 5c, Col2 5d", False, False, False, False, False, True
        AddRow " 5f) Sec 3121(q) Notice and Demand-Tax due on unreported tips", False, False, True, False, False, True
        
        AddRow " 6 ) Total taxes before adjustments", False, False, False, False, False, True
        AddRow " 7 ) Current qtr adj for fractions of cents", False, False, True, False, False, True
        AddRow " 8 ) Current qtr adj for sick pay", False, False, True, False, False, True
        AddRow " 9 ) Current qtr adj for tips and group-term life insurance", False, False, True, False, False, True
        
        AddRow "10 ) Total taxes after adjustments", False, False, False, False, False, True
        AddRow "11a) Qualified small bus. payroll tax credit ", False, False, True, False, False, True
        AddRow "11b) Nonref. sick & family leave wages before 4/1/21", False, False, True, False, False, True
        AddRow "11c) Nonref. portion employee retention credit", False, False, True, False, False, True
        AddRow "11d) Nonref. sick & family leave wages after 3/31/21", False, False, True, False, False, True
        AddRow "11e) Nonref. COBRA prem. assistance", False, False, True, False, False, True
        AddRow "11f) Number of individuals COBRA prem. asst.", False, True, False, False, True, False
        AddRow "11g) Total nonrefundable credits", False, False, False, False, False, True
        
        AddRow "12 ) Total taxes after adjustments and credits ", False, False, False, False, False, True
        
        AddRow "13a) Total deposits for this quarter ", False, False, True, False, False, True
        AddRow "13c) Before 4/1/21 - Refundable portion of credit for sick/family leave ", False, False, True, False, False, True
        AddRow "13d) Refundable portion of employee retention credit ", False, False, True, False, False, True
        AddRow "13e) After 3/31/21 - Refundable portion of credit for sick/family leave  ", False, False, False, False, False, True
        AddRow "13f) Refundable portion of COBRA premium assistance ", False, False, True, False, False, True
        AddRow "13g) Total deposits and refundable credits ", False, False, False, False, False, True
        AddRow "13h) Total Advances received from Form(s) 7200 for the quarter ", False, False, True, False, False, True
        AddRow "13i) Total deposits and refundable credits less advances", False, False, False, False, False, True
        AddRow "14 ) Balance due ", False, False, False, False, False, True
        AddRow "15 ) Overpayment ", False, False, False, False, True, False
        
        AddRow "19 ) Before 4/1/21 Qualified health plan expenses to sick leave ", False, False, True, False, False, True
        AddRow "20 ) Before 4/1/21 Qualified health plan expenses to family leave ", False, False, True, False, False, True
        AddRow "21 ) Qualified wages for employee retention credit ", False, False, True, False, False, True
        AddRow "22 ) Qualified health plan expenses for employee retention credit", False, False, True, False, False, True
        AddRow "23 ) After 3/31/21 Qualified sick leave wages ", False, False, True, False, False, True
        AddRow "24 ) Qualified health plan expenses on line 23 ", False, False, True, False, False, True
        AddRow "25 ) Amounts under CBA allocable to sick leave on line 23", False, False, True, False, False, True
        AddRow "26 ) After 3/31/21 Qualified family leave wages", False, False, True, False, False, True
        AddRow "27 ) Qualified health plan expenses on line 26", False, False, True, False, False, True
        AddRow "28 ) Amounts under CBA qualified family leave line 26", False, False, True, False, False, True
        
        .Rows = RowNum
    
        .ColFormat(GetCol("Amt1")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt2")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt3")) = "##,###,##0.00-"
    
        ' color the grid
        For I = 1 To RowNum
            For J = 1 To 3
                If J = 1 Then K = GetCol("Show1")
                If J = 2 Then K = GetCol("Show2")
                If J = 3 Then K = GetCol("Show3")
                If .TextMatrix(I - 1, K) = "False" Then
                    .Select I - 1, K - 6
                    .CellBackColor = RGB(192, 192, 192)
                    .CellBackColor = RGB(100, 100, 100)
                End If
            Next J
        Next I
    
        ' set the amounts to zero
        Set941Val " 1 )", 3, 0
        Set941Val " 2 )", 3, 0
        Set941Val " 3 )", 3, 0
        Set941Val " 5a)", 1, 0
        Set941Val " 5a)", 2, 0
        Set941Val " 5ai)", 1, 0
        Set941Val " 5ai)", 2, 0
        Set941Val " 5aii)", 1, 0
        Set941Val " 5aii)", 2, 0
        Set941Val " 5b)", 1, 0
        Set941Val " 5b)", 2, 0
        Set941Val " 5c)", 1, 0
        Set941Val " 5c)", 2, 0
        Set941Val " 5d)", 1, 0
        Set941Val " 5d)", 2, 0
        Set941Val " 5e)", 3, 0
        Set941Val " 5f)", 3, 0
        Set941Val " 6 )", 3, 0
        Set941Val " 7 )", 3, 0
        Set941Val " 8 )", 3, 0
        Set941Val " 9 )", 3, 0
        Set941Val "10 )", 3, 0
        Set941Val "11a)", 3, 0
        Set941Val "11b)", 3, 0
        Set941Val "11c)", 3, 0
        Set941Val "11d)", 3, 0
        Set941Val "11e)", 3, 0
        Set941Val "11f)", 2, 0
        Set941Val "11g)", 3, 0
        Set941Val "12 )", 3, 0
        Set941Val "13a)", 3, 0
        Set941Val "13c)", 3, 0
        Set941Val "13d)", 3, 0
        Set941Val "13e)", 3, 0
        Set941Val "13f)", 3, 0
        Set941Val "13g)", 3, 0
        Set941Val "13h)", 3, 0
        Set941Val "13i)", 3, 0
        Set941Val "14 )", 3, 0
        Set941Val "15 )", 2, 0
        
        Set941Val "19 )", 3, 0
        Set941Val "20 )", 3, 0
        Set941Val "21 )", 3, 0
        Set941Val "22 )", 3, 0
        Set941Val "23 )", 3, 0
        Set941Val "24 )", 3, 0
        Set941Val "25 )", 3, 0
        Set941Val "26 )", 3, 0
        Set941Val "27 )", 3, 0
        Set941Val "28 )", 3, 0
    
    End With

End Sub

Private Sub BInitGrid(ByRef fg As VSFlexGrid)
        
Dim I, J As Integer
Dim K, m As Integer
        
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
        
        For I = 1 To 8
            .TextMatrix(I, 0) = I
            .TextMatrix(I, 1) = 0
            .TextMatrix(I, 2) = I + 8
            .TextMatrix(I, 3) = 0
            .TextMatrix(I, 4) = I + 16
            .TextMatrix(I, 5) = 0
            .TextMatrix(I, 6) = I + 24
            .TextMatrix(I, 7) = 0
        Next I
    
        .ColFormat(1) = "$###,###,##0.00"
        .ColFormat(3) = "$###,###,##0.00"
        .ColFormat(5) = "$###,###,##0.00"
        .ColFormat(7) = "$###,###,##0.00"


    For K = 0 To 7 Step 2
        .ColWidth(K) = 400
        .TextMatrix(0, K) = "Day"
    Next K
            
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

Private Sub Line16Check1_Click()
    If Line16Check1 = 1 Then
        Me.AlphaCheckLine16a = "X"
        Me.Line16Check2 = 0
        Me.Line16Check3 = 0
    Else
        Me.AlphaCheckLine16a = ""
    End If
    
    Me.Line16Check2.TabStop = False
    Me.Line16Mo1.TabStop = False
    Me.Line16Mo2.TabStop = False
    Me.Line16Mo3.TabStop = False
    Me.Line16Check3.TabStop = False
    
    Me.Line16Mo1.Visible = False
    Me.Line16Mo2.Visible = False
    Me.Line16Mo3.Visible = False
    Me.Line16Total.Visible = False
    Me.Line10Show.Visible = False
    Me.Line16Diff.Visible = False
    Me.Label40.Visible = False
    Me.Label45.Visible = False

End Sub

Private Sub Line16Check2_Click()
    
    If Line16Check2 = 1 Then
        Me.AlphaCheckLine16b = "X"
        Me.Line16Check1 = 0
        Me.Line16Check3 = 0
    Else
        Me.AlphaCheckLine16b = ""
        Me.Line16Mo1.TabStop = False
        Me.Line16Mo2.TabStop = False
        Me.Line16Mo3.TabStop = False
    End If
    Me.Line16Mo1.Visible = True
    Me.Line16Mo2.Visible = True
    Me.Line16Mo3.Visible = True
    Me.Line16Total.Visible = True
    Me.Line10Show.Visible = True
    Me.Line16Diff.Visible = True
    Me.Label40.Visible = True
    Me.Label45.Visible = True


End Sub

Private Sub Line16Check3_Click()
    If Line16Check3 = 1 Then
        Me.AlphaCheckLine16c = "X"
        Me.Line16Check1 = 0
        Me.Line16Check2 = 0
    Else
        Me.AlphaCheckLine16c = ""
    End If
    Me.Line16Mo1.Visible = False
    Me.Line16Mo2.Visible = False
    Me.Line16Mo3.Visible = False
    Me.Line16Total.Visible = False
    Me.Line10Show.Visible = False
    Me.Line16Diff.Visible = False
    Me.Label40.Visible = False
    Me.Label45.Visible = False

End Sub

Private Sub Line17Check_Click()
    If Line17Check = 1 Then
        Me.AlphaCheckLine17 = "X"
        Line17Date.Visible = True
        Line17Date = Int(Now())
    Else
        Me.AlphaCheckLine17 = ""
        Line17Date.Visible = False
    End If
End Sub

Private Sub Line18a_Click()
    
    If Line18a = 1 Then
        Me.AlphaCheckLine18a = "X"
    Else
        Me.AlphaCheckLine18a = ""
    End If

End Sub

Private Sub Line18b_Click()
    
    If Line18b = 1 Then
        Me.AlphaCheckLine18b = "X"
    Else
        Me.AlphaCheckLine18b = ""
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
    
    PrtTest = (User.Logon = "jim")
    
    PrtInit ("Port")
    SetFont 10, Equate.Portrait

    HorzNudge = Me.tdbNumHorzNudge.Value
    VertNudge = Me.tdbNumVertNudge.Value
    
    SaveNudge User.ID, "941_2021June"
    
    Me.KeyPreview = True
    
    Form941A2021Jun
        
    If Me.Line16Check3 = 1 Or PrtTest Then
        FormFeed
        
        VertNudge = VertNudge + 2
        HorzNudge = HorzNudge + 0
        
        Form941BHdr_2017 Me, Me.cmbYear.text
        
'        Form941BPrint 2300, Me.fgMo1, BMo1Tax
'        Form941BPrint 6400, Me.fgMo2, BMo2Tax
'        Form941BPrint 10500, Me.fgMo3, BMo3Tax
    
        ' twk for eagl 07/02/10
'        Form941BPrint_2017 4785, Me.fgMo1, BMo1Tax
'        Form941BPrint_2017 7855, Me.fgMo2, BMo2Tax
'        Form941BPrint_2017 11010, Me.fgMo3, BMo3Tax
    
        ' 2021 941
        Form941BPrint_2017 4580, Me.fgMo1, BMo1Tax
        Form941BPrint_2017 7650, Me.fgMo2, BMo2Tax
        Form941BPrint_2017 10805, Me.fgMo3, BMo3Tax
    
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
    For I = 1 To fg.Rows - 1
        For J = 1 To 3
            If J = 1 Then K = GetCol("Show1")
            If J = 2 Then K = GetCol("Show2")
            If J = 3 Then K = GetCol("Show3")
            If fg.TextMatrix(I, K) = "True" Then
                fg.TextMatrix(I, J) = "0.00"
            End If
        Next J
    Next I
                
    SSTax = 0
    MatchSS = 0
    MedTax = 0
    MedAddTl = 0
    MatchSSTotal = 0
                
    Part4Pin = " "
    PrepDate = Int(Now())
                
    Me.Line16Mo1 = 0
    Me.Line16Mo2 = 0
    Me.Line16Mo3 = 0
    
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
    StartYM = CLng(Me.cmbYear.text) * 100 + (Me.cmbQtr.ListIndex * 3) + 1
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
    ct = 0
    
    Do
    
        If ct Mod 10 = 1 Then
            frmProgress.lblMsg2 = Me.cmbQtr & Me.cmbYear & " " & _
                                 Format(ct, CountFormat) & " of: " & Format(Recs, CountFormat)
            frmProgress.Refresh
        End If
    
        SSTax = SSTax + PRHist.SSTax
        
        ' bozo
        MedTax = MedTax + PRHist.MedTax - PRHist.MedAddAmt
        MedAddTl = MedAddTl + PRHist.MedAddAmt
        
        ' match SS# logic
        ' 2011 ER is still 6.2% / EE is 4.2%
        ' 2013 - straight match for SS
'        ERSSTax = (Round(PRHist.SSWage * 0.062, 2))
'        MatchSS = MatchSS + ERSSTax
'        MatchSS = MatchSS + PRHist.SSTax
        
        
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
        Add941Val " 5d)", 1, PRHist.MedAddAmt / 0.009
        
        ' *** Line7b - sick pay ***
        ' *** Line7c - tips and group ins ***
        ' *** Line9 EIC payments ***
    
        ' tax liability per month
        ' use the 2011 match logic
        ' 2013 - SS straight match
        TaxLiab = PRHist.FWTTax + (PRHist.SSTax * 2) + ((PRHist.MedTax - PRHist.MedAddAmt) * 2) + PRHist.MedAddAmt
        '' TaxLiab = PRHist.FWTTax + PRHist.SSTax + ERSSTax + PRHist.MedTax * 2
        
        If PRHist.YearMonth Mod 100 = Month1 Then
            Line16Mo1 = Line16Mo1 + TaxLiab
            BGridPop Me.fgMo1, TaxLiab, Day(PRHist.CheckDate)
        ElseIf PRHist.YearMonth Mod 100 = Month2 Then
            Line16Mo2 = Line16Mo2 + TaxLiab
            BGridPop Me.fgMo2, TaxLiab, Day(PRHist.CheckDate)
        Else
            Line16Mo3 = Line16Mo3 + TaxLiab
            BGridPop Me.fgMo3, TaxLiab, Day(PRHist.CheckDate)
        End If
        If Line17Check = 1 Then
            Line17Date.Visible = True
            Line17Date = Int(Now())
        Else
            Line17Date.Visible = False
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

    Set941Val " 5a)", 2, Round(Get941Val(" 5a)", 1) * 0.124, 2)
    Set941Val " 5ai)", 2, Round(Get941Val(" 5ai)", 1) * 0.062, 2)
    Set941Val " 5aii)", 2, Round(Get941Val(" 5aii)", 1) * 0.062, 2)
    Set941Val " 5b)", 2, Round(Get941Val(" 5b)", 1) * 0.124, 2)
    Set941Val " 5c)", 2, Round(Get941Val(" 5c)", 1) * 0.029, 2)
    Set941Val " 5d)", 2, Round(Get941Val(" 5d)", 1) * 0.009, 2)
            
    ' 2021-10-28 - subtract 5ai & 5aii
    Cur = Get941Val(" 5a)", 2) - Get941Val(" 5ai)", 2) - Get941Val(" 5aii)", 2) + Get941Val(" 5b)", 2) + Get941Val(" 5c)", 2) + Get941Val(" 5d)", 2)
    Set941Val " 5e)", 3, Cur
    
    Cur = Get941Val(" 3 )", 3) + Get941Val(" 5e)", 3) + Get941Val(" 5f)", 3)
    Set941Val " 6 )", 3, Cur
    
    ' 7) fraction of cents
    If Me.chkManualFractions = 0 Then
        Cur = SSTax * 2 - Get941Val(" 5a)", 2) - Get941Val(" 5b)", 2) + MedTax * 2 - Get941Val(" 5c)", 2) + MedAddTl - Get941Val(" 5d)", 2)
        Set941Val " 7 )", 3, Round(Cur, 2)
    End If
    
'    If Me.chkCents = 0 Then
'        Me.Line7a = Round(SSTax * 2 - Me.Line5aa - Me.Line5bb + MedTax * 2 - Me.Line5cc, 2)
'    End If
    
    ' 10) total of taxes after adjustments
    Cur = Get941Val(" 6 )", 3) + Get941Val(" 7 )", 3) + Get941Val(" 8 )", 3) + Get941Val(" 9 )", 3)
    Set941Val "10 )", 3, Cur
    
    ' 11g) Total Refundable Credits
    Cur = Get941Val("11a)", 3) + Get941Val("11b)", 3) + Get941Val("11c)", 3) + Get941Val("11d)", 3) + Get941Val("11e)", 3)
    Set941Val "11g)", 3, Cur
    
    ' 12) total of taxes after adjustments and credits
    Cur = Get941Val("10 )", 3) - Get941Val("11g)", 3)
    Set941Val "12 )", 3, Cur
    
    ' 13g) Total Deposits and refundable credits
    Cur = Get941Val("13a)", 3) + Get941Val("13c)", 3) + Get941Val("13d)", 3) + Get941Val("13e)", 3) + Get941Val("13f)", 3)
    Set941Val "13g)", 3, Cur
    
    ' 13i) Total deposits and refundable credits less advances
    Cur = Get941Val("13g)", 3) - Get941Val("13h)", 3)
    Set941Val "13i)", 3, Cur
    
    ' 14) Balance Due
    If Get941Val("12 )", 3) > Get941Val("13i)", 3) Then
        Cur = Get941Val("12 )", 3) - Get941Val("13i)", 3)
        Set941Val "14 )", 3, Cur
        Set941Val "15 )", 2, 0
        Me.Line15Check1.Enabled = False
        Me.Line15Check2.Enabled = False
    Else        ' overpayment
        Cur = Get941Val("13i)", 3) - Get941Val("12 )", 3)
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
    Line16Total = Me.Line16Mo1 + Me.Line16Mo2 + Me.Line16Mo3
    Line16Diff = Get941Val("10 )", 3) - Line16Total

    BLine10Show = Get941Val("10 )", 3)
    BDifference = BLine10Show - BTotalTax

    BGridUpdate Me.fgMo1, BMo1Tax
    BGridUpdate Me.fgMo2, BMo2Tax
    BGridUpdate Me.fgMo3, BMo3Tax
    
    BTotalTax = BMo1Tax + BMo2Tax + BMo3Tax
    
    If Get941Val("10 )", 3) < 2500 Then
        Me.Line16Check1 = 1
        ' Me.Line17Check2.Enabled = False
    End If

End Sub

Private Sub BGridUpdate(ByRef fg As VSFlexGrid, ByRef MonthTotal As TDBNumber)
    
Dim CellValue As Currency
Dim I, J As Integer
    
    MonthTotal = 0
    For I = 1 To 8
        For J = 1 To 7 Step 2
            If fg.TextMatrix(I, J) <> "" Then
                CellValue = 0
                On Error Resume Next        ' turn of error handling
                CellValue = CCur(fg.TextMatrix(I, J))
                On Error GoTo 0             ' turn error handling back on
                MonthTotal = MonthTotal + CellValue
            End If
        Next J
    Next I

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
    Date12 = DateSerial(Me.cmbYear.text, (Me.cmbQtr.ListIndex * 3) + 3, 12)
    DateDiff = 99999
    
    ' get start/end date for the quarter
    StartYM = CLng(Me.cmbYear.text) * 100 + (Me.cmbQtr.ListIndex * 3) + 1
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
Dim I As Long
    
    Part4ID = 0
    Part5ID = 0
    PaidPrepID = 0
    txtTradeName = ""
    
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
        Me.txtTradeName = PRGlobal.Var4
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
        Me.cmbPrepName.text = PRGlobal.Var9
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
Private Sub Line16Mo1_lostfocus()
    Me.Line16Total = Me.Line16Mo1 + Me.Line16Mo2 + Me.Line16Mo3
    Calc941Data
End Sub
Private Sub Line17Mo2_Change()
    Me.Line16Total = Me.Line16Mo1 + Me.Line16Mo2 + Me.Line16Mo3
    Calc941Data
End Sub

Private Sub Line17Mo3_lostfocus()
    Me.Line16Total = Me.Line16Mo1 + Me.Line16Mo2 + Me.Line16Mo3
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
' Rev Jun 2021
'
Public Sub Form941A2021Jun()

Dim VertSpace, VertPosn, HorzPosn, HorzPosn1, HorzPosn2 As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString, TelFmtString, ReportTitle As String
Dim ff, FedID As String
Dim Xincr, XXpos As Long
 
    TestPattern941
 
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
    
    With frm941_2021_June
    
        TestPattern941
        
        ' %%%%% Pg 1 Panel 1 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        PanelVert = 800
        
        ' PosPrint 3200, 1020, PRCompany.FederalID
        ' formatting for the GAY fed id boxes
        HorzPosn = 2410
        Xincr = 499
        FedID = Trim(PRCompany.FederalID)
        For XXpos = 1 To Len(FedID)
            ff = Mid(FedID, XXpos, 1)
            If ff <> "-" Then
                HorzPosn = HorzPosn + Xincr
                PosPrint HorzPosn, 170 + PanelVert, ff
            End If
            If XXpos = 2 Then
                HorzPosn = HorzPosn + 315
            End If
        Next XXpos
        
        ' PosPrint 2500, 1490, PRCompany.Name
        PosPrint 2530, 630 + PanelVert, Me.txtName
        
        PosPrint 2120, 1120 + PanelVert, Me.txtTradeName
        
        PosPrint 1430, 1580 + PanelVert, Trim(PRCompany.Address1) & " " & Trim(PRCompany.Address2)
        
'        If PRCompany.Address2 <> "" Then
'            PosPrint 2400, 2450, PRCompany.Address2
'        End If
        
        PosPrint 1430, 2150 + PanelVert, PRCompany.City
        If PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
            PosPrint 5290, 2150 + PanelVert, PRState.StateAbbrev
        End If
        PosPrint 6130, 2150 + PanelVert, PRCompany.ZipCode
    
        ' checkbox for the quarter
        Dim qNum As Integer
        qNum = .cmbQtr
        ' qNum = 2
        Dim qtr941 As Integer
        VertPosn = 610 + PanelVert + (qNum - 1) * 360
        PosPrint 8263, VertPosn, "X"
        If PrtTest Then
            For qtr941 = 0 To 3
                VertPosn = 610 + PanelVert + qtr941 * 360
                PosPrint 8263, VertPosn, "X"
            Next qtr941
        End If
        
        ' %%%%% Pg 1 Panel 1 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        ' june 2021
        VertPosn = 4910
        VertSpace = 490
        
        PosPrint 10200, VertPosn, PadRight(Format(Get941Val(" 1 )", 3), "##,##0"), 6)
        
        VertPosn = VertPosn + VertSpace
        ' PosPrint 9400, VertPosn, PadRight(Format(Get941Val(" 2 )", 3), FmtString), 13)
        PosPrint 9320, VertPosn, PadRight(DollarAndCents(Get941Val(" 2 )", 3)), 15)
        
        VertPosn = VertPosn + VertSpace
        PosPrint 9320, VertPosn, PadRight(DollarAndCents(Get941Val(" 3 )", 3)), 15)
        
        VertPosn = VertPosn + VertSpace + 35
        PosPrint 8700, VertPosn, IIf(Not PrtTest, .AlphaCheckLine4, "X")
    
        VertPosn = 6990
        
        VertSpace = 375
        HorzPosn1 = 3980
        HorzPosn2 = 6745
    
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5a)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5a)", 2)), 15)
        
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5ai)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5ai)", 2)), 15)
        
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5aii)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5aii)", 2)), 15)
        
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5b)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5b)", 2)), 15)
        
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5c)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5c)", 2)), 15)
        
        ' new for 2013 - addl med fields
        VertPosn = VertPosn + VertSpace + 70
        PosPrint HorzPosn1, VertPosn, PadRight(DollarAndCents(Get941Val(" 5d)", 1)), 15)
        PosPrint HorzPosn2, VertPosn, PadRight(DollarAndCents(Get941Val(" 5d)", 2)), 15)
        
        VertPosn = 9400
        VertSpace = 482
        HorzPosn = 9270
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 5e)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 5f)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 6 )", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        '' PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 7 )", 3)), 15)
        PosPrint HorzPosn - 80, VertPosn, PadRight(Get941Val(" 7 )", 3), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 8 )", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val(" 9 )", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("10 )", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11a)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        VertPosn = VertPosn + 50
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11b)", 3)), 15)
        VertPosn = VertPosn + VertSpace + 30
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11c)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        '   #######################  FORM 941 - PAGE 2  ####################################
        '
        FormFeed
        
        TestPattern941
        
        PosPrint 900, 800, PRCompany.Name
        PosPrint 8490, 800, PRCompany.FederalID

        VertPosn = 1580
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11d)", 3)), 15)
        VertPosn = VertPosn + VertSpace + 90
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11e)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint 5720, VertPosn, PadRight(Format(Get941Val("11f)", 2), "##,##0"), 6)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("11g)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("12 )", 3)), 15)
        VertPosn = VertPosn + VertSpace + 90
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13a)", 3)), 15)
        VertPosn = VertPosn + (VertSpace * 2) + 90
        
        ' 13b reserved for future use
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13c)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13d)", 3)), 15)
        VertPosn = VertPosn + VertSpace + 90
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13e)", 3)), 15)
        VertPosn = VertPosn + VertSpace + 90
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13f)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13g)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13h)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("13i)", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("14 )", 3)), 15)
        VertPosn = VertPosn + VertSpace
        
        PosPrint 5720, VertPosn, PadRight(DollarAndCents(Get941Val("15 )", 2)), 15)

'        PosPrint 9730, 13900, .AlphaCheckLine15a
'        PosPrint 9730, 14170, .AlphaCheckLine15b

'        PosPrint 9730, VertPosn - 70, .AlphaCheckLine15a
'        PosPrint 9730, VertPosn + 50, .AlphaCheckLine15b

        PosPrint 8780, VertPosn, IIf(Not PrtTest, AlphaCheckLine15a, "X")
        PosPrint 10170, VertPosn, IIf(Not PrtTest, AlphaCheckLine15b, "X")

        ' *** FIX ***
        'VertNudge = VertNudge + 6
        'HorzNudge = HorzNudge + 2

        ' PosPrint 950, 2100, .Line16 - was the state on 2011 form

        ' %%%%%%%%%%%%% Part 2 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        PanelVert = 10000

        ' stuff it!!!!
        If Me.Line16Check1 Then
            AlphaCheckLine16a = "X"
            AlphaCheckLine16b = ""
            AlphaCheckLine16c = ""
        End If
        If Me.Line16Check2 Then
            AlphaCheckLine16a = ""
            AlphaCheckLine16b = "X"
            AlphaCheckLine16c = ""
        End If
        If Me.Line16Check3 Then
            AlphaCheckLine16a = ""
            AlphaCheckLine16b = ""
            AlphaCheckLine16c = "X"
        End If

        PosPrint 1980, 430 + PanelVert, AlphaCheckLine16a
        PosPrint 1980, 1370 + PanelVert, AlphaCheckLine16b
        PosPrint 1980, 3630 + PanelVert, AlphaCheckLine16c
        If PrtTest Then
            PosPrint 1980, 430 + PanelVert, "X"
            PosPrint 1980, 1370 + PanelVert, "X"
            PosPrint 1980, 3630 + PanelVert, "X"
        End If

        If .Line16Check2 = 1 Or PrtTest Then
            HorzPosn = 5010
            VertPosn = 1990 + PanelVert
            VertSpace = 440
            PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(IIf(Not PrtTest, .Line16Mo1, 99999.99)), 15)
            VertPosn = VertPosn + VertSpace
            PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(IIf(Not PrtTest, .Line16Mo2, 99999.99)), 15)
            VertPosn = VertPosn + VertSpace - 15  ' nice...
            PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(IIf(Not PrtTest, .Line16Mo3, 99999.99)), 15)
            VertPosn = VertPosn + VertSpace
            PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(IIf(Not PrtTest, .Line16Total, 99999.99)), 15)
        End If
        
        ' %%%%%%%%%%%%% Part 2 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        '   #######################  FORM 941 - PAGE 3  ####################################
        '
        
        FormFeed

        TestPattern941
        
        ' %%%%%% page 3 panel 1 - lines 17 to 28 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        PanelVert = 410
        
        PosPrint 900, 365 + PanelVert, PRCompany.Name
        PosPrint 8490, 365 + PanelVert, PRCompany.FederalID

        ' Part 3
        PosPrint 9390, 1000 + PanelVert, IIf(Not PrtTest, .AlphaCheckLine17, "X")
        If .Line17Check = 1 Then
            If IsNull(.Line17Date) = False Then
                PosPrint 3890, 1400 + PanelVert, DateSplit(.Line17Date)
            End If
        End If
        If PrtTest Then PosPrint 3890, 1400 + PanelVert, DateSplit(Date)
        
        ' seasonal
        PosPrint 9390, 1825 + PanelVert, IIf(Not PrtTest, .AlphaCheckLine18a, "X")
        ' startup
        PosPrint 9390, 2180 + PanelVert, IIf(Not PrtTest, .AlphaCheckLine18b, "X")

        VertSpace = 360
        VertPosn = 2175 + PanelVert
        HorzPosn = 9330
        
        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("19 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("20 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("21 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("22 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("23 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("24 )", 3)), 15)

        VertPosn = VertPosn + VertSpace + 90
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("25 )", 3)), 15)

        VertPosn = VertPosn + VertSpace + 120
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("26 )", 3)), 15)

        VertPosn = VertPosn + VertSpace
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("27 )", 3)), 15)

        VertPosn = VertPosn + VertSpace + 110
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Get941Val("28 )", 3)), 15)
    
        ' %%%%%% page 3 panel 1 - lines 17 to 28 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        Form941Pt4Pt5_2017 frm941_2021_June


' ????????
'        PosPrint 2600, 14230, .AlphaCheckPart5
        
        ' *** put it back (for sched B) ***
        'VertNudge = VertNudge - 6
        'HorzNudge = HorzNudge - 2
    
    End With

End Sub

Public Sub Form941Pt4Pt5_2017(ByRef frm As Form)
    
Dim VertSpace, VertPosn, HorzPosn As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString, TelFmtString, ReportTitle As String
Dim HorzPosn1, HorzPosn2, Xincr As Long
    
    PRGlobal.Var1 = ""
    PRGlobal.Var2 = ""
    PRGlobal.Var3 = ""
    PRGlobal.Var4 = ""
    PRGlobal.Var5 = ""
    PRGlobal.Var6 = ""
    PRGlobal.Var7 = ""
    
    ' %%%%%%%%%%% Part 4 - Third Party Designee - Per User %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    
    PanelVert = 7100
    
    If Not PrtTest Then
        If frm.Part4CheckYes = 1 Then
            PosPrint 1060, 635 + PanelVert, "X"
        Else
            PosPrint 1060, 1430 + PanelVert, "X"
        End If
    Else
        PosPrint 1060, 635 + PanelVert, "X"
        PosPrint 1060, 1430 + PanelVert, "X"
    End If
    
    If frm.Part4ID <> 0 Then
        If PRGlobal.GetByID(frm.Part4ID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part4
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
           
    PRGlobal.Var1 = frm.Part4Name
    PRGlobal.Var2 = frm.Part4Phone
    PRGlobal.Var3 = frm.Part4Pin
    PRGlobal.Var4 = frm.txtTradeName
    PRGlobal.Save (Equate.RecPut)
    
    If frm.Part4CheckYes Or PrtTest Then
        VertPosn = 650 + PanelVert
        PosPrint 5210, VertPosn, PRGlobal.Var1
        PosPrint 9060, VertPosn, PRGlobal.Var2
        
        HorzPosn = 8480
        Xincr = 445
        VertPosn = 1070 + PanelVert
        
        ' part 4 PIN in gay boxes
        x = Trim(PRGlobal.Var3)
        For I = 1 To 5
            If Len(x) >= I Then
                PosPrint HorzPosn, VertPosn, Mid(x, I, 1)
            End If
            HorzPosn = HorzPosn + Xincr
        Next I
        
    End If
           
    ' Part 5 - Company Signature - Per Company
    If frm.Part5ID <> 0 Then
        If PRGlobal.GetByID(frm.Part5ID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part5
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If

    PRGlobal.Var1 = frm.Part5NameTitle

    PRGlobal.Save (Equate.RecPut)

    ' part 5 - name and title - split the string on the slash
    
    PanelVert = 6970
    
    PosPrint 8565, 2720 + PanelVert, SlashSplit(PRGlobal.Var1, 1)
    PosPrint 8565, 3230 + PanelVert, SlashSplit(PRGlobal.Var1, 2)
    PosPrint 2890, 3840 + PanelVert, DateSplit(frm.Part5Date)
    PosPrint 9100, 3840 + PanelVert, Trim(frm.Part5Phone)
        
    'Paid Preparer - Per User
    If frm.PaidPrepID <> 0 Then
        If PRGlobal.GetByID(frm.PaidPrepID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941PaidPrep
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
               
    PRGlobal.Var1 = frm.PrepFirm
    PRGlobal.Var2 = frm.PrepAddr1
    PRGlobal.Var3 = frm.PrepAddr2
    PRGlobal.Var4 = frm.PrepPhone
    PRGlobal.Var5 = frm.PrepEIN
    PRGlobal.Var6 = frm.PrepZip
    PRGlobal.Var7 = frm.PrepSSN
    If frm.PrepCheck Then
        PRGlobal.Var8 = "1"
    Else
        PRGlobal.Var8 = "0"
    End If
    PRGlobal.Var9 = frm.cmbPrepName.text
    PRGlobal.Save (Equate.RecPut)
    
    VertPosn = 4800 + PanelVert
    VertSpace = 485
    HorzPosn1 = 2405
    HorzPosn2 = 8805
    
    PosPrint HorzPosn1, VertPosn, frm.cmbPrepName
    PosPrint HorzPosn2, VertPosn, frm.PrepSSN               ' PTIN
    
    VertPosn = VertPosn + VertSpace
    PosPrint HorzPosn2, VertPosn, " " & DateSplit(frm.PrepDate)
    
    VertPosn = VertPosn + VertSpace
    PosPrint HorzPosn1, VertPosn, PRGlobal.Var1      ' firm
    PosPrint HorzPosn2, VertPosn, PRGlobal.Var5      ' EIN
    
    VertPosn = VertPosn + VertSpace
    PosPrint HorzPosn1, VertPosn, PRGlobal.Var2      ' addr
    PosPrint HorzPosn2, VertPosn, PRGlobal.Var4      ' phone
    
    VertPosn = VertPosn + VertSpace
    PosPrint HorzPosn1, VertPosn, SlashSplit(PRGlobal.Var3, 1)              ' city
    PosPrint HorzPosn1 + 4300, VertPosn, SlashSplit(PRGlobal.Var3, 2)        ' state
    PosPrint HorzPosn2, VertPosn, PRGlobal.Var6                             ' zip code
    
   
End Sub

Public Sub Form941BHdr_2017(ByRef frm As Form, ByVal TaxYear As String)
    
Dim FmtString As String
    
    FmtString = "##,###,##0.00"

Dim HP As Integer
Dim yy1, yy2, FedID, ff As String
Dim HorzPosn, Xincr, XXpos As Integer

    With frm

        HP = 8220
        CurrYear = Year(Now())
        If .cmbQtr = 1 Then
            PosPrint HP, 1900, "X"
        ElseIf .cmbQtr = 2 Then
            PosPrint HP, 2150, "X"
        ElseIf .cmbQtr = 3 Then
            PosPrint HP, 2430, "X"
        ElseIf .cmbQtr = 4 Then
            PosPrint HP, 2660, "X"
        End If
    
        ' PosPrint 3380, 900, PRCompany.FederalID
        ' formatting for the GAY fed id boxes
        HorzPosn = 2310
        Xincr = 499
        FedID = Trim(PRCompany.FederalID)
        For XXpos = 1 To Len(FedID)
            ff = Mid(FedID, XXpos, 1)
            If ff <> "-" Then
                HorzPosn = HorzPosn + Xincr
                PosPrint HorzPosn, 1400, ff
            End If
            If XXpos = 2 Then
                HorzPosn = HorzPosn + 315
            End If
        Next XXpos
        
        PosPrint 2500, 1870, PRCompany.Name
        
        ' GAY boxes for tax year
        ' PosPrint 3380, 1430, TaxYear
        HorzPosn = 2310
        Xincr = 499
        
        ' WTF
        yy1 = Trim(TaxYear)
        For XXpos = 1 To Len(yy1)
            yy2 = Mid(yy1, XXpos, 1)
            HorzPosn = HorzPosn + Xincr
            PosPrint HorzPosn, 2360, yy2
        Next XXpos
        
        PosPrint 9355, 13880, PadRight(DollarAndCents(.BTotalTax), 15)
        
    End With
    
End Sub

Public Sub Form941BPrint_2017(ByVal VertPos As Long, ByRef fg As VSFlexGrid, ByVal BMoTax As Currency)

Dim VertSpace As Long
Dim Col1X, Col2X, Col3X, Col4X As Long
Dim FmtString As String

    ' SetEquates
    ' Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)

'    SQLString = "SELECT * FROM PREmployee"
'    rsInit SQLString, cn, rs941

    VertSpace = 360
    FmtString = "##,###,##0.00"

    Col1X = 795
    Col2X = 2835
    Col3X = 4830
    Col4X = 6875

    ' month total
    PosPrint 9355, VertPos + 520, PadRight(DollarAndCents(BMoTax), 15)
    
    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(1, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(1, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(1, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(1, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(2, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(2, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(2, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(2, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(3, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(3, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(3, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(3, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(4, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(4, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(4, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(4, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(5, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(5, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(5, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(5, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(6, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(6, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(6, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(6, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(7, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(7, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(7, 5)), 13)
    PosPrint Col4X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(7, 7)), 13)
    VertPos = VertPos + VertSpace

    PosPrint Col1X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(8, 1)), 13)
    PosPrint Col2X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(8, 3)), 13)
    PosPrint Col3X, VertPos, PadRight(DollarAndCents(fg.TextMatrix(8, 5)), 13)

End Sub

Sub TestPattern941()
    If Not PrtTest Then Exit Sub
    Exit Sub
    Dim trow, tcol, CCount As Integer
    For trow = 100 To 15000 Step 100
        If trow Mod 2000 = 0 Then
            CCount = 0
            For tcol = 100 To 11000 Step 100
                CCount = CCount + 1
                PosPrint tcol, trow, CCount Mod 10
            Next tcol
            CCount = 0
            For tcol = 100 To 12000 Step 1000
                CCount = CCount + 1
                If CCount > 1 Then
                    PosPrint tcol - 100, trow + 100, CCount - 1 Mod 10
                End If
            Next tcol
        End If
        PosPrint 100, trow, Int(trow / 100)
    Next trow
End Sub
