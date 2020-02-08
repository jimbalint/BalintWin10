VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm941Entry 
   Caption         =   "Form 941 for 2008"
   ClientHeight    =   9285
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   11745
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCalc 
      Caption         =   "CALC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   9360
      TabIndex        =   118
      Top             =   65
      Width           =   735
   End
   Begin VB.ComboBox cmbChkDate12 
      BackColor       =   &H00FFFF80&
      Height          =   315
      Left            =   7560
      TabIndex        =   117
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
      TabIndex        =   33
      Top             =   0
      Width           =   3855
   End
   Begin VB.ComboBox cmbQtr 
      Height          =   315
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   -10
      Width           =   615
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   250
      Left            =   3960
      TabIndex        =   65
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   11040
      TabIndex        =   69
      Top             =   65
      Width           =   690
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   10200
      TabIndex        =   67
      Top             =   65
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   63
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   16748
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
      TabCaption(0)   =   "Form 941"
      TabPicture(0)   =   "frm941.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(6)=   "Label33"
      Tab(0).Control(7)=   "Label37"
      Tab(0).Control(8)=   "Label38"
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(12)=   "Label21"
      Tab(0).Control(13)=   "Image1"
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(15)=   "Label5"
      Tab(0).Control(16)=   "Label18"
      Tab(0).Control(17)=   "Line1"
      Tab(0).Control(18)=   "Line5a"
      Tab(0).Control(19)=   "Line15"
      Tab(0).Control(20)=   "Line13"
      Tab(0).Control(21)=   "Line10"
      Tab(0).Control(22)=   "Line12b"
      Tab(0).Control(23)=   "Line12a"
      Tab(0).Control(24)=   "Line7c"
      Tab(0).Control(25)=   "Line7b"
      Tab(0).Control(26)=   "Line6"
      Tab(0).Control(27)=   "Line7d"
      Tab(0).Control(28)=   "Line7a"
      Tab(0).Control(29)=   "Line14"
      Tab(0).Control(30)=   "Line11"
      Tab(0).Control(31)=   "Line9"
      Tab(0).Control(32)=   "Line8"
      Tab(0).Control(33)=   "Line5d"
      Tab(0).Control(34)=   "Line5cc"
      Tab(0).Control(35)=   "Line5c"
      Tab(0).Control(36)=   "Line5bb"
      Tab(0).Control(37)=   "Line5b"
      Tab(0).Control(38)=   "Line2"
      Tab(0).Control(39)=   "Line3"
      Tab(0).Control(40)=   "Line5aa"
      Tab(0).Control(41)=   "Line4"
      Tab(0).Control(42)=   "Line15Check2"
      Tab(0).Control(43)=   "Line10Total"
      Tab(0).Control(44)=   "Line15Check1"
      Tab(0).ControlCount=   45
      TabCaption(1)   =   "Form 941   Page 2"
      TabPicture(1)   =   "frm941.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Part4Phone"
      Tab(1).Control(1)=   "Part4Pin"
      Tab(1).Control(2)=   "Part4Name"
      Tab(1).Control(3)=   "Part4CheckYes"
      Tab(1).Control(4)=   "Line17Diff"
      Tab(1).Control(5)=   "Line10Show"
      Tab(1).Control(6)=   "txtName"
      Tab(1).Control(7)=   "Line19"
      Tab(1).Control(8)=   "Line18Date"
      Tab(1).Control(9)=   "Line18Check"
      Tab(1).Control(10)=   "Part4CheckNo"
      Tab(1).Control(11)=   "Line17Check3"
      Tab(1).Control(12)=   "Line17Check2"
      Tab(1).Control(13)=   "Line17Check1"
      Tab(1).Control(14)=   "Line16"
      Tab(1).Control(15)=   "txtEIN"
      Tab(1).Control(16)=   "Line17Mo2"
      Tab(1).Control(17)=   "Line17Mo3"
      Tab(1).Control(18)=   "Line17Total"
      Tab(1).Control(19)=   "Line17Mo1"
      Tab(1).Control(20)=   "Label54"
      Tab(1).Control(21)=   "Label53"
      Tab(1).Control(22)=   "Label52"
      Tab(1).Control(23)=   "Label51"
      Tab(1).Control(24)=   "Label50"
      Tab(1).Control(25)=   "Label49"
      Tab(1).Control(26)=   "Label47"
      Tab(1).Control(27)=   "Label46"
      Tab(1).Control(28)=   "Label45"
      Tab(1).Control(29)=   "Label40"
      Tab(1).Control(30)=   "Label39"
      Tab(1).Control(31)=   "Label44"
      Tab(1).Control(32)=   "Label43"
      Tab(1).Control(33)=   "Label42"
      Tab(1).Control(34)=   "Label36"
      Tab(1).Control(35)=   "Label17"
      Tab(1).Control(36)=   "Label6"
      Tab(1).Control(37)=   "Label48"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "Form 941   Pg 2  (Cont'd)"
      TabPicture(2)   =   "frm941.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label55"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label56"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label14"
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
      TabPicture(3)   =   "frm941.frx":0054
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
      Begin TDBText6Ctl.TDBText Part4Phone 
         Height          =   375
         Left            =   -73560
         TabIndex        =   128
         Top             =   7080
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "frm941.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":00CE
         Key             =   "frm941.frx":00EC
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
         TabIndex        =   127
         Top             =   7080
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   661
         Caption         =   "frm941.frx":0130
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":01CC
         Key             =   "frm941.frx":01EA
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
         TabIndex        =   51
         Top             =   1740
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   609
         Caption         =   "frm941.frx":022E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":028C
         Key             =   "frm941.frx":02AA
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
         Left            =   2080
         TabIndex        =   55
         Top             =   5000
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   609
         Caption         =   "frm941.frx":02EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":035A
         Key             =   "frm941.frx":0378
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
         TabIndex        =   52
         Top             =   3600
         Width           =   5750
      End
      Begin TDBText6Ctl.TDBText Part4Name 
         Height          =   375
         Left            =   -73560
         TabIndex        =   48
         Top             =   6600
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   661
         Caption         =   "frm941.frx":03BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":042E
         Key             =   "frm941.frx":044C
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
         TabIndex        =   47
         Top             =   6660
         Width           =   735
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Diff 
         Height          =   300
         Left            =   -66240
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0490
         Caption         =   "frm941.frx":04B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0516
         Keys            =   "frm941.frx":0534
         Spin            =   "frm941.frx":057E
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
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   529
         Calculator      =   "frm941.frx":05A6
         Caption         =   "frm941.frx":05C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0630
         Keys            =   "frm941.frx":064E
         Spin            =   "frm941.frx":0698
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
         TabIndex        =   114
         Top             =   600
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":06C0
         Caption         =   "frm941.frx":06E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0760
         Keys            =   "frm941.frx":077E
         Spin            =   "frm941.frx":07C8
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
         Left            =   -74520
         TabIndex        =   34
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
         TabIndex        =   113
         Top             =   -480
         Width           =   3855
      End
      Begin VB.CheckBox Line15Check1 
         Caption         =   "Apply to next return."
         Height          =   255
         Left            =   -65280
         TabIndex        =   31
         Top             =   8180
         Width           =   1750
      End
      Begin TDBNumber6Ctl.TDBNumber Line10Total 
         Height          =   300
         Left            =   -65910
         TabIndex        =   24
         Top             =   5700
         Width           =   1935
         _Version        =   65536
         _ExtentX        =   3404
         _ExtentY        =   529
         Calculator      =   "frm941.frx":07F0
         Caption         =   "frm941.frx":0810
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":087C
         Keys            =   "frm941.frx":089A
         Spin            =   "frm941.frx":08E4
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
         ValueVT         =   25427969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VSFlex8Ctl.VSFlexGrid fgMo1 
         Height          =   2325
         Left            =   -73800
         TabIndex        =   2
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
         FormatString    =   $"frm941.frx":090C
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
         TabIndex        =   61
         Top             =   5460
         Width           =   3495
      End
      Begin VB.CheckBox Line19 
         Caption         =   "Check here."
         Height          =   255
         Left            =   -65280
         TabIndex        =   46
         Top             =   5620
         Width           =   1215
      End
      Begin TDBDate6Ctl.TDBDate Line18Date 
         Height          =   285
         Left            =   -74400
         TabIndex        =   45
         Top             =   5220
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   503
         Calendar        =   "frm941.frx":09E6
         Caption         =   "frm941.frx":0AFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0BA0
         Keys            =   "frm941.frx":0BBE
         Spin            =   "frm941.frx":0C1C
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
         Caption         =   "Check here, and"
         Height          =   255
         Left            =   -65280
         TabIndex        =   44
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
         TabIndex        =   62
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
         TabIndex        =   43
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
         TabIndex        =   38
         Top             =   2040
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
         TabIndex        =   37
         Top             =   1755
         Width           =   7215
      End
      Begin TDBText6Ctl.TDBText Line16 
         Height          =   375
         Left            =   -74520
         TabIndex        =   36
         Top             =   1185
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Caption         =   "frm941.frx":0C44
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0CB0
         Key             =   "frm941.frx":0CCE
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
         TabIndex        =   35
         Top             =   735
         Width           =   3855
      End
      Begin VB.CheckBox Line15Check2 
         Caption         =   "Send refund check."
         Height          =   255
         Left            =   -65280
         TabIndex        =   32
         Top             =   8385
         Width           =   1750
      End
      Begin VB.CheckBox Line4 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -65940
         TabIndex        =   8
         Top             =   1485
         Width           =   255
      End
      Begin TDBNumber6Ctl.TDBNumber Line5aa 
         Height          =   300
         Left            =   -68160
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1935
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0D12
         Caption         =   "frm941.frx":0D32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0D9E
         Keys            =   "frm941.frx":0DBC
         Spin            =   "frm941.frx":0E06
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
         ValueVT         =   42467329
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber Line3 
         Height          =   300
         Left            =   -66255
         TabIndex        =   7
         Top             =   1065
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0E2E
         Caption         =   "frm941.frx":0E4E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0EB4
         Keys            =   "frm941.frx":0ED2
         Spin            =   "frm941.frx":0F1C
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
         ShowContextMenu =   1
         ValueVT         =   24641537
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line2 
         Height          =   300
         Left            =   -66255
         TabIndex        =   6
         Top             =   765
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0F44
         Caption         =   "frm941.frx":0F64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0FCA
         Keys            =   "frm941.frx":0FE8
         Spin            =   "frm941.frx":1032
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
         ShowContextMenu =   1
         ValueVT         =   42467329
         Value           =   0
         MaxValueVT      =   5636101
         MinValueVT      =   3342341
      End
      Begin TDBNumber6Ctl.TDBNumber Line5b 
         Height          =   300
         Left            =   -74325
         TabIndex        =   11
         Top             =   2220
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   529
         Calculator      =   "frm941.frx":105A
         Caption         =   "frm941.frx":107A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":111A
         Keys            =   "frm941.frx":1138
         Spin            =   "frm941.frx":1182
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
         Format          =   "$ ###,###.##"
         HighlightText   =   1
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
         ValueVT         =   25427969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line5bb 
         Height          =   300
         Left            =   -68160
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2220
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":11AA
         Caption         =   "frm941.frx":11CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1236
         Keys            =   "frm941.frx":1254
         Spin            =   "frm941.frx":129E
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
         ValueVT         =   34144257
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber Line5c 
         Height          =   300
         Left            =   -74325
         TabIndex        =   13
         Top             =   2505
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   529
         Calculator      =   "frm941.frx":12C6
         Caption         =   "frm941.frx":12E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":138C
         Keys            =   "frm941.frx":13AA
         Spin            =   "frm941.frx":13F4
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
         Format          =   "$ ###,###.##"
         HighlightText   =   1
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
         ValueVT         =   25427969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line5cc 
         Height          =   300
         Left            =   -68160
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2505
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":141C
         Caption         =   "frm941.frx":143C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":14A8
         Keys            =   "frm941.frx":14C6
         Spin            =   "frm941.frx":1510
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
         ValueVT         =   34144257
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo2 
         Height          =   300
         Left            =   -71760
         TabIndex        =   40
         Top             =   2940
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1538
         Caption         =   "frm941.frx":1558
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":15C2
         Keys            =   "frm941.frx":15E0
         Spin            =   "frm941.frx":162A
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
         ValueVT         =   32047105
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo3 
         Height          =   300
         Left            =   -71760
         TabIndex        =   41
         Top             =   3300
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1652
         Caption         =   "frm941.frx":1672
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":16DC
         Keys            =   "frm941.frx":16FA
         Spin            =   "frm941.frx":1744
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
         ValueVT         =   32047105
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Total 
         Height          =   300
         Left            =   -73245
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3660
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8184
         _ExtentY        =   529
         Calculator      =   "frm941.frx":176C
         Caption         =   "frm941.frx":178C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":181E
         Keys            =   "frm941.frx":183C
         Spin            =   "frm941.frx":1886
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
         ValueVT         =   32047105
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line17Mo1 
         Height          =   300
         Left            =   -71760
         TabIndex        =   39
         Top             =   2580
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941.frx":18AE
         Caption         =   "frm941.frx":18CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1938
         Keys            =   "frm941.frx":1956
         Spin            =   "frm941.frx":19A0
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
         ValueVT         =   32047105
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText Part5NameTitle 
         Height          =   345
         Left            =   120
         TabIndex        =   49
         Top             =   1260
         Width           =   11055
         _Version        =   65536
         _ExtentX        =   19500
         _ExtentY        =   609
         Caption         =   "frm941.frx":19C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1A4C
         Key             =   "frm941.frx":1A6A
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
         TabIndex        =   53
         Top             =   4020
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941.frx":1AAE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1B20
         Key             =   "frm941.frx":1B3E
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
         TabIndex        =   54
         Top             =   4500
         Width           =   7575
         _Version        =   65536
         _ExtentX        =   13361
         _ExtentY        =   609
         Caption         =   "frm941.frx":1B82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1BEC
         Key             =   "frm941.frx":1C0A
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
         TabIndex        =   57
         Top             =   4020
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   609
         Caption         =   "frm941.frx":1C4E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1CB0
         Key             =   "frm941.frx":1CCE
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
         TabIndex        =   58
         Top             =   4500
         Width           =   2895
         _Version        =   65536
         _ExtentX        =   5106
         _ExtentY        =   609
         Caption         =   "frm941.frx":1D12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1D7E
         Key             =   "frm941.frx":1D9C
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
         TabIndex        =   60
         Top             =   5460
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   609
         Calendar        =   "frm941.frx":1DE0
         Caption         =   "frm941.frx":1EF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1F5C
         Keys            =   "frm941.frx":1F7A
         Spin            =   "frm941.frx":1FD8
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
         TabIndex        =   59
         Top             =   4980
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   609
         Caption         =   "frm941.frx":2000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":206C
         Key             =   "frm941.frx":208A
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
         TabIndex        =   50
         Top             =   1740
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   609
         Calendar        =   "frm941.frx":20CE
         Caption         =   "frm941.frx":21E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":224A
         Keys            =   "frm941.frx":2268
         Spin            =   "frm941.frx":22C6
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
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":22EE
         Caption         =   "frm941.frx":230E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2398
         Keys            =   "frm941.frx":23B6
         Spin            =   "frm941.frx":2400
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
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   4920
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":2428
         Caption         =   "frm941.frx":2448
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":24D2
         Keys            =   "frm941.frx":24F0
         Spin            =   "frm941.frx":253A
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
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   6840
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":2562
         Caption         =   "frm941.frx":2582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":260C
         Keys            =   "frm941.frx":262A
         Spin            =   "frm941.frx":2674
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
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   7560
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":269C
         Caption         =   "frm941.frx":26BC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":274A
         Keys            =   "frm941.frx":2768
         Spin            =   "frm941.frx":27B2
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
      Begin VSFlex8Ctl.VSFlexGrid fgMo2 
         Height          =   2325
         Left            =   -73800
         TabIndex        =   3
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
         FormatString    =   $"frm941.frx":27DA
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
         TabIndex        =   4
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
         FormatString    =   $"frm941.frx":28B4
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
      Begin TDBNumber6Ctl.TDBNumber Line5d 
         Height          =   300
         Left            =   -74325
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2820
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   529
         Calculator      =   "frm941.frx":298E
         Caption         =   "frm941.frx":29AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2ADE
         Keys            =   "frm941.frx":2AFC
         Spin            =   "frm941.frx":2B46
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   78053377
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line8 
         Height          =   300
         Left            =   -74760
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4800
         Width           =   10800
         _Version        =   65536
         _ExtentX        =   19050
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2B6E
         Caption         =   "frm941.frx":2B8E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2D02
         Keys            =   "frm941.frx":2D20
         Spin            =   "frm941.frx":2D6A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   78053377
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line9 
         Height          =   300
         Left            =   -74760
         TabIndex        =   22
         Top             =   5100
         Width           =   10800
         _Version        =   65536
         _ExtentX        =   19050
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2D92
         Caption         =   "frm941.frx":2DB2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2EFE
         Keys            =   "frm941.frx":2F1C
         Spin            =   "frm941.frx":2F66
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   78053377
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line11 
         Height          =   300
         Left            =   -74760
         TabIndex        =   25
         Top             =   5970
         Width           =   8745
         _Version        =   65536
         _ExtentX        =   15434
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2F8E
         Caption         =   "frm941.frx":2FAE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":30AA
         Keys            =   "frm941.frx":30C8
         Spin            =   "frm941.frx":3112
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   7208961
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line14 
         Height          =   300
         Left            =   -74760
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   7680
         Width           =   10845
         _Version        =   65536
         _ExtentX        =   19129
         _ExtentY        =   529
         Calculator      =   "frm941.frx":313A
         Caption         =   "frm941.frx":315A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":32BA
         Keys            =   "frm941.frx":32D8
         Spin            =   "frm941.frx":3322
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   79101953
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line7a 
         Height          =   300
         Left            =   -74325
         TabIndex        =   17
         Top             =   3615
         Width           =   8325
         _Version        =   65536
         _ExtentX        =   14684
         _ExtentY        =   529
         Calculator      =   "frm941.frx":334A
         Caption         =   "frm941.frx":336A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":349A
         Keys            =   "frm941.frx":34B8
         Spin            =   "frm941.frx":3502
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   128319489
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line7d 
         Height          =   300
         Left            =   -74325
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4515
         Width           =   10365
         _Version        =   65536
         _ExtentX        =   18283
         _ExtentY        =   529
         Calculator      =   "frm941.frx":352A
         Caption         =   "frm941.frx":354A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":367A
         Keys            =   "frm941.frx":3698
         Spin            =   "frm941.frx":36E2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   79101953
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line6 
         Height          =   300
         Left            =   -74760
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3120
         Width           =   10800
         _Version        =   65536
         _ExtentX        =   19050
         _ExtentY        =   529
         Calculator      =   "frm941.frx":370A
         Caption         =   "frm941.frx":372A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":38AA
         Keys            =   "frm941.frx":38C8
         Spin            =   "frm941.frx":3912
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   79101953
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line7b 
         Height          =   300
         Left            =   -74325
         TabIndex        =   18
         Top             =   3915
         Width           =   8325
         _Version        =   65536
         _ExtentX        =   14684
         _ExtentY        =   529
         Calculator      =   "frm941.frx":393A
         Caption         =   "frm941.frx":395A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":3A96
         Keys            =   "frm941.frx":3AB4
         Spin            =   "frm941.frx":3AFE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   128319489
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line7c 
         Height          =   300
         Left            =   -74325
         TabIndex        =   19
         Top             =   4215
         Width           =   8325
         _Version        =   65536
         _ExtentX        =   14684
         _ExtentY        =   529
         Calculator      =   "frm941.frx":3B26
         Caption         =   "frm941.frx":3B46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":3C36
         Keys            =   "frm941.frx":3C54
         Spin            =   "frm941.frx":3C9E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   128319489
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line12a 
         Height          =   300
         Left            =   -74760
         TabIndex        =   26
         Top             =   6540
         Width           =   8745
         _Version        =   65536
         _ExtentX        =   15434
         _ExtentY        =   529
         Calculator      =   "frm941.frx":3CC6
         Caption         =   "frm941.frx":3CE6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":3DE8
         Keys            =   "frm941.frx":3E06
         Spin            =   "frm941.frx":3E50
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   128319489
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line12b 
         Height          =   300
         Left            =   -74325
         TabIndex        =   27
         Top             =   7080
         Width           =   6255
         _Version        =   65536
         _ExtentX        =   11033
         _ExtentY        =   529
         Calculator      =   "frm941.frx":3E78
         Caption         =   "frm941.frx":3E98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":3F52
         Keys            =   "frm941.frx":3F70
         Spin            =   "frm941.frx":3FBA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0;;0"
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
         ShowContextMenu =   1
         ValueVT         =   128319489
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line10 
         Height          =   300
         Left            =   -74760
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5400
         Width           =   10800
         _Version        =   65536
         _ExtentX        =   19050
         _ExtentY        =   529
         Calculator      =   "frm941.frx":3FE2
         Caption         =   "frm941.frx":4002
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":4160
         Keys            =   "frm941.frx":417E
         Spin            =   "frm941.frx":41C8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   78053377
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line13 
         Height          =   300
         Left            =   -74760
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   7380
         Width           =   10845
         _Version        =   65536
         _ExtentX        =   19129
         _ExtentY        =   529
         Calculator      =   "frm941.frx":41F0
         Caption         =   "frm941.frx":4210
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":43BE
         Keys            =   "frm941.frx":43DC
         Spin            =   "frm941.frx":4426
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   79101953
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line15 
         Height          =   300
         Left            =   -74760
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   8160
         Width           =   8745
         _Version        =   65536
         _ExtentX        =   15434
         _ExtentY        =   529
         Calculator      =   "frm941.frx":444E
         Caption         =   "frm941.frx":446E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":4574
         Keys            =   "frm941.frx":4592
         Spin            =   "frm941.frx":45DC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
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
         ShowContextMenu =   1
         ValueVT         =   74579969
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber Line5a 
         Height          =   300
         Left            =   -74325
         TabIndex        =   9
         Top             =   1935
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   529
         Calculator      =   "frm941.frx":4604
         Caption         =   "frm941.frx":4624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":46C4
         Keys            =   "frm941.frx":46E2
         Spin            =   "frm941.frx":472C
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
         Format          =   "$ ###,###.##"
         HighlightText   =   1
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
         ValueVT         =   74579969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbNumVertNudge 
         Height          =   615
         Left            =   -65640
         TabIndex        =   116
         Top             =   1320
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   1085
         Calculator      =   "frm941.frx":4754
         Caption         =   "frm941.frx":4774
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":47F4
         Keys            =   "frm941.frx":4812
         Spin            =   "frm941.frx":485C
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
      Begin TDBNumber6Ctl.TDBNumber Line1 
         Height          =   300
         Left            =   -66255
         TabIndex        =   5
         Top             =   480
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":4884
         Caption         =   "frm941.frx":48A4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":490A
         Keys            =   "frm941.frx":4928
         Spin            =   "frm941.frx":4972
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16776960
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
         ShowContextMenu =   1
         ValueVT         =   6619137
         Value           =   0
         MaxValueVT      =   5636101
         MinValueVT      =   3342341
      End
      Begin TDBNumber6Ctl.TDBNumber BLine10Show 
         Height          =   300
         Left            =   -66600
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   8205
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   529
         Calculator      =   "frm941.frx":499A
         Caption         =   "frm941.frx":49BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":4A24
         Keys            =   "frm941.frx":4A42
         Spin            =   "frm941.frx":4A8C
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
         ValueVT         =   59506689
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber BDifference 
         Height          =   300
         Left            =   -70560
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   8520
         Width           =   6945
         _Version        =   65536
         _ExtentX        =   12259
         _ExtentY        =   529
         Calculator      =   "frm941.frx":4AB4
         Caption         =   "frm941.frx":4AD4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":4B9E
         Keys            =   "frm941.frx":4BBC
         Spin            =   "frm941.frx":4C06
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
         ValueVT         =   17629185
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText PrepPhone 
         Height          =   345
         Left            =   8400
         TabIndex        =   56
         Top             =   3600
         Width           =   2655
         _Version        =   65536
         _ExtentX        =   4683
         _ExtentY        =   609
         Caption         =   "frm941.frx":4C2E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":4C8C
         Key             =   "frm941.frx":4CAA
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
         TabIndex        =   126
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
         TabIndex        =   125
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
         TabIndex        =   124
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
         TabIndex        =   123
         Top             =   945
         Width           =   950
      End
      Begin VB.Label Label18 
         Caption         =   "12b  Number of individuals provided COBRA"
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
         Left            =   -74760
         TabIndex        =   112
         Top             =   6885
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   "One"
         Height          =   225
         Left            =   -65820
         TabIndex        =   111
         Top             =   8385
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   " You MUST complete both pages of Form 941 and SIGN it."
         Height          =   255
         Left            =   -74400
         TabIndex        =   110
         Top             =   8460
         Width           =   4215
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   -74865
         Picture         =   "frm941.frx":4CEE
         Top             =   8340
         Width           =   480
      End
      Begin VB.Label Label21 
         Caption         =   "     Inlcuding: Mar. 12 (Quarter 1), June 12 (Quarter 2), Sept. 12 (Quarter 3), Dec. 12 (Quarter 4) "
         Height          =   255
         Left            =   -74520
         TabIndex        =   109
         Top             =   660
         Width           =   8175
      End
      Begin VB.Label Label3 
         Caption         =   "   prior quarter and overpayment applied from Form 941-X or Form 944-X"
         Height          =   240
         Left            =   -74490
         TabIndex        =   108
         Top             =   6245
         Width           =   6495
      End
      Begin VB.Label Label2 
         Caption         =   "  7      CURRENT QUARTER'S ADJUSTMENTS, for example, a fractions of cents adjustment"
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   3420
         Width           =   7095
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74880
         TabIndex        =   106
         Top             =   3225
         Width           =   2295
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
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         Left            =   120
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
         Top             =   4980
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   $"frm941.frx":4FF8
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
         TabIndex        =   93
         Top             =   4930
         Width           =   9375
      End
      Begin VB.Label Label46 
         Caption         =   "Report of Tax Liability for Semiweekly Schedule Depositors, and attach it to this form."
         Height          =   255
         Left            =   -73005
         TabIndex        =   92
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
         TabIndex        =   91
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
         TabIndex        =   90
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   77
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Employer Identification number (EIN)"
         Height          =   255
         Left            =   -67320
         TabIndex        =   84
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label48 
         Caption         =   "Name (not your trade name)"
         Height          =   255
         Left            =   -74520
         TabIndex        =   83
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label38 
         Caption         =   "Check"
         Height          =   225
         Left            =   -65820
         TabIndex        =   82
         Top             =   8190
         Width           =   495
      End
      Begin VB.Label Label37 
         Caption         =   "  Follow the Instructions for Form 941-V, Payment Voucher."
         Height          =   165
         Left            =   -74445
         TabIndex        =   81
         Top             =   7965
         Width           =   4185
      End
      Begin VB.Label Label33 
         Height          =   180
         Left            =   -73605
         TabIndex        =   80
         Top             =   7230
         Width           =   7035
      End
      Begin VB.Label Label13 
         Caption         =   "  5     Taxable social security and Medicare wages and tips:"
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
         Left            =   -74760
         TabIndex        =   76
         Top             =   1665
         Width           =   5535
      End
      Begin VB.Label Label8 
         Caption         =   $"frm941.frx":508A
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   885
         Width           =   8295
      End
      Begin VB.Label Label7 
         Caption         =   "  1      Number of employees who received wages, tips, or other components for the pay period"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   375
         Width           =   7815
      End
      Begin VB.Label Label11 
         Caption         =   "  3     Total income tax withhold from wage, tips, and other compensation . . . . . . . . . . . . . . . . . . . . . . . . ."
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
         Left            =   -74760
         TabIndex        =   73
         Top             =   1140
         Width           =   8295
      End
      Begin VB.Label Label12 
         Caption         =   "  4      If no wages, tips, and other compensation are subject to social security or Medicare tax . . . . "
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
         Left            =   -74760
         TabIndex        =   72
         Top             =   1440
         Width           =   8295
      End
      Begin VB.Label Label15 
         Caption         =   "Check and go to line 6"
         Height          =   285
         Left            =   -65640
         TabIndex        =   71
         Top             =   1515
         Width           =   1695
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
      Height          =   615
      Left            =   0
      TabIndex        =   115
      Top             =   0
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   1085
      Calculator      =   "frm941.frx":512C
      Caption         =   "frm941.frx":514C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm941.frx":51CC
      Keys            =   "frm941.frx":51EA
      Spin            =   "frm941.frx":5234
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
   Begin VB.Label Label10 
      Caption         =   "Qtr"
      Height          =   195
      Left            =   6360
      TabIndex        =   79
      Top             =   60
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Year"
      Height          =   195
      Left            =   4920
      TabIndex        =   78
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "frm941Entry"
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
Dim StartYM, EndYM As Long
Dim LoadFlag As Boolean

Public Part4ID, Part5ID, PaidPrepID As Long

Dim SSTax, MedTax As Currency


Private Sub Form_Load()
    
    LoadFlag = True
    
    tdbAmountSet Me.Line2
    tdbAmountSet Me.Line3
    tdbAmountSet Me.Line5a
    tdbAmountSet Me.Line5aa
    tdbAmountSet Me.Line5b
    tdbAmountSet Me.Line5bb
    tdbAmountSet Me.Line5c
    tdbAmountSet Me.Line5cc
    tdbAmountSet Me.Line5d
    tdbAmountSet Me.Line6
    tdbAmountSet Me.Line7a
    tdbAmountSet Me.Line7b
    tdbAmountSet Me.Line7c
    tdbAmountSet Me.Line7d
    tdbAmountSet Me.Line8
    tdbAmountSet Me.Line9
    tdbAmountSet Me.Line10
    tdbAmountSet Me.Line11
    tdbAmountSet Me.Line12a
    tdbAmountSet Me.Line13
    tdbAmountSet Me.Line14
    tdbAmountSet Me.Line15
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
    
    ' don't allow input for total fields
    Line5d.ReadOnly = True
    Line6.ReadOnly = True
    Line7d.ReadOnly = True
    Line8.ReadOnly = True
    Line10.ReadOnly = True
    Line13.ReadOnly = True
    Line14.ReadOnly = True
    Line15.ReadOnly = True
    Line17Total.ReadOnly = True
    Line10Show.ReadOnly = True
    Line17Diff.ReadOnly = True
    
    tdbDateSet Me.Part5Date, Int(Now())
    
    Me.cmbChkDate12.ToolTipText = "Check Date for EE Count - Line1"
    
    tdbIntegerSet Me.Line1
    tdbIntegerSet Line12b
    
    ' init the year qtr combo
    If cmbYrQtrSet(Me.cmbYear, Me.cmbQtr) = False Then GoBack
    
' *** stuff 1st qtr
Me.cmbQtr.ListIndex = 0                                     ''''''''''''  TAKE OUT  '''''''''''''
    
    LoadFlag = False
    
    ' pop ChkDate12 combo
    PopChkDate12
    
    ' load the data
    Get941Data
    
    frm941Entry.AlphaCheckLine4 = " "
    frm941Entry.AlphaCheckLine15a = " "
    frm941Entry.AlphaCheckLine15b = " "

    EmployerName = UCase(PRCompany.Name)
    frm941Entry.txtEIN = PRCompany.FederalID
    frm941Entry.txtName = UCase(PRCompany.Name)
    If Not PRState.GetBySQL("Select * from PRState where PRState.StateID = " & PRCompany.AddrStateID) Then
       MsgBox "Company Info Not Found!!!", vbCritical, "Form 941 Entry"
       End
    Else
        frm941Entry.Line16 = PRState.StateAbbrev
    End If
    CurrYear = Year(Now())

    SetNudge Me.tdbNumHorzNudge
    SetNudge Me.tdbNumVertNudge
    GetNudge User.ID, "941B"
    Me.tdbNumHorzNudge = HorzNudge
    Me.tdbNumVertNudge = VertNudge

    Me.KeyPreview = True
    
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

Private Sub Line4_click()
    If frm941Entry.Line4 = 1 Then
        frm941Entry.AlphaCheckLine4 = "X"
        frm941Entry.Line5a.TabStop = False
        frm941Entry.Line5b.TabStop = False
        frm941Entry.Line5c.TabStop = False
        frm941Entry.Line5d.TabStop = False
        
        If IsNull(frm941Entry.Line3) Then
            frm941Entry.Line3 = 0
        End If
        
        frm941Entry.Line6 = frm941Entry.Line3
    Else
        frm941Entry.AlphaCheckLine4 = ""
    End If

End Sub

Private Sub Line5a_LostFocus()
    Line5aa = Line5a * 0.124
End Sub

Private Sub Line5b_LostFocus()
    Line5bb = Line5b * 0.124
End Sub

Private Sub Line5c_LostFocus()
    Line5cc = Line5c * 0.029
    frm941Entry.Line5d = frm941Entry.Line5a + frm941Entry.Line5b + frm941Entry.Line5c
    frm941Entry.Line6 = frm941Entry.Line3 + frm941Entry.Line5d
End Sub

Private Sub Line7c_lostfocus()
    frm941Entry.Line7d = frm941Entry.Line7a + frm941Entry.Line7b + frm941Entry.Line7c
    frm941Entry.Line8 = frm941Entry.Line6 + frm941Entry.Line7d
End Sub

Private Sub Line9_lostfocus()
    frm941Entry.Line10 = frm941Entry.Line8 - frm941Entry.Line9
End Sub


Private Sub Line12a_LostFocus()
    frm941Entry.Line13 = frm941Entry.Line11 + frm941Entry.Line12a
End Sub

Private Sub Line12b_LostFocus()
    
    If frm941Entry.Line10 > frm941Entry.Line13 Then
        frm941Entry.Line15 = 0
        frm941Entry.Line14 = frm941Entry.Line10 - frm941Entry.Line13
    Else
        frm941Entry.Line14 = 0
        frm941Entry.Line15 = frm941Entry.Line13 - frm941Entry.Line10
    End If
End Sub

Private Sub Line15Check1_Click()
    If Line15Check1 = 1 And Line15Check2 = 1 Then
        MsgBox "Please check EITHER Apply to Next Return or Send Refund", vbCritical, "Form 941"
    ElseIf Line15Check1 = 1 Then
        frm941Entry.AlphaCheckLine15a = "X"
    Else
        frm941Entry.AlphaCheckLine15a = ""
    End If
End Sub

Private Sub Line15Check2_Click()
    If Line15Check1 = 1 And Line15Check2 = 1 Then
        MsgBox "Please check EITHER Apply to Next Return or Send Refund", vbCritical, "Form 941"
    ElseIf Line15Check2 = 1 Then
        frm941Entry.AlphaCheckLine15b = "X"
    Else
        frm941Entry.AlphaCheckLine15b = ""
    End If
End Sub

Private Sub Line16__KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Line17Check1_Click()
    
    If Line17Check1 = 1 Then
        frm941Entry.AlphaCheckLine17a = "X"
        Me.Line17Check2 = 0
        Me.Line17Check3 = 0
    Else
        frm941Entry.AlphaCheckLine17a = ""
    End If
    
    frm941Entry.Line17Check2.TabStop = False
    frm941Entry.Line17Mo1.TabStop = False
    frm941Entry.Line17Mo2.TabStop = False
    frm941Entry.Line17Mo3.TabStop = False
    frm941Entry.Line17Check3.TabStop = False

End Sub

Private Sub Line17Check2_Click()
    
    If Line17Check2 = 1 Then
        frm941Entry.AlphaCheckLine17b = "X"
        Me.Line17Check1 = 0
        Me.Line17Check3 = 0
    Else
        frm941Entry.AlphaCheckLine17b = ""
        frm941Entry.Line17Mo1.TabStop = False
        frm941Entry.Line17Mo2.TabStop = False
        frm941Entry.Line17Mo3.TabStop = False
    End If

End Sub

Private Sub Line17Check3_Click()
    If Line17Check3 = 1 Then
        frm941Entry.AlphaCheckLine17c = "X"
        Me.Line17Check1 = 0
        Me.Line17Check2 = 0
    Else
        frm941Entry.AlphaCheckLine17c = ""
    End If
End Sub

Private Sub Line18Check_Click()
    If Line18Check = 1 Then
        frm941Entry.AlphaCheckLine18 = "X"
    Else
        frm941Entry.AlphaCheckLine18 = ""
    End If
End Sub

Private Sub Line19_Click()
    If Line19 = 1 Then
        frm941Entry.AlphaCheckLine19 = "X"
    Else
        frm941Entry.AlphaCheckLine19 = ""
    End If
End Sub

Private Sub Part4CheckYes_Click()

    If Part4CheckYes = 1 And Part4CheckNo = 1 Then
        MsgBox "Please check EITHER [YES] or [NO]", vbCritical, "Form 941"
    ElseIf Part4CheckYes = 1 Then
        frm941Entry.AlphaCheckPart4Yes = "X"
        frm941Entry.Part4CheckNo = 0
    ElseIf Part4CheckYes = 0 Then
        frm941Entry.AlphaCheckPart4Yes = 0
    End If
    
End Sub

Private Sub Part4Name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Part4CheckNo_Click()

    If Part4CheckNo = 1 And Part4CheckYes = 1 Then
        MsgBox "Please check EITHER [YES] or [NO]", vbCritical, "Form 941"
    ElseIf Part4CheckNo = 1 Then
        frm941Entry.AlphaCheckPart4No = "X"
        frm941Entry.Part4CheckYes = 0
    ElseIf Part4CheckNo = 0 Then
        frm941Entry.AlphaCheckPart4No = 0
    End If
    
End Sub




Private Sub prepCheck_Click()
    If Part5Check = 1 Then
        frm941Entry.AlphaCheckPart5 = "X"
    Else
        frm941Entry.AlphaCheckPart5 = ""
    End If
End Sub

Private Sub cmdExit_Click()
   GoBack
End Sub

Private Sub cmdPrint_Click()
    PrtInit ("Port")
    SetFont 10, Equate.Portrait

    HorzNudge = Me.tdbNumHorzNudge
    VertNudge = Me.tdbNumVertNudge
    SaveNudge User.ID, "941B"
    
    Me.KeyPreview = True
    Form941APrint
    
    Form941BHdr
    
    Form941BPrint 2300, Me.fgMo1, BMo1Tax
    Form941BPrint 6400, Me.fgMo2, BMo2Tax
    Form941BPrint 10500, Me.fgMo3, BMo3Tax

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
                
    SSTax = 0
    MedTax = 0
                
    Line2 = 0
    Line3 = 0
    Line5a = 0
    Line5b = 0
    Line5c = 0
    Line7b = 0
    Line7c = 0
    Line9 = 0
    Line18Date = Int(Now())
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
    
    ' get the PRHist data
    SQLString = "SELECT * FROM PRHist " & _
                " WHERE YearMonth >= " & StartYM & " AND YearMonth <= " & EndYM & _
                " ORDER BY CheckDate, EmployeeID"
                    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No Payroll History Found!!", vbExclamation
        Exit Sub
    End If
    
    Do
    
        SSTax = SSTax + PRHist.SSTax
        MedTax = MedTax + PRHist.MedTax
        
        Line2 = Line2 + PRHist.FWTWage
        Line3 = Line3 + PRHist.FWTTax
        Line5a = Line5a + PRHist.SSWage
        ' *** Line5b - Tips ***
        Line5c = Line5c + PRHist.MEDWage
        ' *** Line7b - sick pay ***
        ' *** Line7c - tips and group ins ***
        ' *** Line9 EIC payments ***
    
        ' tax liability per month
        TaxLiab = PRHist.FWTTax + PRHist.SSTax * 2 + PRHist.MedTax * 2
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
    
        If Not PRHist.GetNext Then Exit Do
    
    Loop
    
    Calc941Data
    PopChkDate12
    PopPart4Part5

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
    
    ' calculated lines
    ' *** Line17Mo3 = Line10 - Line17Mo1 - Line17Mo2 ***
    Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Line10Show = Line10
    Line17Diff = Line10 - Line17Total

    BLine10Show = Line10
    BDifference = Line10 - BLine10Show
    
    Me.Line5aa = Round(Line5a * 0.124, 2)
    Me.Line5bb = Round(Line5b * 0.124, 2)
    Me.Line5cc = Round(Line5c * 0.029, 2)
    Me.Line5d = Me.Line5aa + Me.Line5bb + Me.Line5cc
    Me.Line6 = Me.Line3 + Me.Line5d
    Me.Line7a = Round(SSTax * 2 - Me.Line5aa - Me.Line5bb + MedTax * 2 - Me.Line5cc, 2)
    Me.Line7d = Line7a + Line7b + Line7c
    
    Me.Line8 = Me.Line6 + Me.Line7d
    Me.Line10 = Me.Line8 - Me.Line9
    Me.Line13 = Me.Line11 + Me.Line12a
    If Me.Line10 >= Me.Line13 Then   ' balance due
        Me.Line14 = Me.Line10 - Me.Line13
        Me.Line15 = 0
        Me.Line15Check1.Enabled = False
        Me.Line15Check2.Enabled = False
    Else                            ' overpayment
        Me.Line14 = 0
        Me.Line15 = Me.Line13 - Me.Line10
        Me.Line15Check1.Enabled = True
        Me.Line15Check2.Enabled = True
    End If

    Line17Total = Me.Line17Mo1 + Me.Line17Mo2 + Me.Line17Mo3
    Line10Show = Line10
    Line17Diff = Line10 - Line17Total
    
    BLine10Show = Line10
    BDifference = Line10 - BLine10Show

    BGridUpdate Me.fgMo1, BMo1Tax
    BGridUpdate Me.fgMo2, BMo2Tax
    BGridUpdate Me.fgMo3, BMo3Tax
    
    BTotalTax = BMo1Tax + BMo2Tax + BMo3Tax

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

    If fgMo1.Col = 7 Then
        If fgMo1.Row = 8 Then
        Else
            fgMo1.Row = fgMo1.Row + 1
            fgMo1.Col = 0
        End If
    Else
        fgMo1.Col = fgMo1.Col + 1
    End If

    Me.BTotalTax = Me.BMo1Tax + Me.BMo2Tax + Me.BMo3Tax
    BDifference = Me.Line10 - Me.BTotalTax
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
        MsgBox "No Payroll History Found!!", vbExclamation
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
    SQLString = "SELECT * FROM Users ORDER BY NAME"
    If Not User.GetSQL(SQLString) Then
       MsgBox "Users not found: " & UserID, vbCritical, "Form941 Entry"
       End
    End If

    Do
        cmbPrepName.AddItem UCase(User.Name)
        If Not User.GetNext Then Exit Do
    Loop
    
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
    End If


End Sub

Private Sub cmbChkDate12_Click()
    Me.Line1 = Mid(Me.cmbChkDate12, 10, 10)
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
    frm941Entry.Line17Total = frm941Entry.Line17Mo1 + frm941Entry.Line17Mo2 + frm941Entry.Line17Mo3
    Calc941Data
End Sub
Private Sub Line17Mo2_Change()
    frm941Entry.Line17Total = frm941Entry.Line17Mo1 + frm941Entry.Line17Mo2 + frm941Entry.Line17Mo3
    Calc941Data
End Sub

Private Sub Line17Mo3_lostfocus()
    frm941Entry.Line17Total = frm941Entry.Line17Mo1 + frm941Entry.Line17Mo2 + frm941Entry.Line17Mo3
    Calc941Data
End Sub

