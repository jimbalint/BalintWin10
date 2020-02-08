VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frm941 
   Caption         =   "Form 941 for 2008"
   ClientHeight    =   9240
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin TDBNumber6Ctl.TDBNumber TDBNumber26 
      Height          =   300
      Left            =   3000
      TabIndex        =   91
      Top             =   3000
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   529
      Calculator      =   "frm941.frx":0000
      Caption         =   "frm941.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm941.frx":008A
      Keys            =   "frm941.frx":00A8
      Spin            =   "frm941.frx":00F2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "00,000,000;-00,000,000;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "[###,##0.00];[-###,##0.00]"
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
      ValueVT         =   30081025
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.ComboBox cmbQtr 
      Height          =   315
      Left            =   7185
      TabIndex        =   52
      Text            =   "Combo2"
      Top             =   35
      Width           =   735
   End
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   5385
      TabIndex        =   51
      Text            =   "Combo1"
      Top             =   35
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   250
      Left            =   3960
      TabIndex        =   48
      Top             =   70
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
      Height          =   300
      Left            =   10320
      TabIndex        =   1
      Top             =   70
      Width           =   810
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
      Height          =   300
      Left            =   9120
      TabIndex        =   0
      Top             =   70
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   15901
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Form 941 for 2008"
      TabPicture(0)   =   "frm941.frx":011A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check4"
      Tab(0).Control(1)=   "Check3"
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(3)=   "TDBNumber3"
      Tab(0).Control(4)=   "TDBNumber1"
      Tab(0).Control(5)=   "TDB2"
      Tab(0).Control(6)=   "tdbOne"
      Tab(0).Control(7)=   "TDBNumber4"
      Tab(0).Control(8)=   "TDBNumber5"
      Tab(0).Control(9)=   "TDBNumber6"
      Tab(0).Control(10)=   "TDBNumber7"
      Tab(0).Control(11)=   "TDBNumber10"
      Tab(0).Control(12)=   "TDBNumber8"
      Tab(0).Control(13)=   "TDBNumber9"
      Tab(0).Control(14)=   "TDBNumber2"
      Tab(0).Control(15)=   "TDBNumber12"
      Tab(0).Control(16)=   "TDBNumber14"
      Tab(0).Control(17)=   "TDBNumber16"
      Tab(0).Control(18)=   "TDBNumber18"
      Tab(0).Control(19)=   "TDBNumber20"
      Tab(0).Control(20)=   "TDBNumber21"
      Tab(0).Control(21)=   "TDBNumber22"
      Tab(0).Control(22)=   "TDBNumber23"
      Tab(0).Control(23)=   "TDBNumber24"
      Tab(0).Control(24)=   "TDBNumber25"
      Tab(0).Control(25)=   "TDBNumber11"
      Tab(0).Control(26)=   "TDBNumber13"
      Tab(0).Control(27)=   "TDBNumber15"
      Tab(0).Control(28)=   "TDBNumber17"
      Tab(0).Control(29)=   "TDBNumber19"
      Tab(0).Control(30)=   "Label3"
      Tab(0).Control(31)=   "Label38"
      Tab(0).Control(32)=   "Label37"
      Tab(0).Control(33)=   "Label35"
      Tab(0).Control(34)=   "Label34"
      Tab(0).Control(35)=   "Label33"
      Tab(0).Control(36)=   "Label32"
      Tab(0).Control(37)=   "Label31"
      Tab(0).Control(38)=   "Label30"
      Tab(0).Control(39)=   "Label29"
      Tab(0).Control(40)=   "Label28"
      Tab(0).Control(41)=   "Label27"
      Tab(0).Control(42)=   "Label26"
      Tab(0).Control(43)=   "Label25"
      Tab(0).Control(44)=   "Label5"
      Tab(0).Control(45)=   "Label22"
      Tab(0).Control(46)=   "Label24"
      Tab(0).Control(47)=   "Label4"
      Tab(0).Control(48)=   "Label2"
      Tab(0).Control(49)=   "Label23"
      Tab(0).Control(50)=   "Label13"
      Tab(0).Control(51)=   "Label8"
      Tab(0).Control(52)=   "Label7"
      Tab(0).Control(53)=   "Label11"
      Tab(0).Control(54)=   "Label12"
      Tab(0).Control(55)=   "Label14"
      Tab(0).Control(56)=   "Label15"
      Tab(0).Control(57)=   "Label18"
      Tab(0).Control(58)=   "Label19"
      Tab(0).Control(59)=   "Label20"
      Tab(0).Control(60)=   "Label21"
      Tab(0).ControlCount=   61
      TabCaption(1)   =   "16"
      TabPicture(1)   =   "frm941.frx":0136
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label48"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label36"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label42"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label43"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label44"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label39"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label40"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label45"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label46"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label47"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label49"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label50"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label51"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label52"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label53"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label54"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TDBNumber29"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TDBNumber28"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TDBNumber27"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text1"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text2"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TDBText1"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Check5"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Check6"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Check7"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Check9"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Check10"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Check11"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Check12"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "TDBDate1"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Check2"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "DesigneeName"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "TDBMask1"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Schedule B (Form 941)"
      TabPicture(2)   =   "frm941.frx":0152
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin TDBMask6Ctl.TDBMask TDBMask1 
         Height          =   255
         Left            =   3240
         TabIndex        =   107
         Top             =   6360
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   450
         Caption         =   "frm941.frx":016E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frm941.frx":01D4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         AllowSpace      =   -1
         AutoConvert     =   -1
         BackColor       =   -2147483643
         BorderStyle     =   1
         ClipMode        =   0
         CursorPosition  =   -1
         DataProperty    =   0
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "999(99)9999"
         HighlightText   =   0
         IMEMode         =   0
         IMEStatus       =   0
         LookupMode      =   0
         LookupTable     =   ""
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "___(__)____"
         Value           =   ""
      End
      Begin VB.TextBox DesigneeName 
         Height          =   285
         Left            =   3240
         TabIndex        =   106
         Text            =   "Text3"
         Top             =   5895
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check here."
         Height          =   255
         Left            =   9720
         TabIndex        =   102
         Top             =   5040
         Width           =   1935
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   285
         Left            =   600
         TabIndex        =   99
         Top             =   4785
         Width           =   4575
         _Version        =   65536
         _ExtentX        =   8070
         _ExtentY        =   494
         Calendar        =   "frm941.frx":0216
         Caption         =   "frm941.frx":032E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":03D0
         Keys            =   "frm941.frx":03EE
         Spin            =   "frm941.frx":044C
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
         Text            =   "08/26/2008"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39686
         CenturyMode     =   0
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Check here, and"
         Height          =   255
         Left            =   9720
         TabIndex        =   88
         Top             =   4500
         Width           =   1935
      End
      Begin VB.CheckBox Check11 
         Height          =   375
         Left            =   480
         TabIndex        =   87
         Top             =   7320
         Width           =   1215
      End
      Begin VB.CheckBox Check10 
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
         Left            =   600
         TabIndex        =   86
         Top             =   6720
         Width           =   615
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Yes.   Designee's Name"
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
         Left            =   600
         TabIndex        =   85
         Top             =   5895
         Width           =   2655
      End
      Begin VB.CheckBox Check7 
         Caption         =   "You were a semiweekly schedule depositor for any part of this quarter.  Fill out Schedule B (Form 941):"
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
         Left            =   1560
         TabIndex        =   84
         Top             =   3720
         Width           =   9855
      End
      Begin VB.CheckBox Check6 
         Caption         =   "You were a monthly schedule depositor for the entire quarter.  Fill out your tax liability for each month."
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
         Left            =   1560
         TabIndex        =   83
         Top             =   2040
         Width           =   9375
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Line 10 is less than $2,500.  Go to Part 3."
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
         Left            =   1560
         TabIndex        =   82
         Top             =   1755
         Width           =   7215
      End
      Begin TDBText6Ctl.TDBText TDBText1 
         Height          =   255
         Left            =   480
         TabIndex        =   76
         Top             =   1185
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   450
         Caption         =   "frm941.frx":0474
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":04E0
         Key             =   "frm941.frx":04FE
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
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6960
         TabIndex        =   74
         Top             =   560
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   73
         Top             =   560
         Width           =   3855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Send refund check."
         Height          =   255
         Left            =   -65760
         TabIndex        =   71
         Top             =   7800
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Apply to next return."
         Height          =   255
         Left            =   -65760
         TabIndex        =   70
         Top             =   7560
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   -66480
         TabIndex        =   63
         Top             =   1330
         Width           =   255
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber3 
         Height          =   300
         Left            =   -69000
         TabIndex        =   3
         Top             =   1860
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0542
         Caption         =   "frm941.frx":0562
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":05CE
         Keys            =   "frm941.frx":05EC
         Spin            =   "frm941.frx":0636
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#######0;(#######0);Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   300
         Left            =   -66450
         TabIndex        =   4
         Top             =   1060
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":065E
         Caption         =   "frm941.frx":067E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":06DE
         Keys            =   "frm941.frx":06FC
         Spin            =   "frm941.frx":0746
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDB2 
         Height          =   300
         Left            =   -66450
         TabIndex        =   5
         Top             =   760
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":076E
         Caption         =   "frm941.frx":078E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":07EE
         Keys            =   "frm941.frx":080C
         Spin            =   "frm941.frx":0856
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#######0;(#######0);Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#######0;(#######0)"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5636101
         MinValueVT      =   3342341
      End
      Begin TDBNumber6Ctl.TDBNumber tdbOne 
         Height          =   300
         Left            =   -66450
         TabIndex        =   6
         Top             =   465
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   529
         Calculator      =   "frm941.frx":087E
         Caption         =   "frm941.frx":089E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":08FE
         Keys            =   "frm941.frx":091C
         Spin            =   "frm941.frx":0966
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber4 
         Height          =   300
         Left            =   -74640
         TabIndex        =   7
         Top             =   2160
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   529
         Calculator      =   "frm941.frx":098E
         Caption         =   "frm941.frx":09AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0A4C
         Keys            =   "frm941.frx":0A6A
         Spin            =   "frm941.frx":0AB4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber5 
         Height          =   300
         Left            =   -69000
         TabIndex        =   8
         Top             =   2160
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0ADC
         Caption         =   "frm941.frx":0AFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0B68
         Keys            =   "frm941.frx":0B86
         Spin            =   "frm941.frx":0BD0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber6 
         Height          =   300
         Left            =   -74640
         TabIndex        =   9
         Top             =   2460
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0BF8
         Caption         =   "frm941.frx":0C18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0CBC
         Keys            =   "frm941.frx":0CDA
         Spin            =   "frm941.frx":0D24
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber7 
         Height          =   300
         Left            =   -69000
         TabIndex        =   10
         Top             =   2460
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0D4C
         Caption         =   "frm941.frx":0D6C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0DD8
         Keys            =   "frm941.frx":0DF6
         Spin            =   "frm941.frx":0E40
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber10 
         Height          =   300
         Left            =   -66555
         TabIndex        =   11
         Top             =   2820
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0E68
         Caption         =   "frm941.frx":0E88
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0EE8
         Keys            =   "frm941.frx":0F06
         Spin            =   "frm941.frx":0F50
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber8 
         Height          =   300
         Left            =   -66480
         TabIndex        =   12
         Top             =   3120
         Width           =   2310
         _Version        =   65536
         _ExtentX        =   4083
         _ExtentY        =   529
         Calculator      =   "frm941.frx":0F78
         Caption         =   "frm941.frx":0F98
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":0FF8
         Keys            =   "frm941.frx":1016
         Spin            =   "frm941.frx":1060
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber9 
         Height          =   300
         Left            =   -74640
         TabIndex        =   23
         Top             =   3900
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18486
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1088
         Caption         =   "frm941.frx":10A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":123A
         Keys            =   "frm941.frx":1258
         Spin            =   "frm941.frx":12A2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
         Height          =   300
         Left            =   -74640
         TabIndex        =   24
         Top             =   1860
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9754
         _ExtentY        =   529
         Calculator      =   "frm941.frx":12CA
         Caption         =   "frm941.frx":12EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":138A
         Keys            =   "frm941.frx":13A8
         Spin            =   "frm941.frx":13F2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber12 
         Height          =   300
         Left            =   -74640
         TabIndex        =   26
         Top             =   4200
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18486
         _ExtentY        =   529
         Calculator      =   "frm941.frx":141A
         Caption         =   "frm941.frx":143A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1584
         Keys            =   "frm941.frx":15A2
         Spin            =   "frm941.frx":15EC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber14 
         Height          =   300
         Left            =   -74640
         TabIndex        =   27
         Top             =   4500
         Width           =   3645
         _Version        =   65536
         _ExtentX        =   6429
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1614
         Caption         =   "frm941.frx":1634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":16E4
         Keys            =   "frm941.frx":1702
         Spin            =   "frm941.frx":174C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   36241409
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber16 
         Height          =   300
         Left            =   -74640
         TabIndex        =   29
         Top             =   5115
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1774
         Caption         =   "frm941.frx":1794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":184C
         Keys            =   "frm941.frx":186A
         Spin            =   "frm941.frx":18B4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   36241409
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber18 
         Height          =   300
         Left            =   -66240
         TabIndex        =   34
         Top             =   4500
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         Calculator      =   "frm941.frx":18DC
         Caption         =   "frm941.frx":18FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1958
         Keys            =   "frm941.frx":1976
         Spin            =   "frm941.frx":19C0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   0
         MarginRight     =   0
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber20 
         Height          =   300
         Left            =   -66540
         TabIndex        =   37
         Top             =   5700
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":19E8
         Caption         =   "frm941.frx":1A08
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1A68
         Keys            =   "frm941.frx":1A86
         Spin            =   "frm941.frx":1AD0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber21 
         Height          =   300
         Left            =   -66540
         TabIndex        =   40
         Top             =   6000
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1AF8
         Caption         =   "frm941.frx":1B18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1B7A
         Keys            =   "frm941.frx":1B98
         Spin            =   "frm941.frx":1BE2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber22 
         Height          =   300
         Left            =   -66540
         TabIndex        =   42
         Top             =   6300
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1C0A
         Caption         =   "frm941.frx":1C2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1C8C
         Keys            =   "frm941.frx":1CAA
         Spin            =   "frm941.frx":1CF4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber23 
         Height          =   300
         Left            =   -66540
         TabIndex        =   45
         Top             =   6600
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1D1C
         Caption         =   "frm941.frx":1D3C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1D9E
         Keys            =   "frm941.frx":1DBC
         Spin            =   "frm941.frx":1E06
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber24 
         Height          =   300
         Left            =   -66540
         TabIndex        =   47
         Top             =   6900
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1E2E
         Caption         =   "frm941.frx":1E4E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1EB0
         Keys            =   "frm941.frx":1ECE
         Spin            =   "frm941.frx":1F18
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber25 
         Height          =   300
         Left            =   -66540
         TabIndex        =   57
         Top             =   7200
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   529
         Calculator      =   "frm941.frx":1F40
         Caption         =   "frm941.frx":1F60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":1FC2
         Keys            =   "frm941.frx":1FE0
         Spin            =   "frm941.frx":202A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber11 
         Height          =   300
         Left            =   -74640
         TabIndex        =   64
         Top             =   3600
         Width           =   10485
         _Version        =   65536
         _ExtentX        =   18486
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2052
         Caption         =   "frm941.frx":2072
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":21F6
         Keys            =   "frm941.frx":2214
         Spin            =   "frm941.frx":225E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   1
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   33619969
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber13 
         Height          =   300
         Left            =   -66240
         TabIndex        =   65
         Top             =   4800
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2286
         Caption         =   "frm941.frx":22A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2302
         Keys            =   "frm941.frx":2320
         Spin            =   "frm941.frx":236A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   0
         MarginRight     =   0
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber15 
         Height          =   300
         Left            =   -66240
         TabIndex        =   66
         Top             =   5100
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         Calculator      =   "frm941.frx":2392
         Caption         =   "frm941.frx":23B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":240E
         Keys            =   "frm941.frx":242C
         Spin            =   "frm941.frx":2476
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   0
         MarginRight     =   0
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber17 
         Height          =   300
         Left            =   -66240
         TabIndex        =   68
         Top             =   5400
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         Calculator      =   "frm941.frx":249E
         Caption         =   "frm941.frx":24BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":251A
         Keys            =   "frm941.frx":2538
         Spin            =   "frm941.frx":2582
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   0
         MarginRight     =   0
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
         ValueVT         =   33685505
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber19 
         Height          =   300
         Left            =   -68900
         TabIndex        =   69
         Top             =   7560
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   529
         Calculator      =   "frm941.frx":25AA
         Caption         =   "frm941.frx":25CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2626
         Keys            =   "frm941.frx":2644
         Spin            =   "frm941.frx":268E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   0
         MarginRight     =   0
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
         ValueVT         =   30081025
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber27 
         Height          =   300
         Left            =   3000
         TabIndex        =   92
         Top             =   2820
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941.frx":26B6
         Caption         =   "frm941.frx":26D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":2740
         Keys            =   "frm941.frx":275E
         Spin            =   "frm941.frx":27A8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   30081025
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber28 
         Height          =   300
         Left            =   3000
         TabIndex        =   93
         Top             =   3120
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   529
         Calculator      =   "frm941.frx":27D0
         Caption         =   "frm941.frx":27F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":285A
         Keys            =   "frm941.frx":2878
         Spin            =   "frm941.frx":28C2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   30081025
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber29 
         Height          =   300
         Left            =   1800
         TabIndex        =   94
         Top             =   3440
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   529
         Calculator      =   "frm941.frx":28EA
         Caption         =   "frm941.frx":290A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frm941.frx":299C
         Keys            =   "frm941.frx":29BA
         Spin            =   "frm941.frx":2A04
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "00,000,000;-00,000,000;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "[###,##0.00];[-###,##0.00]"
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
         ValueVT         =   30081025
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
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
         Left            =   600
         TabIndex        =   105
         Top             =   5640
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
         Left            =   120
         TabIndex        =   104
         Top             =   5400
         Width           =   4455
      End
      Begin VB.Label Label52 
         Caption         =   "Part 3:  Applies to your business.  If a question does NOT apply to your business, leave it blank."
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
         TabIndex        =   103
         Top             =   4200
         Width           =   10095
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
         Left            =   600
         TabIndex        =   101
         Top             =   5040
         Width           =   9375
      End
      Begin VB.Label Label50 
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
         Left            =   120
         TabIndex        =   100
         Top             =   5085
         Width           =   255
      End
      Begin VB.Label Label49 
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
         Left            =   120
         TabIndex        =   98
         Top             =   4500
         Width           =   255
      End
      Begin VB.Label Label47 
         Caption         =   $"frm941.frx":2A2C
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
         Left            =   600
         TabIndex        =   97
         Top             =   4485
         Width           =   9375
      End
      Begin VB.Label Label46 
         Caption         =   "Report of Tax Liability for Semiweekly Schedule Depositors, and attach it to this form."
         Height          =   255
         Left            =   1875
         TabIndex        =   96
         Top             =   3975
         Width           =   6255
      End
      Begin VB.Label Label45 
         Caption         =   "Total must equal line 10."
         Height          =   255
         Left            =   6240
         TabIndex        =   95
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label40 
         Caption         =   "Tax liability:      "
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
         Left            =   1800
         TabIndex        =   90
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "15"
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
         Left            =   0
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
         Left            =   1800
         TabIndex        =   81
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
         Left            =   360
         TabIndex        =   50
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
         Left            =   1200
         TabIndex        =   80
         Top             =   1440
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
         Left            =   1200
         TabIndex        =   78
         Top             =   1200
         Width           =   9975
      End
      Begin VB.Label Label17 
         Caption         =   "14"
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
         Left            =   120
         TabIndex        =   77
         Top             =   1215
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Part 2:  Deposit Schedule and tax liability for this quarter."
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
         TabIndex        =   75
         Top             =   900
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "Employer Identification number (EIN)"
         Height          =   255
         Left            =   6960
         TabIndex        =   72
         Top             =   320
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "(attach Form 941c) . . . . . . . . . . . . . . . . .   7g"
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
         Left            =   -69840
         TabIndex        =   67
         Top             =   5460
         Width           =   3585
      End
      Begin VB.Label Label48 
         Caption         =   "Name (not your trade name)"
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   320
         Width           =   3375
      End
      Begin VB.Label Label38 
         Caption         =   "Check one"
         Height          =   255
         Left            =   -66645
         TabIndex        =   61
         Top             =   7590
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Follow the Instructions for Form 941-V, Payment Voucher."
         Height          =   165
         Left            =   -74685
         TabIndex        =   60
         Top             =   7440
         Width           =   4185
      End
      Begin VB.Label Label35 
         Caption         =   "(If line 10 is more than line 11, write the difference here.)  . . . . ."
         Height          =   180
         Left            =   -73455
         TabIndex        =   59
         Top             =   7650
         Width           =   5025
      End
      Begin VB.Label Label34 
         Caption         =   "13 Overpayment"
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
         Left            =   -74925
         TabIndex        =   58
         Top             =   7620
         Width           =   1485
      End
      Begin VB.Label Label33 
         Caption         =   "(If line 10 is more than line 11, write the difference here.) . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
         Height          =   180
         Left            =   -73605
         TabIndex        =   56
         Top             =   7230
         Width           =   7035
      End
      Begin VB.Label Label32 
         Caption         =   "12 Balance due"
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
         Left            =   -74925
         TabIndex        =   55
         Top             =   7200
         Width           =   1245
      End
      Begin VB.Label Label31 
         Caption         =   "11 Total deposits for this quarter, including overpayment applied from a prior quarter  . . . . . . . . . . . . . . . "
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
         Left            =   -74925
         TabIndex        =   46
         Top             =   6915
         Width           =   8535
      End
      Begin VB.Label Label30 
         Caption         =   "(line 8 - line 9 = line 10)  . . . . . . . . . . . . . . . . . . . . . . . . . "
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
         Left            =   -70935
         TabIndex        =   44
         Top             =   6615
         Width           =   4455
      End
      Begin VB.Label Label29 
         Caption         =   "10 Total taxes after adjustments for advance EIC"
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
         Left            =   -74925
         TabIndex        =   43
         Top             =   6615
         Width           =   3975
      End
      Begin VB.Label Label28 
         Caption         =   "9  Advance earned income credit (EIC) payments made to employees . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
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
         Left            =   -74880
         TabIndex        =   41
         Top             =   6315
         Width           =   8295
      End
      Begin VB.Label Label27 
         Caption         =   "(Combine lines 6 and 7h) . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . "
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
         Left            =   -72000
         TabIndex        =   39
         Top             =   6020
         Width           =   5535
      End
      Begin VB.Label Label26 
         Caption         =   "8  Total taxes after adjustments"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   6030
         Width           =   2775
      End
      Begin VB.Label Label25 
         Caption         =   "(Combine allamounts:  lines 7a through 7g)  . . . . . . . . . . . . . . . . ."
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
         Left            =   -72000
         TabIndex        =   36
         Top             =   5730
         Width           =   5535
      End
      Begin VB.Label Label5 
         Caption         =   "7h   TOTAL ADJUSTMENTS "
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
         Left            =   -74640
         TabIndex        =   35
         Top             =   5740
         Width           =   2655
      End
      Begin VB.Label Label22 
         Caption         =   "(attach Form 941c) . . . . . . . . . . . . . . . . . . . . . . . . . . . . .   7d"
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
         Left            =   -70935
         TabIndex        =   33
         Top             =   4560
         Width           =   4605
      End
      Begin VB.Label Label24 
         Caption         =   "(attach Form 941c). . . . . . . . . . . . . . . . .   7e"
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
         Left            =   -69810
         TabIndex        =   32
         Top             =   4860
         Width           =   3585
      End
      Begin VB.Label Label4 
         Caption         =   "7g   Special additions to social security and Medicare"
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
         Left            =   -74640
         TabIndex        =   31
         Top             =   5460
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "(attach Form 941c) . . . . . . . . . . . . . . . . . . . . . . . . . . .   7f"
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
         Left            =   -70740
         TabIndex        =   30
         Top             =   5160
         Width           =   4515
      End
      Begin VB.Label Label23 
         Caption         =   "7e   Prior quarter's social security and Medicare taxes"
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
         Left            =   -74640
         TabIndex        =   28
         Top             =   4860
         Width           =   4815
      End
      Begin VB.Label Label13 
         Caption         =   "5  Taxable social security and Medicare wages and tips:"
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
         Left            =   -74880
         TabIndex        =   25
         Top             =   1615
         Width           =   5055
      End
      Begin VB.Label Label8 
         Caption         =   $"frm941.frx":2ABE
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
         Left            =   -74880
         TabIndex        =   22
         Top             =   810
         Width           =   8295
      End
      Begin VB.Label Label7 
         Caption         =   "1   Number of employees who received wages, tips, or other components for the pay period"
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
         Left            =   -74880
         TabIndex        =   21
         Top             =   330
         Width           =   7815
      End
      Begin VB.Label Label11 
         Caption         =   "3  Total income tax withhold from wage, tips, and other compensation . . . . . . . . . . . . . . . . . . . . . . . . ."
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   1050
         Width           =   8295
      End
      Begin VB.Label Label12 
         Caption         =   "4   If no wages, tips, and other compensation are subject to social security or Medicare tax . . . . "
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
         Left            =   -74880
         TabIndex        =   19
         Top             =   1400
         Width           =   8295
      End
      Begin VB.Label Label14 
         Caption         =   "including: Mar. 12 (Quarter 1), June 12 (Quarter 2), Sept. 12(Quarter 3), Dec. 12 (Quarter 4) . . . . . . . . . . . . . . . . . . "
         Height          =   255
         Left            =   -74640
         TabIndex        =   18
         Top             =   570
         Width           =   8175
      End
      Begin VB.Label Label15 
         Caption         =   "Check and go to line 6"
         Height          =   285
         Left            =   -66220
         TabIndex        =   17
         Top             =   1430
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "5d   Total social security and Medicare taxes "
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
         Left            =   -74640
         TabIndex        =   16
         Top             =   2820
         Width           =   3975
      End
      Begin VB.Label Label19 
         Caption         =   $"frm941.frx":2B5B
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   3120
         Width           =   8295
      End
      Begin VB.Label Label20 
         Caption         =   "(Column 2, lines 5a + 5b + 5c = line 5d). . . . . . . . . . . . ."
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
         Left            =   -70560
         TabIndex        =   14
         Top             =   2835
         Width           =   3975
      End
      Begin VB.Label Label21 
         Caption         =   "7  TAX ADJUSTMENTS (Read instructions for line 7 before completing lines 7a through 7h.):"
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
         Left            =   -74880
         TabIndex        =   13
         Top             =   3360
         Width           =   7095
      End
   End
   Begin VB.Label Label41 
      Caption         =   "Label41"
      Height          =   495
      Left            =   3360
      TabIndex        =   79
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label10 
      Caption         =   "Quarter"
      Height          =   195
      Left            =   6480
      TabIndex        =   54
      Top             =   75
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Year"
      Height          =   195
      Left            =   4905
      TabIndex        =   53
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   49
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frm941"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   GoBack
End Sub

