VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmPayer 
   Caption         =   "1099 Payer Information"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbHorz 
      Height          =   375
      Left            =   7680
      TabIndex        =   27
      Top             =   7680
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calculator      =   "frmPayer.frx":030A
      Caption         =   "frmPayer.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0394
      Keys            =   "frmPayer.frx":03B2
      Spin            =   "frmPayer.frx":03FC
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
   Begin TDBDate6Ctl.TDBDate tdbDate 
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   7680
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calendar        =   "frmPayer.frx":0424
      Caption         =   "frmPayer.frx":0524
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0582
      Keys            =   "frmPayer.frx":05A0
      Spin            =   "frmPayer.frx":05FE
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
      Text            =   "01/16/2012"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40924
      CenturyMode     =   0
   End
   Begin VB.CheckBox chkFinal 
      Caption         =   "Form 1099-MISC with NEC in box 7, check"
      Height          =   615
      Left            =   1440
      TabIndex        =   19
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ComboBox cmbForm 
      Height          =   360
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   6240
      Width           =   1335
   End
   Begin TDBNumber6Ctl.TDBNumber tdbFWT 
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   6960
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   661
      Calculator      =   "frmPayer.frx":0626
      Caption         =   "frmPayer.frx":0646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":06BA
      Keys            =   "frmPayer.frx":06D8
      Spin            =   "frmPayer.frx":0722
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
      MaxValue        =   99999999
      MinValue        =   -9999999
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
      MaxValueVT      =   6356997
      MinValueVT      =   5242885
   End
   Begin TDBNumber6Ctl.TDBNumber tdbNumForms 
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   6960
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      Calculator      =   "frmPayer.frx":074A
      Caption         =   "frmPayer.frx":076A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":07DE
      Keys            =   "frmPayer.frx":07FC
      Spin            =   "frmPayer.frx":0846
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
      MaxValueVT      =   7274501
      MinValueVT      =   6356997
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoadTotals 
      Caption         =   "&LOAD TOTALS"
      Height          =   615
      Left            =   6840
      TabIndex        =   15
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmd1096 
      Caption         =   "&PRINT 1096"
      Height          =   615
      Left            =   8863
      TabIndex        =   23
      Top             =   8400
      Width           =   1575
   End
   Begin TDBText6Ctl.TDBText txtName 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   12495
      _Version        =   65536
      _ExtentX        =   22040
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":086E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":08D8
      Key             =   "frmPayer.frx":08F6
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   4096
      TabIndex        =   21
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6479
      TabIndex        =   22
      Top             =   8400
      Width           =   1575
   End
   Begin TDBText6Ctl.TDBText txtAddr1 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   12495
      _Version        =   65536
      _ExtentX        =   22040
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":093A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":09A2
      Key             =   "frmPayer.frx":09C0
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
   Begin TDBText6Ctl.TDBText txtAddr2 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   12495
      _Version        =   65536
      _ExtentX        =   22040
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0A04
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0A6C
      Key             =   "frmPayer.frx":0A8A
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
   Begin TDBText6Ctl.TDBText txtCity 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   12495
      _Version        =   65536
      _ExtentX        =   22040
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0ACE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0B2C
      Key             =   "frmPayer.frx":0B4A
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
   Begin TDBText6Ctl.TDBText txtState 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0B8E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0BEE
      Key             =   "frmPayer.frx":0C0C
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
   Begin TDBText6Ctl.TDBText txtZip 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   3360
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0C50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0CAC
      Key             =   "frmPayer.frx":0CCA
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
   Begin TDBText6Ctl.TDBText txtFedID 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3840
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0D0E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0D78
      Key             =   "frmPayer.frx":0D96
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
   Begin TDBText6Ctl.TDBText tdbSSNumber 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   3840
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0DDA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0E42
      Key             =   "frmPayer.frx":0E60
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
   Begin TDBText6Ctl.TDBText tdbContactPerson 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   4320
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0EA4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0F16
      Key             =   "frmPayer.frx":0F34
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
   Begin TDBText6Ctl.TDBText tdbPhone 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   4800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":0F78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":0FD8
      Key             =   "frmPayer.frx":0FF6
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
   Begin TDBText6Ctl.TDBText tdbFax 
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   4800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":103A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":1096
      Key             =   "frmPayer.frx":10B4
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
   Begin TDBText6Ctl.TDBText tdbEMail 
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   5280
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":10F8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":1158
      Key             =   "frmPayer.frx":1176
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
   Begin TDBText6Ctl.TDBText tdbTitle 
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   5280
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   661
      Caption         =   "frmPayer.frx":11BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":121A
      Key             =   "frmPayer.frx":1238
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
   Begin TDBNumber6Ctl.TDBNumber tdbTotalAmount 
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   6960
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   661
      Calculator      =   "frmPayer.frx":127C
      Caption         =   "frmPayer.frx":129C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":130A
      Keys            =   "frmPayer.frx":1328
      Spin            =   "frmPayer.frx":1372
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
      MaxValueVT      =   6356997
      MinValueVT      =   5242885
   End
   Begin TDBNumber6Ctl.TDBNumber tdbVertical 
      Height          =   375
      Left            =   10560
      TabIndex        =   28
      Top             =   7680
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Calculator      =   "frmPayer.frx":139A
      Caption         =   "frmPayer.frx":13BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayer.frx":1424
      Keys            =   "frmPayer.frx":1442
      Spin            =   "frmPayer.frx":148C
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
   Begin VB.Label Label2 
      Caption         =   "Form Type:"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Year:"
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   6360
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
      Height          =   615
      Left            =   240
      TabIndex        =   24
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmPayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim NudgeID, GlobalID As Long

Dim rs As New ADODB.Recordset
Dim Count99 As Long
Dim LastID As Long
Dim Amt, Tax, TotAmt As Currency
Dim FormID As Long
Dim TaxBoxName As String

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.txtName = GLCompany.Name
    Me.KeyPreview = True

    Me.tdbDate = Now()

    With Me
        
        .cmbForm.AddItem "1099-NEC"
        .cmbForm.AddItem "1099-MISC"
        .cmbForm.AddItem "1099-R"
        .cmbForm.AddItem "1099-INT"
        .cmbForm.AddItem "1099-DIV"
        .cmbForm.ListIndex = 0
    
        PopTaxYear .cmbTaxYear
    
    End With
    
    ' name and address info from GLCompany
    ' federal id and SSN from GLCompany
    ' submitter info from PRGlobal

    ' submitter info
    GlobalID = 0
    ' PREquate.GlobalTypeW3E = 30
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = 30 AND " & _
                "UserID = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        GlobalID = PRGlobal.GlobalID
        Me.tdbContactPerson = PRGlobal.Var1
        Me.tdbEMail = PRGlobal.Var2
        Me.tdbPhone = PRGlobal.Var3
        Me.tdbFax = PRGlobal.Var4
        Me.tdbTitle = PRGlobal.Var5
        Me.chkFinal = PRGlobal.Byte1
    End If
    
    With Me
        
        .txtName.MaxLength = 60
        .txtAddr1.MaxLength = 60
        .txtAddr2.MaxLength = 60
        .txtCity.MaxLength = 30
        .txtState.MaxLength = 2
        .txtZip.MaxLength = 10
        .txtFedID.MaxLength = 15
        
        .txtName.text = GLCompany.Name
        .txtAddr1.text = GLCompany.Address1
        .txtAddr2.text = GLCompany.Address2
        .txtCity.text = GLCompany.City
        .txtState.text = GLCompany.State
        .txtZip.text = GLCompany.ZipCode
        .txtFedID.text = GLCompany.FederalID
        .tdbSSNumber = GLCompany.SSN
    
        .tdbContactPerson.MaxLength = 30
        .tdbEMail.MaxLength = 30
        .tdbPhone.MaxLength = 30
        .tdbFax.MaxLength = 30
        .tdbTitle.MaxLength = 30
        .tdbSSNumber.MaxLength = 15
    
    End With
    
    ' load nudge
    NudgeID = 0
    With Me
        SQLString = " SELECT * FROM PRGlobal WHERE UserID = " & User.ID & _
                    " AND Description = '1096'"
        If PRGlobal.GetBySQL(SQLString) = False Then
            .tdbHorz.Value = 0
            .tdbVertical.Value = 0
        Else
            NudgeID = PRGlobal.GlobalID
            .tdbHorz.Value = PRGlobal.Var1
            .tdbVertical.Value = PRGlobal.Var2
        End If
    End With

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdSave_Click()

    SaveData
    GoBack

End Sub
Private Sub cmd1096_Click()

    SaveData

    With Me
        Form96_NumForms = .tdbNumForms.Value
        Form96_FWT = .tdbFWT.Value
        Form96_TotalAmt = .tdbTotalAmount
        Form96_Title = .tdbTitle
        Form96_Type = .cmbForm.ListIndex
        Form96_TaxYear = .cmbTaxYear.text
        Form96_Date = .tdbDate.text
        If .chkFinal Then
            Form96_Final = "X"
        Else
            Form96_Final = ""
        End If
    
        Form96_NECX = ""
        Form96_MiscX = ""
        Form96_RX = ""
        Form96_IntX = ""
        Form96_DivX = ""
        Select Case .cmbForm
            Case "1099-NEC"
                Form96_NECX = "XXX"
            Case "1099-MISC"
                Form96_MiscX = "XXX"
            Case "1099-R"
                Form96_RX = "XXX"
            Case "1099-INT"
                Form96_IntX = "XXX"
            Case "1099-DIV"
                Form96_DivX = "XXX"
        End Select
    
        HorzNudge = .tdbHorz.Value
        VertNudge = .tdbVertical.Value
    
    End With

    PrintForm99 "1096", Form96_TaxYear, False

End Sub

Private Sub cmdLoadTotals_Click()
    
    ' get the formid
    SQLString = " SELECT * FROM Form99 WHERE TaxYear = " & Me.cmbTaxYear.text & _
                " AND FormType = '" & Mid(Me.cmbForm.text, 6) & "'"
    If Form99.GetBySQL(SQLString) = False Then
        MsgBox "Form Not Found: " & Me.cmbTaxYear & " " & Me.cmbForm, vbExclamation
        GoBack
    End If
    
    FormID = Form99.FormID
    
    ' get the form count
    SQLString = " SELECT DISTINCT(PayeeID) FROM Detail99 " & _
                " WHERE FormType = '" & Mid(Me.cmbForm.text, 6) & "' " & _
                " AND TaxYear = " & Me.cmbTaxYear.text
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveLast
    Count99 = rs.RecordCount
    rs.Close
    Me.tdbNumForms = Count99

'    ' which FieldID is the tax amount?
'    TaxBoxName = ""
'    SQLString = " SELECT * FROM Field99 WHERE FormType = '" & Mid(Me.cmbForm.Text, 6) & "' " & _
'                " AND TaxYear = " & Me.cmbTaxYear.Text
'    If Field99.GetBySQL(SQLString) = False Then
'        MsgBox "Field data not found???", vbExclamation
'        GoBack
'    End If
'    Do
'        If InStr(1, "income tax withheld", LCase(Field99.FieldTitle), vbTextCompare) > 0 Then
'            TaxBoxName = Field99.BoxName
'            Exit Do
'        End If
'        If Field99.GetNext = False Then Exit Do
'    Loop
'    If TaxBoxName = "" Then
'        MsgBox "Tax field not found!!!", vbExclamation
'        GoBack
'    End If
'
    ' get the amounts
    Amt = 0
    TotAmt = 0
    Tax = 0
    SQLString = " SELECT * FROM Detail99 WHERE FormType = '" & Mid(Me.cmbForm.text, 6) & "' " & _
                " AND TaxYear = " & Me.cmbTaxYear.text
    If Detail99.GetBySQL(SQLString) = True Then
        Do
            SQLString = " SELECT * FROM Field99 WHERE FormType = '" & Detail99.FormType & "' " & _
                        " AND TaxYear = " & Detail99.TaxYear & _
                        " AND BoxName = '" & Detail99.BoxName & "' " & _
                        " AND FieldFormat = " & Equate.fmtAmount
            If Field99.GetBySQL(SQLString) Then
                Amt = ParseAmt(Detail99.FieldValue)
                If InStr(1, LCase(Field99.FieldTitle), "tax withheld", vbTextCompare) Then
                    Tax = Tax + Amt
                Else
                    TotAmt = TotAmt + Amt
                End If
            End If
                
            If Detail99.GetNext = False Then Exit Do
        Loop
    End If

    With Me
        tdbAmountSet .tdbFWT
        tdbAmountSet .tdbTotalAmount
        .tdbTotalAmount.Value = TotAmt
        .tdbFWT.Value = Tax
    End With

End Sub


Private Sub SaveData()

    With Me
        
        GLCompany.Name = .txtName.text
        GLCompany.Address1 = .txtAddr1.text
        GLCompany.Address2 = .txtAddr2.text
        GLCompany.City = .txtCity.text
        GLCompany.State = .txtState.text
        GLCompany.ZipCode = .txtZip.text
        GLCompany.FederalID = .txtFedID.text
        GLCompany.SSN = .tdbSSNumber
        GLCompany.Save (Equate.RecPut)
            
        If GlobalID = 0 Then
            PRGlobal.Clear
            PRGlobal.UserID = GLCompany.ID
            PRGlobal.TypeCode = 30
            PRGlobal.Save (Equate.RecAdd)
        Else
            If PRGlobal.GetByID(GlobalID) = False Then
                MsgBox "PRGlobal Error?:", vbExclamation
                GoBack
            End If
        End If
        
        PRGlobal.Var1 = Me.tdbContactPerson & ""
        PRGlobal.Var2 = Me.tdbEMail & ""
        PRGlobal.Var3 = Me.tdbPhone & ""
        PRGlobal.Var4 = Me.tdbFax & ""
        PRGlobal.Var5 = Me.tdbTitle & ""
        If Me.chkFinal Then
            PRGlobal.Byte1 = 1
        Else
            PRGlobal.Byte1 = 0
        End If
        
        PRGlobal.Save (Equate.RecPut)
    
    End With

    SaveNudge

End Sub

Private Sub SaveNudge()

    If NudgeID = 0 Then
        PRGlobal.OpenRS
        PRGlobal.Clear
        PRGlobal.UserID = User.ID
        PRGlobal.Description = "1096"
        PRGlobal.Save (Equate.RecAdd)
    Else
        If PRGlobal.GetByID(NudgeID) = False Then
            MsgBox "PRGlobal Error?:", vbExclamation
            GoBack
        End If
    End If
    
    PRGlobal.Var1 = Me.tdbHorz.Value
    PRGlobal.Var2 = Me.tdbVertical.Value
    PRGlobal.Save (Equate.RecPut)

End Sub

