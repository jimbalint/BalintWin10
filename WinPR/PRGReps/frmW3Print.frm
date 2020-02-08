VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmW3Print 
   Caption         =   "W3 Form Print"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9720
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate tdbDate 
      Height          =   615
      Left            =   5640
      TabIndex        =   35
      Top             =   8880
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1085
      Calendar        =   "frmW3Print.frx":0000
      Caption         =   "frmW3Print.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":015E
      Keys            =   "frmW3Print.frx":017C
      Spin            =   "frmW3Print.frx":01DA
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
      Text            =   "01/11/2010"
      ValidateMode    =   0
      ValueVT         =   997457927
      Value           =   40189
      CenturyMode     =   0
   End
   Begin TDBText6Ctl.TDBText tdbTitle 
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   8880
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":0202
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0262
      Key             =   "frmW3Print.frx":0280
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
   Begin TDBNumber6Ctl.TDBNumber tdbHorzNudge 
      Height          =   375
      Left            =   7560
      TabIndex        =   36
      Top             =   8520
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":02C4
      Caption         =   "frmW3Print.frx":02E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":035A
      Keys            =   "frmW3Print.frx":0378
      Spin            =   "frmW3Print.frx":03C2
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   10920
      TabIndex        =   39
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   10920
      TabIndex        =   38
      Top             =   8280
      Width           =   1335
   End
   Begin TDBText6Ctl.TDBText tdbBoxG2 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print.frx":03EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":044E
      Key             =   "frmW3Print.frx":046C
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
   Begin TDBText6Ctl.TDBText tdbContactPerson 
      Height          =   615
      Left            =   6120
      TabIndex        =   28
      Top             =   4920
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":04B0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0522
      Key             =   "frmW3Print.frx":0540
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
   Begin TDBText6Ctl.TDBText tdbBoxE 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":0584
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0602
      Key             =   "frmW3Print.frx":0620
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
   Begin VB.ComboBox cmbPayer 
      Height          =   360
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin TDBText6Ctl.TDBText tdbBoxD 
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":0664
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":06D4
      Key             =   "frmW3Print.frx":06F2
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
   Begin TDBNumber6Ctl.TDBNumber tdbBoxC 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1085
      Calculator      =   "frmW3Print.frx":0736
      Caption         =   "frmW3Print.frx":0756
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":07CC
      Keys            =   "frmW3Print.frx":07EA
      Spin            =   "frmW3Print.frx":0834
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
      MaxValueVT      =   3014661
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbBox1 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":085C
      Caption         =   "frmW3Print.frx":087C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":08F4
      Keys            =   "frmW3Print.frx":0912
      Spin            =   "frmW3Print.frx":095C
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox2 
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   480
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0984
      Caption         =   "frmW3Print.frx":09A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0A14
      Keys            =   "frmW3Print.frx":0A32
      Spin            =   "frmW3Print.frx":0A7C
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox3 
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   960
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0AA4
      Caption         =   "frmW3Print.frx":0AC4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0B2E
      Keys            =   "frmW3Print.frx":0B4C
      Spin            =   "frmW3Print.frx":0B96
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox4 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   960
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0BBE
      Caption         =   "frmW3Print.frx":0BDE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0C44
      Keys            =   "frmW3Print.frx":0C62
      Spin            =   "frmW3Print.frx":0CAC
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox5 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   1440
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0CD4
      Caption         =   "frmW3Print.frx":0CF4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0D62
      Keys            =   "frmW3Print.frx":0D80
      Spin            =   "frmW3Print.frx":0DCA
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox6 
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0DF2
      Caption         =   "frmW3Print.frx":0E12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0E7A
      Keys            =   "frmW3Print.frx":0E98
      Spin            =   "frmW3Print.frx":0EE2
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox7 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   1920
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":0F0A
      Caption         =   "frmW3Print.frx":0F2A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":0F92
      Keys            =   "frmW3Print.frx":0FB0
      Spin            =   "frmW3Print.frx":0FFA
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox8 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":1022
      Caption         =   "frmW3Print.frx":1042
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":10B0
      Keys            =   "frmW3Print.frx":10CE
      Spin            =   "frmW3Print.frx":1118
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox9 
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   2400
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":1140
      Caption         =   "frmW3Print.frx":1160
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":11D2
      Keys            =   "frmW3Print.frx":11F0
      Spin            =   "frmW3Print.frx":123A
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox10 
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   2400
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":1262
      Caption         =   "frmW3Print.frx":1282
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":12FA
      Keys            =   "frmW3Print.frx":1318
      Spin            =   "frmW3Print.frx":1362
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox11 
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   2880
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":138A
      Caption         =   "frmW3Print.frx":13AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":141C
      Keys            =   "frmW3Print.frx":143A
      Spin            =   "frmW3Print.frx":1484
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox12 
      Height          =   375
      Left            =   8520
      TabIndex        =   16
      Top             =   2880
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":14AC
      Caption         =   "frmW3Print.frx":14CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":153C
      Keys            =   "frmW3Print.frx":155A
      Spin            =   "frmW3Print.frx":15A4
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox13 
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   3360
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":15CC
      Caption         =   "frmW3Print.frx":15EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1666
      Keys            =   "frmW3Print.frx":1684
      Spin            =   "frmW3Print.frx":16CE
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox14 
      Height          =   375
      Left            =   8520
      TabIndex        =   18
      Top             =   3360
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":16F6
      Caption         =   "frmW3Print.frx":1716
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":178A
      Keys            =   "frmW3Print.frx":17A8
      Spin            =   "frmW3Print.frx":17F2
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox16 
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3840
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":181A
      Caption         =   "frmW3Print.frx":183A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":18AC
      Keys            =   "frmW3Print.frx":18CA
      Spin            =   "frmW3Print.frx":1914
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox17 
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   3840
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":193C
      Caption         =   "frmW3Print.frx":195C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":19D2
      Keys            =   "frmW3Print.frx":19F0
      Spin            =   "frmW3Print.frx":1A3A
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox18 
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4320
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":1A62
      Caption         =   "frmW3Print.frx":1A82
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1AF4
      Keys            =   "frmW3Print.frx":1B12
      Spin            =   "frmW3Print.frx":1B5C
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
      MinValue        =   -99999999
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox19 
      Height          =   375
      Left            =   8520
      TabIndex        =   22
      Top             =   4320
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":1B84
      Caption         =   "frmW3Print.frx":1BA4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1C1A
      Keys            =   "frmW3Print.frx":1C38
      Spin            =   "frmW3Print.frx":1C82
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
      MinValue        =   -99999999
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
   Begin TDBText6Ctl.TDBText tdbEMail 
      Height          =   615
      Left            =   6120
      TabIndex        =   29
      Top             =   5640
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":1CAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1D1A
      Key             =   "frmW3Print.frx":1D38
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
   Begin TDBText6Ctl.TDBText tdbPhone 
      Height          =   615
      Left            =   6120
      TabIndex        =   30
      Top             =   6360
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":1D7C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1DF2
      Key             =   "frmW3Print.frx":1E10
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
   Begin TDBText6Ctl.TDBText tdbFaxNumber 
      Height          =   615
      Left            =   6120
      TabIndex        =   31
      Top             =   7080
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":1E54
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1EBE
      Key             =   "frmW3Print.frx":1EDC
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
   Begin TDBText6Ctl.TDBText tdbBoxF 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":1F20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":1F94
      Key             =   "frmW3Print.frx":1FB2
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
   Begin TDBText6Ctl.TDBText tdbBoxG1 
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":1FF6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":2082
      Key             =   "frmW3Print.frx":20A0
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
   Begin TDBText6Ctl.TDBText tdbBoxG3 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print.frx":20E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":2148
      Key             =   "frmW3Print.frx":2166
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
   Begin TDBText6Ctl.TDBText tdbBoxG4 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6480
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print.frx":21AA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":220E
      Key             =   "frmW3Print.frx":222C
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
   Begin TDBText6Ctl.TDBText tdbBoxH 
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":2270
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":22FA
      Key             =   "frmW3Print.frx":2318
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
   Begin TDBText6Ctl.TDBText tdbBox15A 
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   8040
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":235C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":23C2
      Key             =   "frmW3Print.frx":23E0
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
   Begin TDBText6Ctl.TDBText tdbBox15B 
      Height          =   615
      Left            =   1560
      TabIndex        =   33
      Top             =   8040
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   1085
      Caption         =   "frmW3Print.frx":2424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":24AC
      Key             =   "frmW3Print.frx":24CA
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
   Begin TDBNumber6Ctl.TDBNumber tdbVertNudge 
      Height          =   375
      Left            =   7560
      TabIndex        =   37
      Top             =   9000
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmW3Print.frx":250E
      Caption         =   "frmW3Print.frx":252E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print.frx":25A0
      Keys            =   "frmW3Print.frx":25BE
      Spin            =   "frmW3Print.frx":2608
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
   Begin VB.Label lblTaxYear 
      Caption         =   "TxYr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Height          =   255
      Left            =   1725
      TabIndex        =   41
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      Caption         =   "Kind of Payer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmW3Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TaxYear As Long
Dim GID(4) As Long
Dim i, j As Long

Private Sub Form_Load()

    PrvwReturn = True

    Me.tdbDate.Value = Now()

    TaxYear = frmW2Print.TaxYear
    Me.lblTaxYear = TaxYear
    Me.lblCompanyName = PRCompany.Name

    SetNudge Me.tdbHorzNudge
    SetNudge Me.tdbVertNudge
    
    GetNudge User.ID, "W3"
    Me.tdbHorzNudge = HorzNudge
    Me.tdbVertNudge = VertNudge

    ' kind of payer drop-down
    With Me.cmbPayer
        .AddItem "941"
        .AddItem "Military"
        .AddItem "943"
        .AddItem "944"
        .AddItem "CT-1"
        .AddItem "Hshld emp"
        .AddItem "Med Govt Emp"
        .AddItem "Third Party Sick Pay"
        .ListIndex = 0
    End With
    
    tdbAmountSet Me.tdbBox1
    tdbAmountSet Me.tdbBox2
    tdbAmountSet Me.tdbBox3
    tdbAmountSet Me.tdbBox4
    tdbAmountSet Me.tdbBox5
    tdbAmountSet Me.tdbBox6
    tdbAmountSet Me.tdbBox7
    tdbAmountSet Me.tdbBox8
    tdbAmountSet Me.tdbBox9
    tdbAmountSet Me.tdbBox10
    tdbAmountSet Me.tdbBox11
    tdbAmountSet Me.tdbBox12
    tdbAmountSet Me.tdbBox13
    tdbAmountSet Me.tdbBox14
    tdbAmountSet Me.tdbBox16
    tdbAmountSet Me.tdbBox17
    tdbAmountSet Me.tdbBox18
    tdbAmountSet Me.tdbBox19

    ' load the Company data from PRGlobal
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3A & " AND " & _
                "Year = " & TaxYear & " AND " & _
                "UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        
        GID(1) = PRGlobal.GlobalID
        Me.tdbBox1 = PRGlobal.Var1
        Me.tdbBox2 = PRGlobal.Var2
        Me.tdbBox3 = PRGlobal.Var3
        Me.tdbBox4 = PRGlobal.Var4
        Me.tdbBox5 = PRGlobal.Var5
        Me.tdbBox6 = PRGlobal.Var6
        Me.tdbBox7 = PRGlobal.Var7
        Me.tdbBox8 = PRGlobal.Var8
        Me.tdbBox9 = PRGlobal.Var9
        Me.tdbBox10 = PRGlobal.Var10

        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3B & " AND " & _
                    "Year = " & TaxYear & " AND " & _
                    "UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = False Then PRGlobal.Clear
        GID(2) = PRGlobal.GlobalID
        Me.tdbBox11 = PRGlobal.Var1
        Me.tdbBox12 = PRGlobal.Var2
        Me.tdbBox13 = PRGlobal.Var3
        Me.tdbBox14 = PRGlobal.Var4
        Me.tdbBox15A = PRGlobal.Var5
        Me.tdbBox16 = PRGlobal.Var6
        Me.tdbBox17 = PRGlobal.Var7
        Me.tdbBox18 = PRGlobal.Var8
        Me.tdbBox19 = PRGlobal.Var9
        Me.tdbBox15B = PRGlobal.Var10

        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3C & " AND " & _
                    "Year = " & TaxYear & " AND " & _
                    "UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = False Then PRGlobal.Clear
        GID(3) = PRGlobal.GlobalID
        Me.cmbPayer.ListIndex = PRGlobal.Var1
        Me.tdbBoxC = PRGlobal.Var2
        Me.tdbBoxD = PRGlobal.Var3
        Me.tdbBoxE = PRGlobal.Var4
        Me.tdbBoxF = PRGlobal.Var5
        Me.tdbBoxG1 = PRGlobal.Var6
        Me.tdbBoxG2 = PRGlobal.Var7
        Me.tdbBoxG3 = PRGlobal.Var8
        Me.tdbBoxG4 = PRGlobal.Var9
        Me.tdbBoxH = PRGlobal.Var10

    Else
        ' ?????
    End If

    ' submitter info
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3E & " AND " & _
                "UserID = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        GID(4) = PRGlobal.GlobalID
        Me.tdbContactPerson = PRGlobal.Var1
        Me.tdbEMail = PRGlobal.Var2
        Me.tdbPhone = PRGlobal.Var3
        Me.tdbFaxNumber = PRGlobal.Var4
        Me.tdbTitle = PRGlobal.Var5
    End If

    Me.KeyPreview = True

End Sub
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    ' save info to PRGlobal
    For i = 1 To 4
        
        If GID(i) = 0 Then
            MsgBox "PRGlobal Save Err?", vbExclamation
            GoBack
        End If
        If PRGlobal.GetByID(GID(i)) = False Then
            MsgBox "PRGlobal Save Err?", vbExclamation
            GoBack
        End If
        
        If i = 1 Then
            PRGlobal.Month = Me.cmbPayer.ListIndex
            PRGlobal.Var1 = nNull(Me.tdbBox1)
            PRGlobal.Var2 = nNull(Me.tdbBox2)
            PRGlobal.Var3 = nNull(Me.tdbBox3)
            PRGlobal.Var4 = nNull(Me.tdbBox4)
            PRGlobal.Var5 = nNull(Me.tdbBox5)
            PRGlobal.Var6 = nNull(Me.tdbBox6)
            PRGlobal.Var7 = nNull(Me.tdbBox7)
            PRGlobal.Var8 = nNull(Me.tdbBox8)
            PRGlobal.Var9 = nNull(Me.tdbBox9)
            PRGlobal.Var10 = nNull(Me.tdbBox10)
        ElseIf i = 2 Then
            PRGlobal.Var1 = nNull(Me.tdbBox11)
            PRGlobal.Var2 = nNull(Me.tdbBox12)
            PRGlobal.Var3 = nNull(Me.tdbBox13)
            PRGlobal.Var4 = nNull(Me.tdbBox14)
            PRGlobal.Var5 = Me.tdbBox15A & ""
            PRGlobal.Var6 = nNull(Me.tdbBox16)
            PRGlobal.Var7 = nNull(Me.tdbBox17)
            PRGlobal.Var8 = nNull(Me.tdbBox18)
            PRGlobal.Var9 = nNull(Me.tdbBox19)
            PRGlobal.Var10 = Me.tdbBox15B & ""
        ElseIf i = 3 Then
            PRGlobal.Var1 = Me.cmbPayer.ListIndex
            PRGlobal.Var2 = nNull(Me.tdbBoxC)
            PRGlobal.Var3 = Me.tdbBoxD & ""
            PRGlobal.Var4 = Me.tdbBoxE & ""
            PRGlobal.Var5 = Me.tdbBoxF & ""
            PRGlobal.Var6 = Me.tdbBoxG1 & ""
            PRGlobal.Var7 = Me.tdbBoxG2 & ""
            PRGlobal.Var8 = Me.tdbBoxG3 & ""
            PRGlobal.Var9 = Me.tdbBoxG4 & ""
            PRGlobal.Var10 = Me.tdbBoxH & ""
        ElseIf i = 4 Then
            PRGlobal.Var1 = Me.tdbContactPerson & ""
            PRGlobal.Var2 = Me.tdbEMail & ""
            PRGlobal.Var3 = Me.tdbPhone & ""
            PRGlobal.Var4 = Me.tdbFaxNumber & ""
            PRGlobal.Var5 = Me.tdbTitle & ""
        End If
    
        PRGlobal.Save (Equate.RecPut)
    
    Next i
    
    HorzNudge = Me.tdbHorzNudge
    VertNudge = Me.tdbVertNudge
    SaveNudge User.ID, "W3"
    
    PrtInit ("Port")
    SetFont 10, Equate.Portrait

    W3Print

End Sub

Private Sub W3Print()
    
Dim CurX, CurY As Long
Dim BxG(4) As String
    
    ' kind of payer
    If Me.cmbPayer.ListIndex <= 3 Then
        CurY = 1180
    Else
        CurY = 1670
    End If
    Select Case Me.cmbPayer.ListIndex
        Case 0, 4
            CurX = 2160
        Case 1, 5
            CurX = 2820
        Case 2, 6
            CurX = 3540
        Case 3, 7
            CurX = 4440
    End Select
    PosPrint CurX, CurY, "X"
    
    ' Box G - ER address
    j = 0
    For i = 1 To 4
        If i = 1 Then x = Me.tdbBoxG1
        If i = 2 Then x = Me.tdbBoxG2
        If i = 3 Then x = Me.tdbBoxG3
        If i = 4 Then x = Me.tdbBoxG4
        If x <> "" Then
            j = j + 1
            BxG(j) = x
        End If
    Next i
    
    YUnits = 240
    Ln = 5
    
    PrintValue(1) = " ":                        FormatString(1) = "a55"
    PrintValue(2) = Me.tdbBox1:                 FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Me.tdbBox2:                 FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a55"
    PrintValue(2) = Me.tdbBox3:                 FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Me.tdbBox4:                 FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a6"
    PrintValue(2) = Me.tdbBoxC:                 FormatString(2) = "n5"
    PrintValue(3) = " ":                        FormatString(3) = "a44"
    PrintValue(4) = Me.tdbBox5:                 FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a10"
    PrintValue(6) = Me.tdbBox6:                 FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = Me.tdbBoxE:                 FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "a38"
    PrintValue(4) = Me.tdbBox7:                 FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a10"
    PrintValue(6) = Me.tdbBox8:                 FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = Me.tdbBoxF:                 FormatString(2) = "a30":
    PrintValue(3) = " ":                        FormatString(3) = "a18"
    PrintValue(4) = Me.tdbBox9:                 FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a10"
    PrintValue(6) = Me.tdbBox10:                FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = BxG(1):                     FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Me.tdbBox11:                FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "a10"
    PrintValue(6) = Me.tdbBox12:                FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = BxG(2):                     FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1

    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = BxG(3):                     FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Me.tdbBox13:                FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 1
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = BxG(4):                     FormatString(2) = "a50"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 1
        
    PrintValue(1) = " ":                        FormatString(1) = "a55"
    PrintValue(2) = Me.tdbBox14:                FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a7"
    PrintValue(2) = Me.tdbBoxH:                 FormatString(2) = "a10"
    PrintValue(3) = " ":                        FormatString(3) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a5"
    PrintValue(2) = Me.tdbBox15A:               FormatString(2) = "a2"  '  Line 15
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Me.tdbBox15B:               FormatString(4) = "a10"
    PrintValue(5) = " ":                        FormatString(5) = "a28"
    PrintValue(6) = Me.tdbBox16:                FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "a10"
    PrintValue(8) = Me.tdbBox17:                FormatString(8) = "d12"
    PrintValue(9) = " ":                        FormatString(9) = "~"
    FormatPrint
    Ln = Ln + 2
    
    PrintValue(1) = " ":                        FormatString(1) = "a55"
    PrintValue(2) = Me.tdbBox18:                FormatString(2) = "d12"
    PrintValue(3) = " ":                        FormatString(3) = "a10"
    PrintValue(4) = Me.tdbBox19:                FormatString(4) = "d12"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2

    PrintValue(1) = " ":                        FormatString(1) = "a4"
    PrintValue(2) = Me.tdbContactPerson:        FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Me.tdbPhone:                FormatString(4) = "a20"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    
    ' 2020-01-18 switch order of fax#/email
    PrintValue(1) = " ":                        FormatString(1) = "a4"
    PrintValue(2) = Me.tdbFaxNumber:            FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Me.tdbEMail:                FormatString(4) = "a40"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 5
    
    PrintValue(1) = " ":                        FormatString(1) = "a47"
    PrintValue(2) = Me.tdbTitle:                FormatString(2) = "a33"
    PrintValue(3) = " ":                        FormatString(3) = "a2"
    PrintValue(4) = Format(Me.tdbDate, "mm/dd/yyyy"): FormatString(4) = "a10"
    PrintValue(5) = " ":                        FormatString(5) = "~"
    FormatPrint
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

