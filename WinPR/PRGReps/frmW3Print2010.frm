VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmW3Print2010 
   Caption         =   "W3 Form Print"
   ClientHeight    =   10170
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
   ScaleHeight     =   10170
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate tdbDate 
      Height          =   615
      Left            =   6120
      TabIndex        =   36
      Top             =   9240
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1085
      Calendar        =   "frmW3Print2010.frx":0000
      Caption         =   "frmW3Print2010.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":015E
      Keys            =   "frmW3Print2010.frx":017C
      Spin            =   "frmW3Print2010.frx":01DA
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
   Begin TDBNumber6Ctl.TDBNumber tdbHorzNudge 
      Height          =   375
      Left            =   8160
      TabIndex        =   37
      Top             =   8760
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":0202
      Caption         =   "frmW3Print2010.frx":0222
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0298
      Keys            =   "frmW3Print2010.frx":02B6
      Spin            =   "frmW3Print2010.frx":0300
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
      Left            =   11280
      TabIndex        =   40
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   11280
      TabIndex        =   39
      Top             =   8760
      Width           =   1335
   End
   Begin TDBText6Ctl.TDBText tdbBoxG2 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6120
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print2010.frx":0328
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":038C
      Key             =   "frmW3Print2010.frx":03AA
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
      TabIndex        =   29
      Top             =   5520
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":03EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0460
      Key             =   "frmW3Print2010.frx":047E
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
      Caption         =   "frmW3Print2010.frx":04C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0540
      Key             =   "frmW3Print2010.frx":055E
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
      Caption         =   "frmW3Print2010.frx":05A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0612
      Key             =   "frmW3Print2010.frx":0630
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
      Calculator      =   "frmW3Print2010.frx":0674
      Caption         =   "frmW3Print2010.frx":0694
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":070A
      Keys            =   "frmW3Print2010.frx":0728
      Spin            =   "frmW3Print2010.frx":0772
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
      Calculator      =   "frmW3Print2010.frx":079A
      Caption         =   "frmW3Print2010.frx":07BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0832
      Keys            =   "frmW3Print2010.frx":0850
      Spin            =   "frmW3Print2010.frx":089A
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
      Calculator      =   "frmW3Print2010.frx":08C2
      Caption         =   "frmW3Print2010.frx":08E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0952
      Keys            =   "frmW3Print2010.frx":0970
      Spin            =   "frmW3Print2010.frx":09BA
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
      Calculator      =   "frmW3Print2010.frx":09E2
      Caption         =   "frmW3Print2010.frx":0A02
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0A6C
      Keys            =   "frmW3Print2010.frx":0A8A
      Spin            =   "frmW3Print2010.frx":0AD4
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
      Calculator      =   "frmW3Print2010.frx":0AFC
      Caption         =   "frmW3Print2010.frx":0B1C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0B82
      Keys            =   "frmW3Print2010.frx":0BA0
      Spin            =   "frmW3Print2010.frx":0BEA
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
      Calculator      =   "frmW3Print2010.frx":0C12
      Caption         =   "frmW3Print2010.frx":0C32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0CA0
      Keys            =   "frmW3Print2010.frx":0CBE
      Spin            =   "frmW3Print2010.frx":0D08
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
      Calculator      =   "frmW3Print2010.frx":0D30
      Caption         =   "frmW3Print2010.frx":0D50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0DB8
      Keys            =   "frmW3Print2010.frx":0DD6
      Spin            =   "frmW3Print2010.frx":0E20
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
      Calculator      =   "frmW3Print2010.frx":0E48
      Caption         =   "frmW3Print2010.frx":0E68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0ED0
      Keys            =   "frmW3Print2010.frx":0EEE
      Spin            =   "frmW3Print2010.frx":0F38
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
      Calculator      =   "frmW3Print2010.frx":0F60
      Caption         =   "frmW3Print2010.frx":0F80
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":0FEE
      Keys            =   "frmW3Print2010.frx":100C
      Spin            =   "frmW3Print2010.frx":1056
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
      Calculator      =   "frmW3Print2010.frx":107E
      Caption         =   "frmW3Print2010.frx":109E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1110
      Keys            =   "frmW3Print2010.frx":112E
      Spin            =   "frmW3Print2010.frx":1178
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
      Calculator      =   "frmW3Print2010.frx":11A0
      Caption         =   "frmW3Print2010.frx":11C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1238
      Keys            =   "frmW3Print2010.frx":1256
      Spin            =   "frmW3Print2010.frx":12A0
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
      Calculator      =   "frmW3Print2010.frx":12C8
      Caption         =   "frmW3Print2010.frx":12E8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":135A
      Keys            =   "frmW3Print2010.frx":1378
      Spin            =   "frmW3Print2010.frx":13C2
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
      Calculator      =   "frmW3Print2010.frx":13EA
      Caption         =   "frmW3Print2010.frx":140A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":147C
      Keys            =   "frmW3Print2010.frx":149A
      Spin            =   "frmW3Print2010.frx":14E4
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
      TabIndex        =   18
      Top             =   3840
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":150C
      Caption         =   "frmW3Print2010.frx":152C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":15A6
      Keys            =   "frmW3Print2010.frx":15C4
      Spin            =   "frmW3Print2010.frx":160E
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
      TabIndex        =   19
      Top             =   3840
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":1636
      Caption         =   "frmW3Print2010.frx":1656
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":16CA
      Keys            =   "frmW3Print2010.frx":16E8
      Spin            =   "frmW3Print2010.frx":1732
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
      TabIndex        =   20
      Top             =   4320
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":175A
      Caption         =   "frmW3Print2010.frx":177A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":17EC
      Keys            =   "frmW3Print2010.frx":180A
      Spin            =   "frmW3Print2010.frx":1854
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
      TabIndex        =   21
      Top             =   4320
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":187C
      Caption         =   "frmW3Print2010.frx":189C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1912
      Keys            =   "frmW3Print2010.frx":1930
      Spin            =   "frmW3Print2010.frx":197A
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
      TabIndex        =   22
      Top             =   4800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":19A2
      Caption         =   "frmW3Print2010.frx":19C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1A34
      Keys            =   "frmW3Print2010.frx":1A52
      Spin            =   "frmW3Print2010.frx":1A9C
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
      TabIndex        =   23
      Top             =   4800
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":1AC4
      Caption         =   "frmW3Print2010.frx":1AE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1B5A
      Keys            =   "frmW3Print2010.frx":1B78
      Spin            =   "frmW3Print2010.frx":1BC2
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
      TabIndex        =   30
      Top             =   6240
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":1BEA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1C5A
      Key             =   "frmW3Print2010.frx":1C78
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
      TabIndex        =   31
      Top             =   6960
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":1CBC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1D32
      Key             =   "frmW3Print2010.frx":1D50
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
      TabIndex        =   32
      Top             =   7680
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":1D94
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1DFE
      Key             =   "frmW3Print2010.frx":1E1C
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
      Caption         =   "frmW3Print2010.frx":1E60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1ED4
      Key             =   "frmW3Print2010.frx":1EF2
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
      TabIndex        =   24
      Top             =   5400
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":1F36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":1FC2
      Key             =   "frmW3Print2010.frx":1FE0
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
      TabIndex        =   26
      Top             =   6600
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print2010.frx":2024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":2088
      Key             =   "frmW3Print2010.frx":20A6
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
      TabIndex        =   27
      Top             =   7080
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   661
      Caption         =   "frmW3Print2010.frx":20EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":214E
      Key             =   "frmW3Print2010.frx":216C
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
      TabIndex        =   28
      Top             =   7560
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":21B0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":223A
      Key             =   "frmW3Print2010.frx":2258
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
      TabIndex        =   33
      Top             =   8400
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":229C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":2302
      Key             =   "frmW3Print2010.frx":2320
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
      TabIndex        =   34
      Top             =   8400
      Width           =   5535
      _Version        =   65536
      _ExtentX        =   9763
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":2364
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":23EC
      Key             =   "frmW3Print2010.frx":240A
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
      Left            =   8160
      TabIndex        =   38
      Top             =   9360
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":244E
      Caption         =   "frmW3Print2010.frx":246E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":24E0
      Keys            =   "frmW3Print2010.frx":24FE
      Spin            =   "frmW3Print2010.frx":2548
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
   Begin TDBNumber6Ctl.TDBNumber tdbBox12b 
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   3360
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmW3Print2010.frx":2570
      Caption         =   "frmW3Print2010.frx":2590
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":2604
      Keys            =   "frmW3Print2010.frx":2622
      Spin            =   "frmW3Print2010.frx":266C
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
   Begin TDBText6Ctl.TDBText tdbTitle 
      Height          =   615
      Left            =   120
      TabIndex        =   35
      Top             =   9240
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1085
      Caption         =   "frmW3Print2010.frx":2694
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmW3Print2010.frx":26F4
      Key             =   "frmW3Print2010.frx":2712
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
      TabIndex        =   43
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
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmW3Print2010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TaxYear As Long
Dim GID(5) As Long
Dim i, j As Long
Dim HIREAmt As Currency

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
    tdbAmountSet Me.tdbBox12b
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

        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3D & " AND " & _
                    "Year = " & TaxYear & " AND " & _
                    "UserID = " & PRCompany.CompanyID
        If PRGlobal.GetBySQL(SQLString) = False Then PRGlobal.Clear
        GID(4) = PRGlobal.GlobalID
        Me.tdbBox12b = PRGlobal.Var1
        Me.tdbBox12.Value = Me.tdbBox12.Value - Me.tdbBox12b.Value
    
    Else
        ' ?????
    End If

    ' submitter info
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW3E & " AND " & _
                "UserID = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        GID(5) = PRGlobal.GlobalID
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
    For i = 1 To 5
        
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
            PRGlobal.Var1 = Me.tdbBox12b
        ElseIf i = 5 Then
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
        If i = 1 Then X = Me.tdbBoxG1
        If i = 2 Then X = Me.tdbBoxG2
        If i = 3 Then X = Me.tdbBoxG3
        If i = 4 Then X = Me.tdbBoxG4
        If X <> "" Then
            j = j + 1
            BxG(j) = X
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
    PrintValue(5) = " ":                        FormatString(5) = "a10"
    PrintValue(6) = Me.tdbBox12b:               FormatString(6) = "d12"
    PrintValue(7) = " ":                        FormatString(7) = "~"
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
    
    PrintValue(1) = " ":                        FormatString(1) = "a4"
    PrintValue(2) = Me.tdbEMail:                FormatString(2) = "a40"
    PrintValue(3) = " ":                        FormatString(3) = "a8"
    PrintValue(4) = Me.tdbFaxNumber:            FormatString(4) = "a20"
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

