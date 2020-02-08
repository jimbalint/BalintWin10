VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmUpdateCustomer 
   Caption         =   "Add/Update Customer"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin TDBNumber6Ctl.TDBNumber TDBSign1Left 
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   6150
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":0000
      Caption         =   "frmAddCustomer.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":007C
      Keys            =   "frmAddCustomer.frx":009A
      Spin            =   "frmAddCustomer.frx":00E4
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
   Begin TDBNumber6Ctl.TDBNumber TDBAcctSpaces 
      Height          =   330
      Left            =   6360
      TabIndex        =   17
      Top             =   5160
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   582
      Calculator      =   "frmAddCustomer.frx":010C
      Caption         =   "frmAddCustomer.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0196
      Keys            =   "frmAddCustomer.frx":01B4
      Spin            =   "frmAddCustomer.frx":01FE
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
      MaxValueVT      =   5636101
      MinValueVT      =   3342341
   End
   Begin VB.CheckBox chkBoldAddr1 
      Caption         =   "Bold Addr1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8400
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin TDBNumber6Ctl.TDBNumber tdbCompanyID 
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   582
      Calculator      =   "frmAddCustomer.frx":0226
      Caption         =   "frmAddCustomer.frx":0246
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":02AE
      Keys            =   "frmAddCustomer.frx":02CC
      Spin            =   "frmAddCustomer.frx":0316
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
      EditMode        =   1
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
      MaxValueVT      =   5636101
      MinValueVT      =   3342341
   End
   Begin VB.CheckBox chkBoldAddr4 
      Caption         =   "Bold Addr4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8400
      TabIndex        =   33
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox chkBoldAddr3 
      Caption         =   "Bold Addr3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8400
      TabIndex        =   32
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox chkBoldAddr2 
      Caption         =   "Bold Addr2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   8400
      TabIndex        =   31
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox chkTwoSigs 
      Caption         =   "     Two Signatures"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   34
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelAddCust 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   30
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveCust 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   29
      Top             =   1680
      Width           =   1215
   End
   Begin TDBText6Ctl.TDBText TDBAddr1 
      Height          =   330
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":033E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":03A4
      Key             =   "frmAddCustomer.frx":03C2
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
   Begin TDBText6Ctl.TDBText TDBAddr2 
      Height          =   330
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0406
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":046C
      Key             =   "frmAddCustomer.frx":048A
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBAddr3 
      Height          =   330
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":04CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0534
      Key             =   "frmAddCustomer.frx":0552
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBAddr4 
      Height          =   330
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0596
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":05FC
      Key             =   "frmAddCustomer.frx":061A
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBBank1 
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":065E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":06BE
      Key             =   "frmAddCustomer.frx":06DC
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBBank2 
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0780
      Key             =   "frmAddCustomer.frx":079E
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
   Begin TDBText6Ctl.TDBText TDBBank3 
      Height          =   330
      Left            =   360
      TabIndex        =   11
      Top             =   3720
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":07E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0842
      Key             =   "frmAddCustomer.frx":0860
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBBank4 
      Height          =   330
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":08A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0904
      Key             =   "frmAddCustomer.frx":0922
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBBankFraction 
      Height          =   330
      Left            =   360
      TabIndex        =   13
      Top             =   4440
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0966
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":09D4
      Key             =   "frmAddCustomer.frx":09F2
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBBankAccount 
      Height          =   330
      Left            =   360
      TabIndex        =   15
      Top             =   5160
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0A36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0AA2
      Key             =   "frmAddCustomer.frx":0AC0
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBSignImage1 
      Height          =   330
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0B04
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0B6E
      Key             =   "frmAddCustomer.frx":0B8C
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBSignImage2 
      Height          =   330
      Left            =   360
      TabIndex        =   23
      Top             =   6650
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0BD0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0C3A
      Key             =   "frmAddCustomer.frx":0C58
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBLogo 
      Height          =   330
      Left            =   360
      TabIndex        =   28
      Top             =   7750
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0C9C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0D04
      Key             =   "frmAddCustomer.frx":0D22
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBText6Ctl.TDBText TDBCustName 
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   7815
      _Version        =   65536
      _ExtentX        =   13785
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0D66
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0DD4
      Key             =   "frmAddCustomer.frx":0DF2
      BackColor       =   -2147483643
      EditMode        =   1
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
   Begin TDBText6Ctl.TDBText TDBBankABA 
      Height          =   330
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Width           =   5895
      _Version        =   65536
      _ExtentX        =   10398
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":0E36
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0E9A
      Key             =   "frmAddCustomer.frx":0EB8
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign1Top 
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   6150
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":0EFC
      Caption         =   "frmAddCustomer.frx":0F1C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":0F78
      Keys            =   "frmAddCustomer.frx":0F96
      Spin            =   "frmAddCustomer.frx":0FE0
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign1Height 
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   6150
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":1008
      Caption         =   "frmAddCustomer.frx":1028
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":1084
      Keys            =   "frmAddCustomer.frx":10A2
      Spin            =   "frmAddCustomer.frx":10EC
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign1Width 
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   6150
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":1114
      Caption         =   "frmAddCustomer.frx":1134
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":1190
      Keys            =   "frmAddCustomer.frx":11AE
      Spin            =   "frmAddCustomer.frx":11F8
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign2Left 
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   7265
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":1220
      Caption         =   "frmAddCustomer.frx":1240
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":129C
      Keys            =   "frmAddCustomer.frx":12BA
      Spin            =   "frmAddCustomer.frx":1304
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
   Begin TDBNumber6Ctl.TDBNumber tdbSign2Top 
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   7265
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":132C
      Caption         =   "frmAddCustomer.frx":134C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":13A8
      Keys            =   "frmAddCustomer.frx":13C6
      Spin            =   "frmAddCustomer.frx":1410
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign2Height 
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   7265
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":1438
      Caption         =   "frmAddCustomer.frx":1458
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":14B4
      Keys            =   "frmAddCustomer.frx":14D2
      Spin            =   "frmAddCustomer.frx":151C
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
   Begin TDBNumber6Ctl.TDBNumber TDBSign2Width 
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   7265
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":1544
      Caption         =   "frmAddCustomer.frx":1564
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":15C0
      Keys            =   "frmAddCustomer.frx":15DE
      Spin            =   "frmAddCustomer.frx":1628
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
   Begin TDBText6Ctl.TDBText TDBClientIdName 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   7215
      _Version        =   65536
      _ExtentX        =   12726
      _ExtentY        =   873
      Caption         =   "frmAddCustomer.frx":1650
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":16C0
      Key             =   "frmAddCustomer.frx":16DE
      BackColor       =   -2147483629
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
   Begin TDBText6Ctl.TDBText tdbBankAccountAdd 
      Height          =   330
      Left            =   9000
      TabIndex        =   18
      Top             =   5160
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   582
      Caption         =   "frmAddCustomer.frx":1722
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":178A
      Key             =   "frmAddCustomer.frx":17A8
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
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   40
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
   Begin TDBNumber6Ctl.TDBNumber tdbAddressAdjust 
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   661
      Calculator      =   "frmAddCustomer.frx":17EC
      Caption         =   "frmAddCustomer.frx":180C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddCustomer.frx":187A
      Keys            =   "frmAddCustomer.frx":1898
      Spin            =   "frmAddCustomer.frx":18E2
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
   Begin VB.Label Label8 
      Caption         =   "Sign2 Width"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   42
      Top             =   7000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Sign2 Height"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   41
      Top             =   7000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Sign2 Top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   40
      Top             =   7000
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Sign2 Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   39
      Top             =   7000
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sign1 Width"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   38
      Top             =   5900
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Sign1 Height"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   37
      Top             =   5900
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sign1 Top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   5900
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sign1 Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   35
      Top             =   5900
      Width           =   975
   End
End
Attribute VB_Name = "frmUpdateCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    FieldFormat Me.TDBCustName
    FieldFormat Me.TDBAddr1
    FieldFormat Me.TDBAddr2
    FieldFormat Me.TDBAddr3
    FieldFormat Me.TDBAddr4
    FieldFormat Me.TDBBank1
    FieldFormat Me.TDBBank2
    FieldFormat Me.TDBBank3
    FieldFormat Me.TDBBank4

    TDBClientIdName = ClientID & "  -  " & ClientName
    
    If EditSw = True Then
        If Not Customer.GetByID(CustID) Then
            MsgBox "Customer ID Not Found: " & SelID, vbExclamation
            End
        Else
            Me.tdbCompanyID = Customer.PRCompanyID
            Me.TDBCustName = Customer.CustomerName
            Me.TDBAddr1 = Customer.Address1
            Me.TDBAddr2 = Customer.Address2
            Me.TDBAddr3 = Customer.Address3
            Me.TDBAddr4 = Customer.Address4
            Me.chkBoldAddr1 = Customer.Addr1Bold
            Me.chkBoldAddr2 = Customer.Addr2Bold
            Me.chkBoldAddr3 = Customer.Addr3Bold
            Me.chkBoldAddr4 = Customer.Addr4Bold
            Me.TDBBank1 = Customer.Bank1
            Me.TDBBank2 = Customer.Bank2
            Me.TDBBank3 = Customer.Bank3
            Me.TDBBank4 = Customer.Bank4
            Me.TDBBankFraction = Customer.BankFraction
            Me.TDBBankAccount = Customer.BankAccount
            Me.chkTwoSigs = Customer.TwoSignLines
            Me.TDBAcctSpaces = Customer.AccountSpace
            Me.TDBBankABA = Customer.BankABA           '''' NEEDS TO be DEFINED As A NUMBER
            Me.TDBSignImage1 = Customer.SignImage1
            Me.TDBSign1Left = Customer.Sign1Left
            Me.TDBSign1Top = Customer.Sign1Top
            Me.TDBSign1Height = Customer.Sign1Height
            Me.TDBSign1Width = Customer.Sign1Width
            Me.TDBSignImage2 = Customer.SignImage2
            Me.TDBSign2Left = Customer.Sign2Left
            Me.tdbSign2Top = Customer.Sign2Top
            Me.TDBSign2Height = Customer.Sign2Height
            Me.TDBSign2Width = Customer.Sign2Width
            Me.TDBLogo = Customer.LogoImage
            
            Me.tdbBankAccountAdd = Customer.BankAccountAdd
            Me.tdbAddressAdjust = Customer.AddressAdjust
            
        End If
    Else
   
        Customer.OpenRS
        Customer.Clear
    End If
 End Sub
    
Private Sub cmdCancelAddCust_Click()
    Me.Hide
End Sub

Private Sub cmdSaveCust_Click()

    '  add/Update it
    
    Customer.ClientID = ClientID
    Customer.PRCompanyID = Me.tdbCompanyID
    Customer.CustomerName = Me.TDBCustName
    Customer.Address1 = Me.TDBAddr1
    Customer.Address2 = Me.TDBAddr2
    Customer.Address3 = Me.TDBAddr3
    Customer.Address4 = Me.TDBAddr4
    Customer.Addr1Bold = Me.chkBoldAddr1
    Customer.Addr2Bold = Me.chkBoldAddr2
    Customer.Addr3Bold = Me.chkBoldAddr3
    Customer.Addr4Bold = Me.chkBoldAddr4
    Customer.Bank1 = Me.TDBBank1
    Customer.Bank2 = Me.TDBBank2
    Customer.Bank3 = Me.TDBBank3
    Customer.Bank4 = Me.TDBBank4
    Customer.BankFraction = Me.TDBBankFraction
    Customer.BankABA = Me.TDBBankABA
    Customer.BankAccount = Me.TDBBankAccount
    Customer.AccountSpace = Me.TDBAcctSpaces
    Customer.SignImage1 = Me.TDBSignImage1
    Customer.Sign1Left = Me.TDBSign1Left
    Customer.Sign1Top = Me.TDBSign1Top
    Customer.Sign1Height = Me.TDBSign1Height
    Customer.Sign1Width = Me.TDBSign1Width
    Customer.SignImage2 = Me.TDBSignImage2
    Customer.Sign2Left = Me.TDBSign2Left
    Customer.Sign2Top = Me.tdbSign2Top
    Customer.Sign2Height = Me.TDBSign2Height
    Customer.Sign2Width = Me.TDBSign2Width
    Customer.LogoImage = Me.TDBLogo
    Customer.TwoSignLines = Me.chkTwoSigs

    Customer.BankAccountAdd = Me.tdbBankAccountAdd
    Customer.AddressAdjust = Me.tdbAddressAdjust

    If EditSw = False Then
        Customer.Save (RecAdd)
    Else
        Customer.Save (RecPut)
    End If
 
    Me.TDBCustName.Text = Customer.CustomerName
    Me.Hide

End Sub


Private Sub FieldFormat(ByRef tdbTxt As TDBText)
    With tdbTxt
        .MaxLength = 40
        .Text = ""
        .Key.Clear = ""       ' no key to clear field
        .FormatMode = dbiIncludeFormat
        .Format = "A9#@"
    End With
'    FormatString = UCase(tdbTxt)
    
End Sub

