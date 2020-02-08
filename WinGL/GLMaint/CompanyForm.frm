VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form CompanyForm 
   Caption         =   " Company Data Fields"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CompanyForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbRetEarnAcct 
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":030A
      Caption         =   "CompanyForm.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0396
      Keys            =   "CompanyForm.frx":03B4
      Spin            =   "CompanyForm.frx":03FE
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
   Begin TDBNumber6Ctl.TDBNumber tdbLowBranch 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0426
      Caption         =   "CompanyForm.frx":0446
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":04B2
      Keys            =   "CompanyForm.frx":04D0
      Spin            =   "CompanyForm.frx":051A
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
   Begin TDBNumber6Ctl.TDBNumber tdbLastClose 
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0542
      Caption         =   "CompanyForm.frx":0562
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":05CE
      Keys            =   "CompanyForm.frx":05EC
      Spin            =   "CompanyForm.frx":0636
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "00000000;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "00000000"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBText6Ctl.TDBText txtZipCode 
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   6120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":065E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":06CA
      Key             =   "CompanyForm.frx":06E8
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
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
      Left            =   6000
      TabIndex        =   19
      Top             =   6120
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":072C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0798
      Key             =   "CompanyForm.frx":07B6
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
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
      Left            =   5280
      TabIndex        =   17
      Top             =   5640
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":07FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0866
      Key             =   "CompanyForm.frx":0884
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtAddress3 
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   6600
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":08C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0934
      Key             =   "CompanyForm.frx":0952
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtAddress2 
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   6120
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":0996
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0A02
      Key             =   "CompanyForm.frx":0A20
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtAddress1 
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   5640
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":0A64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0AD0
      Key             =   "CompanyForm.frx":0AEE
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText txtName 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "CompanyForm.frx":0B32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0B9E
      Key             =   "CompanyForm.frx":0BBC
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
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber txtFirstPAcct 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0C00
      Caption         =   "CompanyForm.frx":0C20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0C8C
      Keys            =   "CompanyForm.frx":0CAA
      Spin            =   "CompanyForm.frx":0CF4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtNetProfitAcct 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0D1C
      Caption         =   "CompanyForm.frx":0D3C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0DA8
      Keys            =   "CompanyForm.frx":0DC6
      Spin            =   "CompanyForm.frx":0E10
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   6553605
      MinValueVT      =   6619141
   End
   Begin TDBNumber6Ctl.TDBNumber txtSuspAcct 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0E38
      Caption         =   "CompanyForm.frx":0E58
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0EC4
      Keys            =   "CompanyForm.frx":0EE2
      Spin            =   "CompanyForm.frx":0F2C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtFirstFiscalYear 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   4200
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":0F54
      Caption         =   "CompanyForm.frx":0F74
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":0FE0
      Keys            =   "CompanyForm.frx":0FFE
      Spin            =   "CompanyForm.frx":1048
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   6553605
      MinValueVT      =   6619141
   End
   Begin TDBNumber6Ctl.TDBNumber txtCurFiscalYear 
      Height          =   375
      Left            =   2600
      TabIndex        =   13
      Top             =   4200
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":1070
      Caption         =   "CompanyForm.frx":1090
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":10FC
      Keys            =   "CompanyForm.frx":111A
      Spin            =   "CompanyForm.frx":1164
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   2050
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtCurPeriod 
      Height          =   375
      Left            =   4240
      TabIndex        =   14
      Top             =   4200
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":118C
      Caption         =   "CompanyForm.frx":11AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":1218
      Keys            =   "CompanyForm.frx":1236
      Spin            =   "CompanyForm.frx":1280
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   13
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   1
      MaxValueVT      =   6553605
      MinValueVT      =   6619141
   End
   Begin TDBNumber6Ctl.TDBNumber txtFirstPeriod 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":12A8
      Caption         =   "CompanyForm.frx":12C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":1334
      Keys            =   "CompanyForm.frx":1352
      Spin            =   "CompanyForm.frx":139C
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   13
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   1
      MaxValueVT      =   6553605
      MinValueVT      =   6619141
   End
   Begin TDBNumber6Ctl.TDBNumber txtNumberPds 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":13C4
      Caption         =   "CompanyForm.frx":13E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":1450
      Keys            =   "CompanyForm.frx":146E
      Spin            =   "CompanyForm.frx":14B8
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   13
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   12
      MaxValueVT      =   5242885
      MinValueVT      =   3014661
   End
   Begin TDBNumber6Ctl.TDBNumber txtSubDigits 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":14E0
      Caption         =   "CompanyForm.frx":1500
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":156C
      Keys            =   "CompanyForm.frx":158A
      Spin            =   "CompanyForm.frx":15D4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "0;;0;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   39976961
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   22
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   23
      Top             =   7560
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiBranch 
      Height          =   375
      Left            =   4040
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":15FC
      Caption         =   "CompanyForm.frx":161C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":1688
      Keys            =   "CompanyForm.frx":16A6
      Spin            =   "CompanyForm.frx":16F0
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
   Begin TDBNumber6Ctl.TDBNumber tdbLowConsolidated 
      Height          =   375
      Left            =   5800
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":1718
      Caption         =   "CompanyForm.frx":1738
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":17A4
      Keys            =   "CompanyForm.frx":17C2
      Spin            =   "CompanyForm.frx":180C
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
   Begin TDBNumber6Ctl.TDBNumber tdbHiConsolidated 
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "CompanyForm.frx":1834
      Caption         =   "CompanyForm.frx":1854
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "CompanyForm.frx":18C0
      Keys            =   "CompanyForm.frx":18DE
      Spin            =   "CompanyForm.frx":1928
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
   Begin VB.Label Label20 
      Caption         =   "Retained Earnings Account"
      Height          =   615
      Left            =   8040
      TabIndex        =   44
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label19 
      Caption         =   "High Consol."
      Height          =   255
      Left            =   7560
      TabIndex        =   43
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Low Consol."
      Height          =   255
      Left            =   5880
      TabIndex        =   42
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "High Branch"
      Height          =   255
      Left            =   4080
      TabIndex        =   41
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Low Branch"
      Height          =   255
      Left            =   2280
      TabIndex        =   40
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Last Close Date (YYYYMMDD)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   39
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Digits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "1st Fiscal Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   37
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Current Fiscal Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   36
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Current Period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   35
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   34
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "First Period"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Number of Periods"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   32
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "First P Acct."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   31
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Net profit Acct."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Susp. Acct."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Zip Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   28
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblDBName 
      Caption         =   "Data File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "CompanyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public userOK As Boolean

Dim LastClose As Date
Private Sub Form_Load()

Dim tdbNum As TDBNumber
       
    ' settings for tdb controls
    tdbSet

    ' populate fields
    lblDBName = "File Name: " & GLCompany.FileName
    txtName = GLCompany.Name & ""
    txtAddress1 = GLCompany.Address1 & ""
    txtAddress2 = GLCompany.Address2 & ""
    txtAddress3 = GLCompany.Address3 & ""
    txtCity = GLCompany.City & ""
    txtState = GLCompany.State & ""
    txtZipCode = GLCompany.ZipCode & ""
   
    If GLCompany.NumberPds = 0 Then GLCompany.NumberPds = 12
    txtNumberPds = GLCompany.NumberPds
    
    If GLCompany.FirstPeriod = 0 Then GLCompany.FirstPeriod = 1
    txtFirstPeriod = GLCompany.FirstPeriod
    
    txtCurPeriod = GLCompany.CurPeriod
    txtCurFiscalYear = GLCompany.CurFiscalYear
    
    If GLCompany.FirstFiscalYear < 1990 Or GLCompany.FirstFiscalYear > Year(Now()) + 10 Then
        GLCompany.FirstFiscalYear = Year(Now()) - 5
    End If
    txtFirstFiscalYear = GLCompany.FirstFiscalYear
    
    Me.txtSuspAcct = CStr(GLCompany.SuspAcct)
    Me.txtNetProfitAcct = CStr(GLCompany.NetProfitAcct)
    Me.txtFirstPAcct = CStr(GLCompany.FirstPAcct)
    Me.txtSubDigits = CStr(GLCompany.SubDigits)
    Me.tdbLastClose = CStr(GLCompany.LastClose)
    Me.tdbLowBranch = CStr(GLCompany.LowBranch)
    Me.tdbHiBranch = CStr(GLCompany.HiBranch)
    Me.tdbLowConsolidated = CStr(GLCompany.LowConsolidated)
    Me.tdbHiConsolidated = CStr(GLCompany.HiConsolidated)
    Me.tdbRetEarnAcct = CStr(GLCompany.RetEarnAcct)

End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdOk_Click()
        
    On Error GoTo glErr
    If txtName = "" Then
        MsgBox "Name Not Filled In", vbExclamation, "GL Company Maintenance"
        txtName.SetFocus
        Exit Sub
    End If
    
    If Len(txtState) > 2 Then
        MsgBox "State Must Be a 2 Letter Abbreviation", vbExclamation, "GL Company Maintenance"
        txtState.SetFocus
        Exit Sub
    End If

    GLCompany.Name = txtName
    GLCompany.Address1 = txtAddress1
    GLCompany.Address2 = txtAddress2
    GLCompany.Address3 = txtAddress3
    GLCompany.City = txtCity
    GLCompany.State = txtState
    GLCompany.ZipCode = txtZipCode

'    glcompany.lastUpdate = Now
'    glcompany.lastClose = lastClose
    
    If IsNumeric(txtSuspAcct) Then
        GLCompany.SuspAcct = CLng(txtSuspAcct)
    Else
        GLCompany.SuspAcct = 0
    End If
    
    If IsNumeric(txtNetProfitAcct) Then
        GLCompany.NetProfitAcct = CLng(txtNetProfitAcct)
    Else
        GLCompany.NetProfitAcct = 0
    End If
    
    If IsNumeric(txtNumberPds) Then
        GLCompany.NumberPds = CByte(txtNumberPds)
    Else
        GLCompany.NumberPds = 0
    End If
    
    If IsNumeric(txtFirstPeriod) Then
        GLCompany.FirstPeriod = CByte(txtFirstPeriod)
    Else
        GLCompany.FirstPeriod = 0
    End If
    
    If IsNumeric(txtCurFiscalYear) Then
        GLCompany.CurFiscalYear = CInt(txtCurFiscalYear)
    Else
        GLCompany.CurFiscalYear = 0
    End If
    
    If IsNumeric(txtFirstFiscalYear) Then
        GLCompany.FirstFiscalYear = CInt(txtFirstFiscalYear)
    Else
        GLCompany.FirstFiscalYear = 0
    End If
    
    If IsNumeric(txtCurPeriod) Then
        GLCompany.CurPeriod = CByte(txtCurPeriod)
    Else
        GLCompany.CurPeriod = 0
    End If
    
    If IsNumeric(txtFirstPAcct) Then
        GLCompany.FirstPAcct = CLng(txtFirstPAcct)
    Else
        GLCompany.FirstPAcct = 0
    End If
    
    If IsNumeric(txtSubDigits) Then
       GLCompany.SubDigits = CByte(txtSubDigits)
    Else
       GLCompany.SubDigits = 0
    End If
    
    If IsNumeric(Me.tdbLastClose) Then
       GLCompany.LastClose = CLng(Me.tdbLastClose)
    Else
       GLCompany.LastClose = 0
    End If
    
    If IsNumeric(Me.tdbLowBranch) Then
       GLCompany.LowBranch = CLng(Me.tdbLowBranch)
    Else
       GLCompany.LowBranch = 0
    End If
    
    If IsNumeric(Me.tdbHiBranch) Then
       GLCompany.HiBranch = CLng(Me.tdbHiBranch)
    Else
       GLCompany.HiBranch = 0
    End If
    
    If IsNumeric(Me.tdbLowConsolidated) Then
       GLCompany.LowConsolidated = CLng(Me.tdbLowConsolidated)
    Else
       GLCompany.LowConsolidated = 0
    End If
    
    If IsNumeric(Me.tdbHiConsolidated) Then
       GLCompany.HiConsolidated = CLng(Me.tdbHiConsolidated)
    Else
       GLCompany.HiConsolidated = 0
    End If
    
    If IsNumeric(Me.tdbRetEarnAcct) Then
       GLCompany.RetEarnAcct = CLng(Me.tdbRetEarnAcct)
    Else
       GLCompany.RetEarnAcct = 0
    End If
    
    If ID = 0 Then
        GLCompany.LastBatch = 0
    End If
    
    GLCompany.Save (Equate.RecPut)
    
    GoBack
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub txtName_Change()
    Me.Caption = txtName & " Data Fields"
End Sub

Private Sub tdbSet()

Dim MinV As Long
Dim MaxV As Long
Dim Fmt As String

    MinV = 0
    MaxV = 999999999
    Fmt = "########0"

    tdbLowBranch.HighlightText = True
    tdbLowBranch.MinValue = MinV
    tdbLowBranch.MaxValue = MaxV
    tdbLowBranch.Format = Fmt
    tdbLowBranch.DisplayFormat = ""

    tdbHiBranch.HighlightText = True
    tdbHiBranch.MinValue = MinV
    tdbHiBranch.MaxValue = MaxV
    tdbHiBranch.Format = Fmt
    tdbHiBranch.DisplayFormat = ""

    tdbLowConsolidated.HighlightText = True
    tdbLowConsolidated.MinValue = MinV
    tdbLowConsolidated.MaxValue = MaxV
    tdbLowConsolidated.Format = Fmt
    tdbLowConsolidated.DisplayFormat = ""

    tdbHiConsolidated.HighlightText = True
    tdbHiConsolidated.MinValue = MinV
    tdbHiConsolidated.MaxValue = MaxV
    tdbHiConsolidated.Format = Fmt
    tdbHiConsolidated.DisplayFormat = ""
    
    tdbRetEarnAcct.HighlightText = True
    tdbRetEarnAcct.MinValue = MinV
    tdbRetEarnAcct.MaxValue = MaxV
    tdbRetEarnAcct.Format = Fmt
    tdbRetEarnAcct.DisplayFormat = ""
    
End Sub
