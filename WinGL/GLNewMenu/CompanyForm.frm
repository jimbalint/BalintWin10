VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
   Begin MSComDlg.CommonDialog msDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
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

Dim com As New rCompany
Dim lastClose As Date

Private Sub cmdBrowse_Click()
    On Error GoTo glErr:
    msDialog.FileName = txtFileName
    msDialog.Filter = "Client Files|*.mdb"
    msDialog.ShowOpen
    txtFileName = msDialog.FileName
    Exit Sub
glErr:
End Sub

Private Sub cmdCloseDate_Click()
    GetDate.Prompt = "Date of Last Closing"
    GetDate.Show vbModal
    lastClose = GetDate.Calendar.Value
    txtLastClose = "Last Closing Date " & CStr(lastClose)
End Sub

Private Sub CmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdFirstPAcct_Click()
    GetAccount.Account = txtFirstPAcct
    GetAccount.Init
    GetAccount.Prompt = "FIRST P ACCOUNT"
    GetAccount.Show vbModal
    txtFirstPAcct = GetAccount.Account
End Sub

Private Sub cmdNetProfitAcct_Click()
    GetAccount.Account = txtNetProfitAcct
    GetAccount.Init
    GetAccount.Prompt = "NET PROFIT ACCOUNT"
    GetAccount.Show vbModal
    txtNetProfitAcct = GetAccount.Account
End Sub

Private Sub cmdOK_Click()
    On Error GoTo glErr
    If txtName = "" Then
        MsgBox "Name Not Filled In"
        txtName.SetFocus
        Exit Sub
    End If
    If Len(txtState) > 2 Then
        MsgBox "State Must Be a 2 Letter Abbreviation"
        txtState.SetFocus
        Exit Sub
    End If
'    com.FileName = txtFileName
    
    com.name = txtName
    com.address1 = txtAddress1
    com.address2 = txtAddress2
    com.address3 = txtAddress3
    com.city = txtCity
    com.state = txtState
    com.zipcode = txtZipCode

'    com.lastUpdate = Now
'    com.lastClose = lastClose
    
    If IsNumeric(txtSuspAcct) Then
        com.SuspAcct = CLng(txtSuspAcct)
    Else
        com.SuspAcct = 0
    End If
    
    If IsNumeric(txtNetProfitAcct) Then
        com.NetProfitAcct = CLng(txtNetProfitAcct)
    Else
        com.NetProfitAcct = 0
    End If
    
    If IsNumeric(txtNumberPds) Then
        com.NumberPds = CByte(txtNumberPds)
    Else
        com.NumberPds = 0
    End If
    
    If IsNumeric(txtFirstPeriod) Then
        com.FirstPeriod = CByte(txtFirstPeriod)
    Else
        com.FirstPeriod = 0
    End If
    
    If IsNumeric(txtCurFiscalYear) Then
        com.curFiscalYear = CInt(txtCurFiscalYear)
    Else
        com.curFiscalYear = 0
    End If
    
    If IsNumeric(txtFirstFiscalYear) Then
        com.FirstFiscalYear = CInt(txtFirstFiscalYear)
    Else
        com.FirstFiscalYear = 0
    End If
    
    If IsNumeric(txtCurPeriod) Then
        com.curPeriod = CByte(txtCurPeriod)
    Else
        com.curPeriod = 0
    End If
    
    If IsNumeric(txtFirstPAcct) Then
        com.FirstPAcct = CLng(txtFirstPAcct)
    Else
        com.FirstPAcct = 0
    End If
    
    If IsNumeric(txtSubDigits) Then
       com.SubDigits = CByte(txtSubDigits)
    Else
       com.SubDigits = 0
    End If
    
    If IsNumeric(Me.tdbLastClose) Then
       com.lastClose = CLng(Me.tdbLastClose)
    Else
       com.lastClose = 0
    End If
    
    If IsNumeric(Me.tdbLowBranch) Then
       com.LowBranch = CLng(Me.tdbLowBranch)
    Else
       com.LowBranch = 0
    End If
    
    If IsNumeric(Me.tdbHiBranch) Then
       com.HiBranch = CLng(Me.tdbHiBranch)
    Else
       com.HiBranch = 0
    End If
    
    If IsNumeric(Me.tdbLowConsolidated) Then
       com.LowConsolidated = CLng(Me.tdbLowConsolidated)
    Else
       com.LowConsolidated = 0
    End If
    
    If IsNumeric(Me.tdbHiConsolidated) Then
       com.HiConsolidated = CLng(Me.tdbHiConsolidated)
    Else
       com.HiConsolidated = 0
    End If
    
    If IsNumeric(Me.tdbRetEarnAcct) Then
       com.RetEarnAcct = CLng(Me.tdbRetEarnAcct)
    Else
       com.RetEarnAcct = 0
    End If
    
    If ID = 0 Then
        com.LastBatch = 0
    End If
    
    ID = com.PutRecord(ID)
    userOK = True
    Me.Hide
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub Init()
    userOK = False
    txtName = ""
    txtAddress1 = ""
    txtAddress2 = ""
    txtAddress3 = ""
    txtCity = ""
    txtState = ""
    txtZipCode = ""
    txtLastClose = ""
    txtLastUpdate = ""
    txtSuspAcct = ""
    txtNetProfitAcct = ""
    txtNumberPds = ""
    txtCurFiscalYear = ""
    txtFirstFiscalYear = ""
    txtCurPeriod = ""
    txtFirstPeriod = ""
    txtFirstPAcct = ""
    
    If ID = 0 Then
    Else
        If com.GetRecord(ID) = True Then
            txtName = com.name
            lblDBName = com.FileName
            txtAddress1 = com.address1
            txtAddress2 = com.address2
            txtAddress3 = com.address3
            txtCity = com.city
            txtState = com.state
            txtZipCode = com.zipcode
'            If IsDate(com.lastUpdate) Then
'                txtLastUpdate = "Last Updated " & CStr(com.lastUpdate)
'            End If
'            lastClose = com.lastClose
'            If IsDate(lastClose) Then
'                txtLastClose = "Last Closing Date " & CStr(lastClose)
'            End If
            txtSuspAcct = CStr(com.SuspAcct)
            txtNetProfitAcct = CStr(com.NetProfitAcct)
            txtNumberPds = CStr(com.NumberPds)
            txtFirstPeriod = CStr(com.FirstPeriod)
            txtCurFiscalYear = CStr(com.curFiscalYear)
            txtFirstFiscalYear = CStr(com.FirstFiscalYear)
            txtFirstPAcct = CStr(com.FirstPAcct)
            txtCurPeriod = CStr(com.curPeriod)
        End If
    End If
    Me.Caption = txtName & " Data Fields"
End Sub


Private Sub cmdSuspAcct_Click()
    GetAccount.Account = txtSuspAcct
    GetAccount.Init
    GetAccount.Prompt = "SUSPEND ACCOUNT"
    GetAccount.Show vbModal
    txtSuspAcct = GetAccount.Account
End Sub

Private Sub Form_Load()

Dim tdbNum As TDBNumber
       
    ' settings for tdb controls
    tdbSet

    If ID = 0 Then
       userOK = False
       txtName = ""
       txtAddress1 = ""
       txtAddress2 = ""
       txtAddress3 = ""
       txtCity = ""
       txtState = ""
       txtZipCode = ""
       txtLastClose = ""
       txtLastUpdate = ""
       txtSuspAcct = ""
       txtNetProfitAcct = ""
       txtNumberPds = ""
       txtCurFiscalYear = ""
       txtFirstFiscalYear = ""
       txtCurPeriod = ""
       txtFirstPeriod = ""
       txtFirstPAcct = ""
       txtSubDigits = 0
    
       tdbRetEarnAcct = 0
       tdbLowBranch = 0
       tdbHiBranch = 0
       tdbLowConsolidated = 0
       tdbHiConsolidated = 0
    
    Else
       
       Set com = New rCompany
       com.GetRecord (ID)
       
       ' populate fields
       lblDBName = "File Name: " & com.FileName
       txtName = com.name & ""
       txtAddress1 = com.address1 & ""
       txtAddress2 = com.address2 & ""
       txtAddress3 = com.address3 & ""
       txtCity = com.city & ""
       txtState = com.state & ""
       txtZipCode = com.zipcode & ""
       
       If Not IsNull(com.NumberPds) Then
          Me.txtNumberPds = CStr(com.NumberPds)
       Else
          Me.txtNumberPds = "12"
       End If
       
       If com.NumberPds = 0 Then
          Me.txtNumberPds = "12"
       End If
       
       If Not IsNull(com.FirstPeriod) Then
          If com.FirstPeriod = 0 Then com.FirstPeriod = 1
          Me.txtFirstPeriod = CStr(com.FirstPeriod)
       Else
          Me.txtFirstPeriod = "1"
       End If
       
       If Not IsNull(com.curPeriod) Then
          Me.txtCurPeriod = CStr(com.curPeriod)
       Else
          Me.txtCurPeriod = "0"
       End If
       
       If Not IsNull(com.curFiscalYear) Then
          Me.txtCurFiscalYear = CStr(com.curFiscalYear)
       Else
          Me.txtCurFiscalYear = "0"
       End If
       
       If Not IsNull(com.FirstFiscalYear) Then
          If com.FirstFiscalYear < 1990 Or com.FirstFiscalYear > 2020 Then
             com.FirstFiscalYear = 2000
          End If
          Me.txtFirstFiscalYear = com.FirstFiscalYear
       Else
          Me.txtFirstFiscalYear = "2000"
       End If
       
       If Not IsNull(com.SuspAcct) Then
          Me.txtSuspAcct = CStr(com.SuspAcct)
       Else
          Me.txtSuspAcct = "0"
       End If
       
       If Not IsNull(com.NetProfitAcct) Then
          Me.txtNetProfitAcct = CStr(com.NetProfitAcct)
       Else
          Me.txtNetProfitAcct = "0"
       End If
       
       If Not IsNull(com.FirstPAcct) Then
          Me.txtFirstPAcct = CStr(com.FirstPAcct)
       Else
          Me.txtFirstPAcct = "0"
       End If
       
       If Not IsNull(com.SubDigits) Then
          Me.txtSubDigits = CStr(com.SubDigits)
       Else
          Me.txtSubDigits = "0"
       End If
    
       If Not IsNull(com.lastClose) Then
          Me.tdbLastClose = CStr(com.lastClose)
       Else
          Me.tdbLastClose = "0"
       End If
    
       If Not IsNull(com.LowBranch) Then
          Me.tdbLowBranch = CStr(com.LowBranch)
       Else
          Me.tdbLowBranch = "0"
       End If
    
       If Not IsNull(com.HiBranch) Then
          Me.tdbHiBranch = CStr(com.HiBranch)
       Else
          Me.tdbHiBranch = "0"
       End If
    
       If Not IsNull(com.LowConsolidated) Then
          Me.tdbLowConsolidated = CStr(com.LowConsolidated)
       Else
          Me.tdbLowConsolidated = "0"
       End If
    
       If Not IsNull(com.HiConsolidated) Then
          Me.tdbHiConsolidated = CStr(com.HiConsolidated)
       Else
          Me.tdbHiConsolidated = "0"
       End If
    
       If Not IsNull(com.RetEarnAcct) Then
          Me.tdbRetEarnAcct = CStr(com.RetEarnAcct)
       Else
          Me.tdbRetEarnAcct = "0"
       End If
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload GetDate
    Unload GetAccount
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
