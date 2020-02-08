VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmJCEdit 
   Caption         =   "Customer / Job Maintenance"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13185
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
   ScaleHeight     =   9885
   ScaleWidth      =   13185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraJobInfo 
      Caption         =   "  J O B   I N F O R M A T I O N  "
      Height          =   2295
      Left            =   360
      TabIndex        =   31
      Top             =   6600
      Width           =   12375
      Begin VB.ComboBox cmbCityTax 
         Height          =   360
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   6135
      End
      Begin TDBText6Ctl.TDBText tdbtxtJobType 
         Height          =   375
         Left            =   6480
         TabIndex        =   25
         Top             =   1800
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   661
         Caption         =   "frmJCEdit.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmJCEdit.frx":0066
         Key             =   "frmJCEdit.frx":0084
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
      Begin TDBDate6Ctl.TDBDate tdbdateJobEnd 
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   1800
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         Calendar        =   "frmJCEdit.frx":00C8
         Caption         =   "frmJCEdit.frx":01C8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmJCEdit.frx":022E
         Keys            =   "frmJCEdit.frx":024C
         Spin            =   "frmJCEdit.frx":02AA
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
         Text            =   "11/27/2009"
         ValidateMode    =   0
         ValueVT         =   3342343
         Value           =   40144
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tdbdateJobStart 
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         Calendar        =   "frmJCEdit.frx":02D2
         Caption         =   "frmJCEdit.frx":03D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmJCEdit.frx":043C
         Keys            =   "frmJCEdit.frx":045A
         Spin            =   "frmJCEdit.frx":04B8
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
         Text            =   "11/27/2009"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   40144
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText tdbtxtJobDescription 
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   1200
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   661
         Caption         =   "frmJCEdit.frx":04E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmJCEdit.frx":054C
         Key             =   "frmJCEdit.frx":056A
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
      Begin TDBText6Ctl.TDBText tdbtxtJobStatus 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "frmJCEdit.frx":05AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmJCEdit.frx":0610
         Key             =   "frmJCEdit.frx":062E
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
         Text            =   "TDBText8"
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
      Begin VB.Label Label3 
         Caption         =   "City Tax:"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
   End
   Begin TDBText6Ctl.TDBText tdbtxtBillZip 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   5880
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0672
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":06D8
      Key             =   "frmJCEdit.frx":06F6
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
   Begin TDBText6Ctl.TDBText tdbtxtBillState 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":073A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":079A
      Key             =   "frmJCEdit.frx":07B8
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
   Begin TDBText6Ctl.TDBText tdbtxtBillCity 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":07FC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":085A
      Key             =   "frmJCEdit.frx":0878
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
   Begin TDBText6Ctl.TDBText tdbtxtBillAddr1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":08BC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0924
      Key             =   "frmJCEdit.frx":0942
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
   Begin TDBText6Ctl.TDBText tdbtxtLastName 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   2400
      Width           =   4335
      _Version        =   65536
      _ExtentX        =   7646
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0986
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":09EE
      Key             =   "frmJCEdit.frx":0A0C
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
   Begin TDBText6Ctl.TDBText tdbtxtMidInit 
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0A50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0AB6
      Key             =   "frmJCEdit.frx":0AD4
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
   Begin TDBText6Ctl.TDBText tdbtxtFirstName 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   5175
      _Version        =   65536
      _ExtentX        =   9128
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0B18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0B82
      Key             =   "frmJCEdit.frx":0BA0
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
   Begin TDBText6Ctl.TDBText tdbtxtName 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0BE4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0C42
      Key             =   "frmJCEdit.frx":0C60
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   7845
      TabIndex        =   27
      Top             =   9240
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   3885
      TabIndex        =   26
      Top             =   9240
      Width           =   1455
   End
   Begin TDBText6Ctl.TDBText tdbtxtFullName 
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0CA4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0D0C
      Key             =   "frmJCEdit.frx":0D2A
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
   Begin TDBText6Ctl.TDBText tdbtxtCompanyName 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0D6E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0DDC
      Key             =   "frmJCEdit.frx":0DFA
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
   Begin TDBText6Ctl.TDBText tdbtxtBillAddr2 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0E3E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0EA6
      Key             =   "frmJCEdit.frx":0EC4
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
   Begin TDBText6Ctl.TDBText tdbtxtBillAddr3 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0F08
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":0F70
      Key             =   "frmJCEdit.frx":0F8E
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
   Begin TDBText6Ctl.TDBText tdbtxtBillAddr4 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":0FD2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":103A
      Key             =   "frmJCEdit.frx":1058
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
   Begin TDBText6Ctl.TDBText tdbtxtShipAddr1 
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   3480
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":109C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":1104
      Key             =   "frmJCEdit.frx":1122
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
   Begin TDBText6Ctl.TDBText tdbtxtShipAddr2 
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   3960
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":1166
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":11CE
      Key             =   "frmJCEdit.frx":11EC
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
   Begin TDBText6Ctl.TDBText tdbtxtShipAddr3 
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   4440
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":1230
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":1298
      Key             =   "frmJCEdit.frx":12B6
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
   Begin TDBText6Ctl.TDBText tdbtxtShipAddr4 
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   4920
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":12FA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":1362
      Key             =   "frmJCEdit.frx":1380
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
   Begin TDBText6Ctl.TDBText tdbtxtShipCity 
      Height          =   375
      Left            =   6840
      TabIndex        =   17
      Top             =   5400
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":13C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":1422
      Key             =   "frmJCEdit.frx":1440
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
   Begin TDBText6Ctl.TDBText tdbtxtShipState 
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Top             =   5880
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   3625
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":1484
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":14E4
      Key             =   "frmJCEdit.frx":1502
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
   Begin TDBText6Ctl.TDBText tdbtxtShipZip 
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   5880
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Caption         =   "frmJCEdit.frx":1546
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCEdit.frx":15AC
      Key             =   "frmJCEdit.frx":15CA
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
   Begin VB.Label Label2 
      Caption         =   "SHIP TO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   30
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "BILL TO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   12255
   End
End
Attribute VB_Name = "frmJCEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EditJob As Boolean
Public Action As Byte
Dim X As String
Public ParentID As Long
Dim i As Long

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name

    ' what action??
    If EditJob = True Then
        X = "Job"
    Else
        X = "Customer"
        Me.fraJobInfo.Visible = False
    End If
    
    If Action = PREquate.ActionAdd Then
        Me.Caption = "Add a " & X
        JCJob.Clear
        JCCustomer.Clear
    Else
        Me.Caption = "Edit a " & X
    End If

    ' text fields
    tdbTextSet Me.tdbtxtName
    tdbTextSet Me.tdbtxtFullName
    tdbTextSet Me.tdbtxtCompanyName
    tdbTextSet Me.tdbtxtFirstName
    tdbTextSet Me.tdbtxtMidInit
    tdbTextSet Me.tdbtxtLastName
    
    tdbTextSet Me.tdbtxtBillAddr1
    tdbTextSet Me.tdbtxtBillAddr2
    tdbTextSet Me.tdbtxtBillAddr3
    tdbTextSet Me.tdbtxtBillAddr4
    tdbTextSet Me.tdbtxtBillCity
    tdbTextSet Me.tdbtxtBillState
    tdbTextSet Me.tdbtxtBillZip

    tdbTextSet Me.tdbtxtShipAddr1
    tdbTextSet Me.tdbtxtShipAddr2
    tdbTextSet Me.tdbtxtShipAddr3
    tdbTextSet Me.tdbtxtShipAddr4
    tdbTextSet Me.tdbtxtShipCity
    tdbTextSet Me.tdbtxtShipState
    tdbTextSet Me.tdbtxtShipZip

    tdbTextSet Me.tdbtxtJobDescription
    tdbTextSet Me.tdbtxtJobType
    tdbTextSet Me.tdbtxtJobStatus

    ' get the data
    If Action = PREquate.ActionEdit Then
        If EditJob = True Then
            
            If JCJob.GetByID(TaskID) = False Then
                MsgBox "Job not found: " & TaskID, vbExclamation
                Unload Me
            End If
            
            Me.tdbtxtName = JCJob.Name
            Me.tdbtxtFullName = JCJob.FullName
            Me.tdbtxtCompanyName = JCJob.CompanyName
            Me.tdbtxtFirstName = JCJob.FirstName
            
            Me.tdbtxtMidInit = JCJob.MidInit
            Me.tdbtxtLastName = JCJob.LastName
            
            Me.tdbtxtBillAddr1 = JCJob.BillAddr1
            Me.tdbtxtBillAddr2 = JCJob.BillAddr2
            Me.tdbtxtBillAddr3 = JCJob.BillAddr3
            Me.tdbtxtBillAddr4 = JCJob.BillAddr4
            Me.tdbtxtBillCity = JCJob.BillCity
            Me.tdbtxtBillState = JCJob.BillState
            Me.tdbtxtBillZip = JCJob.BillZip
            
            Me.tdbtxtShipAddr1 = JCJob.ShipAddr1
            Me.tdbtxtShipAddr2 = JCJob.ShipAddr2
            Me.tdbtxtShipAddr3 = JCJob.ShipAddr3
            Me.tdbtxtShipAddr4 = JCJob.ShipAddr4
            Me.tdbtxtShipCity = JCJob.ShipCity
            Me.tdbtxtShipState = JCJob.ShipState
            Me.tdbtxtShipZip = JCJob.ShipZip
    
            Me.tdbtxtJobDescription = JCJob.Description
            Me.tdbtxtJobStatus = JCJob.Status
            Me.tdbtxtJobType = JCJob.TypeName
            
            tdbDateSet Me.tdbdateJobStart, nNull(JCJob.StartDate)
            tdbDateSet Me.tdbdateJobEnd, nNull(JCJob.EndDate)
    
        Else
        
            If JCCustomer.GetByID(TaskID) = False Then
                MsgBox "Customer not found: " & TaskID, vbExclamation
                Unload Me
            End If
        
            Me.tdbtxtName = JCCustomer.Name
            Me.tdbtxtFullName = JCCustomer.FullName
            Me.tdbtxtCompanyName = JCCustomer.CompanyName
            Me.tdbtxtFirstName = JCCustomer.FirstName
            Me.tdbtxtMidInit = JCCustomer.MidInit
            Me.tdbtxtLastName = JCCustomer.LastName
            
            Me.tdbtxtBillAddr1 = JCCustomer.BillAddr1
            Me.tdbtxtBillAddr2 = JCCustomer.BillAddr2
            Me.tdbtxtBillAddr3 = JCCustomer.BillAddr3
            Me.tdbtxtBillAddr4 = JCCustomer.BillAddr4
            Me.tdbtxtBillCity = JCCustomer.BillCity
            Me.tdbtxtBillState = JCCustomer.BillState
            Me.tdbtxtBillZip = JCCustomer.BillZip
            
            Me.tdbtxtShipAddr1 = JCCustomer.ShipAddr1
            Me.tdbtxtShipAddr2 = JCCustomer.ShipAddr2
            Me.tdbtxtShipAddr3 = JCCustomer.ShipAddr3
            Me.tdbtxtShipAddr4 = JCCustomer.ShipAddr4
            Me.tdbtxtShipCity = JCCustomer.ShipCity
            Me.tdbtxtShipState = JCCustomer.ShipState
            Me.tdbtxtShipZip = JCCustomer.ShipZip
    
        End If
            
    Else
    
        TaskID = 0
    
        ' add new Job - get dflt info from Customer
        If EditJob = True Then
                    
            If JCCustomer.GetByID(ParentID) = False Then
                MsgBox "Customer not found: " & TaskID, vbExclamation
                Unload Me
            End If
            
            Me.tdbtxtName = JCCustomer.Name
            Me.tdbtxtFullName = JCCustomer.FullName
            Me.tdbtxtCompanyName = JCCustomer.CompanyName
            Me.tdbtxtFirstName = JCCustomer.FirstName
            Me.tdbtxtMidInit = JCCustomer.MidInit
            Me.tdbtxtLastName = JCCustomer.LastName
            
            Me.tdbtxtBillAddr1 = JCCustomer.BillAddr1
            Me.tdbtxtBillAddr2 = JCCustomer.BillAddr2
            Me.tdbtxtBillAddr3 = JCCustomer.BillAddr3
            Me.tdbtxtBillAddr4 = JCCustomer.BillAddr4
            Me.tdbtxtBillCity = JCCustomer.BillCity
            Me.tdbtxtBillState = JCCustomer.BillState
            Me.tdbtxtBillZip = JCCustomer.BillZip
            
            Me.tdbtxtShipAddr1 = JCCustomer.ShipAddr1
            Me.tdbtxtShipAddr2 = JCCustomer.ShipAddr2
            Me.tdbtxtShipAddr3 = JCCustomer.ShipAddr3
            Me.tdbtxtShipAddr4 = JCCustomer.ShipAddr4
            Me.tdbtxtShipCity = JCCustomer.ShipCity
            Me.tdbtxtShipState = JCCustomer.ShipState
            Me.tdbtxtShipZip = JCCustomer.ShipZip
                    
        End If
    
    End If
    
    ' populate and point the City Tax Grid
    If EditJob = True Then
        With Me.cmbCityTax
            .AddItem "NONE"
            .ItemData(.NewIndex) = 0
            SQLString = "SELECT * FROM PRCity ORDER BY CityName"
            If PRCity.GetBySQL(SQLString) Then
                Do
                    .AddItem PRCity.CityName & " " & Format(PRCity.CityRate, "##0.00") & "%"
                    .ItemData(.NewIndex) = PRCity.CityID
                    If PRCity.GetNext = False Then Exit Do
                Loop
            End If
            .ListIndex = 0
            For i = 0 To .ListCount - 1
                If .ItemData(i) = JCJob.CityID Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End With
    End If

    
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    TaskID = 0
    Unload Me
End Sub
Private Sub cmdSave_Click()
         
    If Action = PREquate.ActionAdd Then
        If EditJob = True Then
            JCJob.OpenRS
            JCJob.Clear
            JCJob.Name = Me.tdbtxtName
            JCJob.ParentID = ParentID
            JCJob.Save (Equate.RecAdd)
            TaskID = JCJob.JobID
        Else
            JCCustomer.OpenRS
            JCCustomer.Clear
            JCCustomer.Name = Me.tdbtxtName
            JCCustomer.Save (Equate.RecAdd)
            TaskID = JCCustomer.CustomerID
        End If
    End If
    
    If EditJob = True Then
        JCJob.Name = Me.tdbtxtName
        JCJob.FullName = Me.tdbtxtFullName
        JCJob.CompanyName = Me.tdbtxtCompanyName
        JCJob.FirstName = Me.tdbtxtFirstName
        JCJob.MidInit = Me.tdbtxtMidInit
        JCJob.LastName = Me.tdbtxtLastName
                                                                   
        JCJob.BillAddr1 = Me.tdbtxtBillAddr1
        JCJob.BillAddr2 = Me.tdbtxtBillAddr2
        JCJob.BillAddr3 = Me.tdbtxtBillAddr3
        JCJob.BillAddr4 = Me.tdbtxtBillAddr4
        JCJob.BillCity = Me.tdbtxtBillCity
        JCJob.BillState = Me.tdbtxtBillState
        JCJob.BillZip = Me.tdbtxtBillZip
                                                                   
        JCJob.ShipAddr1 = Me.tdbtxtShipAddr1
        JCJob.ShipAddr2 = Me.tdbtxtShipAddr2
        JCJob.ShipAddr3 = Me.tdbtxtShipAddr3
        JCJob.ShipAddr4 = Me.tdbtxtShipAddr4
        JCJob.ShipCity = Me.tdbtxtShipCity
        JCJob.ShipState = Me.tdbtxtShipState
        JCJob.ShipZip = Me.tdbtxtShipZip
                                                                   
        JCJob.Description = Me.tdbtxtJobDescription
        JCJob.Status = Me.tdbtxtJobStatus
        JCJob.TypeName = Me.tdbtxtJobType
        JCJob.StartDate = nNull(Me.tdbdateJobStart)
        JCJob.EndDate = nNull(Me.tdbdateJobEnd)
        JCJob.CityID = Me.cmbCityTax.ItemData(Me.cmbCityTax.ListIndex)
        JCJob.Save (Equate.RecPut)
    Else
        JCCustomer.Name = Me.tdbtxtName
        JCCustomer.FullName = Me.tdbtxtFullName
        JCCustomer.CompanyName = Me.tdbtxtCompanyName
        JCCustomer.FirstName = Me.tdbtxtFirstName
        JCCustomer.MidInit = Me.tdbtxtMidInit
        JCCustomer.LastName = Me.tdbtxtLastName
                                                                   
        JCCustomer.BillAddr1 = Me.tdbtxtBillAddr1
        JCCustomer.BillAddr2 = Me.tdbtxtBillAddr2
        JCCustomer.BillAddr3 = Me.tdbtxtBillAddr3
        JCCustomer.BillAddr4 = Me.tdbtxtBillAddr4
        JCCustomer.BillCity = Me.tdbtxtBillCity
        JCCustomer.BillState = Me.tdbtxtBillState
        JCCustomer.BillZip = Me.tdbtxtBillZip
                                                                   
        JCCustomer.ShipAddr1 = Me.tdbtxtShipAddr1
        JCCustomer.ShipAddr2 = Me.tdbtxtShipAddr2
        JCCustomer.ShipAddr3 = Me.tdbtxtShipAddr3
        JCCustomer.ShipAddr4 = Me.tdbtxtShipAddr4
        JCCustomer.ShipCity = Me.tdbtxtShipCity
        JCCustomer.ShipState = Me.tdbtxtShipState
        JCCustomer.ShipZip = Me.tdbtxtShipZip
                                                                   
        JCCustomer.Save (Equate.RecPut)
    
    End If

    Unload Me

End Sub


