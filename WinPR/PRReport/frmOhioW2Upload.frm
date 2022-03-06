VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOHW2 
   Caption         =   "Ohio W2 Upload"
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14505
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
   ScaleHeight     =   10290
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlTextOutput 
      Left            =   14880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   11640
      TabIndex        =   22
      Top             =   120
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Submitter Information"
      TabPicture(0)   =   "frmOhioW2Upload.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "txtUserID"
      Tab(0).Control(3)=   "txtContactPhn"
      Tab(0).Control(4)=   "txtContactName"
      Tab(0).Control(5)=   "txtZipCodeExt"
      Tab(0).Control(6)=   "txtZipCode"
      Tab(0).Control(7)=   "txtState"
      Tab(0).Control(8)=   "txtCity"
      Tab(0).Control(9)=   "txtDeliveryAddress"
      Tab(0).Control(10)=   "txtLocationAddress"
      Tab(0).Control(11)=   "txtCompanyName"
      Tab(0).Control(12)=   "txtEIN"
      Tab(0).Control(13)=   "txtContactPhnExt"
      Tab(0).Control(14)=   "txtContactEmail"
      Tab(0).Control(15)=   "txtContactFax"
      Tab(0).Control(16)=   "cmbPreparerCode"
      Tab(0).Control(17)=   "cmdSaveSubm"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Submit OH W2 File"
      TabPicture(1)   =   "frmOhioW2Upload.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "txtResubID"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkResubIndicator"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtOutputFile"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdFileName"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdCreateFile"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtTaxYear"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin TDBText6Ctl.TDBText txtTaxYear 
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   600
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0038
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":009C
         Key             =   "frmOhioW2Upload.frx":00BA
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
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
      Begin VB.CommandButton cmdCreateFile 
         BackColor       =   &H00FFFF00&
         Caption         =   "Create OH W2 Upload"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   11640
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton cmdFileName 
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin TDBText6Ctl.TDBText txtOutputFile 
         Height          =   375
         Left            =   360
         TabIndex        =   23
         Top             =   1560
         Width           =   9375
         _Version        =   65536
         _ExtentX        =   16536
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":00FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0172
         Key             =   "frmOhioW2Upload.frx":0190
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
         Format          =   "Aa9@"
         FormatMode      =   0
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
      Begin VB.CheckBox chkResubIndicator 
         Caption         =   "Resub Indicator"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton cmdSaveSubm 
         Caption         =   "&SAVE"
         Height          =   615
         Left            =   -63000
         TabIndex        =   18
         Top             =   8040
         Width           =   1455
      End
      Begin VB.ComboBox cmbPreparerCode 
         Height          =   360
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   8280
         Width           =   4575
      End
      Begin TDBText6Ctl.TDBText txtContactFax 
         Height          =   375
         Left            =   -74640
         TabIndex        =   15
         Top             =   7800
         Width           =   8055
         _Version        =   65536
         _ExtentX        =   14208
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":01D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":023E
         Key             =   "frmOhioW2Upload.frx":025C
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   10
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
      Begin TDBText6Ctl.TDBText txtContactEmail 
         Height          =   375
         Left            =   -74640
         TabIndex        =   14
         Top             =   7320
         Width           =   8775
         _Version        =   65536
         _ExtentX        =   15478
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":02A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0312
         Key             =   "frmOhioW2Upload.frx":0330
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
         AllowSpace      =   0
         Format          =   "Aa9@"
         FormatMode      =   0
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
      Begin TDBText6Ctl.TDBText txtContactPhnExt 
         Height          =   375
         Left            =   -74640
         TabIndex        =   13
         Top             =   6840
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0374
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":03F2
         Key             =   "frmOhioW2Upload.frx":0410
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
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
      Begin TDBText6Ctl.TDBText txtEIN 
         Height          =   375
         Left            =   -74640
         TabIndex        =   2
         Top             =   1320
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0454
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":04B6
         Key             =   "frmOhioW2Upload.frx":04D4
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   9
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
      Begin TDBText6Ctl.TDBText txtCompanyName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   3
         Top             =   2520
         Width           =   13695
         _Version        =   65536
         _ExtentX        =   24156
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0518
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0588
         Key             =   "frmOhioW2Upload.frx":05A6
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
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   57
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
      Begin TDBText6Ctl.TDBText txtLocationAddress 
         Height          =   375
         Left            =   -74640
         TabIndex        =   4
         Top             =   3000
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":05EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":065E
         Key             =   "frmOhioW2Upload.frx":067C
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
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   22
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
      Begin TDBText6Ctl.TDBText txtDeliveryAddress 
         Height          =   375
         Left            =   -74640
         TabIndex        =   5
         Top             =   3480
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":06C0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0738
         Key             =   "frmOhioW2Upload.frx":0756
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
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   22
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
         Left            =   -74640
         TabIndex        =   6
         Top             =   3960
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":079A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":07FA
         Key             =   "frmOhioW2Upload.frx":0818
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
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   22
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
         Left            =   -74640
         TabIndex        =   7
         Top             =   4440
         Width           =   3615
         _Version        =   65536
         _ExtentX        =   6376
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":085C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":08BE
         Key             =   "frmOhioW2Upload.frx":08DC
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
         AllowSpace      =   0
         Format          =   "A"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   2
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
      Begin TDBText6Ctl.TDBText txtZipCode 
         Height          =   375
         Left            =   -74640
         TabIndex        =   8
         Top             =   4920
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0920
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0988
         Key             =   "frmOhioW2Upload.frx":09A6
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   5
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
      Begin TDBText6Ctl.TDBText txtZipCodeExt 
         Height          =   375
         Left            =   -74640
         TabIndex        =   9
         Top             =   5400
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":09EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0A62
         Key             =   "frmOhioW2Upload.frx":0A80
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   4
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
      Begin TDBText6Ctl.TDBText txtContactName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   10
         Top             =   5880
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0AC4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0B34
         Key             =   "frmOhioW2Upload.frx":0B52
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
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   27
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
      Begin TDBText6Ctl.TDBText txtContactPhn 
         Height          =   375
         Left            =   -74640
         TabIndex        =   11
         Top             =   6360
         Width           =   6495
         _Version        =   65536
         _ExtentX        =   11456
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0B96
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0C08
         Key             =   "frmOhioW2Upload.frx":0C26
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
         AllowSpace      =   0
         Format          =   "9"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   15
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
      Begin TDBText6Ctl.TDBText txtUserID 
         Height          =   375
         Left            =   -74640
         TabIndex        =   20
         Top             =   1920
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0C6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0CD8
         Key             =   "frmOhioW2Upload.frx":0CF6
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
         AllowSpace      =   0
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   8
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
      Begin TDBText6Ctl.TDBText txtResubID 
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   1080
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Caption         =   "frmOhioW2Upload.frx":0D3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOhioW2Upload.frx":0DA8
         Key             =   "frmOhioW2Upload.frx":0DC6
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
         AllowSpace      =   0
         Format          =   "Aa9@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   6
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
      Begin VB.Label Label3 
         Caption         =   "* Preparer Code"
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   8400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "* = Required Field"
         Height          =   255
         Left            =   -65400
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Ohio W2 Upload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmOHW2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim TextFileName As String
Dim TextChannel2 As Integer
Dim sOut As String
Public FileExt As String
Dim ii As Integer
Dim xx As String
Dim ReportTitle As String
Dim OutOfStateEmp As String
Dim OHStateID As Integer

Dim PRW2 As New cPRW2
Dim PRW2State As New cPRW2State
Dim PRW2City As New cPRW2City
Dim W2TL As New cOHW2Totals

Dim frmp As New frmProgress

Private Sub Form_Load()

    strSQL = "select *" & _
            " from PRState " & _
            " where StateAbbrev = 'OH'"
    If Not PRState.GetBySQL(strSQL) Then
        MsgBox "State record not found of OH!!"
        End
    End If
    OHStateID = PRState.StateID

    W2TL.RWTotalCount = 0
    OutOfStateEmp = ""

    ' create PRGlobal Records
    InitPRGlobal PREquate.GlobalTypeOHW2Company
    InitPRGlobal PREquate.GlobalTypeOHW2Contact
    InitPRGlobal PREquate.GlobalTypeOHW2Submit
    
    With Me
        
        .KeyPreview = True
        
        .cmbPreparerCode.AddItem ("A Accounting Firm")
        .cmbPreparerCode.AddItem ("L Self-Prepared")
        .cmbPreparerCode.AddItem ("S Service Bureau")
        .cmbPreparerCode.AddItem ("P Parent Company")
        .cmbPreparerCode.AddItem ("O Other")
        
        ' Company Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Company
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        .txtEIN = PRGlobal.Var1
        .txtUserID = PRGlobal.Var2
        .txtCompanyName = PRGlobal.Var3
        .txtLocationAddress = PRGlobal.Var4
        .txtDeliveryAddress = PRGlobal.Var5
        .txtCity = PRGlobal.Var6
        .txtZipCode = PRGlobal.Var7
        
        .txtZipCodeExt = PRGlobal.Var8
        If PRGlobal.Var9 = "A" Then .cmbPreparerCode.ListIndex = 0
        If PRGlobal.Var9 = "L" Then .cmbPreparerCode.ListIndex = 1
        If PRGlobal.Var9 = "S" Then .cmbPreparerCode.ListIndex = 2
        If PRGlobal.Var9 = "P" Then .cmbPreparerCode.ListIndex = 3
        If PRGlobal.Var9 = "O" Then .cmbPreparerCode.ListIndex = 4
        
        ' Contact Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Contact
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        .txtContactName = PRGlobal.Var1
        .txtContactPhn = PRGlobal.Var2
        .txtContactPhnExt = PRGlobal.Var3
        .txtContactEmail = PRGlobal.Var4
        .txtContactFax = PRGlobal.Var5
        
        ' Submit Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Submit
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        .chkResubIndicator = PRGlobal.Byte1
        .txtResubID = PRGlobal.Var1
        .txtOutputFile = PRGlobal.Var2
        .txtTaxYear = PRGlobal.Var3
    
        If .txtState = "" Then .txtState = "OH"
    
    End With

End Sub
Private Sub cmdCreateFile_Click()
    
    If Not PreCheck Then Exit Sub
    SaveSubmitterInfo
    If Not InitOutputFile Then Exit Sub
    
    If CInt(Me.txtTaxYear.text) <> Year(Date) - 1 Then
        Dim bResult As Integer
        bResult = MsgBox("Are you SURE you want to submit for the Tax Year of: " & Me.txtTaxYear.text, vbExclamation + vbYesNo, "OH W2 Upload")
        If bResult = vbNo Then Exit Sub
    End If
   
    frmp.lblMsg1 = "Now Running OH W2 Upload"
    frmp.Show
    
    InitReport
    WriteRA
    
    strSQL = "select * from PRCompany where OHeW2 = True"
    If Not PRCompany.GetBySQL(strSQL) Then End
    Do

        frmp.lblMsg2 = PRCompany.Name
        frmp.Refresh

        W2TL.Clear
        WriteCompany (PRCompany.CompanyID)
        strSQL = "select * from PRW2 where TaxYear = " & Me.txtTaxYear & _
                " and Void = 0" & _
                " and Skip = 0"
        If PRW2.GetBySQL(strSQL) Then
            Do
                WriteRW
                WriteRO
                WriteRS
                If Not PRW2.GetNext Then Exit Do
            Loop
        End If
        WriteRT
        If W2TL.ROCount > 0 Then WriteRU
        CompanyReport
        If Not PRCompany.GetNext Then Exit Do
        FormFeed
    Loop
    WriteRF
    Close #TextChannel2
    frmProgress.Hide
    
    modPRGlobal.Prvw.vsp.EndDoc
    modPRGlobal.Prvw.Show vbModal
    
    GoBack

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdSaveSubm_Click()
    If Not PreCheck Then Exit Sub
    SaveSubmitterInfo
End Sub

Sub SaveSubmitterInfo()
    With Me
    
        ' Company Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Company
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        PRGlobal.Var1 = .txtEIN.text
        PRGlobal.Var2 = .txtUserID.text
        PRGlobal.Var3 = .txtCompanyName.text
        PRGlobal.Var4 = .txtLocationAddress.text
        PRGlobal.Var5 = .txtDeliveryAddress.text
        PRGlobal.Var6 = .txtCity.text
        PRGlobal.Var7 = .txtZipCode.text
        PRGlobal.Var8 = .txtZipCodeExt.text
        PRGlobal.Var9 = Left(.cmbPreparerCode.text, 1)
        PRGlobal.Save (Equate.RecPut)
        
        ' Contact Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Contact
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        PRGlobal.Var1 = .txtContactName.text
        PRGlobal.Var2 = .txtContactPhn.text
        PRGlobal.Var3 = .txtContactPhnExt.text
        PRGlobal.Var4 = .txtContactEmail.text
        PRGlobal.Var5 = .txtContactFax.text
        PRGlobal.Save (Equate.RecPut)
        
        ' Submit Info
        strSQL = "select * from PRGlobal where TypeCode = " & PREquate.GlobalTypeOHW2Submit
        If Not PRGlobal.GetBySQL(strSQL) Then
            MsgBox "Internal err - PRGlobal ..."
            End
        End If
        PRGlobal.Byte1 = .chkResubIndicator
        PRGlobal.Var1 = .txtResubID.text
        PRGlobal.Var2 = .txtOutputFile.text
        PRGlobal.Var3 = .txtTaxYear.text
        PRGlobal.Save (Equate.RecPut)
    
    End With

End Sub

Function PreCheck() As Boolean
    Dim msg As String
    msg = ""
    With Me
        If .txtEIN.text = "" Then msg = msg & "EIN must be entered" & vbCrLf
        If .txtUserID.text = "" Then msg = msg & "User ID must be entered" & vbCrLf
        If .txtCompanyName = "" Then msg = msg & "Company Name must be entered" & vbCrLf
        If .txtDeliveryAddress = "" Then msg = msg & "Delivery Address must be entered" & vbCrLf
        If .txtCity = "" Then msg = msg & "City must be entered" & vbCrLf
        If .txtState = "" Then msg = msg & "State must be entered" & vbCrLf
        If .txtZipCode = "" Then msg = msg & "Zip Code must be entered" & vbCrLf
        If .txtContactName = "" Then msg = msg & "Contact Name must be entered" & vbCrLf
        If .txtContactPhn = "" Then msg = msg & "Contact Phn must be entered" & vbCrLf
        If .txtContactEmail = "" Then msg = msg & "Contact Email must be entered" & vbCrLf
        If .cmbPreparerCode = "" Then msg = msg & "Preparer Code must be entered" & vbCrLf
        If .txtTaxYear = "" Then msg = msg & "Tax Year must be entered" & vbCrLf
        If Not (IsNumeric(.txtTaxYear.text)) Then msg = msg & "Tax Year must be entered" & vbCrLf
    End With
    If msg <> "" Then
        MsgBox msg, vbExclamation
        PreCheck = False
    Else
        PreCheck = True
    End If
End Function

Sub InitPRGlobal(ByVal TypeCode As Integer)
    strSQL = "select * from PRGlobal where TypeCode = " & TypeCode
    If Not PRGlobal.GetBySQL(strSQL) Then
        PRGlobal.Clear
        PRGlobal.TypeCode = TypeCode
        PRGlobal.Save (Equate.RecAdd)
    End If
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdFileName_Click()
            
    cdlTextOutput.CancelError = True
    
    ' set to current
    cdlTextOutput.Flags = cdlCFBoth Or cdlCFEffects
    cdlTextOutput.Filter = "Text File|*.txt"
    cdlTextOutput.DialogTitle = "Select a file for the Ohio W2 Upload"
    cdlTextOutput.CancelError = True

    ' call the file dialog
    cdlTextOutput.ShowOpen
    Me.txtOutputFile.text = cdlTextOutput.FileName

End Sub


Sub CompanyReport()
    PageHeader ReportTitle, "Tax Year", Me.txtTaxYear.text, "", 1, False, True
    Ln = Ln + 1
    PrintInfo "Submitter", Me.txtCompanyName, 2
    PrintInfo PRCompany.Name, "", 1
    PrintInfo "W2 Count", W2TL.RWCount, 1
    PrintInfo "Box 1 Wages", AmtPrt(W2TL.Box1_Wages), 1
    PrintInfo "Box 2 Fed Tax", AmtPrt(W2TL.Box2_FedTax), 1
    PrintInfo "Box 3 SS Wages", AmtPrt(W2TL.Box3_SSWages), 1
    PrintInfo "Box 4 SS Tax", AmtPrt(W2TL.Box4_SSTax), 1
    PrintInfo "Box 5 Med Wages", AmtPrt(W2TL.Box5_MedWages), 1
    PrintInfo "Box 6 Med Tax", AmtPrt(W2TL.Box6_MedTax), 1
    PrintInfo "Box 7 SS Tips", AmtPrt(W2TL.Box7_SSTips), 1
    PrintInfo "Box 8 Alloc Tips", AmtPrt(W2TL.Box8_AllocTips), 1
    PrintInfo "Box 9 EIC", AmtPrt(W2TL.Box9_EIC), 1
    PrintInfo "Box 10 DC Benefits", AmtPrt(W2TL.Box10_DCBen), 1
    PrintInfo "Box 11 NQ Plans", AmtPrt(W2TL.Box11_NQPlans), 1
    PrintInfo "Box 12 Code A", AmtPrt(W2TL.CodeA), 1
    PrintInfo "Box 12 Code B", AmtPrt(W2TL.CodeB), 1
    PrintInfo "Box 12 Code C", AmtPrt(W2TL.CodeC), 1
    PrintInfo "Box 12 Code D", AmtPrt(W2TL.CodeD), 1
    PrintInfo "Box 12 Code E", AmtPrt(W2TL.CodeE), 1
    PrintInfo "Box 12 Code F", AmtPrt(W2TL.CodeF), 1
    PrintInfo "Box 12 Code G", AmtPrt(W2TL.CodeG), 1
    PrintInfo "Box 12 Code H", AmtPrt(W2TL.CodeH), 1
    PrintInfo "Box 12 Code M", AmtPrt(W2TL.CodeM), 1
    PrintInfo "Box 12 Code N", AmtPrt(W2TL.CodeN), 1
    PrintInfo "Box 12 Code Q", AmtPrt(W2TL.CodeQ), 1
    PrintInfo "Box 12 Code R", AmtPrt(W2TL.CodeR), 1
    PrintInfo "Box 12 Code S", AmtPrt(W2TL.CodeS), 1
    PrintInfo "Box 12 Code T", AmtPrt(W2TL.CodeT), 1
    PrintInfo "Box 12 Code V", AmtPrt(W2TL.CodeV), 1
    PrintInfo "Box 12 Code W", AmtPrt(W2TL.CodeW), 1
    PrintInfo "Box 12 Code Y", AmtPrt(W2TL.CodeY), 1
    PrintInfo "Box 12 Code Z", AmtPrt(W2TL.CodeZ), 1
    PrintInfo "Box 12 Code AA", AmtPrt(W2TL.CodeAA), 1
    PrintInfo "Box 12 Code BB", AmtPrt(W2TL.CodeBB), 1
    PrintInfo "Box 12 Code DD", AmtPrt(W2TL.CodeDD), 1
    PrintInfo "Box 12 Code EE", AmtPrt(W2TL.CodeEE), 1
    PrintInfo "Box 12 Code FF", AmtPrt(W2TL.CodeFF), 1
    PrintInfo "Box 12 Code GG", AmtPrt(W2TL.CodeGG), 1
    PrintInfo "Box 12 Code HH", AmtPrt(W2TL.CodeHH), 1
    PrintInfo "Box 14 Retirement", AmtPrt(W2TL.RetireAmt), 1
    PrintInfo "Box 16 State Wage", AmtPrt(W2TL.Box16_StateWages), 1
    PrintInfo "Box 17 State Tax", AmtPrt(W2TL.Box17_StateTax), 1
    PrintInfo "SD Tax Wage", AmtPrt(W2TL.Box18_LocalWages), 1
    PrintInfo "SD Tax", AmtPrt(W2TL.Box19_LocalTax), 1
    
    If OutOfStateEmp <> "" Then
        PrintInfo "Employee(s) skipped - No OH W2", "", 1
        PrintInfo OutOfStateEmp, "", 1
    End If
    OutOfStateEmp = ""
    
End Sub

Sub InitReport()
    PrtInit ("Port")    ' "Port" = Portrait
    SetFont 10, Equate.Portrait
    
    ReportTitle = "OH W2 Upload"
    Prvw.Caption = ReportTitle
    
    Ln = 2
    PageHeader ReportTitle, "Tax Year", Me.txtTaxYear.text, "", 1, False, True
    
    PrintInfo "Upload File Name", Me.txtOutputFile.text, 2
    PrintInfo "Submitter Info:", "", 1
    PrintInfo "E I N", Me.txtEIN.text, 1
    PrintInfo "BSO UserID", Me.txtUserID, 1
    PrintInfo "Company Name", Me.txtCompanyName, 1
    PrintInfo "Location Address", Me.txtLocationAddress, 1
    PrintInfo "Delivery Address", Me.txtDeliveryAddress, 1
    PrintInfo "City", Me.txtCity, 1
    PrintInfo "State", Me.txtState, 1
    PrintInfo "Zip Code", Me.txtZipCode, 1
    PrintInfo "Zip Code Ext", Me.txtZipCodeExt, 1
    PrintInfo "Contact Name", Me.txtContactName, 1
    PrintInfo "Contact Phn#", Me.txtContactPhn, 1
    PrintInfo "Contact Phn Ext", Me.txtContactPhnExt, 1
    PrintInfo "Contact Email", Me.txtContactEmail, 1
    PrintInfo "Preparer Code", Me.cmbPreparerCode.text, 2
    
    PrintInfo "Companies Reported", "", 1
    strSQL = "select * from PRCompany where OHeW2 = True"
    If Not PRCompany.GetBySQL(strSQL) Then End
    Do
        PrintInfo PRCompany.Name, "", 1
        If Not PRCompany.GetNext Then Exit Do
    Loop
    
    FormFeed

End Sub
Function AmtPrt(ByVal Amt As Currency) As String
    AmtPrt = Format(Amt, "##,###,##0.00")
    AmtPrt = Space(20 - Len(AmtPrt)) & AmtPrt
End Function

Sub PrintInfo(ByVal strLabel As String, ByVal strInfo As String, ByVal lnfd As Integer)
    PrintValue(1) = strLabel: FormatString(1) = "a30"
    PrintValue(2) = strInfo: FormatString(2) = "a60"
    PrintValue(3) = " ": FormatString(3) = "~"
    FormatPrint
    Ln = Ln + lnfd
End Sub

Sub WriteRF()
    sOut = "RF"
    sOut = sOut & Wrt("", 5)
    sOut = sOut & Format(W2TL.RWTotalCount, "000000000")
    sOut = sOut & Wrt("", 496)
    Print #TextChannel2, sOut
End Sub

Sub WriteRU()
    sOut = "RU"
    sOut = sOut & Format(W2TL.ROCount, "0000000")
    sOut = sOut & AmtFmt15(W2TL.Box8_AllocTips)
    sOut = sOut & AmtFmt15(W2TL.CodeA + W2TL.CodeB)
    sOut = sOut & AmtFmt15(W2TL.CodeR)
    sOut = sOut & AmtFmt15(W2TL.CodeS)
    sOut = sOut & AmtFmt15(W2TL.CodeT)
    sOut = sOut & AmtFmt15(W2TL.CodeM)
    sOut = sOut & AmtFmt15(W2TL.CodeN)
    sOut = sOut & AmtFmt15(W2TL.CodeZ)
    sOut = sOut & Wrt("", 15)
    sOut = sOut & AmtFmt15(W2TL.CodeEE)
    sOut = sOut & AmtFmt15(W2TL.CodeGG)
    sOut = sOut & AmtFmt15(W2TL.CodeHH)
    sOut = sOut & Wrt("", 165)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(W2TL.RetireAmt)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & AmtFmt15(0)
    sOut = sOut & Wrt("", 23)
    Print #TextChannel2, sOut
End Sub

Sub WriteRT()
    sOut = "RT"
    sOut = sOut & Format(W2TL.RWCount, "0000000")
    sOut = sOut & AmtFmt15(W2TL.Box1_Wages)
    sOut = sOut & AmtFmt15(W2TL.Box2_FedTax)
    sOut = sOut & AmtFmt15(W2TL.Box3_SSWages)
    sOut = sOut & AmtFmt15(W2TL.Box4_SSTax)
    sOut = sOut & AmtFmt15(W2TL.Box5_MedWages)
    sOut = sOut & AmtFmt15(W2TL.Box6_MedTax)
    sOut = sOut & AmtFmt15(W2TL.Box7_SSTips)
    sOut = sOut & Wrt("", 15)
    sOut = sOut & AmtFmt15(W2TL.Box10_DCBen)
    sOut = sOut & AmtFmt15(W2TL.CodeD)
    sOut = sOut & AmtFmt15(W2TL.CodeE)
    sOut = sOut & AmtFmt15(W2TL.CodeF)
    sOut = sOut & AmtFmt15(W2TL.CodeG)
    sOut = sOut & AmtFmt15(W2TL.CodeH)
    sOut = sOut & Wrt("", 15)
    sOut = sOut & AmtFmt15(0)       ' ToDo 457 nq
    sOut = sOut & AmtFmt15(W2TL.CodeW)
    sOut = sOut & AmtFmt15(0)       ' ToDo 457 nq ns
    sOut = sOut & AmtFmt15(W2TL.CodeQ)
    sOut = sOut & AmtFmt15(W2TL.CodeDD)
    sOut = sOut & AmtFmt15(W2TL.CodeC)
    sOut = sOut & AmtFmt15(W2TL.IncTax3rdSick)  ' ToDo ???
    sOut = sOut & AmtFmt15(W2TL.CodeV)
    sOut = sOut & AmtFmt15(W2TL.CodeY)
    sOut = sOut & AmtFmt15(W2TL.CodeAA)
    sOut = sOut & AmtFmt15(W2TL.CodeBB)
    sOut = sOut & AmtFmt15(W2TL.CodeFF)
    sOut = sOut & Wrt("", 98)
    Print #TextChannel2, sOut
End Sub


Sub WriteRS()
    
    strSQL = "select *" & _
            " from PRW2State " & _
            " where W2ID = " & PRW2.W2ID & _
            " and TaxYear = " & Me.txtTaxYear & _
            " and StateID = " & OHStateID
    If Not PRW2State.GetBySQL(strSQL) Then Exit Sub ' ???
    
    Dim NameLast As String
    Dim NameSuffix As String
    ii = InStr(PRW2.BoxE_EELastName, ",")
    If ii > 0 Then
        NameLast = Left(PRW2.BoxE_EELastName, ii - 1)
        NameSuffix = Right(PRW2.BoxE_EELastName, Len(PRW2.BoxE_EELastName) - ii)
    Else
        NameLast = PRW2.BoxE_EELastName
        NameSuffix = ""
    End If
    
    Dim ZipExt As String
    If Len(Trim(PRW2.BoxE_EEZip)) > 5 Then
        ZipExt = Right(PRW2.BoxE_EEZip, 4)
    Else
        ZipExt = ""
    End If
    
    sOut = "RS39"
    sOut = sOut & Wrt("", 5)
    sOut = sOut & Wrt(Right("000000000" & PRW2.BoxA_SSNumber, 9), 9)
    sOut = sOut & Wrt(PRW2.BoxE_EEFirstName, 15)
    sOut = sOut & Wrt(Replace(PRW2.BoxE_EEMidInit, ".", ""), 15)
    sOut = sOut & Wrt(NameLast, 20)
    sOut = sOut & Wrt(NameSuffix, 4)
    sOut = sOut & Wrt(PRW2.BoxE_EEAddr2, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EEAddr1, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EECity, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EEState, 2)
    sOut = sOut & Wrt(Left(PRW2.BoxE_EEZip, 5), 5)
    sOut = sOut & Wrt(ZipExt, 4)
    sOut = sOut & Wrt("", 5)
    sOut = sOut & Wrt("", 23)       ' foreign state
    sOut = sOut & Wrt("", 15)       ' foreign postal code
    sOut = sOut & Wrt("", 2)        ' country code
    
    ' unemployment reporting
    sOut = sOut & Wrt("", 2)
    sOut = sOut & Wrt("", 6)
    sOut = sOut & Wrt("", 11)
    sOut = sOut & Wrt("", 11)
    sOut = sOut & Wrt("", 2)
    sOut = sOut & Wrt("", 8)
    sOut = sOut & Wrt("", 8)
    sOut = sOut & Wrt("", 5)
    
    sOut = sOut & Wrt(PRW2State.ERStateID, 20)
    sOut = sOut & Wrt("", 6)
    sOut = sOut & Wrt("39", 2)
    sOut = sOut & AmtFmt(PRW2State.StateWage)
    sOut = sOut & AmtFmt(PRW2State.StateTax)
    sOut = sOut & Right(AmtFmt(PRW2.Box1_Wages), 10)
    
    W2TL.Box16_StateWages = W2TL.Box16_StateWages + PRW2State.StateWage
    W2TL.Box17_StateTax = W2TL.Box17_StateTax + PRW2State.StateTax
    
    ' SD Tax?
    Dim LocalWages As Currency
    Dim LocalTax As Currency
    Dim TaxTypeCode As String
    Dim SDNumber As String
    TaxTypeCode = ""
    LocalWages = 0
    LocalTax = 0
    SDNumber = ""
    strSQL = "select *" & _
            " from PRW2City" & _
            " where W2ID = " & PRW2.W2ID & _
            " and TaxYear = " & Me.txtTaxYear & _
            " and SDTax = 1"
    If PRW2City.GetBySQL(strSQL) Then
        TaxTypeCode = "E"
        LocalWages = PRW2City.CityWage
        LocalTax = PRW2City.CityTax
        If PRItem.GetByID(PRW2City.CityID) Then
            SDNumber = PRItem.Abbreviation
        Else
            MsgBox "Item ID not found for SD Tax: " & PRW2City.CityID
            End
        End If
    End If
    
    sOut = sOut & Wrt(TaxTypeCode, 1)
    sOut = sOut & AmtFmt(LocalWages)
    sOut = sOut & AmtFmt(LocalTax)
    
    ' this is SD tax only ...
    W2TL.Box18_LocalWages = W2TL.Box18_LocalWages + LocalWages
    W2TL.Box19_LocalTax = W2TL.Box19_LocalTax + LocalTax
    
    ' right justory
    SDNumber = Trim(SDNumber)
    If SDNumber <> "" Then
        sOut = sOut & Wrt(Space(7 - Len(SDNumber)) & SDNumber, 7)
    Else
        sOut = sOut & Wrt("", 7)
    End If
    
    sOut = sOut & Wrt("", 75)
    sOut = sOut & Wrt("", 75)
    sOut = sOut & Wrt("", 25)
    
    Print #TextChannel2, sOut

End Sub

Sub WriteRO()
    
    Dim roamt As Currency
    roamt = 0
    roamt = roamt + PRW2.Box8_AllocTips
    roamt = roamt + GetBox12Amt("A")
    roamt = roamt + GetBox12Amt("B")
    roamt = roamt + GetBox12Amt("R")
    roamt = roamt + GetBox12Amt("S")
    roamt = roamt + GetBox12Amt("T")
    roamt = roamt + GetBox12Amt("M")
    roamt = roamt + GetBox12Amt("N")
    roamt = roamt + GetBox12Amt("Z")
    roamt = roamt + GetBox12Amt("EE")
    roamt = roamt + GetBox12Amt("GG")
    roamt = roamt + GetBox12Amt("HH")
    
    Dim RetireAmt As Currency
    RetireAmt = 0
    If PRGlobal.GetByID(PRW2.Box14A_ID) Then
        If PRGlobal.Description = "RETIREMENT" Then
            roamt = roamt + PRW2.Box14A_Amount
            RetireAmt = RetireAmt + PRW2.Box14A_Amount
        End If
    End If
    If PRGlobal.GetByID(PRW2.Box14B_ID) Then
        If PRGlobal.Description = "RETIREMENT" Then
            roamt = roamt + PRW2.Box14B_Amount
            RetireAmt = RetireAmt + PRW2.Box14B_Amount
        End If
    End If
    If PRGlobal.GetByID(PRW2.Box14C_ID) Then
        If PRGlobal.Description = "RETIREMENT" Then
            roamt = roamt + PRW2.Box14C_Amount
            RetireAmt = RetireAmt + PRW2.Box14C_Amount
        End If
    End If
    If PRGlobal.GetByID(PRW2.Box14D_ID) Then
        If PRGlobal.Description = "RETIREMENT" Then
            roamt = roamt + PRW2.Box14D_Amount
            RetireAmt = RetireAmt + PRW2.Box14D_Amount
        End If
    End If
    
    If roamt <= 0 Then Exit Sub
    
    W2TL.ROCount = W2TL.ROCount + 1
    W2TL.Box8_AllocTips = W2TL.Box8_AllocTips + PRW2.Box8_AllocTips
    W2TL.CodeA = W2TL.CodeA + GetBox12Amt("A")
    W2TL.CodeB = W2TL.CodeB + GetBox12Amt("B")
    W2TL.CodeR = W2TL.CodeR + GetBox12Amt("R")
    W2TL.CodeS = W2TL.CodeS + GetBox12Amt("S")
    W2TL.CodeT = W2TL.CodeT + GetBox12Amt("T")
    W2TL.CodeM = W2TL.CodeM + GetBox12Amt("M")
    W2TL.CodeN = W2TL.CodeN + GetBox12Amt("N")
    W2TL.CodeZ = W2TL.CodeZ + GetBox12Amt("Z")
    W2TL.CodeEE = W2TL.CodeEE + GetBox12Amt("EE")
    W2TL.CodeGG = W2TL.CodeGG + GetBox12Amt("GG")
    W2TL.CodeHH = W2TL.CodeHH + GetBox12Amt("HH")
    W2TL.RetireAmt = W2TL.RetireAmt + RetireAmt
    
    sOut = "RO"
    sOut = sOut & Wrt("", 9)
    sOut = sOut & AmtFmt(PRW2.Box8_AllocTips)
    sOut = sOut & AmtFmt(GetBox12Amt("A") + GetBox12Amt("B"))
    sOut = sOut & AmtFmt(GetBox12Amt("R"))
    sOut = sOut & AmtFmt(GetBox12Amt("S"))
    sOut = sOut & AmtFmt(GetBox12Amt("T"))
    sOut = sOut & AmtFmt(GetBox12Amt("M"))
    sOut = sOut & AmtFmt(GetBox12Amt("N"))
    sOut = sOut & AmtFmt(GetBox12Amt("Z"))
    sOut = sOut & Wrt("", 11)
    sOut = sOut & AmtFmt(GetBox12Amt("EE"))
    sOut = sOut & AmtFmt(GetBox12Amt("GG"))
    sOut = sOut & AmtFmt(GetBox12Amt("HH"))
    sOut = sOut & Wrt("", 131)
    
    ' Puerto Rico fields
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    
    sOut = sOut & AmtFmt(RetireAmt)
        
    sOut = sOut & Wrt("", 11)
    
    ' vi/guam
    sOut = sOut & AmtFmt(0)
    sOut = sOut & AmtFmt(0)
    
    sOut = sOut & Wrt("", 128)
    
    Print #TextChannel2, sOut

End Sub

Sub WriteRW()
    
    ' must have OH W2
    strSQL = "select *" & _
            " from PRW2State" & _
            " where W2ID = " & PRW2.W2ID & _
            " and TaxYear = " & PRW2.TaxYear & _
            " and StateID = " & OHStateID
    If Not PRW2State.GetBySQL(strSQL) Then
        If PREmployee.GetByID(PRW2.EmployeeID) Then
            OutOfStateEmp = OutOfStateEmp & PREmployee.EmployeeNumber & " " & PREmployee.FirstName & " " & PREmployee.LastName & vbCr
        End If
        Exit Sub
    End If
    
    Dim NameLast As String
    Dim NameSuffix As String
    ii = InStr(PRW2.BoxE_EELastName, ",")
    If ii > 0 Then
        NameLast = Left(PRW2.BoxE_EELastName, ii - 1)
        NameSuffix = Right(PRW2.BoxE_EELastName, Len(PRW2.BoxE_EELastName) - ii)
    Else
        NameLast = PRW2.BoxE_EELastName
        NameSuffix = ""
    End If
    
    Dim ZipExt As String
    If Len(Trim(PRW2.BoxE_EEZip)) > 5 Then
        ZipExt = Right(PRW2.BoxE_EEZip, 4)
    Else
        ZipExt = ""
    End If
    
    W2TL.Box1_Wages = W2TL.Box1_Wages + PRW2.Box1_Wages
    W2TL.Box2_FedTax = W2TL.Box2_FedTax + PRW2.Box2_FedTax
    W2TL.Box3_SSWages = W2TL.Box3_SSWages + PRW2.Box3_SSWages
    W2TL.Box4_SSTax = W2TL.Box4_SSTax + PRW2.Box4_SSTax
    W2TL.Box5_MedWages = W2TL.Box5_MedWages + PRW2.Box5_MedWages
    W2TL.Box6_MedTax = W2TL.Box6_MedTax + PRW2.Box6_MedTax
    W2TL.Box7_SSTips = W2TL.Box7_SSTips + PRW2.Box7_SSTips
    W2TL.Box10_DCBen = W2TL.Box10_DCBen + PRW2.Box10_DCBen
    
    W2TL.CodeD = W2TL.CodeD + GetBox12Amt("D")
    W2TL.CodeE = W2TL.CodeE + GetBox12Amt("E")
    W2TL.CodeF = W2TL.CodeF + GetBox12Amt("F")
    W2TL.CodeG = W2TL.CodeG + GetBox12Amt("G")
    W2TL.CodeH = W2TL.CodeH + GetBox12Amt("H")
    W2TL.CodeW = W2TL.CodeW + GetBox12Amt("W")
    W2TL.CodeQ = W2TL.CodeQ + GetBox12Amt("Q")
    W2TL.CodeC = W2TL.CodeC + GetBox12Amt("C")
    W2TL.CodeV = W2TL.CodeV + GetBox12Amt("V")
    W2TL.CodeY = W2TL.CodeY + GetBox12Amt("Y")
    W2TL.CodeAA = W2TL.CodeAA + GetBox12Amt("AA")
    W2TL.CodeDD = W2TL.CodeDD + GetBox12Amt("DD")
    W2TL.CodeBB = W2TL.CodeBB + GetBox12Amt("BB")
    W2TL.CodeFF = W2TL.CodeFF + GetBox12Amt("FF")
    ' ToDo - 457 ???
    
    sOut = "RW"
    sOut = sOut & Wrt(Right("000000000" & PRW2.BoxA_SSNumber, 9), 9)
    sOut = sOut & Wrt(PRW2.BoxE_EEFirstName, 15)
    sOut = sOut & Wrt(Replace(PRW2.BoxE_EEMidInit, ".", ""), 15)
    sOut = sOut & Wrt(NameLast, 20)
    sOut = sOut & Wrt(NameSuffix, 4)
    sOut = sOut & Wrt(PRW2.BoxE_EEAddr2, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EEAddr1, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EECity, 22)
    sOut = sOut & Wrt(PRW2.BoxE_EEState, 2)
    sOut = sOut & Wrt(Left(PRW2.BoxE_EEZip, 5), 5)
    sOut = sOut & Wrt(ZipExt, 4)
    sOut = sOut & Wrt("", 5)
    sOut = sOut & Wrt("", 23)       ' foreign state
    sOut = sOut & Wrt("", 15)       ' foreign postal code
    sOut = sOut & Wrt("", 2)        ' country code
    sOut = sOut & AmtFmt(PRW2.Box1_Wages)
    sOut = sOut & AmtFmt(PRW2.Box2_FedTax)
    sOut = sOut & AmtFmt(PRW2.Box3_SSWages)
    sOut = sOut & AmtFmt(PRW2.Box4_SSTax)
    sOut = sOut & AmtFmt(PRW2.Box5_MedWages)
    sOut = sOut & AmtFmt(PRW2.Box6_MedTax)
    sOut = sOut & AmtFmt(PRW2.Box7_SSTips)
    sOut = sOut & Wrt("", 11)
    sOut = sOut & AmtFmt(PRW2.Box10_DCBen)
    sOut = sOut & AmtFmt(GetBox12Amt("D"))
    sOut = sOut & AmtFmt(GetBox12Amt("E"))
    sOut = sOut & AmtFmt(GetBox12Amt("F"))
    sOut = sOut & AmtFmt(GetBox12Amt("G"))
    sOut = sOut & AmtFmt(GetBox12Amt("H"))
    sOut = sOut & Wrt("", 11)
    sOut = sOut & AmtFmt(0)                 ' ????? sec 457 - not code G  ToDo
    sOut = sOut & AmtFmt(GetBox12Amt("W"))
    sOut = sOut & AmtFmt(0)                 ' ????? NOT sec 457 - not code G  ToDo
    sOut = sOut & AmtFmt(GetBox12Amt("Q"))
    sOut = sOut & Wrt("", 11)
    sOut = sOut & AmtFmt(GetBox12Amt("C"))
    sOut = sOut & AmtFmt(GetBox12Amt("V"))
    sOut = sOut & AmtFmt(GetBox12Amt("Y"))
    sOut = sOut & AmtFmt(GetBox12Amt("AA"))
    sOut = sOut & AmtFmt(GetBox12Amt("BB"))
    sOut = sOut & AmtFmt(GetBox12Amt("DD"))
    sOut = sOut & AmtFmt(GetBox12Amt("FF"))
    sOut = sOut & Wrt("", 1)
        
    sOut = sOut & Wrt(PRW2.Box13_StatEmp, 1)
    sOut = sOut & Wrt("", 1)
    sOut = sOut & Wrt(PRW2.Box13_RetirePlan, 1)
    sOut = sOut & Wrt(PRW2.Box13_3rdParty, 1)
    sOut = sOut & Wrt("", 23)
    Print #TextChannel2, sOut
    W2TL.RWCount = W2TL.RWCount + 1
    W2TL.RWTotalCount = W2TL.RWTotalCount + 1
End Sub

Sub WriteRA()
    With Me
        sOut = ""
        sOut = sOut & Wrt("RA", 2)
        sOut = sOut & Wrt(.txtEIN.text, 9)
        sOut = sOut & Wrt(.txtUserID, 8)
        sOut = sOut & Wrt("", 4)    ' software vendor code
        sOut = sOut & Wrt("", 5)
        sOut = sOut & Wrt(IIf(.chkResubIndicator.Value = True, "1", "0"), 1)
        sOut = sOut & Wrt(.txtResubID, 6)
        sOut = sOut & Wrt("98", 2)  ' software code - in house program
        
        sOut = sOut & Wrt(.txtCompanyName, 57)
        sOut = sOut & Wrt(.txtLocationAddress, 22)
        sOut = sOut & Wrt(.txtDeliveryAddress, 22)
        sOut = sOut & Wrt(.txtCity, 22)
        sOut = sOut & Wrt(.txtState, 2)
        sOut = sOut & Wrt(.txtZipCode, 5)
        sOut = sOut & Wrt(.txtZipCodeExt, 4)
        sOut = sOut & Wrt("", 5)
        sOut = sOut & Wrt("", 23)       ' foreign state
        sOut = sOut & Wrt("", 15)       ' foreign postal code
        sOut = sOut & Wrt("", 2)        ' country code
        
        sOut = sOut & Wrt(.txtCompanyName, 57)
        sOut = sOut & Wrt(.txtLocationAddress, 22)
        sOut = sOut & Wrt(.txtDeliveryAddress, 22)
        sOut = sOut & Wrt(.txtCity, 22)
        sOut = sOut & Wrt(.txtState, 2)
        sOut = sOut & Wrt(.txtZipCode, 5)
        sOut = sOut & Wrt(.txtZipCodeExt, 4)
        sOut = sOut & Wrt("", 5)
        sOut = sOut & Wrt("", 23)       ' foreign state
        sOut = sOut & Wrt("", 15)       ' foreign postal code
        sOut = sOut & Wrt("", 2)        ' country code
        
        sOut = sOut & Wrt(.txtContactName, 27)
        sOut = sOut & Wrt(.txtContactPhn, 15)
        sOut = sOut & Wrt(.txtContactPhnExt, 5)
        sOut = sOut & Wrt("", 3)
        sOut = sOut & Wrt(.txtContactEmail, 40)
        sOut = sOut & Wrt("", 3)
        sOut = sOut & Wrt(.txtContactFax, 10)
        sOut = sOut & Wrt("", 1)
        sOut = sOut & Wrt(.cmbPreparerCode.text, 1)
        sOut = sOut & Wrt("", 12)
    End With
    Print #TextChannel2, sOut
End Sub

Function AmtFmt(Amt As Currency) As String
    If Amt < 0 Then Amt = 0
    AmtFmt = Format(Amt * 100, "00000000000")
End Function
Function AmtFmt15(Amt As Currency) As String
    If Amt < 0 Then Amt = 0
    AmtFmt15 = Format(Amt * 100, "000000000000000")
End Function

Function GetBox12Amt(ByVal Code12 As String) As Currency
    GetBox12Amt = 0
    If PRW2.Box12A_ID <> 0 Then GetBox12Amt = GetBox12Amt + GetBox12CodeAmt(Code12, PRW2.Box12A_ID, 1)
    If PRW2.Box12B_ID <> 0 Then GetBox12Amt = GetBox12Amt + GetBox12CodeAmt(Code12, PRW2.Box12B_ID, 2)
    If PRW2.Box12C_ID <> 0 Then GetBox12Amt = GetBox12Amt + GetBox12CodeAmt(Code12, PRW2.Box12C_ID, 3)
    If PRW2.Box12D_ID <> 0 Then GetBox12Amt = GetBox12Amt + GetBox12CodeAmt(Code12, PRW2.Box12D_ID, 4)
End Function

Function GetBox12CodeAmt(ByVal Code12 As String, ByVal ID12 As Integer, ByVal BucketNum As Integer) As Currency
    If Not PRGlobal.GetByID(ID12) Then
        MsgBox "PRGlobal - Box 12 Error!! " & ID12
        End
    End If
    Dim ThisCode As String
    ThisCode = Mid(PRGlobal.Description, 2, IIf(Mid(PRGlobal.Description, 3, 1) = ")", 1, 2))
    If ThisCode = Code12 Then
        If BucketNum = 1 Then GetBox12CodeAmt = PRW2.Box12A_Amount
        If BucketNum = 2 Then GetBox12CodeAmt = PRW2.Box12B_Amount
        If BucketNum = 3 Then GetBox12CodeAmt = PRW2.Box12C_Amount
        If BucketNum = 4 Then GetBox12CodeAmt = PRW2.Box12D_Amount
    End If
End Function

Sub WriteCompany(PRCompanyID As Integer)
    ' todo - validate eer ein / state id
    ' prcompany zip+4
    cn.Close
    
    ' open the company database
    If BalintFolder = "" Then
        X = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
        ' 2016-04-23
        X = "\Balint\Data\" & FNameOnly(PRCompany.FileName)
    Else
        X = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(PRCompany.FileName)
    End If
    If FileExt = ".accdb" Then X = Replace(LCase(X), ".mdb", ".accdb")
    CNOpen X, ""
    
    strSQL = "select * from PREmployee"
    If Not PREmployee.GetBySQL(strSQL) Then
        MsgBox "wtf..."
        End
    End If
    WriteRE
    
End Sub

Sub WriteRE()
    
    Dim ZipExt As String
    If Len(Trim(PRCompany.ZipCode)) > 5 Then
        ZipExt = Right(PRCompany.ZipCode, 4)
    Else
        ZipExt = ""
    End If
    
    sOut = ""
    sOut = sOut & Wrt("RE", 2)
    sOut = sOut & Wrt(Me.txtTaxYear.text, 4)
    sOut = sOut & Wrt("", 1)        ' agent indicator code
    sOut = sOut & Wrt(Replace(PRCompany.FederalID, "-", ""), 9)
    sOut = sOut & Wrt("", 9)        ' agent for EIN
    sOut = sOut & Wrt(PRCompany.TermBiz, 1)
    sOut = sOut & Wrt(PRCompany.EstablishmentNumber, 4)
    sOut = sOut & Wrt(PRCompany.OtherEIN, 9)
    sOut = sOut & Wrt(PRCompany.Name, 57)
    sOut = sOut & Wrt(PRCompany.Address2, 22)
    sOut = sOut & Wrt(PRCompany.Address1, 22)
    sOut = sOut & Wrt(PRCompany.City, 22)
    
    If Not PRState.GetByID(PRCompany.AddrStateID) Then
        sOut = sOut & "OH"
    Else
        sOut = sOut & PRState.StateAbbrev
    End If
    
    sOut = sOut & Wrt(Left(PRCompany.ZipCode, 5), 5)
    sOut = sOut & Wrt(ZipExt, 4)
    sOut = sOut & Wrt(PRCompany.KindOfEmployer, 1)
    sOut = sOut & Wrt("", 4)
    sOut = sOut & Wrt("", 23)                   ' foreign state/prov.
    sOut = sOut & Wrt("", 15)                   ' foreign postal code
    sOut = sOut & Wrt("", 2)                    ' country code
    sOut = sOut & Wrt(PRCompany.EmploymentCode, 1)
    sOut = sOut & Wrt("", 1)                    ' tax jurid. code
    sOut = sOut & Wrt(PRCompany.ThirdPartySickPay, 1)
    sOut = sOut & Wrt(PRCompany.ContactName, 27)
    sOut = sOut & Wrt(PRCompany.ContactPhoneNum, 15)
    sOut = sOut & Wrt(PRCompany.ContactPhoneExt, 5)
    sOut = sOut & Wrt(PRCompany.ContactFasNum, 10)
    sOut = sOut & Wrt(PRCompany.ContactEmail, 40)
    sOut = sOut & Wrt("", 194)
    Print #TextChannel2, sOut
End Sub

Function Wrt(ByVal strng As String, ByVal sLen As Integer) As String
    Wrt = RTrim(strng)
    If Len(strng) > sLen Then
        Wrt = Left(strng, sLen)
    Else
        Wrt = strng & Space(sLen - Len(strng))
    End If
End Function

Function InitOutputFile() As Boolean
    InitOutputFile = False
    
    With Me
        If Dir(.txtOutputFile) <> "" Then
            If MsgBox(.txtOutputFile & vbCr & "Already exists - OK to overwrite?", vbQuestion + vbYesNo, "Ohio W2 Upload") = vbNo Then Exit Function
        End If
    End With
    
    ' assign
    TextFileName = Me.txtOutputFile.text
    TextChannel2 = FreeFile

    Do
        
        On Error Resume Next
        Open TextFileName For Output As #TextChannel2
        
        If Err.Number <> 0 Then
            
            ErrMsg = "Error Opening: " & TextFileName & vbCr & vbCr & _
                " " & Err.Number & " " & Err.Description
                
            MsgResponse = MsgBox(ErrMsg, vbRetryCancel + vbExclamation, "File Open Error")
            If MsgResponse <> vbRetry Then
                TextChannel2 = 0
                TextFileName = ""
                Exit Do
            End If
            
        Else
            Exit Do
        End If
    
    Loop
    On Error GoTo 0
    If TextChannel2 <> 0 Then
        InitOutputFile = True
    End If

End Function



