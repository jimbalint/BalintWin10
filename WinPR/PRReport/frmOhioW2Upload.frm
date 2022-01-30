VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOHW2 
   Caption         =   "Ohio W2 Upload"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15795
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
   ScaleWidth      =   15795
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
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtUserID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtContactPhn"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtContactName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtZipCodeExt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtZipCode"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtState"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCity"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDeliveryAddress"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtLocationAddress"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCompanyName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEIN"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtContactPhnExt"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtContactEmail"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtContactFax"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmbPreparerCode"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdSaveSubm"
      Tab(0).Control(17).Enabled=   0   'False
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
         Left            =   11520
         TabIndex        =   25
         Top             =   720
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

Private Sub Form_Load()

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

Private Sub cmdCreateFile_Click()
    
    If Not PreCheck Then Exit Sub
    SaveSubmitterInfo
    If Not InitOutputFile Then Exit Sub
    
    WriteRA
    
    strSQL = "select * from PRCompany where OHeW2 = True"
    If Not PRCompany.GetBySQL(strSQL) Then End
    Do
        WriteCompany (65)
        If Not PRCompany.GetNext Then Exit Do
    Loop
    Close #TextChannel2
    GoBack
    
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

Sub WriteCompany(PRCompanyID As Integer)
    
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
    MsgBox (PRCompany.ContactEmail)
    
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



