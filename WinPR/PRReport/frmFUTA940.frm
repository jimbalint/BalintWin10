VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmFUTA940 
   Caption         =   "FUTA 940"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
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
   ScaleHeight     =   10695
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReload 
      Caption         =   "&RELOAD"
      Height          =   495
      Left            =   9120
      TabIndex        =   95
      ToolTipText     =   "Reload PR Data"
      Top             =   10080
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber numHorzNudge 
      Height          =   375
      Left            =   2880
      TabIndex        =   88
      Top             =   10200
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmFUTA940.frx":0000
      Caption         =   "frmFUTA940.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFUTA940.frx":009E
      Keys            =   "frmFUTA940.frx":00BC
      Spin            =   "frmFUTA940.frx":0106
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   10920
      TabIndex        =   84
      Top             =   10080
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frmFUTA940.frx":012E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk4e"
      Tab(0).Control(1)=   "chk4d"
      Tab(0).Control(2)=   "chk4c"
      Tab(0).Control(3)=   "chk4b"
      Tab(0).Control(4)=   "chk4a"
      Tab(0).Control(5)=   "num4"
      Tab(0).Control(6)=   "num3"
      Tab(0).Control(7)=   "chkCreditReduction"
      Tab(0).Control(8)=   "chkMultiState"
      Tab(0).Control(9)=   "cmbState"
      Tab(0).Control(10)=   "txtForeignZip"
      Tab(0).Control(11)=   "txtForeignProv"
      Tab(0).Control(12)=   "txtForeignCountry"
      Tab(0).Control(13)=   "txtZip"
      Tab(0).Control(14)=   "txtState"
      Tab(0).Control(15)=   "txtCity"
      Tab(0).Control(16)=   "txtAddr2"
      Tab(0).Control(17)=   "txtAddr1"
      Tab(0).Control(18)=   "txtTradeName"
      Tab(0).Control(19)=   "txtName"
      Tab(0).Control(20)=   "txtEIN"
      Tab(0).Control(21)=   "Frame1"
      Tab(0).Control(22)=   "num5"
      Tab(0).Control(23)=   "num6"
      Tab(0).Control(24)=   "num7"
      Tab(0).Control(25)=   "num8"
      Tab(0).Control(26)=   "Label8"
      Tab(0).Control(27)=   "Label7"
      Tab(0).Control(28)=   "Label6"
      Tab(0).Control(29)=   "Label5"
      Tab(0).Control(30)=   "Label3"
      Tab(0).Control(31)=   "Label2"
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frmFUTA940.frx":014A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "num15"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "num14"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "num13"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "num12"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "num11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "num10"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "num9"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkApplyToNext"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkRefund"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkLine9"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkCrRedPct"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "numCrRedPct"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Page 3"
      TabPicture(2)   =   "frmFUTA940.frx":0166
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSignNameTitle"
      Tab(2).Control(1)=   "dteSignDate"
      Tab(2).Control(2)=   "txt3rdPartyName"
      Tab(2).Control(3)=   "chk3rdPartyNo"
      Tab(2).Control(4)=   "chk3rdPartyYes"
      Tab(2).Control(5)=   "num16a"
      Tab(2).Control(6)=   "num16b"
      Tab(2).Control(7)=   "num16c"
      Tab(2).Control(8)=   "num16d"
      Tab(2).Control(9)=   "num17"
      Tab(2).Control(10)=   "txt3rdPartyPhone"
      Tab(2).Control(11)=   "txtSignPhone"
      Tab(2).Control(12)=   "txt3rdPartyPIN"
      Tab(2).Control(13)=   "Label19"
      Tab(2).Control(14)=   "Label18"
      Tab(2).Control(15)=   "Label17"
      Tab(2).Control(16)=   "Label16"
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Paid Preparer Use Only"
      TabPicture(3)   =   "frmFUTA940.frx":0182
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label21"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label22"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtPPZip"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtPPCityState"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtPPPhone"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtPPAddr"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtPPEIN"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtPPFirmName"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtPPPTIN"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "dtePPDate"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "chkPPSE"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cmbPPName"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).ControlCount=   13
      Begin TDBNumber6Ctl.TDBNumber numCrRedPct 
         Height          =   375
         Left            =   -71280
         TabIndex        =   92
         Top             =   3480
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":019E
         Caption         =   "frmFUTA940.frx":01BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0222
         Keys            =   "frmFUTA940.frx":0240
         Spin            =   "frmFUTA940.frx":028A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.0000%; ##0.0000%"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.0000%; ##0.0000%"
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
         ValueVT         =   25165825
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5177349
      End
      Begin VB.CheckBox chkCrRedPct 
         Caption         =   "Use Pct for credit reduction"
         Height          =   255
         Left            =   -74280
         TabIndex        =   91
         Top             =   3600
         Width           =   2895
      End
      Begin VB.CheckBox chkLine9 
         Height          =   375
         Left            =   -61440
         TabIndex        =   90
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox cmbPPName 
         Height          =   360
         Left            =   2640
         TabIndex        =   87
         Top             =   1680
         Width           =   6495
      End
      Begin VB.CheckBox chkPPSE 
         Caption         =   "Check if you are self-employed"
         Height          =   375
         Left            =   5520
         TabIndex        =   85
         Top             =   600
         Width           =   3735
      End
      Begin TDBDate6Ctl.TDBDate dtePPDate 
         Height          =   375
         Left            =   9840
         TabIndex        =   76
         Top             =   2280
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         Calendar        =   "frmFUTA940.frx":02B2
         Caption         =   "frmFUTA940.frx":03B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0410
         Keys            =   "frmFUTA940.frx":042E
         Spin            =   "frmFUTA940.frx":048C
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
         Text            =   "10/13/2016"
         ValidateMode    =   0
         ValueVT         =   3342343
         Value           =   42656
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText txtSignNameTitle 
         Height          =   375
         Left            =   -71160
         TabIndex        =   72
         Top             =   6300
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":04B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0544
         Key             =   "frmFUTA940.frx":0562
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
      Begin TDBDate6Ctl.TDBDate dteSignDate 
         Height          =   375
         Left            =   -74400
         TabIndex        =   71
         Top             =   6300
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         Calendar        =   "frmFUTA940.frx":05A6
         Caption         =   "frmFUTA940.frx":06A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0704
         Keys            =   "frmFUTA940.frx":0722
         Spin            =   "frmFUTA940.frx":0780
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
         Text            =   "10/13/2016"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   42656
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText txt3rdPartyName 
         Height          =   375
         Left            =   -73200
         TabIndex        =   68
         Top             =   4500
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":07A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":081E
         Key             =   "frmFUTA940.frx":083C
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
      Begin VB.CheckBox chk3rdPartyNo 
         Caption         =   "No"
         Height          =   375
         Left            =   -74400
         TabIndex        =   67
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox chk3rdPartyYes 
         Caption         =   "Yes"
         Height          =   375
         Left            =   -74400
         TabIndex        =   66
         Top             =   4500
         Width           =   855
      End
      Begin VB.CheckBox chkRefund 
         Caption         =   "Send a refund"
         Height          =   255
         Left            =   -64200
         TabIndex        =   57
         Top             =   7380
         Width           =   2055
      End
      Begin VB.CheckBox chkApplyToNext 
         Caption         =   "Apply to next return"
         Height          =   255
         Left            =   -64200
         TabIndex        =   56
         Top             =   7020
         Width           =   2055
      End
      Begin VB.CheckBox chk4e 
         Caption         =   "4e - Other"
         Height          =   255
         Left            =   -63360
         TabIndex        =   37
         Top             =   7020
         Width           =   1335
      End
      Begin VB.CheckBox chk4d 
         Caption         =   "4d - Dependent care"
         Height          =   255
         Left            =   -65520
         TabIndex        =   36
         Top             =   7020
         Width           =   2175
      End
      Begin VB.CheckBox chk4c 
         Caption         =   "4c - Retirement/Pension"
         Height          =   255
         Left            =   -68040
         TabIndex        =   35
         Top             =   7020
         Width           =   2535
      End
      Begin VB.CheckBox chk4b 
         Caption         =   "4b Group-term Life Ins"
         Height          =   255
         Left            =   -70440
         TabIndex        =   34
         Top             =   7020
         Width           =   2295
      End
      Begin VB.CheckBox chk4a 
         Caption         =   "4a Fringe Benefits"
         Height          =   255
         Left            =   -72480
         TabIndex        =   33
         Top             =   7020
         Width           =   1935
      End
      Begin TDBNumber6Ctl.TDBNumber num4 
         Height          =   375
         Left            =   -74640
         TabIndex        =   31
         Top             =   6540
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":0880
         Caption         =   "frmFUTA940.frx":08A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":09BC
         Keys            =   "frmFUTA940.frx":09DA
         Spin            =   "frmFUTA940.frx":0A24
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num3 
         Height          =   375
         Left            =   -74640
         TabIndex        =   30
         Top             =   6180
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":0A4C
         Caption         =   "frmFUTA940.frx":0A6C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0BF6
         Keys            =   "frmFUTA940.frx":0C14
         Spin            =   "frmFUTA940.frx":0C5E
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
      Begin VB.CheckBox chkCreditReduction 
         Caption         =   "Check Here "
         Height          =   255
         Left            =   -63120
         TabIndex        =   28
         Top             =   5460
         Width           =   1455
      End
      Begin VB.CheckBox chkMultiState 
         Caption         =   "Check Here "
         Height          =   255
         Left            =   -63120
         TabIndex        =   26
         Top             =   5100
         Width           =   1455
      End
      Begin VB.ComboBox cmbState 
         Height          =   360
         Left            =   -63120
         TabIndex        =   21
         Top             =   4620
         Width           =   975
      End
      Begin TDBText6Ctl.TDBText txtForeignZip 
         Height          =   375
         Left            =   -65400
         TabIndex        =   18
         Top             =   3780
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":0C86
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0CE2
         Key             =   "frmFUTA940.frx":0D00
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
      Begin TDBText6Ctl.TDBText txtForeignProv 
         Height          =   375
         Left            =   -69480
         TabIndex        =   17
         Top             =   3780
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":0D44
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0DAA
         Key             =   "frmFUTA940.frx":0DC8
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
      Begin TDBText6Ctl.TDBText txtForeignCountry 
         Height          =   375
         Left            =   -74640
         TabIndex        =   16
         Top             =   3780
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":0E0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0E7A
         Key             =   "frmFUTA940.frx":0E98
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
         Left            =   -67560
         TabIndex        =   15
         Top             =   3300
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":0EDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0F38
         Key             =   "frmFUTA940.frx":0F56
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
         Left            =   -69360
         TabIndex        =   14
         Top             =   3300
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":0F9A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":0FFA
         Key             =   "frmFUTA940.frx":1018
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   3300
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":105C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":10BA
         Key             =   "frmFUTA940.frx":10D8
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
         Left            =   -74640
         TabIndex        =   12
         Top             =   2820
         Width           =   9495
         _Version        =   65536
         _ExtentX        =   16748
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":111C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1182
         Key             =   "frmFUTA940.frx":11A0
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
      Begin TDBText6Ctl.TDBText txtAddr1 
         Height          =   375
         Left            =   -74640
         TabIndex        =   11
         Top             =   2340
         Width           =   9495
         _Version        =   65536
         _ExtentX        =   16748
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":11E4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":124A
         Key             =   "frmFUTA940.frx":1268
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
      Begin TDBText6Ctl.TDBText txtTradeName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   10
         Top             =   1860
         Width           =   9495
         _Version        =   65536
         _ExtentX        =   16748
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":12AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1316
         Key             =   "frmFUTA940.frx":1334
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
      Begin TDBText6Ctl.TDBText txtName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   9
         Top             =   1380
         Width           =   9495
         _Version        =   65536
         _ExtentX        =   16748
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":1378
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":13D6
         Key             =   "frmFUTA940.frx":13F4
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
      Begin TDBText6Ctl.TDBText txtEIN 
         Height          =   375
         Left            =   -74640
         TabIndex        =   8
         Top             =   900
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":1438
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1498
         Key             =   "frmFUTA940.frx":14B6
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
      Begin VB.Frame Frame1 
         Caption         =   "Type of Return "
         Height          =   2055
         Left            =   -64800
         TabIndex        =   2
         Top             =   1020
         Width           =   3615
         Begin VB.CheckBox chkTypeD 
            Caption         =   "d. Final: Business Closed"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   1680
            Width           =   2775
         End
         Begin VB.CheckBox chkTypeC 
            Caption         =   "c. No payments to employees"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox chkTypeB 
            Caption         =   "b. Successor employer"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkTypeA 
            Caption         =   "a. Amended"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Check all that apply"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
      End
      Begin TDBNumber6Ctl.TDBNumber num5 
         Height          =   375
         Left            =   -74640
         TabIndex        =   38
         Top             =   7380
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":14FA
         Caption         =   "frmFUTA940.frx":151A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1626
         Keys            =   "frmFUTA940.frx":1644
         Spin            =   "frmFUTA940.frx":168E
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num6 
         Height          =   375
         Left            =   -74640
         TabIndex        =   39
         Top             =   7740
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":16B6
         Caption         =   "frmFUTA940.frx":16D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1868
         Keys            =   "frmFUTA940.frx":1886
         Spin            =   "frmFUTA940.frx":18D0
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
      Begin TDBNumber6Ctl.TDBNumber num7 
         Height          =   375
         Left            =   -74640
         TabIndex        =   40
         Top             =   8220
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":18F8
         Caption         =   "frmFUTA940.frx":1918
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1A8A
         Keys            =   "frmFUTA940.frx":1AA8
         Spin            =   "frmFUTA940.frx":1AF2
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
      Begin TDBNumber6Ctl.TDBNumber num8 
         Height          =   375
         Left            =   -74640
         TabIndex        =   41
         Top             =   8700
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":1B1A
         Caption         =   "frmFUTA940.frx":1B3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1CB4
         Keys            =   "frmFUTA940.frx":1CD2
         Spin            =   "frmFUTA940.frx":1D1C
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
      Begin TDBNumber6Ctl.TDBNumber num9 
         Height          =   375
         Left            =   -74760
         TabIndex        =   42
         Top             =   1380
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":1D44
         Caption         =   "frmFUTA940.frx":1D64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":1EA6
         Keys            =   "frmFUTA940.frx":1EC4
         Spin            =   "frmFUTA940.frx":1F0E
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
      Begin TDBNumber6Ctl.TDBNumber num10 
         Height          =   375
         Left            =   -74760
         TabIndex        =   45
         Top             =   2100
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":1F36
         Caption         =   "frmFUTA940.frx":1F56
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":2094
         Keys            =   "frmFUTA940.frx":20B2
         Spin            =   "frmFUTA940.frx":20FC
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
      Begin TDBNumber6Ctl.TDBNumber num11 
         Height          =   375
         Left            =   -74760
         TabIndex        =   48
         Top             =   3060
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2124
         Caption         =   "frmFUTA940.frx":2144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":22B2
         Keys            =   "frmFUTA940.frx":22D0
         Spin            =   "frmFUTA940.frx":231A
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
      Begin TDBNumber6Ctl.TDBNumber num12 
         Height          =   375
         Left            =   -74760
         TabIndex        =   50
         Top             =   4440
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2342
         Caption         =   "frmFUTA940.frx":2362
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":24D0
         Keys            =   "frmFUTA940.frx":24EE
         Spin            =   "frmFUTA940.frx":2538
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
      Begin TDBNumber6Ctl.TDBNumber num13 
         Height          =   375
         Left            =   -74760
         TabIndex        =   51
         Top             =   4920
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2560
         Caption         =   "frmFUTA940.frx":2580
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":26DC
         Keys            =   "frmFUTA940.frx":26FA
         Spin            =   "frmFUTA940.frx":2744
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
      Begin TDBNumber6Ctl.TDBNumber num14 
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   5520
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":276C
         Caption         =   "frmFUTA940.frx":278C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":28F8
         Keys            =   "frmFUTA940.frx":2916
         Spin            =   "frmFUTA940.frx":2960
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
      Begin TDBNumber6Ctl.TDBNumber num15 
         Height          =   375
         Left            =   -74760
         TabIndex        =   54
         Top             =   6480
         Width           =   12975
         _Version        =   65536
         _ExtentX        =   22886
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2988
         Caption         =   "frmFUTA940.frx":29A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":2AFC
         Keys            =   "frmFUTA940.frx":2B1A
         Spin            =   "frmFUTA940.frx":2B64
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
      Begin TDBNumber6Ctl.TDBNumber num16a 
         Height          =   375
         Left            =   -72840
         TabIndex        =   59
         Top             =   1500
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2B8C
         Caption         =   "frmFUTA940.frx":2BAC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":2CCC
         Keys            =   "frmFUTA940.frx":2CEA
         Spin            =   "frmFUTA940.frx":2D34
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num16b 
         Height          =   375
         Left            =   -72840
         TabIndex        =   60
         Top             =   1980
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2D5C
         Caption         =   "frmFUTA940.frx":2D7C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":2EA2
         Keys            =   "frmFUTA940.frx":2EC0
         Spin            =   "frmFUTA940.frx":2F0A
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num16c 
         Height          =   375
         Left            =   -72840
         TabIndex        =   61
         Top             =   2460
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":2F32
         Caption         =   "frmFUTA940.frx":2F52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3072
         Keys            =   "frmFUTA940.frx":3090
         Spin            =   "frmFUTA940.frx":30DA
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num16d 
         Height          =   375
         Left            =   -72840
         TabIndex        =   62
         Top             =   2940
         Width           =   10095
         _Version        =   65536
         _ExtentX        =   17806
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":3102
         Caption         =   "frmFUTA940.frx":3122
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":323C
         Keys            =   "frmFUTA940.frx":325A
         Spin            =   "frmFUTA940.frx":32A4
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber num17 
         Height          =   375
         Left            =   -74160
         TabIndex        =   64
         Top             =   3420
         Width           =   11415
         _Version        =   65536
         _ExtentX        =   20135
         _ExtentY        =   661
         Calculator      =   "frmFUTA940.frx":32CC
         Caption         =   "frmFUTA940.frx":32EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3422
         Keys            =   "frmFUTA940.frx":3440
         Spin            =   "frmFUTA940.frx":348A
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
         MaxValueVT      =   6881285
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText txt3rdPartyPhone 
         Height          =   375
         Left            =   -67440
         TabIndex        =   69
         Top             =   4500
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":34B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3520
         Key             =   "frmFUTA940.frx":353E
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
      Begin TDBText6Ctl.TDBText txtSignPhone 
         Height          =   375
         Left            =   -71160
         TabIndex        =   73
         Top             =   6840
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":3582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":35FC
         Key             =   "frmFUTA940.frx":361A
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
      Begin TDBText6Ctl.TDBText txtPPPTIN 
         Height          =   375
         Left            =   9840
         TabIndex        =   75
         Top             =   1440
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":365E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":36BC
         Key             =   "frmFUTA940.frx":36DA
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
      Begin TDBText6Ctl.TDBText txtPPFirmName 
         Height          =   375
         Left            =   480
         TabIndex        =   77
         Top             =   3000
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":371E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":378A
         Key             =   "frmFUTA940.frx":37A8
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
      Begin TDBText6Ctl.TDBText txtPPEIN 
         Height          =   375
         Left            =   9840
         TabIndex        =   79
         Top             =   3000
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":37EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3848
         Key             =   "frmFUTA940.frx":3866
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
      Begin TDBText6Ctl.TDBText txtPPAddr 
         Height          =   375
         Left            =   480
         TabIndex        =   80
         Top             =   3960
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":38AA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":390E
         Key             =   "frmFUTA940.frx":392C
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
      Begin TDBText6Ctl.TDBText txtPPPhone 
         Height          =   375
         Left            =   9840
         TabIndex        =   81
         Top             =   3960
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":3970
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":39D0
         Key             =   "frmFUTA940.frx":39EE
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
      Begin TDBText6Ctl.TDBText txtPPCityState 
         Height          =   375
         Left            =   480
         TabIndex        =   82
         Top             =   4800
         Width           =   8655
         _Version        =   65536
         _ExtentX        =   15266
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":3A32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3AA0
         Key             =   "frmFUTA940.frx":3ABE
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
      Begin TDBText6Ctl.TDBText txtPPZip 
         Height          =   375
         Left            =   480
         TabIndex        =   83
         Top             =   5280
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":3B02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3B66
         Key             =   "frmFUTA940.frx":3B84
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
      Begin TDBText6Ctl.TDBText txt3rdPartyPIN 
         Height          =   375
         Left            =   -73200
         TabIndex        =   94
         Top             =   5040
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "frmFUTA940.frx":3BC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFUTA940.frx":3C24
         Key             =   "frmFUTA940.frx":3C42
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
      Begin VB.Label Label22 
         Caption         =   "Preparer's name:"
         Height          =   255
         Left            =   480
         TabIndex        =   86
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "(or yours if self-employed)"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "Paid Preparer Use Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   74
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label19 
         Caption         =   "Part 7 - Sign Here.  You MUST complete both pages of this form and SIGN it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   70
         Top             =   5940
         Width           =   13935
      End
      Begin VB.Label Label18 
         Caption         =   "Part 6 - May we speak with your third-party designee?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   65
         Top             =   4020
         Width           =   13935
      End
      Begin VB.Label Label17 
         Caption         =   "Part 5 - Report your FUTA tax liability by quarter only if line 12 is more than $500.  If not, go to Part 6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   63
         Top             =   780
         Width           =   13935
      End
      Begin VB.Label Label16 
         Caption         =   $"frmFUTA940.frx":3C86
         Height          =   375
         Left            =   -74880
         TabIndex        =   58
         Top             =   1140
         Width           =   13935
      End
      Begin VB.Label Label15 
         Caption         =   ">> You MUST complete both pages of this form and SIGN it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   55
         Top             =   7320
         Width           =   6615
      End
      Begin VB.Label Label14 
         Caption         =   $"frmFUTA940.frx":3D2A
         Height          =   255
         Left            =   -74400
         TabIndex        =   53
         Top             =   6000
         Width           =   12375
      End
      Begin VB.Label Label13 
         Caption         =   "Part 4 - Determine your FUTA tax and balance due or overpayment.  If any line does NOT apply, leave it blank."
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
         Left            =   -74760
         TabIndex        =   49
         Top             =   4080
         Width           =   11295
      End
      Begin VB.Label Label12 
         Caption         =   "in the instructions.  Enter the amount from line 7 of the worksheet"
         Height          =   255
         Left            =   -74400
         TabIndex        =   47
         Top             =   2700
         Width           =   9255
      End
      Begin VB.Label Label11 
         Caption         =   "you paid ANY state unemployment tax late (after the due date for filing form 940), complete the worksheet"
         Height          =   255
         Left            =   -74400
         TabIndex        =   46
         Top             =   2460
         Width           =   9255
      End
      Begin VB.Label Label10 
         Caption         =   "(line 7 x .054 = line9) Go to line 12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67080
         TabIndex        =   44
         Top             =   1740
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Part 3 - Determine your adjustments.  If any line does NOT apply, leave it blank"
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
         Left            =   -74760
         TabIndex        =   43
         Top             =   1020
         Width           =   11295
      End
      Begin VB.Label Label8 
         Caption         =   "Check all that apply:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   7020
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Part 2 - Determine your FUTA tax before adustments.  If any line does NOT apply, leave it blank."
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
         Left            =   -74640
         TabIndex        =   29
         Top             =   5820
         Width           =   11295
      End
      Begin VB.Label Label6 
         Caption         =   $"frmFUTA940.frx":3DB4
         Height          =   255
         Left            =   -74640
         TabIndex        =   27
         Top             =   5460
         Width           =   11415
      End
      Begin VB.Label Label5 
         Caption         =   $"frmFUTA940.frx":3E52
         Height          =   255
         Left            =   -74640
         TabIndex        =   25
         Top             =   5100
         Width           =   11415
      End
      Begin VB.Label Label3 
         Caption         =   $"frmFUTA940.frx":3EE4
         Height          =   255
         Left            =   -74640
         TabIndex        =   20
         Top             =   4740
         Width           =   11415
      End
      Begin VB.Label Label2 
         Caption         =   "Part 1 - Tell us about your return.  If any line does NOT apply, leave it blank."
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
         Left            =   -74640
         TabIndex        =   19
         Top             =   4380
         Width           =   7695
      End
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   1320
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   12720
      TabIndex        =   0
      Top             =   10080
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber numVertNudge 
      Height          =   375
      Left            =   5880
      TabIndex        =   89
      Top             =   10200
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmFUTA940.frx":3F81
      Caption         =   "frmFUTA940.frx":3FA1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFUTA940.frx":401D
      Keys            =   "frmFUTA940.frx":403B
      Spin            =   "frmFUTA940.frx":4085
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
   Begin VB.Label lblUnemRate 
      Caption         =   "Label23"
      Height          =   255
      Left            =   12960
      TabIndex        =   93
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   120
      Width           =   12255
   End
   Begin VB.Label Label4 
      Caption         =   "Tax Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   10200
      Width           =   1095
   End
End
Attribute VB_Name = "frmFUTA940"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SQLString As String
Dim rs As New ADODB.Recordset
Dim rsState As New ADODB.Recordset

Dim ThirdPartyID As Integer
Dim SignID As Integer
Dim PaidPrepID As Integer
Dim FedUnempMax, FedUnempPct As Double
Dim QtrThreshold As Double

Dim State1 As String

Dim InitFlag As Boolean

Dim QtrWage(4) As Currency
Dim FUTA_PCT As Double

Dim i, j, k As Integer
Dim X, Y, Z As String



Private Sub Form_Load()

    ' ***********************************************************
    ' assumes the right FUN max in place during entry for the year
    ' if not - set max and run taxable wage sweep
    ' ***********************************************************

    Me.KeyPreview = True
    Init

End Sub

Private Sub Init()
    
    QtrThreshold = 500
    Me.numCrRedPct.Value = 0
    
    InitFlag = True
    
    tdbTextSet Me.txtName
    tdbTextSet Me.txtTradeName
    tdbTextSet Me.txtAddr1
    tdbTextSet Me.txtAddr2
    tdbTextSet Me.txtCity
    tdbTextSet Me.txtState
    tdbTextSet Me.txtZip
    tdbTextSet Me.txtForeignCountry
    tdbTextSet Me.txtForeignProv
    tdbTextSet Me.txtForeignZip
    tdbTextSet Me.txt3rdPartyName
    tdbTextSet Me.txt3rdPartyPhone
    tdbTextSet Me.txt3rdPartyPIN
    tdbTextSet Me.txtSignNameTitle
    tdbTextSet Me.txtSignPhone
    tdbTextSet Me.txtPPFirmName
    tdbTextSet Me.txtPPFirmName
    tdbTextSet Me.txtPPAddr
    tdbTextSet Me.txtPPCityState
    tdbTextSet Me.txtPPPTIN
    tdbTextSet Me.txtPPEIN
    tdbTextSet Me.txtPPPhone
    tdbTextSet Me.txtPPZip
    
    tdbAmountSet Me.num3
    tdbAmountSet Me.num4
    tdbAmountSet Me.num5
    tdbAmountSet Me.num6
    tdbAmountSet Me.num7
    tdbAmountSet Me.num8
    tdbAmountSet Me.num9
    tdbAmountSet Me.num10
    tdbAmountSet Me.num11
    tdbAmountSet Me.num12
    tdbAmountSet Me.num13
    tdbAmountSet Me.num14
    tdbAmountSet Me.num15
    
    tdbAmountSet Me.num16a
    tdbAmountSet Me.num16b
    tdbAmountSet Me.num16c
    tdbAmountSet Me.num16d
    tdbAmountSet Me.num17
    
    Me.lblCompanyName.Caption = PRCompany.Name
    
    ' populate the tax year drop down
    SQLString = " SELECT DISTINCT(YEAR(CheckDate)) as TaxYear " & _
                " FROM PRHist "
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then
        MsgBox "No PR History records found!", vbExclamation
        GoBack
    End If
    Dim SelYear As Integer
    rs.MoveFirst
    Do
        Me.cmbTaxYear.AddItem rs!TaxYear
        SelYear = rs!TaxYear
        rs.MoveNext
        If rs.EOF Then Exit Do
    Loop
    
    If Month(Date) >= 1 And Month(Date) <= 3 And rs.RecordCount >= 2 Then
        Me.cmbTaxYear.ListIndex = rs.RecordCount - 2
    Else
        Me.cmbTaxYear.ListIndex = rs.RecordCount - 1
    End If
    
    If Me.cmbTaxYear.Text < "2016" Then Me.cmbTaxYear.Text = "2016"
    
    rs.Close
    
    With Me
        
        .txtEIN = PRCompany.FederalID
        .txtName.Text = PRCompany.Name
        .txtAddr1 = PRCompany.Address1
        .txtAddr2 = PRCompany.Address2
        .txtCity = PRCompany.City
        
        If PRCompany.AddrStateID <> 0 Then
            SQLString = " SELECT * FROM PRState " & _
                        " WHERE StateID = " & PRCompany.AddrStateID
            If PRState.GetBySQL(SQLString) Then
                .txtState = PRState.StateAbbrev
            End If
        End If
        
        .txtZip = PRCompany.ZipCode
    
    End With
    
    ' Third Party Designee - Per User
    ThirdPartyID = 0
    SQLString = "SELECT * FROM PRGlobal WHERE Typecode = " & _
                PREquate.GlobalType941Part4 & " AND UserID = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        ThirdPartyID = PRGlobal.GlobalID
        Me.txt3rdPartyName = PRGlobal.Var1
        Me.txt3rdPartyPhone = PRGlobal.Var2
        Me.txt3rdPartyPIN = PRGlobal.Var3
        Me.chk3rdPartyNo = 1
        Me.chk3rdPartyYes = 0
        If Me.txt3rdPartyName.Text <> "" Then
            Me.chk3rdPartyNo = 0
            Me.chk3rdPartyYes = 1
        End If
        Me.txtTradeName.Text = PRGlobal.Var4
    End If
    
    ' Company Signature - Per Company
    SignID = 0
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & _
                PREquate.GlobalType941Part5 & " AND Userid = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) Then
        SignID = PRGlobal.GlobalID
        Me.txtSignNameTitle = PRGlobal.Var1
        Me.txtSignPhone = PRCompany.PhoneNumber
    End If
    
    ' paid preparer
    PaidPrepID = 0
    
    ' populate the User Combo Box
    Dim UserID1 As Integer
    UserID1 = User.ID
    Me.cmbPPName.Clear
    SQLString = "SELECT * FROM Users ORDER BY NAME"
    If Not User.GetBySQL(SQLString) Then
       MsgBox "Users not found: " & UserID, vbCritical, "Form940 Entry"
       End
    End If

    Do
        Me.cmbPPName.AddItem UCase(User.Name)
        If Not User.GetNext Then Exit Do
    Loop
    
    ' reget the user
    If User.GetByID(UserID1) = True Then
    End If
    
    SQLString = "SELECT * FROM PRglobal WHERE TypeCode = " & _
                PREquate.GlobalType941PaidPrep & " AND Userid = " & User.ID
    If PRGlobal.GetBySQL(SQLString) Then
        PaidPrepID = PRGlobal.GlobalID
        
        ' Me.cmbPrepName.AddItem User.Name
        
        Me.txtPPFirmName = PRGlobal.Var1
        Me.txtPPAddr = PRGlobal.Var2
        Me.txtPPCityState = PRGlobal.Var3
        Me.txtPPPhone = PRGlobal.Var4
        Me.txtPPPTIN = PRGlobal.Var7
        Me.txtPPZip = PRGlobal.Var6
        Me.txtPPEIN = PRGlobal.Var5
        If PRGlobal.Var8 = "1" Then
            Me.chkPPSE = 1
        Else
            Me.chkPPSE = 0
        End If
        
        Me.cmbPPName.Text = PRGlobal.Var9
    
    End If
    
    ' Populate state dropdown box
    PRState.GetBySQL ("SELECT * FROM PRState order by PRState.StateAbbrev")
    Do
        Me.cmbState.AddItem PRState.StateAbbrev
        If Not PRState.GetNext Then
           Exit Do
        End If
    Loop

    ' form nudge
    SetNudge Me.numHorzNudge
    SetNudge Me.numVertNudge
    GetNudge User.ID, "Form940"
    Me.numHorzNudge = nNull(HorzNudge)
    Me.numVertNudge = nNull(VertNudge)

    ' disable lines that have totals
    Dim Grey
    Grey = RGB(192, 192, 192)
    Me.num6.ReadOnly = True
    Me.num6.BackColor = Grey
    Me.num7.ReadOnly = True
    Me.num7.BackColor = Grey
    Me.num8.ReadOnly = True
    Me.num8.BackColor = Grey
    Me.num9.ReadOnly = True
    Me.num9.BackColor = Grey
    Me.num12.ReadOnly = True
    Me.num12.BackColor = Grey
    Me.num14.ReadOnly = True
    Me.num14.BackColor = Grey
    Me.num15.ReadOnly = True
    Me.num15.BackColor = Grey
    Me.num17.ReadOnly = True
    Me.num17.BackColor = Grey

    Me.SSTab1.Tab = 0
    
    ' check box defaults
    SQLString = " SELECT * " & _
                " FROM PRGlobal " & _
                " WHERE UserID = " & PRCompany.CompanyID & _
                " AND Description = 'FUTA940' "
    If PRGlobal.GetBySQL(SQLString) Then
        
        Me.chkTypeA = PRGlobal.Byte1
        Me.chkTypeB = PRGlobal.Byte2
        Me.chkTypeC = PRGlobal.Byte3
        Me.chkTypeD = PRGlobal.Byte4
        
        Me.chkCreditReduction = PRGlobal.Byte5
        
        If PRGlobal.Var1 = "" Then PRGlobal.Var1 = "0.00"
        PRGlobal.Var1 = Replace(PRGlobal.Var1, "%", "")
        Me.numCrRedPct.Value = CDec(PRGlobal.Var1)
        
        Me.chk4a = PRGlobal.Byte6
        Me.chk4b = PRGlobal.Byte7
        Me.chk4c = PRGlobal.Byte8
        Me.chk4d = PRGlobal.Byte9
        Me.chk4e = PRGlobal.Byte10
        
    Else
        PRGlobal.Clear
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Description = "FUTA940"
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    SQLString = " SELECT * " & _
                " FROM PRGlobal " & _
                " WHERE UserID = " & PRCompany.CompanyID & _
                " AND Description = 'FUTA940-B' "
    If PRGlobal.GetBySQL(SQLString) Then
        Me.chkLine9 = PRGlobal.Byte1
        Me.chkCrRedPct = PRGlobal.Byte2
        
        On Error Resume Next
        Me.numCrRedPct.Value = nNull(CDec(PRGlobal.Var1))
        On Error GoTo 0
    
    Else
        PRGlobal.Clear
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Description = "FUTA940-B"
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    Me.numCrRedPct.Visible = Me.chkCrRedPct
    
    Me.dtePPDate.Value = Now()
    Me.dteSignDate.Value = Now()
    
    GetPRData

    InitFlag = False
    
    CalcForm
    
    ' cmdPrint_Click
    
    
End Sub

Private Sub GetPRData()

    Dim progr As New frmProgress
    progr.lblMsg1 = "Now calculating FUTA Taxes ..."
    progr.Show
    
    ' max wages
    FedUnempMax = PRGlobal.GetAmount(PREquate.GlobalTypeFUNMax, frmPRQtrlyRpts.cmbYear)
    FedUnempPct = PRGlobal.GetAmount(PREquate.GlobalTypeFUNPct, CLng(Me.cmbTaxYear.Text))
    Me.num8.Caption = "8 FUTA tax before adjustments (line 7 x " & FedUnempPct / 100 & " = line8)"
    
    ' how many states in PRHist for the year?
    ' get the number of different states in PRHist
    Dim StartYM, EndYM As Long
    StartYM = CLng(Me.cmbTaxYear.Text) * 100 + 1
    EndYM = CLng(Me.cmbTaxYear.Text) * 100 + 12
    SQLString = "SELECT DISTINCT(StateID) as StID FROM PRHist WHERE PRHist.YearMonth BETWEEN " & StartYM & " AND " & EndYM
    rsInit SQLString, cn, rsState
    If rsState.RecordCount = 0 Then
        MsgBox "No PR History records found!", vbExclamation
        GoBack
    End If
    
    If rsState.RecordCount = 1 Then
        rsState.MoveFirst
        SQLString = " SELECT * " & _
                    " FROM PRState " & _
                    " WHERE StateID = " & rsState!StID
        If PRState.GetBySQL(SQLString) Then
            State1 = PRState.StateAbbrev
        Else
            State1 = ""
        End If
    Else
        State1 = ""
    End If
    Me.cmbState.Text = State1
    
    ' loop the history by EE
    Dim TlPmts(4), TlExempt(4), TlOver(4) As Currency
    SQLString = " SELECT * " & _
                " FROM PRHist " & _
                " WHERE PRHist.YearMonth BETWEEN " & StartYM & " AND " & EndYM & _
                " ORDER BY PRHist.EmployeeID, PRHist.CheckDate "
    If PRHist.GetBySQL(SQLString) Then
        
        Dim rct, Recs As Long
        rct = 0
        Recs = PRHist.Records
        Do
            
            rct = rct + 1
            If rct Mod 100 = 1 Then
                progr.lblMsg2 = "On record: " & rct & " of: " & Recs
                progr.Refresh
            End If
            
            ' determine the quarter
            Dim qtr As Integer
            Dim MM As Integer
            MM = PRHist.YearMonth Mod 100
            qtr = Int((MM - 1) / 3) + 1
            
            ' accum values per employee and quarter
            TlPmts(qtr) = TlPmts(qtr) + PRHist.Gross
            TlPmts(0) = TlPmts(0) + PRHist.Gross
            
            If PRHist.FUNWageBase <> PRHist.Gross Then
                TlExempt(qtr) = TlExempt(qtr) + PRHist.Gross - PRHist.FUNWageBase
                TlExempt(0) = TlExempt(0) + PRHist.Gross - PRHist.FUNWageBase
            End If
            
            ' determine overage
            TlOver(qtr) = TlOver(qtr) + PRHist.FUNWageBase - PRHist.FUNWage
            TlOver(0) = TlOver(0) + PRHist.FUNWageBase - PRHist.FUNWage
            
            If Not PRHist.GetNext Then Exit Do
        
        Loop
    
    Else
        MsgBox "No PR History found!", vbExclamation
        GoBack
    End If
    
    ' non taxable wage taken from box 3
    ' box 4 hard coded to 0
    Me.num3.Value = TlPmts(0) - TlExempt(0)
    Me.num4.Value = 0
    
    Me.num5.Value = TlOver(0)
    
    ' >>>> multiply by pct for tax liability
    ' >>> suppress if < $ 500
    QtrWage(0) = 0
    For i = 1 To 4
        QtrWage(i) = TlPmts(i) - TlExempt(i) - TlOver(i)
        QtrWage(0) = QtrWage(0) + QtrWage(i)
    Next i
    
    Me.num16a.Value = Round(QtrWage(1) * FedUnempPct / 100, 2)
    Me.num16b.Value = Round(QtrWage(2) * FedUnempPct / 100, 2)
    Me.num16c.Value = Round(QtrWage(3) * FedUnempPct / 100, 2)
    Me.num16d.Value = Round(QtrWage(4) * FedUnempPct / 100, 2)
    
    progr.Hide

End Sub

Private Sub CalcForm()

    If InitFlag Then Exit Sub

    FUTA_PCT = FedUnempPct / 100
    If Me.chkCrRedPct Then
        On Error Resume Next
        FUTA_PCT = FUTA_PCT + Me.numCrRedPct.Value / 100
        On Error GoTo 0
    End If
    
    Me.num6.Value = Me.num4.Value + Me.num5.Value
    Me.num7.Value = Me.num3.Value - Me.num6.Value
    Me.num8.Value = Round(Me.num7.Value * FedUnempPct / 100, 2)
    
    If Me.chkLine9 Then
        Me.num9.Value = Round(Me.num7.Value * 0.054, 2)
    Else
        Me.num9.Value = 0
    End If
    
    Me.num11.Value = Round(QtrWage(0) * Me.numCrRedPct / 100, 2)
    
    Me.num12.Value = Me.num8.Value + Me.num9.Value + Me.num10.Value + Me.num11.Value
    
    If (Me.num12.Value > Me.num13.Value) Then
        Me.num14.Value = Me.num12.Value - Me.num13.Value
        Me.num15.Value = 0
        InitFlag = True
        Me.chkRefund = 0
        Me.chkRefund.Enabled = False
        Me.chkApplyToNext = 0
        Me.chkApplyToNext.Enabled = False
        InitFlag = False
    Else
        Me.num14.Value = 0
        Me.num15.Value = Me.num13.Value - Me.num12.Value
        Me.chkRefund.Enabled = True
        Me.chkApplyToNext.Enabled = True
        If Me.chkRefund + Me.chkApplyToNext = 0 Then
            InitFlag = True
            Me.chkRefund = 1
            InitFlag = False
        End If
    End If
    
    Me.num17 = Me.num16a + Me.num16b + Me.num16c + Me.num16d
    
'    If Me.num17.Value <= QtrThreshold Then
'        Me.num16a.Visible = False
'        Me.num16b.Visible = False
'        Me.num16c.Visible = False
'        Me.num16d.Visible = False
'        Me.num17.Visible = False
'    Else
'        Me.num16a.Visible = True
'        Me.num16b.Visible = True
'        Me.num16c.Visible = True
'        Me.num16d.Visible = True
'        Me.num17.Visible = True
'    End If
    
    Me.lblUnemRate = FUTA_PCT
    
End Sub

Private Sub SaveSettings()

    ' save nudge
    HorzNudge = Me.numHorzNudge.Value
    VertNudge = Me.numVertNudge.Value
    
    SaveNudge User.ID, "Form940"
    
    PRGlobal.Var1 = ""
    PRGlobal.Var2 = ""
    PRGlobal.Var3 = ""
    PRGlobal.Var4 = ""
    PRGlobal.Var5 = ""
    PRGlobal.Var6 = ""
    PRGlobal.Var7 = ""
    
    ' save check box settings
    SQLString = " SELECT * " & _
                " FROM PRGlobal " & _
                " WHERE UserID = " & PRCompany.CompanyID & _
                " AND Description = 'FUTA940' "
    If PRGlobal.GetBySQL(SQLString) Then
        
        PRGlobal.Byte1 = IIf(Me.chkTypeA, 1, 0)
        PRGlobal.Byte2 = IIf(Me.chkTypeB, 1, 0)
        PRGlobal.Byte3 = IIf(Me.chkTypeC, 1, 0)
        PRGlobal.Byte4 = IIf(Me.chkTypeD, 1, 0)
        
        PRGlobal.Byte5 = IIf(Me.chkCreditReduction, 1, 0)
        PRGlobal.Var1 = Me.numCrRedPct.Text
        
        PRGlobal.Byte6 = IIf(Me.chk4a, 1, 0)
        PRGlobal.Byte7 = IIf(Me.chk4b, 1, 0)
        PRGlobal.Byte8 = IIf(Me.chk4c, 1, 0)
        PRGlobal.Byte9 = IIf(Me.chk4d, 1, 0)
        PRGlobal.Byte10 = IIf(Me.chk4e, 1, 0)
        
        PRGlobal.Save (Equate.RecPut)
    
    End If
    
    SQLString = " SELECT * " & _
                " FROM PRGlobal " & _
                " WHERE UserID = " & PRCompany.CompanyID & _
                " AND Description = 'FUTA940-B' "
    If PRGlobal.GetBySQL(SQLString) Then
        PRGlobal.Byte1 = IIf(Me.chkLine9, 1, 0)
        PRGlobal.Byte2 = IIf(Me.chkCrRedPct, 1, 0)
        PRGlobal.Save (Equate.RecPut)
    End If
    
    ' Part 4 - Third Party Designee - Per User
    If ThirdPartyID <> 0 Then
        If PRGlobal.GetByID(ThirdPartyID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part4
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
           
    PRGlobal.Var1 = Me.txt3rdPartyName
    PRGlobal.Var2 = Me.txt3rdPartyPhone
    PRGlobal.Var3 = Me.txt3rdPartyPIN
    PRGlobal.Var4 = Me.txtTradeName
    PRGlobal.Save (Equate.RecPut)
    
    If Me.chk3rdPartyYes Then
'        VertPosn = 8225
'        PosPrint 5210, VertPosn, PRGlobal.Var1
'        PosPrint 9060, VertPosn, PRGlobal.Var2
'
'        HorzPosn = 8475
'        Xincr = 445
'        VertPosn = 8695
'
'        ' part 4 PIN in gay boxes
'        X = Trim(PRGlobal.Var3)
'        For i = 1 To 5
'            If Len(X) >= i Then
'                PosPrint HorzPosn, VertPosn, Mid(X, i, 1)
'            End If
'            HorzPosn = HorzPosn + Xincr
'        Next i
        
    End If
           
    ' Part 5 - Company Signature - Per Company
    If SignID <> 0 Then
        If PRGlobal.GetByID(SignID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941Part5
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If

    PRGlobal.Var1 = Me.txtSignNameTitle
    PRGlobal.Save (Equate.RecPut)

    ' >>>>>>>>>>>>>>>>>>>>>
    
    'Paid Preparer - Per User
    If PaidPrepID <> 0 Then
        If PRGlobal.GetByID(PaidPrepID) Then
        End If
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalType941PaidPrep
        PRGlobal.UserID = User.ID
        PRGlobal.Save (Equate.RecAdd)
    End If
               
    PRGlobal.Var1 = Me.txtPPFirmName
    PRGlobal.Var2 = Me.txtPPAddr
    PRGlobal.Var3 = Me.txtPPCityState
    PRGlobal.Var4 = Me.txtPPPhone
    PRGlobal.Var7 = Me.txtPPPTIN
    PRGlobal.Var6 = Me.txtPPZip
    PRGlobal.Var5 = Me.txtPPEIN
    If Me.chkPPSE Then
        PRGlobal.Var8 = "1"
    Else
        PRGlobal.Var8 = "0"
    End If
    PRGlobal.Var9 = Me.cmbPPName
    PRGlobal.Save (Equate.RecPut)

End Sub

Private Sub cmdPrint_Click()

    ' make sure lines 12 and 17 match
    If Me.num12.Value <> Me.num17.Value Then
        Dim resp As Integer
        Dim diff As Double
        diff = Round(Me.num12.Value - Me.num17.Value, 2)

        X = "Box 12 = " + Me.num12.Text + vbCr + _
            "Box 17 = " + Me.num17.Text + vbCr + _
            "Off By:  " + CStr(diff) + vbCr + vbCr + _
            "These values must be the same!" + vbCr + _
            "OK to print or cancel?"
        resp = MsgBox(X, vbOKCancel + vbExclamation, "FUTA 940 Print")
        If resp = vbCancel Then Exit Sub
    End If

    SaveSettings

    HorzNudge = Me.numHorzNudge.Value
    VertNudge = Me.numVertNudge.Value

    Dim ReportTitle As String
    ReportTitle = "FUTA 940 Form"
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    MaxLines = 65

    Dim Col1X, Col2X, XXpos As Integer
    Dim HorzPosn, VertPosn, Xincr As Integer
    
    Col1X = 4830
    Col2X = 6875
    
    ' PosPrint 3380, 900, PRCompany.FederalID
    ' formatting for the GAY fed id boxes
    Dim FedID, ff As String
    HorzPosn = 2200
    Xincr = 505
    FedID = Trim(Me.txtEIN)
    For XXpos = 1 To Len(FedID)
        ff = Mid(FedID, XXpos, 1)
        If ff <> "-" Then
            HorzPosn = HorzPosn + Xincr
            PosPrint HorzPosn, 970, ff
        End If
        If XXpos = 2 Then
            HorzPosn = HorzPosn + 315
        End If
    Next XXpos
    
    Dim vincr As Integer
    PosPrint 2390, 1480, Me.txtName
    PosPrint 1990, 1960, Me.txtTradeName
    PosPrint 1240, 2460, Trim(Me.txtAddr1) & " " & Trim(Me.txtAddr2)
    
    VertPosn = 3180
    PosPrint 1240, VertPosn, Me.txtCity
    PosPrint 5270, VertPosn, Me.txtState
    PosPrint 6050, VertPosn, Me.txtZip
    
    VertPosn = 3760
    PosPrint 1240, VertPosn, Me.txtForeignCountry
    PosPrint 3990, VertPosn, Me.txtForeignProv
    PosPrint 6000, VertPosn, Me.txtForeignZip
    
    ' type check boxes
    HorzPosn = 8195
    VertPosn = 1550
    vincr = 360
    If Me.chkTypeA Then PosPrint HorzPosn, VertPosn, "X"
    VertPosn = VertPosn + vincr
    If Me.chkTypeB Then PosPrint HorzPosn, VertPosn, "X"
    VertPosn = VertPosn + vincr
    If Me.chkTypeC Then PosPrint HorzPosn, VertPosn, "X"
    VertPosn = VertPosn + vincr
    If Me.chkTypeD Then PosPrint HorzPosn, VertPosn, "X"
    VertPosn = VertPosn + vincr
    
    ' Part 1
    If Not Me.chkMultiState Then
        PosPrint 8840, 5090, Mid(Me.cmbState.Text, 1, 1)
        PosPrint 9540, 5090, Mid(Me.cmbState.Text, 2, 1)
    End If
    
    HorzPosn = 8730
    If Me.chkMultiState Then PosPrint HorzPosn, 5580, "X"
    If Me.chkCreditReduction Then PosPrint HorzPosn, 5980, "X"
    
    ' Part 2
    Dim hp1, hp2, hp3 As Integer
    hp1 = 6290
    hp2 = 9250
    VertPosn = 6760
    vincr = 425
    
    PosPrint hp2, 6770, PadRight(DollarAndCents(Me.num3.Value), 15)
    PosPrint hp1, 7120, PadRight(DollarAndCents(Me.num4.Value), 15)
    
    ' check boxes
    hp1 = 2740
    hp2 = 5790
    hp3 = 8240
    VertPosn = 7500
    
    If Me.chk4a Then PosPrint hp1, VertPosn, "X"
    If Me.chk4c Then PosPrint hp2, VertPosn, "X"
    If Me.chk4e Then PosPrint hp3, VertPosn, "X"
    
    VertPosn = 7780
    If Me.chk4b Then PosPrint hp1, VertPosn, "X"
    If Me.chk4d Then PosPrint hp2, VertPosn, "X"
    
    hp1 = 6290
    hp2 = 9250
    
    PosPrint hp1, 8190, PadRight(DollarAndCents(Me.num5.Value), 15)
    PosPrint hp2, 8550, PadRight(DollarAndCents(Me.num6.Value), 15)
    PosPrint hp2, 9050, PadRight(DollarAndCents(Me.num7.Value), 15)
    PosPrint hp2, 9530, PadRight(DollarAndCents(Me.num8.Value), 15)
    
    ' part 3
    PosPrint hp2, 10360, PadRight(DollarAndCents(Me.num9.Value), 15)
    PosPrint hp2, 10960, PadRight(DollarAndCents(Me.num10.Value), 15)
    PosPrint hp2, 11410, PadRight(DollarAndCents(Me.num11.Value), 15)
    
    ' part 4
    PosPrint hp2, 12140, PadRight(DollarAndCents(Me.num12.Value), 15)
    PosPrint hp2, 12590, PadRight(DollarAndCents(Me.num13.Value), 15)
    PosPrint hp2, 13340, PadRight(DollarAndCents(Me.num14.Value), 15)
    PosPrint hp2, 13840, PadRight(DollarAndCents(Me.num15.Value), 15)

    VertPosn = 14100
    If Me.chkApplyToNext Then PosPrint 7615, VertPosn, "X"
    If Me.chkRefund Then PosPrint 9515, VertPosn, "X"
    
    FormFeed
    
    ' page 2 -------------------------------------------
    
    VertPosn = 780
    PosPrint 400, VertPosn, Me.txtName
    PosPrint 7800, VertPosn, Me.txtEIN
    
    ' part 5 - qtrly info
    HorzPosn = 7030
    VertPosn = 2112
    vincr = 490
    
    If Me.num17.Value > 500# Then
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Me.num16a.Value), 15)
        VertPosn = VertPosn + vincr
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Me.num16b.Value), 15)
        VertPosn = VertPosn + vincr
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Me.num16c.Value), 15)
        VertPosn = VertPosn + vincr
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Me.num16d.Value), 15)
        VertPosn = VertPosn + vincr
        PosPrint HorzPosn, VertPosn, PadRight(DollarAndCents(Me.num17), 15)
    End If
    
    ' part 6
    ' third party
    HorzPosn = 790
    If Me.chk3rdPartyYes Then
        PosPrint HorzPosn, 5350, "X"
        PosPrint 4810, 5350, Me.txt3rdPartyName
        PosPrint 8180, 5350, Me.txt3rdPartyPhone
    Else
        PosPrint HorzPosn, 6100, "X"
    End If
    
    ' third party pin
    Dim tppin, tpp As String
    HorzPosn = 8270
    Xincr = 535
    tppin = Trim(Me.txt3rdPartyPIN)
    For XXpos = 1 To Len(tppin)
        tpp = Mid(tppin, XXpos, 1)
        PosPrint HorzPosn, 5800, tpp
        HorzPosn = HorzPosn + Xincr
    Next XXpos
    
    ' part 7 - sign here
    PosPrint 7310, 7895, Trim(SlashSplit(Me.txtSignNameTitle, 1))
    PosPrint 7310, 8395, Trim(SlashSplit(Me.txtSignNameTitle, 2))
    
    VertPosn = 9000
    PosPrint 1970, VertPosn, Month(Me.dteSignDate.Value)
    PosPrint 2370, VertPosn, Day(Me.dteSignDate.Value)
    PosPrint 2870, VertPosn, Year(Me.dteSignDate.Value)
    PosPrint 8240, VertPosn - 150, Me.txtSignPhone
    
    ' paid preparer
    hp1 = 2600
    hp2 = 8600
        
    ' -- left side
    PosPrint hp1, 10600, Me.cmbPPName.Text
    PosPrint hp1, 11790, Me.txtPPFirmName
    PosPrint hp1, 12250, Me.txtPPAddr
    
    ' -- right side
    PosPrint hp2, 10600, Me.txtPPPTIN
    PosPrint hp2, 11790, Me.txtPPEIN
    PosPrint hp2, 12250, Me.txtPPPhone
    
    PosPrint hp2 + 50, 11100, Month(Me.dtePPDate.Value)
    PosPrint 9060, 11120, Day(Me.dtePPDate.Value)
    PosPrint 9600, 11100, Year(Me.dtePPDate.Value)
    
    PosPrint 2500, 12670, Trim(SlashSplit(Me.txtPPCityState, 1))
    PosPrint 6000, 12670, Trim(SlashSplit(Me.txtPPCityState, 2))
    PosPrint hp2, 12670, Me.txtPPZip
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If InitFlag = False Then CalcForm
End Sub
Private Sub chkLine9_Click()
    CalcForm
End Sub
Private Sub chkCrRedPct_Click()
    If InitFlag Then Exit Sub
    Me.numCrRedPct.Visible = Me.chkCrRedPct
    If Me.chkCrRedPct = 0 Then
        Me.numCrRedPct.Value = 0
    End If
    CalcForm
End Sub

Private Sub cmbTaxYear_Click()
    GetPRData
    CalcForm
End Sub

Private Sub cmdExit_Click()
   GoBack
End Sub
Private Sub chkApplyToNext_Click()
    If InitFlag Then Exit Sub
    Me.chkRefund = IIf(Me.chkApplyToNext, 0, 1)
End Sub
Private Sub chkRefund_Click()
    If InitFlag Then Exit Sub
    Me.chkApplyToNext = IIf(Me.chkRefund, 0, 1)
End Sub
Private Sub chk3rdPartyNo_Click()
    If InitFlag Then Exit Sub
    Me.chk3rdPartyYes = IIf(Me.chk3rdPartyNo, 0, 1)
End Sub
Private Sub chk3rdPartyYes_Click()
    If InitFlag Then Exit Sub
    Me.chk3rdPartyNo = IIf(Me.chk3rdPartyYes, 0, 1)
End Sub
Private Sub cmdReload_Click()
    InitFlag = True
    GetPRData
    InitFlag = False
    CalcForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GoBack
End Sub
