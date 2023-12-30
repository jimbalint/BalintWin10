VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Begin VB.Form frmCompany 
   Caption         =   "Company Maintenance"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
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
   ScaleHeight     =   10395
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   41
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE AND EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   40
      Top             =   9720
      Width           =   2535
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   120
      TabIndex        =   69
      Top             =   600
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "MAIN"
      TabPicture(0)   =   "frmCompany.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblState"
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(2)=   "txtFileName"
      Tab(0).Control(3)=   "txtFederalID"
      Tab(0).Control(4)=   "txtCity"
      Tab(0).Control(5)=   "txtAddress2"
      Tab(0).Control(6)=   "txtAddress1"
      Tab(0).Control(7)=   "txtCompanyName"
      Tab(0).Control(8)=   "txtStateID"
      Tab(0).Control(9)=   "lngZipCode"
      Tab(0).Control(10)=   "tdbnumSUNPct"
      Tab(0).Control(11)=   "tdbtxtWkcPolicyNum"
      Tab(0).Control(12)=   "tdbtxtPhoneNumber"
      Tab(0).Control(13)=   "cmbState"
      Tab(0).Control(14)=   "tdbtxtStateUnempID"
      Tab(0).Control(15)=   "tdbtxtComment"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "ITEM"
      TabPicture(1)   =   "frmCompany.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "Label13"
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(7)=   "tdbnumAmtPct"
      Tab(1).Control(8)=   "tdbnumMaxPct"
      Tab(1).Control(9)=   "tdbnumMatchPct"
      Tab(1).Control(10)=   "txtAbbrev"
      Tab(1).Control(11)=   "fgERItem"
      Tab(1).Control(12)=   "chkNoSSTax"
      Tab(1).Control(13)=   "chkNoMedTax"
      Tab(1).Control(14)=   "chkNoFWTTax"
      Tab(1).Control(15)=   "chkNoSWTTax"
      Tab(1).Control(16)=   "chkNoCwtTax"
      Tab(1).Control(17)=   "chkNoSunTax"
      Tab(1).Control(18)=   "chkNoFuntax"
      Tab(1).Control(19)=   "chkTips"
      Tab(1).Control(20)=   "chkNotInNet"
      Tab(1).Control(21)=   "txtTitle"
      Tab(1).Control(22)=   "cmbGlAccount"
      Tab(1).Control(23)=   "cmdAddItem"
      Tab(1).Control(24)=   "chkActive"
      Tab(1).Control(25)=   "chkPension"
      Tab(1).Control(26)=   "chkEscrow"
      Tab(1).Control(27)=   "cmdDeleteItem"
      Tab(1).Control(28)=   "cmbType"
      Tab(1).Control(29)=   "cmbW2Box12"
      Tab(1).Control(30)=   "cmbW2Box14"
      Tab(1).Control(31)=   "tdbtxtItemComment"
      Tab(1).Control(32)=   "chkDirDepRpt"
      Tab(1).Control(33)=   "chkSickPay"
      Tab(1).Control(34)=   "cmdBasis"
      Tab(1).Control(35)=   "cmdItemUpdate"
      Tab(1).Control(36)=   "cmbBasis"
      Tab(1).Control(37)=   "cmbRateDiff"
      Tab(1).Control(38)=   "cmbOECity"
      Tab(1).Control(39)=   "cmdApplyToAll"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "PAY INFORMATION"
      TabPicture(2)   =   "frmCompany.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(2)=   "Label10"
      Tab(2).Control(3)=   "tdbnumDfltOTRate"
      Tab(2).Control(4)=   "tdbnumDfltRegHrs"
      Tab(2).Control(5)=   "tdbnumDfltMinWage"
      Tab(2).Control(6)=   "tdbnumCheckDays"
      Tab(2).Control(7)=   "tdbnumLastChkNum"
      Tab(2).Control(8)=   "cmbPPY"
      Tab(2).Control(9)=   "cmbCity"
      Tab(2).Control(10)=   "cmbSortOrder"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "DIRECT DEPOSIT SETUP"
      TabPicture(3)   =   "frmCompany.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(2)=   "tdbtxtBankFraction"
      Tab(3).Control(3)=   "tdbtxtBankAddress2"
      Tab(3).Control(4)=   "tdbtxtBankAddress1"
      Tab(3).Control(5)=   "tdbtxtBankAccount"
      Tab(3).Control(6)=   "tdbtxtBankABA"
      Tab(3).Control(7)=   "tdbtxtBankName"
      Tab(3).Control(8)=   "chkDirDepBalanced"
      Tab(3).Control(9)=   "txtDirDepFolder"
      Tab(3).Control(10)=   "txtDirDepHeader"
      Tab(3).Control(11)=   "chkDirDepUseAltID"
      Tab(3).Control(12)=   "tdbDirDepAltID"
      Tab(3).Control(13)=   "chkDirDepID1"
      Tab(3).Control(14)=   "tdbBatchHeader"
      Tab(3).ControlCount=   15
      TabCaption(4)   =   "Fed/Ohio W2 Upload"
      TabPicture(4)   =   "frmCompany.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label15"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label16"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label17"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label18"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "chkOHeW2"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "chkTermBiz"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtEstablishmentNumber"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txtOtherEIN"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "cmbKindOfEmployer"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "cmbEmploymentCode"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "chkThirdPartySickPay"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtContactName"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "txtContactPhoneNum"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "txtContactPhoneExt"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "txtContactFaxNum"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "txtContactEmail"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).ControlCount=   16
      Begin TDBText6Ctl.TDBText txtContactEmail 
         Height          =   375
         Left            =   3240
         TabIndex        =   102
         Top             =   7560
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":008C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0110
         Key             =   "frmCompany.frx":012E
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
      Begin TDBText6Ctl.TDBText txtContactFaxNum 
         Height          =   375
         Left            =   3240
         TabIndex        =   101
         Top             =   6960
         Width           =   6615
         _Version        =   65536
         _ExtentX        =   11668
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0172
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":01F0
         Key             =   "frmCompany.frx":020E
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
      Begin TDBText6Ctl.TDBText txtContactPhoneExt 
         Height          =   375
         Left            =   3240
         TabIndex        =   100
         Top             =   6360
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0252
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":02D6
         Key             =   "frmCompany.frx":02F4
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
      Begin TDBText6Ctl.TDBText txtContactPhoneNum 
         Height          =   375
         Left            =   3240
         TabIndex        =   99
         Top             =   5760
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0338
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":03BA
         Key             =   "frmCompany.frx":03D8
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
      Begin TDBText6Ctl.TDBText txtContactName 
         Height          =   375
         Left            =   3240
         TabIndex        =   98
         Top             =   5160
         Width           =   7215
         _Version        =   65536
         _ExtentX        =   12726
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":041C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":049E
         Key             =   "frmCompany.frx":04BC
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
         Format          =   "A9"
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
      Begin VB.CheckBox chkThirdPartySickPay 
         Caption         =   "Third Party Sick Pay"
         Height          =   375
         Left            =   3240
         TabIndex        =   97
         Top             =   4560
         Width           =   4575
      End
      Begin VB.ComboBox cmbEmploymentCode 
         Height          =   360
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   3960
         Width           =   5415
      End
      Begin VB.ComboBox cmbKindOfEmployer 
         Height          =   360
         ItemData        =   "frmCompany.frx":0500
         Left            =   5880
         List            =   "frmCompany.frx":0502
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   3360
         Width           =   5415
      End
      Begin TDBText6Ctl.TDBText txtOtherEIN 
         Height          =   375
         Left            =   3240
         TabIndex        =   92
         Top             =   2640
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":056A
         Key             =   "frmCompany.frx":0588
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
      Begin TDBText6Ctl.TDBText txtEstablishmentNumber 
         Height          =   375
         Left            =   3240
         TabIndex        =   90
         Top             =   2040
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":05CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0648
         Key             =   "frmCompany.frx":0666
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
         Format          =   "A9"
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
      Begin VB.CheckBox chkTermBiz 
         Caption         =   "Terminating Business Indicator"
         Height          =   375
         Left            =   3240
         TabIndex        =   89
         Top             =   1560
         Width           =   5895
      End
      Begin VB.CheckBox chkOHeW2 
         Caption         =   "Include in Fed/Ohio W2 Upload"
         Height          =   375
         Left            =   3240
         TabIndex        =   88
         Top             =   1080
         Width           =   3015
      End
      Begin TDBText6Ctl.TDBText tdbBatchHeader 
         Height          =   735
         Left            =   -74760
         TabIndex        =   87
         Top             =   5400
         Width           =   7695
         _Version        =   65536
         _ExtentX        =   13573
         _ExtentY        =   1296
         Caption         =   "frmCompany.frx":06AA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0718
         Key             =   "frmCompany.frx":0736
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
         Format          =   "A9#@"
         FormatMode      =   0
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   16
         LengthAsByte    =   0
         Text            =   "BATCHHEADER"
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
      Begin VB.CommandButton cmdApplyToAll 
         Caption         =   "APPLY TO ALL EMPLOYEES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -63360
         TabIndex        =   86
         Top             =   8160
         Width           =   1335
      End
      Begin VB.ComboBox cmbOECity 
         Height          =   360
         Left            =   -67920
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   8040
         Width           =   3855
      End
      Begin VB.ComboBox cmbRateDiff 
         Height          =   360
         Left            =   -65280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cmbBasis 
         Height          =   360
         Left            =   -68400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmdItemUpdate 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   -72840
         TabIndex        =   2
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CheckBox chkDirDepID1 
         Caption         =   """1"" Before Fed ID"
         Height          =   255
         Left            =   -67440
         TabIndex        =   63
         Top             =   1320
         Width           =   2655
      End
      Begin TDBNumber6Ctl.TDBNumber tdbDirDepAltID 
         Height          =   375
         Left            =   -67440
         TabIndex        =   65
         Top             =   2160
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":077A
         Caption         =   "frmCompany.frx":079A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":07FE
         Keys            =   "frmCompany.frx":081C
         Spin            =   "frmCompany.frx":0866
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
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin VB.CheckBox chkDirDepUseAltID 
         Caption         =   "Use Alternate Fed ID"
         Height          =   255
         Left            =   -67440
         TabIndex        =   64
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtDirDepHeader 
         Height          =   375
         Left            =   -74760
         TabIndex        =   67
         Text            =   "Text2"
         Top             =   4800
         Width           =   9975
      End
      Begin VB.TextBox txtDirDepFolder 
         Height          =   375
         Left            =   -67440
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   3120
         Width           =   4935
      End
      Begin VB.CheckBox chkDirDepBalanced 
         Caption         =   "Balanced Direct Deposit File"
         Height          =   375
         Left            =   -67440
         TabIndex        =   62
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdBasis 
         Caption         =   "BASIS FOR DEDUCTION"
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
         Left            =   -65280
         TabIndex        =   14
         Top             =   3960
         Width           =   2655
      End
      Begin VB.CheckBox chkSickPay 
         Caption         =   "3rd Party Sick Pay"
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
         Left            =   -63000
         TabIndex        =   28
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CheckBox chkDirDepRpt 
         Caption         =   "Direct Dep Rpt"
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
         Left            =   -66960
         TabIndex        =   24
         Top             =   6720
         Width           =   1695
      End
      Begin TDBText6Ctl.TDBText tdbtxtItemComment 
         Height          =   375
         Left            =   -69360
         TabIndex        =   29
         Top             =   7440
         Width           =   5295
         _Version        =   65536
         _ExtentX        =   9340
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":088E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":08F2
         Key             =   "frmCompany.frx":0910
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
      Begin TDBText6Ctl.TDBText tdbtxtComment 
         Height          =   1335
         Left            =   -72720
         TabIndex        =   54
         Top             =   4560
         Width           =   7455
         _Version        =   65536
         _ExtentX        =   13150
         _ExtentY        =   2355
         Caption         =   "frmCompany.frx":0954
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":09B8
         Key             =   "frmCompany.frx":09D6
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
         MultiLine       =   -1
         ScrollBars      =   2
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
      Begin VB.ComboBox cmbSortOrder 
         Height          =   360
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2400
         Width           =   3015
      End
      Begin TDBText6Ctl.TDBText tdbtxtStateUnempID 
         Height          =   375
         Left            =   -74760
         TabIndex        =   50
         Top             =   3480
         Width           =   5175
         _Version        =   65536
         _ExtentX        =   9128
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0A1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0A8C
         Key             =   "frmCompany.frx":0AAA
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
      Begin VB.ComboBox cmbW2Box14 
         Height          =   360
         Left            =   -64440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   6000
         Width           =   2775
      End
      Begin VB.ComboBox cmbW2Box12 
         Height          =   360
         Left            =   -64440
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   5160
         Width           =   2775
      End
      Begin VB.ComboBox cmbCity 
         Height          =   360
         Left            =   -68160
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   -64800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox cmbState 
         Height          =   360
         Left            =   -67680
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2040
         Width           =   1095
      End
      Begin TDBText6Ctl.TDBText tdbtxtPhoneNumber 
         Height          =   375
         Left            =   -68160
         TabIndex        =   53
         Top             =   4080
         Width           =   4935
         _Version        =   65536
         _ExtentX        =   8705
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0AEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0B5C
         Key             =   "frmCompany.frx":0B7A
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
      Begin VB.CommandButton cmdDeleteItem 
         Caption         =   "D&ELETE"
         Height          =   495
         Left            =   -71160
         TabIndex        =   3
         Top             =   8160
         Width           =   1215
      End
      Begin TDBText6Ctl.TDBText tdbtxtWkcPolicyNum 
         Height          =   375
         Left            =   -74760
         TabIndex        =   52
         Top             =   4080
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0BBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0C48
         Key             =   "frmCompany.frx":0C66
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
      Begin VB.ComboBox cmbPPY 
         Height          =   360
         Left            =   -70440
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3840
         Width           =   735
      End
      Begin VB.CheckBox chkEscrow 
         Caption         =   "Include deduction in Employer Escrow"
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
         Left            =   -69240
         TabIndex        =   13
         Top             =   4080
         Width           =   3375
      End
      Begin VB.CheckBox chkPension 
         Caption         =   "Pension"
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
         Left            =   -64320
         TabIndex        =   27
         Top             =   6720
         Width           =   1095
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Active"
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
         Left            =   -69120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "&ADD"
         Height          =   495
         Left            =   -74400
         TabIndex        =   1
         Top             =   8160
         Width           =   1215
      End
      Begin VB.ComboBox cmbGlAccount 
         Height          =   360
         Left            =   -62760
         TabIndex        =   30
         Top             =   7440
         Width           =   1695
      End
      Begin TDBText6Ctl.TDBText txtTitle 
         Height          =   375
         Left            =   -67680
         TabIndex        =   5
         Top             =   960
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0CAA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0D0A
         Key             =   "frmCompany.frx":0D28
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
      Begin VB.CheckBox chkNotInNet 
         Caption         =   "Not In Net Pay"
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
         Left            =   -66960
         TabIndex        =   23
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CheckBox chkTips 
         Caption         =   "Tips"
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
         Left            =   -66960
         TabIndex        =   22
         Top             =   6000
         Width           =   975
      End
      Begin VB.CheckBox chkNoFuntax 
         Caption         =   "No FUN Tax"
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
         Left            =   -66960
         TabIndex        =   20
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CheckBox chkNoSunTax 
         Caption         =   "No SUN Tax"
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
         Left            =   -66960
         TabIndex        =   21
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox chkNoCwtTax 
         Caption         =   "No CWT Tax"
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
         Left            =   -69240
         TabIndex        =   19
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CheckBox chkNoSWTTax 
         Caption         =   "No SWT Tax"
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
         Left            =   -69240
         TabIndex        =   18
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CheckBox chkNoFWTTax 
         Caption         =   "No FWT Tax"
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
         Left            =   -69240
         TabIndex        =   17
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CheckBox chkNoMedTax 
         Caption         =   "No Med Tax"
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
         Left            =   -69240
         TabIndex        =   16
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox chkNoSSTax 
         Caption         =   "No SS Tax"
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
         Left            =   -69240
         TabIndex        =   15
         Top             =   4920
         Width           =   1335
      End
      Begin VSFlex8Ctl.VSFlexGrid fgERItem 
         Height          =   7215
         Left            =   -74520
         TabIndex        =   0
         Top             =   720
         Width           =   4935
         _cx             =   8705
         _cy             =   12726
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumSUNPct 
         Height          =   375
         Left            =   -69240
         TabIndex        =   49
         Top             =   2880
         Width           =   2715
         _Version        =   65536
         _ExtentX        =   4789
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":0D6C
         Caption         =   "frmCompany.frx":0D8C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0E00
         Keys            =   "frmCompany.frx":0E1E
         Spin            =   "frmCompany.frx":0E68
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
         MaxValueVT      =   7208965
         MinValueVT      =   7274501
      End
      Begin TDBNumber6Ctl.TDBNumber lngZipCode 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   -65400
         TabIndex        =   47
         Top             =   2040
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":0E90
         Caption         =   "frmCompany.frx":0EB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":0F16
         Keys            =   "frmCompany.frx":0F34
         Spin            =   "frmCompany.frx":0F7E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   ""
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#########;-#########"
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
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   43450369
         Value           =   0
         MaxValueVT      =   5636101
         MinValueVT      =   3342341
      End
      Begin TDBText6Ctl.TDBText txtStateID 
         Height          =   375
         Left            =   -74760
         TabIndex        =   48
         Top             =   2880
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":0FA6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":100C
         Key             =   "frmCompany.frx":102A
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
      Begin TDBText6Ctl.TDBText txtCompanyName 
         Height          =   375
         Left            =   -74760
         TabIndex        =   42
         Top             =   600
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":106E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":10DC
         Key             =   "frmCompany.frx":10FA
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
      Begin TDBText6Ctl.TDBText txtAddress1 
         Height          =   375
         Left            =   -74760
         TabIndex        =   43
         Top             =   1080
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":113E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":11A6
         Key             =   "frmCompany.frx":11C4
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
      Begin TDBText6Ctl.TDBText txtAddress2 
         Height          =   375
         Left            =   -74760
         TabIndex        =   44
         Top             =   1560
         Width           =   7935
         _Version        =   65536
         _ExtentX        =   13996
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":1208
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1270
         Key             =   "frmCompany.frx":128E
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
         Left            =   -74760
         TabIndex        =   45
         Top             =   2040
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":12D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1330
         Key             =   "frmCompany.frx":134E
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
      Begin TDBText6Ctl.TDBText txtFederalID 
         Height          =   375
         Left            =   -69240
         TabIndex        =   51
         Top             =   3480
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":1392
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":13FC
         Key             =   "frmCompany.frx":141A
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
      Begin TDBText6Ctl.TDBText txtFileName 
         Height          =   375
         Left            =   -74760
         TabIndex        =   55
         Top             =   6000
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":145E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":14C6
         Key             =   "frmCompany.frx":14E4
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
      Begin TDBText6Ctl.TDBText txtAbbrev 
         Height          =   375
         Left            =   -69240
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":1528
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1596
         Key             =   "frmCompany.frx":15B4
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
      Begin TDBNumber6Ctl.TDBNumber tdbnumMatchPct 
         Height          =   375
         Left            =   -69240
         TabIndex        =   11
         Top             =   3360
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":15F8
         Caption         =   "frmCompany.frx":1618
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1680
         Keys            =   "frmCompany.frx":169E
         Spin            =   "frmCompany.frx":16E8
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   ""
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   ""
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
         ShowContextMenu =   1
         ValueVT         =   59310081
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumMaxPct 
         Height          =   375
         Left            =   -65280
         TabIndex        =   12
         Top             =   3360
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1710
         Caption         =   "frmCompany.frx":1730
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1794
         Keys            =   "frmCompany.frx":17B2
         Spin            =   "frmCompany.frx":17FC
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   ""
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   ""
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
         ShowContextMenu =   1
         ValueVT         =   59310081
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumAmtPct 
         Height          =   615
         Left            =   -63360
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   1085
         Calculator      =   "frmCompany.frx":1824
         Caption         =   "frmCompany.frx":1844
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":18A8
         Keys            =   "frmCompany.frx":18C6
         Spin            =   "frmCompany.frx":1910
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   ""
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   ""
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
         ShowContextMenu =   1
         ValueVT         =   59310081
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumLastChkNum 
         Height          =   375
         Left            =   -72720
         TabIndex        =   33
         Top             =   1440
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1938
         Caption         =   "frmCompany.frx":1958
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":19C6
         Keys            =   "frmCompany.frx":19E4
         Spin            =   "frmCompany.frx":1A2E
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
      Begin TDBNumber6Ctl.TDBNumber tdbnumCheckDays 
         Height          =   375
         Left            =   -72960
         TabIndex        =   32
         Top             =   840
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1A56
         Caption         =   "frmCompany.frx":1A76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1AFC
         Keys            =   "frmCompany.frx":1B1A
         Spin            =   "frmCompany.frx":1B64
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
         MaxValueVT      =   5242885
         MinValueVT      =   3014661
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumDfltMinWage 
         Height          =   375
         Left            =   -73080
         TabIndex        =   36
         Top             =   3240
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1B8C
         Caption         =   "frmCompany.frx":1BAC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1C24
         Keys            =   "frmCompany.frx":1C42
         Spin            =   "frmCompany.frx":1C8C
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;Null;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   47513601
         Value           =   0
         MaxValueVT      =   2686981
         MinValueVT      =   4784133
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumDfltRegHrs 
         Height          =   375
         Left            =   -72720
         TabIndex        =   34
         Top             =   2040
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1CB4
         Caption         =   "frmCompany.frx":1CD4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1D48
         Keys            =   "frmCompany.frx":1D66
         Spin            =   "frmCompany.frx":1DB0
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   ""
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   ""
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
         ShowContextMenu =   1
         ValueVT         =   47513601
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumDfltOTRate 
         Height          =   375
         Left            =   -72720
         TabIndex        =   35
         Top             =   2640
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   661
         Calculator      =   "frmCompany.frx":1DD8
         Caption         =   "frmCompany.frx":1DF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1E6C
         Keys            =   "frmCompany.frx":1E8A
         Spin            =   "frmCompany.frx":1ED4
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "$ ###,###.##;;Null;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "$ ###,###.##"
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
         ValueVT         =   47513601
         Value           =   0
         MaxValueVT      =   2686981
         MinValueVT      =   4784133
      End
      Begin TDBText6Ctl.TDBText tdbtxtBankName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   56
         Top             =   720
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":1EFC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":1F64
         Key             =   "frmCompany.frx":1F82
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
      Begin TDBText6Ctl.TDBText tdbtxtBankABA 
         Height          =   375
         Left            =   -74640
         TabIndex        =   57
         Top             =   1320
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":1FC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":202C
         Key             =   "frmCompany.frx":204A
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
      Begin TDBText6Ctl.TDBText tdbtxtBankAccount 
         Height          =   375
         Left            =   -74640
         TabIndex        =   58
         Top             =   1920
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":208E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":20FC
         Key             =   "frmCompany.frx":211A
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
      Begin TDBText6Ctl.TDBText tdbtxtBankAddress1 
         Height          =   375
         Left            =   -74640
         TabIndex        =   59
         Top             =   2520
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":215E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":21D0
         Key             =   "frmCompany.frx":21EE
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
      Begin TDBText6Ctl.TDBText tdbtxtBankAddress2 
         Height          =   375
         Left            =   -74640
         TabIndex        =   60
         Top             =   3120
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":2232
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":22A4
         Key             =   "frmCompany.frx":22C2
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
      Begin TDBText6Ctl.TDBText tdbtxtBankFraction 
         Height          =   375
         Left            =   -74640
         TabIndex        =   61
         Top             =   3720
         Width           =   5895
         _Version        =   65536
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "frmCompany.frx":2306
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCompany.frx":2376
         Key             =   "frmCompany.frx":2394
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
      Begin VB.Label Label18 
         Caption         =   "* = Required Field"
         Height          =   375
         Left            =   10680
         TabIndex        =   104
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Note: EIN && Phone #s are numeric only"
         Height          =   855
         Left            =   10680
         TabIndex        =   103
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Label16 
         Caption         =   "* Employment Code"
         Height          =   255
         Left            =   3240
         TabIndex        =   96
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "* Kind of Employer"
         Height          =   255
         Left            =   3240
         TabIndex        =   94
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "Default City:"
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
         Left            =   -69360
         TabIndex        =   85
         Top             =   8160
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Rate Difference:"
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
         Left            =   -66480
         TabIndex        =   84
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Basis:"
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
         Left            =   -69240
         TabIndex        =   83
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Direct Deposit Header:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   82
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Default Direct Deposit Folder/FileName:"
         Height          =   255
         Left            =   -67440
         TabIndex        =   81
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Label Label10 
         Caption         =   "Default Sort Order:"
         Height          =   255
         Left            =   -68160
         TabIndex        =   80
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "%   - (0.60% = 0.60)"
         Height          =   255
         Left            =   -66360
         TabIndex        =   79
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Default Pays Per Year:"
         Height          =   255
         Left            =   -73080
         TabIndex        =   78
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Default City:"
         Height          =   255
         Left            =   -68160
         TabIndex        =   76
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "W2 Box 14 Code"
         Height          =   255
         Left            =   -64440
         TabIndex        =   75
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "W2 Box 12 Code"
         Height          =   255
         Left            =   -64440
         TabIndex        =   74
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "GL Account:"
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
         Left            =   -63840
         TabIndex        =   72
         Top             =   7440
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -65520
         TabIndex        =   71
         Top             =   1740
         Width           =   615
      End
      Begin VB.Label lblState 
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -68440
         TabIndex        =   70
         Top             =   2100
         Width           =   735
      End
   End
   Begin TrueOleDBList80.TDBCombo TDBCombo1 
      Height          =   390
      Left            =   1920
      TabIndex        =   73
      Top             =   8640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      _DropdownWidth  =   0
      _EDITHEIGHT     =   688
      _GAPHEIGHT      =   53
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).AllowRowSizing=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      Appearance      =   1
      BorderStyle     =   1
      ComboStyle      =   0
      AutoCompletion  =   0   'False
      LimitToList     =   0   'False
      ColumnHeaders   =   -1  'True
      ColumnFooters   =   0   'False
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      FootLines       =   1
      RowDividerStyle =   0
      Caption         =   ""
      EditFont        =   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
      LayoutName      =   ""
      LayoutFileName  =   ""
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   0
      AutoSize        =   -1  'True
      ListField       =   ""
      BoundColumn     =   ""
      IntegralHeight  =   0   'False
      CellTipsWidth   =   0
      CellTipsDelay   =   1000
      AutoDropdown    =   0   'False
      RowTracking     =   -1  'True
      RightToLeft     =   0   'False
      RowMember       =   ""
      MouseIcon       =   0
      MouseIcon.vt    =   3
      MousePointer    =   0
      MatchEntryTimeout=   2000
      OLEDragMode     =   0
      OLEDropMode     =   0
      AnimateWindow   =   0
      AnimateWindowDirection=   0
      AnimateWindowTime=   200
      AnimateWindowClose=   0
      DropdownPosition=   0
      Locked          =   0   'False
      ScrollTrack     =   0   'False
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      AddItemSeparator=   ";"
      _PropDict       =   $"frmCompany.frx":23D8
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Arial"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TDBText6Ctl.TDBText TDBText2 
      Height          =   375
      Left            =   4200
      TabIndex        =   91
      Top             =   8400
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Caption         =   "frmCompany.frx":2482
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCompany.frx":24FE
      Key             =   "frmCompany.frx":251C
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
      FormatMode      =   1
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
   Begin VB.Label Label7 
      Caption         =   "ALL CHANGES WILL NOT BE SAVED !!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   10680
      TabIndex        =   77
      Top             =   9600
      Width           =   2535
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
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
      Left            =   1695
      TabIndex        =   68
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ERItem As New ADODB.Recordset
Dim rsState As New ADODB.Recordset
Dim drs As New ADODB.Recordset

Dim ActiveDrop As String
Dim TypeDrop As String
Dim I, J, K As Long
Dim x As String

Dim PREEItem As New cPRItem

Dim RecFlag, LoadFlag As Boolean

Private Sub cmdItemUpdate_Click()
    ItemUpdate
End Sub

Private Sub Form_Load()
    
    ' +------------------------------------------------------------
    ' | gather employer items
    ' +------------------------------------------------------------
    
    ' ReGet the PRCompany record as disconnected
    '   "All or Nothing" edit
    '   if maint screen is aborted - no changes are saved
    
    LoadFlag = True
    
    ' not necessary - using Class for PRCompany
    ' DisConn = True
    If Not PRCompany.GetByID(PRCompany.CompanyID) Then
        MsgBox "PRCompany Error: " & PRCompany.CompanyID, vbExclamation
        End
    End If
        
    ' get the PRItem records as disconnected also
    ' must set DisConn each time
    DisConn = True
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                "AND PRItem.ItemType > 2 ORDER BY PRItem.ItemType, PRItem.ItemID"
    rsInit SQLString, cn, ERItem
        
    SetGrid ERItem, Me.fgERItem
    
    ' Columns to hide
    fgERItem.ColHidden(0) = True        ' ItemID
    fgERItem.ColHidden(1) = True        ' EmployeeID
    fgERItem.ColHidden(3) = True        ' abbrev
    
    ' no horz scroll
    fgERItem.ScrollBars = flexScrollBarVertical
    
    ' column headers & widths
    fgERItem.TextMatrix(0, 2) = "Item Name"
    fgERItem.ColWidth(2) = 2700
    
    fgERItem.TextMatrix(0, 4) = "Item Type"
    fgERItem.ColWidth(4) = 1300
    
    fgERItem.TextMatrix(0, 5) = "Active"
    fgERItem.ColWidth(5) = 700
    
    ' select by row
    fgERItem.SelectionMode = flexSelectionByRow
    
    ' no edits in the flex grid
    fgERItem.Editable = flexEDNone
    
    ' +------------------------------------------------------------
    ' | gather employer items
    ' +------------------------------------------------------------

    ' set tdbText parameters
    tdbTextSet Me.txtCompanyName
    tdbTextSet Me.txtAddress1
    tdbTextSet Me.txtAddress2
    tdbTextSet Me.txtCity
    tdbTextSet Me.txtStateID
    tdbTextSet Me.txtFederalID
    tdbTextSet Me.txtFileName
    tdbTextSet Me.tdbtxtWkcPolicyNum
    tdbTextSet Me.tdbtxtPhoneNumber
    tdbTextSet Me.tdbtxtStateUnempID
    
    tdbTextSet Me.txtTitle
    tdbTextSet Me.txtAbbrev
    
    SQLString = "SELECT * FROM Notes WHERE NoteType = " & Equate.NoteTypeER & _
                " AND DateTm = 0 AND RelatedID = " & PRCompany.CompanyID
    If Notes.GetBySQL(SQLString) = False Then
        Notes.Clear
        Notes.NoteType = Equate.NoteTypeER
        Notes.DateTm = 0
        Notes.RelatedID = PRCompany.CompanyID
        Notes.Save (Equate.RecAdd)
    End If
    Me.tdbtxtComment = Notes.Notation
    
    tdbTextSet Me.tdbtxtItemComment
    Me.tdbtxtItemComment.MaxLength = 50
    
    ' set tdbNumber parameters
    tdbIntegerSet Me.tdbnumCheckDays
    tdbIntegerSet Me.tdbnumLastChkNum
    Me.lngZipCode.Format = "00000"
    Me.lngZipCode.DisplayFormat = ""
    
    tdbIntegerSet Me.lngZipCode
    
    ' set tdbNumber parameters - two decimal places
    tdbAmountSet Me.tdbnumSUNPct
    tdbAmountSet Me.tdbnumDfltMinWage
    tdbAmountSet Me.tdbnumDfltOTRate
    tdbAmountSet Me.tdbnumDfltRegHrs
    
    tdbAmountSet Me.tdbnumAmtPct
    tdbAmountSet Me.tdbnumMatchPct
    ' tdbAmountSet Me.tdbnumMaxAmt
    tdbAmountSet Me.tdbnumMaxPct
    
    ' +--------------------------------------------------
    ' load info for main tab
    ' +--------------------------------------------------
    Me.lblCompanyName = PRCompany.Name
    
    Me.txtCompanyName.text = PRCompany.Name
    Me.txtAddress1.text = PRCompany.Address1
    Me.txtAddress2.text = PRCompany.Address2
    Me.txtCity.text = PRCompany.City
    
    Me.lngZipCode.MinValue = 0
    Me.lngZipCode.MaxValue = 99999
    Me.lngZipCode.Format = "00000"
    Me.lngZipCode.DisplayFormat = ""
    ZipString = Format(PRCompany.ZipCode, "00000")
    Me.lngZipCode = Mid(PRCompany.ZipCode, 1, 5)
    Me.tdbtxtPhoneNumber = PRCompany.PhoneNumber
    
    Me.txtFederalID.text = PRCompany.FederalID
    Me.txtStateID.text = PRCompany.StateID
    Me.tdbnumSUNPct = PRCompany.StateUnempPct
    Me.tdbtxtStateUnempID = PRCompany.StateUnempID
    
    Me.tdbnumDfltOTRate = PRCompany.DfltOTRate
    Me.tdbnumDfltRegHrs = PRCompany.DfltRegHrs
    Me.tdbnumDfltMinWage = PRCompany.DfltMinWage
    Me.tdbnumCheckDays = PRCompany.CheckDays
    
    ' no comma for check number
    Me.tdbnumLastChkNum.Format = "#######0;(#######0)"
    Me.tdbnumLastChkNum.DisplayFormat = "#######0;(#######0);0"
    Me.tdbnumLastChkNum.HighlightText = True
    Me.tdbnumLastChkNum.Key.Clear = ""
    Me.tdbnumLastChkNum = PRCompany.LastCheckNum
    
    Me.txtFileName.text = PRCompany.FileName
    Me.txtFileName.Enabled = False
    
    Me.tdbtxtWkcPolicyNum = PRCompany.WkcPolicyNum
    
    ' +--------------------------------------------------
    ' load info for main tab
    ' +--------------------------------------------------
    
    ' ***********************************************************************
    ' +--------------------------------------------------
    ' load dir dep info
    ' +--------------------------------------------------
    tdbTextSet Me.tdbtxtBankName
    tdbTextSet Me.tdbtxtBankABA
    tdbTextSet Me.tdbtxtBankAccount
    tdbTextSet Me.tdbtxtBankAddress1
    tdbTextSet Me.tdbtxtBankAddress2
    tdbTextSet Me.tdbtxtBankFraction
    
    Me.tdbtxtBankName = PRCompany.BankName
    Me.tdbtxtBankABA = PRCompany.BankABA
    Me.tdbtxtBankAccount = PRCompany.BankAccount
    Me.tdbtxtBankAddress1 = PRCompany.BankAddr1
    Me.tdbtxtBankAddress2 = PRCompany.BankAddr2
    Me.tdbtxtBankFraction = PRCompany.BankFraction
    
    Me.chkDirDepBalanced = PRCompany.DirDepBalanced
    Me.chkDirDepUseAltID = PRCompany.DirDepUseAltID
    Me.chkDirDepID1 = PRCompany.DirDepID1
    
    With Me.tdbDirDepAltID
        .Format = "000000000"
        .DisplayFormat = ""
        .MinValue = 0
        .MaxValue = 999999999
    End With
    
    Me.tdbDirDepAltID = nNull(PRCompany.DirDepAltID)
    
    If Me.chkDirDepUseAltID = 0 Then
        Me.tdbDirDepAltID.Visible = False
    End If
    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDirDepFolder & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        Me.txtDirDepFolder = PRGlobal.Var1 & ""
        Me.txtDirDepHeader = PRGlobal.Var2 & ""
        Me.tdbBatchHeader.text = PRGlobal.Var3 & ""
    Else
        Me.txtDirDepFolder = ""
        Me.txtDirDepHeader = ""
        Me.tdbBatchHeader.text = ""
    End If
    
    ' +--------------------------------------------------
    ' load info for OH eW2 Tab
    ' +--------------------------------------------------
    Me.chkOHeW2 = PRCompany.OHeW2
    Me.chkTermBiz = PRCompany.TermBiz
    Me.txtEstablishmentNumber.text = PRCompany.EstablishmentNumber
    Me.txtOtherEIN.text = PRCompany.OtherEIN
    Me.chkThirdPartySickPay = PRCompany.ThirdPartySickPay
    Me.txtContactName.text = PRCompany.ContactName
    Me.txtContactPhoneNum.text = PRCompany.ContactPhoneNum
    Me.txtContactPhoneExt.text = PRCompany.ContactPhoneExt
    Me.txtContactFaxNum.text = PRCompany.ContactFasNum
    Me.txtContactEmail.text = PRCompany.ContactEmail
    ' +--------------------------------------------------
    ' load info for OH eW2 Tab
    ' +--------------------------------------------------
    
    
    ' initialize all of the drop downs
    DropDownInit
    
    LoadFlag = False
    
    ' load the screen info for the first item record
    ItemDisplay
    
    Me.SSTab1.Tab = 0

    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True

End Sub

Private Sub cmdAddItem_Click()
    
    LoadFlag = True
    
    ERItem.AddNew
    ERItem!EmployeeID = 0
    ERItem!Active = 1
    ERItem!Title = ""
    ERItem!Abbreviation = ""
    
    ' default to OE
    Me.cmbType.ListIndex = 0
    
    ERItem!NoSSTax = 0
    ERItem!NoMedTax = 0
    ERItem!NoFWTTax = 0
    ERItem!NoSWTTax = 0
    ERItem!NoCWTTax = 0
    ERItem!NoFUNTax = 0
    ERItem!NoSUNTax = 0
    ERItem!Tips = 0
    ERItem!NotInNet = 0
    ERItem!Pension = 0
    ERItem!SickPay = 0
    ERItem!W2Box12Code = 0
    ERItem!W2Box14Code = 0
    ERItem!MatchPct = 0
    ERItem!MaxPct = 0
    ERItem!MaxAmount = 0
    ERItem!AmtPct = 0
    ERItem!GLAccount = 0
    ERItem!Escrow = 0
    
    With Me.cmbType
        ERItem!ItemType = .ItemData(.ListIndex)
    End With
    
    ERItem.Update
    Me.fgERItem.DataRefresh
    
    LoadFlag = False
    
    ItemDisplay
    
    Me.txtTitle.SetFocus

    Me.cmdBasis.ToolTipText = "SET EARNINGS TO INCLUDE AS BASIS FOR DEDUCTION BY PERCENT"

End Sub
Private Sub cmdDeleteItem_Click()

    ' does any history exist?
    If ERItem!ItemType = PREquate.ItemTypeOE Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.EmployerItemID = " & ERItem!ItemID
        If PRDist.GetBySQL(SQLString) Then
            MsgBox "Distribution data exists for this other earning" & vbCr & _
                   "Delete not allowed!", vbExclamation
            Exit Sub
        End If
    Else
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.EmployerItemID = " & ERItem!ItemID
        If PRItemHist.GetBySQL(SQLString) Then
            MsgBox "Detail data exists for this deduction" & vbCr & _
                   "Delete not allowed!", vbExclamation
            Exit Sub
        End If
    End If

    If MsgBox("OK to delete " & Trim(ERItem!Title), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    ERItem!Title = "* DELETED *"
    ERItem.Update
    
End Sub


Private Sub cmdCancel_Click()
    
'    If MsgBox("All changes will be lost - OK to exit?", vbQuestion + vbYesNo, "Employer Maintenance") = vbNo Then
'        Exit Sub
'    End If

    GoBack
    End

End Sub
Private Sub cmdSave_Click()
        
    ' verify data
    If ERItem.RecordCount > 0 Then
        If Me.cmbType.text = "" Then
            MsgBox "Please select an item type!", vbExclamation
            Me.cmbType.SetFocus
            Exit Sub
        End If
    End If
    
    ' verify OH eW2 data
    If Me.chkOHeW2 Then
        If Me.cmbKindOfEmployer.text = "" Then
            MsgBox "OH W2 Upload - Kind of Employer must be entered!", vbExclamation
            Me.cmbKindOfEmployer.SetFocus
            Exit Sub
        End If
        If Me.cmbEmploymentCode.text = "" Then
            MsgBox "OH W2 Upload - Employment Code must be entered!", vbExclamation
            Me.cmbEmploymentCode.SetFocus
            Exit Sub
        End If
        If Me.txtContactName.Caption = "" Then
            MsgBox "OH W2 Upload - Contact Name must be entered!", vbExclamation
            Me.txtContactName.SetFocus
            Exit Sub
        End If
        If Me.txtContactPhoneNum.Caption = "" Then
            MsgBox "OH W2 Upload - Contact Phn# must be entered!", vbExclamation
            Me.txtContactPhoneNum.SetFocus
            Exit Sub
        End If
        If Me.txtContactEmail = "" Then
            MsgBox "OH W2 Upload - Contact Email must be entered!", vbExclamation
            Me.txtContactEmail.SetFocus
            Exit Sub
        End If
    End If
    
    PRCompany.DfltCityID = Me.cmbCity.ItemData(Me.cmbCity.ListIndex)
    If PRCompany.DfltCityID > 0 Then
        If PRCity.GetByID(PRCompany.DfltCityID) Then
            PRCompany.DfltStateID = PRCity.StateID
        End If
    Else
        PRCompany.DfltStateID = 0
    End If
    
    ' save all of the info to PRItem from the disconnected record set
    ItemUpdate
    
    ' look for deleted items
    If ERItem.RecordCount > 0 Then
        LoadFlag = True     ' don't update screen variables to the recordset
        ERItem.MoveFirst
        Do
            If ERItem!Title = "* DELETED *" Then
                ' delete employee items
                SQLString = "DELETE * FROM PRItem WHERE PRItem.EmployerItemID = " & ERItem!ItemID
                rsInit SQLString, cn, drs
                ERItem.Delete
            End If
            ERItem.MoveNext
            If ERItem.EOF Then Exit Do
        Loop
    End If
    
    rsSave ERItem, cn
    
    ' save the PRCompany info
    PRCompany.Name = Me.txtCompanyName
    PRCompany.Address1 = Me.txtAddress1
    PRCompany.Address2 = Me.txtAddress2
    PRCompany.City = Me.txtCity
    PRCompany.ZipCode = Me.lngZipCode
    PRCompany.PhoneNumber = Me.tdbtxtPhoneNumber
    PRCompany.StateID = Me.txtStateID
    PRCompany.FederalID = Me.txtFederalID
    
    PRCompany.StateUnempPct = Me.tdbnumSUNPct
    PRCompany.StateUnempID = Me.tdbtxtStateUnempID
    
    PRCompany.DfltMinWage = Me.tdbnumDfltMinWage.Value
    
    PRCompany.DfltOTRate = Me.tdbnumDfltOTRate
    PRCompany.DfltRegHrs = Me.tdbnumDfltRegHrs
    
    SQLString = "SELECT * FROM Notes WHERE NoteType = " & Equate.NoteTypeER & _
                " AND DateTm = 0 AND RelatedID = " & PRCompany.CompanyID
    If Notes.GetBySQL(SQLString) = False Then
        Notes.Clear
        Notes.NoteType = Equate.NoteTypeER
        Notes.DateTm = 0
        Notes.RelatedID = PRCompany.CompanyID
        Notes.Save (Equate.RecAdd)
    End If
    Notes.Notation = Me.tdbtxtComment
    Notes.Save (Equate.RecPut)
    
    PRCompany.LastCheckNum = Me.tdbnumLastChkNum
    PRCompany.CheckDays = Me.tdbnumCheckDays
    
    ' assign from the drop down xArray's
    PRCompany.AddrStateID = Me.cmbState.ItemData(Me.cmbState.ListIndex)
        
    PRCompany.WkcPolicyNum = Me.tdbtxtWkcPolicyNum.text
    
    ' save dir dep info
    PRCompany.BankName = Me.tdbtxtBankName
    PRCompany.BankABA = Me.tdbtxtBankABA
    PRCompany.BankAccount = Me.tdbtxtBankAccount
    PRCompany.BankAddr1 = Me.tdbtxtBankAddress1
    PRCompany.BankAddr2 = Me.tdbtxtBankAddress2
    PRCompany.BankFraction = Me.tdbtxtBankFraction
    
    ' dflt pays per year
    PRCompany.DfltPaysPerYear = Me.cmbPPY.text
        
    ' dflt sort order
    PRCompany.DfltSortOrder = Me.cmbSortOrder.ListIndex
    
    PRCompany.DirDepBalanced = Me.chkDirDepBalanced
    
    PRCompany.DirDepUseAltID = Me.chkDirDepUseAltID
    PRCompany.DirDepAltID = Me.tdbDirDepAltID
    PRCompany.DirDepID1 = Me.chkDirDepID1
    
    ' dir dep info
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDirDepFolder & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeDirDepFolder
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    PRGlobal.Var1 = Me.txtDirDepFolder & ""
    PRGlobal.Var2 = Me.txtDirDepHeader & ""
    PRGlobal.Var3 = Me.tdbBatchHeader & ""
    PRGlobal.Save (Equate.RecPut)
    
    ' OH eW2 info
    PRCompany.OHeW2 = Me.chkOHeW2.Value
    PRCompany.TermBiz = Me.chkTermBiz
    PRCompany.EstablishmentNumber = Me.txtEstablishmentNumber
    PRCompany.OtherEIN = Me.txtOtherEIN
    PRCompany.KindOfEmployer = Left(Me.cmbKindOfEmployer.text, 1)
    PRCompany.EmploymentCode = Left(Me.cmbEmploymentCode, 1)
    PRCompany.ThirdPartySickPay = Me.chkThirdPartySickPay
    PRCompany.ContactName = Me.txtContactName
    PRCompany.ContactPhoneNum = Me.txtContactPhoneNum
    PRCompany.ContactPhoneExt = Me.txtContactPhoneExt
    PRCompany.ContactFasNum = Me.txtContactFaxNum
    PRCompany.ContactEmail = Me.txtContactEmail
    
    PRCompany.Save (Equate.RecPut)
    ' PRCompany.UpdateBatch
    
    GoBack
    Unload Me

End Sub

Private Sub fgERItem_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
        
    ItemDisplay

End Sub

Private Sub fgERItem_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)

    ItemUpdate

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
    End Select
    
End Sub
Private Sub DropDownInit()

    ' basis drop-down
    With Me.cmbBasis
        .AddItem "Amount"
        .AddItem "Percent"
        .AddItem "Hourly"
    End With
    
    ' rate difference
    With Me.cmbRateDiff
        .AddItem "N / A"
        .AddItem "Amount"
        .AddItem "Percent"
    End With

    ' make a drop down string to translate active = 0 / 1 to No / Yes
    ' used for flex grid display only
    ' no validation needed
    ActiveDrop = "|#0;No|#1;Yes"
    fgERItem.ColComboList(5) = ActiveDrop
    
    ' *********************************************************************************
    ' make a drop down string for item types
    ' used for flex grid and screen combo box
    
    ' skip regular and overtime
    I = 2
    
    Do
        I = I + 1
        x = ItemName(I)     ' modPRGlobal function returns item names
        
        If Mid(x, 1, 1) = "?" Then Exit Do
        
        ' make a string for the flex grid display
        TypeDrop = Trim(TypeDrop) & "|#" & I & ";" & x
        
        ' add to combo box
        With Me.cmbType
            .AddItem x
            .ItemData(.NewIndex) = I
        End With
    
    Loop
    fgERItem.ColComboList(4) = TypeDrop
    
    ' ***********************************************************************
    ' load state Combo Box
    
    I = PRState.Records
    If I = 0 Then
        MsgBox "PRState file empty!", vbExclamation
        End
    End If
    
    With Me.cmbState
        
        .AddItem ""
        .ItemData(.NewIndex) = 0
    
        SQLString = "SELECT * FROM PRState ORDER BY StateAbbrev"
        If Not PRState.GetBySQL(SQLString) Then
            ' ?
        End If
    
        Do
            .AddItem PRState.StateAbbrev
            .ItemData(.NewIndex) = PRState.StateID
            If Not PRState.GetNext Then Exit Do
        Loop
        
    End With
    cmbPoint Me.cmbState, PRCompany.AddrStateID
        
    ' ***********************************************************************
    ' Default City
    With Me.cmbCity
        
        .AddItem "NONE"
        .ItemData(.NewIndex) = 0
        SQLString = "SELECT * FROM PRCity ORDER BY CityName"
        If PRCity.GetBySQL(SQLString) Then
            Do
                x = PRCity.CityName
                If PRState.GetByID(PRCity.StateID) Then
                    x = Trim(x) & ", " & PRState.StateAbbrev
                Else
                    x = Trim(x) & ", ??"
                End If
                x = Trim(x) & "  " & Format(PRCity.CityRate / 100, "##0.00%")
                .AddItem x
                .ItemData(.NewIndex) = PRCity.CityID
                
                If Not PRCity.GetNext Then Exit Do
            Loop
        End If
        
    End With
    cmbPoint Me.cmbCity, PRCompany.DfltCityID
    
    With Me.cmbOECity
        
        .AddItem "NONE"
        .ItemData(.NewIndex) = 0
        SQLString = "SELECT * FROM PRCity ORDER BY CityName"
        If PRCity.GetBySQL(SQLString) Then
            Do
                x = PRCity.CityName
                If PRState.GetByID(PRCity.StateID) Then
                    x = Trim(x) & ", " & PRState.StateAbbrev
                Else
                    x = Trim(x) & ", ??"
                End If
                x = Trim(x) & "  " & Format(PRCity.CityRate / 100, "##0.00%")
                .AddItem x
                .ItemData(.NewIndex) = PRCity.CityID
                
                If Not PRCity.GetNext Then Exit Do
            Loop
        End If
        
    End With
    
    
    ' **********************************************************************************************
    ' W2 Box 12 init
    With Me.cmbW2Box12
        
        .AddItem ""
        .ItemData(.NewIndex) = 0
    
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW2Box12 & _
                    " ORDER BY Description"
        If PRGlobal.GetBySQL(SQLString) Then
            Do
                .AddItem PRGlobal.Description
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If Not PRGlobal.GetNext Then Exit Do
            Loop
        End If
    
    End With
    
    ' **********************************************************************************************
    ' W2 Box 14 init
    
    With Me.cmbW2Box14
        
        .AddItem ""
        .ItemData(.NewIndex) = 0
    
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeW2Box14 & _
                    " ORDER BY Description"
        If PRGlobal.GetBySQL(SQLString) Then
            Do
                .AddItem PRGlobal.Description
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If Not PRGlobal.GetNext Then Exit Do
            Loop
        End If
    
    End With

    ' pays per year Combo
    ' use regular Windows Combo - style #2
    ' user can not type in selection field
    ' setup in PRGlobal module - used by employee screen also
    cmbPPYSet cmbPPY, PRCompany.DfltPaysPerYear


    ' default sort order
    With Me.cmbSortOrder
        .AddItem "By EE Number"
        .AddItem "By EE Name"
        .AddItem "By Dept By EE#"
        .AddItem "By Dept By EE Name"
        .ListIndex = PRCompany.DfltSortOrder
    End With

    ' OH eW2 Combos
    With Me.cmbKindOfEmployer
        .AddItem "N None Apply"
        .AddItem "F Federal govt."
        .AddItem "S State/local non-501c"
        .AddItem "T 501c non-govt."
        .AddItem "Y State/local 501c."
        If PRCompany.KindOfEmployer = "N" Then .ListIndex = 0
        If PRCompany.KindOfEmployer = "F" Then .ListIndex = 1
        If PRCompany.KindOfEmployer = "S" Then .ListIndex = 2
        If PRCompany.KindOfEmployer = "T" Then .ListIndex = 3
        If PRCompany.KindOfEmployer = "Y" Then .ListIndex = 4
    End With
    With Me.cmbEmploymentCode
        .AddItem "R Regular Form 941"
        .AddItem "F Regular Form 944"
        .AddItem "A Agriculture Form 943"
        .AddItem "H Household Schedule H"
        .AddItem "M Military Form 941"
        .AddItem "Q Medicare Qual. Form 941"
        .AddItem "X Railroad Form 944"
        If PRCompany.EmploymentCode = "R" Then .ListIndex = 0
        If PRCompany.EmploymentCode = "F" Then .ListIndex = 1
        If PRCompany.EmploymentCode = "A" Then .ListIndex = 2
        If PRCompany.EmploymentCode = "H" Then .ListIndex = 3
        If PRCompany.EmploymentCode = "M" Then .ListIndex = 4
        If PRCompany.EmploymentCode = "Q" Then .ListIndex = 5
        If PRCompany.EmploymentCode = "X" Then .ListIndex = 6
    End With

End Sub


Private Sub txtTitle_Change()
    ERItem!Title = Trim(txtTitle)
End Sub

Private Sub chkActive_Click()
    ERItem!Active = chkActive
End Sub

Private Sub cmbType_Click()
    
    If LoadFlag = True Then Exit Sub
    
    ' update the grid display
    With Me.cmbType
        
        ERItem!ItemType = .ItemData(.ListIndex)
    
        If ERItem!ItemType = PREquate.ItemTypeOE Then
            
            If Me.cmbBasis.ListIndex <> 2 Then   ' hourly
                Me.cmbRateDiff.Enabled = False
            Else
                Me.cmbRateDiff.Enabled = True
            End If
            
            Me.tdbnumMatchPct.Enabled = False
            Me.tdbnumMaxPct.Enabled = False
            Me.chkEscrow.Enabled = False
            Me.cmdBasis.Enabled = False
        
        
        ElseIf ERItem!ItemType = PREquate.ItemTypeDED Then
        
            Me.tdbnumMatchPct.Enabled = True
            Me.tdbnumMaxPct.Enabled = True
            Me.chkEscrow.Enabled = True
            Me.cmdBasis.Enabled = True
            
            Me.cmbRateDiff.Enabled = False
            Me.cmbRateDiff.ListIndex = 0
            
        ElseIf ERItem!ItemType = PREquate.ItemTypeSDTax Then
        
            Me.chkEscrow.Enabled = True
        
        Else
            
            Me.tdbnumMatchPct.Enabled = False
            Me.tdbnumMaxPct.Enabled = False
            Me.chkEscrow.Enabled = False
            Me.cmdBasis.Enabled = False
            
            Me.cmbRateDiff.Enabled = False
            Me.cmbRateDiff.ListIndex = 0
    
        End If
    
    End With

    ' If GridFlag = True Then Exit Sub

'    With Me.cmbType
'
'        ' something must be selected
'        If .Text = "" Then
'            MsgBox "Please enter an item type", vbExclamation
'            .SetFocus
'            Exit Sub
'        End If
'
'        If IsNull(ERItem!ItemID) Then Exit Sub
'
'        ' not allowed if data exists
'        SQLString = "SELECT * from PRDist WHERE PRDist.EmployerItemID = " & ERItem!ItemID
'        If PRDist.GetBySQL(SQLString) Then
'            MsgBox "Type change not allowed when historical data exists!", vbExclamation
'            Me.cmbType.SetFocus
'            GoBack
'        End If
'
'        SQLString = "SELECT * from PRItemHist WHERE PRItemHist.EmployerItemID = " & ERItem!ItemID
'        If PRItemHist.GetBySQL(SQLString) Then
'            MsgBox "Type change not allowed when historical data exists!", vbExclamation
'            Me.cmbType.SetFocus
'            GoBack
'        End If
'
'
'    End With
    
End Sub

Private Sub ItemDisplay()

Dim SVal As Variant

    If ERItem.RecordCount = 0 Then Exit Sub
    If LoadFlag = True Then Exit Sub

    Me.txtTitle = ERItem!Title & ""
    Me.txtAbbrev = ERItem!Abbreviation & ""

    With Me.cmbBasis
        Select Case nNull(ERItem!Basis)
            Case PREquate.BasisAmount:  .ListIndex = 0
            Case PREquate.BasisPercent: .ListIndex = 1
            Case PREquate.BasisHourly:  .ListIndex = 2
            Case Else:                  .ListIndex = 0
        End Select
    End With

    With Me.cmbRateDiff
        Select Case ERItem!RateDifference
            Case PREquate.BasisAmount:  .ListIndex = 1
            Case PREquate.BasisPercent: .ListIndex = 2
            Case Else:                  .ListIndex = 0
        End Select
    End With

    cmbPoint Me.cmbType, ERItem!ItemType
    
    Me.chkNoSSTax = ERItem!NoSSTax
    Me.chkNoMedTax = ERItem!NoMedTax
    Me.chkNoFWTTax = ERItem!NoFWTTax
    Me.chkNoSWTTax = ERItem!NoSWTTax
    Me.chkNoCwtTax = ERItem!NoCWTTax
    Me.chkNoFuntax = ERItem!NoFUNTax
    Me.chkNoSunTax = ERItem!NoSUNTax
    
    Me.chkTips = nNull(ERItem!Tips)
    Me.chkNotInNet = nNull(ERItem!NotInNet)
    Me.chkActive = nNull(ERItem!Active)
    Me.chkEscrow = nNull(ERItem!Escrow)
    
    Me.chkPension = nNull(ERItem!Pension)
    Me.chkSickPay = nNull(ERItem!SickPay)
    
    Me.tdbnumMatchPct = nNull(ERItem!MatchPct)
    ' Me.tdbnumMaxAmt = nNull(ERItem!MaxAmount)
    Me.tdbnumMaxPct = nNull(ERItem!MaxPct)
    Me.tdbnumAmtPct = nNull(ERItem!AmtPct)
    Me.chkDirDepRpt = nNull(ERItem!DirDepRpt)
    
    cmbPoint Me.cmbW2Box12, ERItem!W2Box12Code
    cmbPoint Me.cmbW2Box14, ERItem!W2Box14Code
    
    Me.tdbtxtItemComment = ERItem!Comment & ""
    
    If ERItem!ItemType = PREquate.ItemTypeOE Then
        Me.chkEscrow.Enabled = False
        Me.cmdBasis.Enabled = False
        Me.tdbnumMatchPct.Enabled = False
        Me.tdbnumMaxPct.Enabled = False
        
        If ERItem!Basis = PREquate.BasisHourly Then
            Me.cmbRateDiff.Enabled = True
        Else
            Me.cmbRateDiff.Enabled = False
        End If
    
    End If
    
    If ERItem!ItemType = PREquate.ItemTypeDED Then
        Me.cmbRateDiff.Enabled = False
        
        Me.chkEscrow.Enabled = True
        Me.cmdBasis.Enabled = True
        Me.tdbnumMatchPct.Enabled = True
        Me.tdbnumMaxPct.Enabled = True
    End If
    
    Me.cmbOECity.ListIndex = -1
    If Not IsNull(ERItem!CityID) Then
        If ERItem!CityID <> 0 Then
            cmbPoint Me.cmbOECity, ERItem!CityID
        End If
    End If
    
    Me.Refresh

End Sub

Private Sub ItemUpdate()
    
    ' skip on initial load
    If LoadFlag = True Then Exit Sub
    If ERItem.RecordCount = 0 Then Exit Sub
    If ERItem!Title = "* DELETED *" Then Exit Sub
    
    ' ***************************************************************
    ' >>>>>>>>> did the item type change?
    If Not IsNull(ERItem!ItemID) Then
    
        If Not PRItem.GetByID(ERItem!ItemID) Then
        End If
        
        If PRItem.ItemType <> Me.cmbType.ItemData(Me.cmbType.ListIndex) Then
            
            ' not allowed if data exists
            SQLString = "SELECT * from PRDist WHERE PRDist.EmployerItemID = " & ERItem!ItemID
            If PRDist.GetBySQL(SQLString) Then
                MsgBox "Type change not allowed when historical data exists!", vbExclamation
                Me.cmbType.SetFocus
                GoBack
            End If
    
            SQLString = "SELECT * from PRItemHist WHERE PRItemHist.EmployerItemID = " & ERItem!ItemID
            If PRItemHist.GetBySQL(SQLString) Then
                MsgBox "Type change not allowed when historical data exists!", vbExclamation
                Me.cmbType.SetFocus
                GoBack
            End If
        
        End If
    End If
    
    '    change it back
    ' ***************************************************************
    
    With Me.cmbType
        If .text = "" Then
            MsgBox "Please select an item type!", vbExclamation
            .SetFocus
            Exit Sub
        Else
            ' add one - selected item is zero based and "NONE" is not an option
            ERItem!ItemType = .ItemData(.ListIndex)
        End If
    End With
    
    ERItem!Active = Me.chkActive
    ERItem!Title = Trim(Me.txtTitle & "")
    ERItem!Abbreviation = Trim(Me.txtAbbrev & "")
    
    ERItem!NoSSTax = Me.chkNoSSTax
    ERItem!NoMedTax = Me.chkNoMedTax
    ERItem!NoFWTTax = Me.chkNoFWTTax
    ERItem!NoSWTTax = Me.chkNoSWTTax
    ERItem!NoCWTTax = Me.chkNoCwtTax
    ERItem!NoFUNTax = Me.chkNoFuntax
    ERItem!NoSUNTax = Me.chkNoSunTax
    ERItem!Escrow = Me.chkEscrow
    ERItem!Pension = Me.chkPension
    ERItem!SickPay = Me.chkSickPay
    
    ERItem!MatchPct = Me.tdbnumMatchPct
    ' ERItem!MaxAmount = Me.tdbnumMaxAmt
    ERItem!MaxPct = Me.tdbnumMaxPct
    ERItem!AmtPct = Me.tdbnumAmtPct
    
    ERItem!Tips = Me.chkTips
    ERItem!NotInNet = Me.chkNotInNet
    ERItem!Comment = Trim(Me.tdbtxtItemComment & "")
    ERItem!DirDepRpt = Me.chkDirDepRpt
    
    ' blank is OK - assign to zero
    With Me.cmbW2Box12
        ERItem!W2Box12Code = .ItemData(.ListIndex)
    End With
    
    With Me.cmbW2Box14
        ERItem!W2Box14Code = .ItemData(.ListIndex)
    End With
    
    With Me.cmbBasis
        Select Case .ListIndex
            Case -1:        ERItem!Basis = PREquate.BasisAmount
            Case 0:         ERItem!Basis = PREquate.BasisAmount
            Case 1:         ERItem!Basis = PREquate.BasisPercent
            Case 2:         ERItem!Basis = PREquate.BasisHourly
        End Select
    End With
        
    With Me.cmbRateDiff
        Select Case .ListIndex
            Case -1:        ERItem!RateDifference = 0
            Case 0:         ERItem!RateDifference = 0
            Case 1:         ERItem!RateDifference = PREquate.BasisAmount
            Case 2:         ERItem!RateDifference = PREquate.BasisPercent
        End Select
    End With

    With Me.cmbOECity
        If .ListIndex > 0 Then
            ERItem!CityID = .ItemData(.ListIndex)
        Else
            ERItem!CityID = 0
        End If
    End With
    
    ERItem.Update

End Sub

Private Sub cmdBasis_Click()
    If Me.cmbType <> "Deduction" Then Exit Sub
    frmDeductBasis.EmployeeID = 0
    frmDeductBasis.ItemID = ERItem!ItemID
    frmDeductBasis.Show vbModal
End Sub

Private Sub chkDirDepUseAltID_Click()
    If Me.chkDirDepUseAltID = 0 Then
        Me.tdbDirDepAltID.Visible = False
    Else
        Me.tdbDirDepAltID.Visible = True
    End If
End Sub

Private Sub cmbBasis_Click()

    With Me.cmbBasis
        If ERItem!ItemType = PREquate.ItemTypeOE Then
            If .ListIndex = PREquate.BasisHourly Then
                Me.cmbRateDiff.Enabled = True
            Else
                Me.cmbRateDiff.Enabled = False
            End If
        Else
            Me.cmbRateDiff.ListIndex = 0
            Me.cmbRateDiff.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdApplyToAll_Click()

Dim EECount As Long

    x = "This item must be SAVED first!"
    
    If IsNull(ERItem!ItemID) Then
        MsgBox x, vbExclamation
        Exit Sub
    End If
    
    If ERItem!ItemID = 0 Then
        MsgBox x, vbExclamation
        Exit Sub
    End If
    
    x = "OK to apply " & ERItem!Title & vbCr & "to all ACTIVE employees?"
    If MsgBox(x, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    EECount = 0
    
    SQLString = "SELECT * FROM PREmployee WHERE Inactive = 0"
    If PREmployee.GetBySQL(SQLString) = True Then
        Do
            SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & PREmployee.EmployeeID & _
                        " AND EmployerItemID = " & ERItem!ItemID
            If PREEItem.GetBySQL(SQLString) = False Then
                PREEItem.Clear
                PREEItem.EmployeeID = PREmployee.EmployeeID
                PREEItem.EmployerItemID = ERItem!ItemID
                PREEItem.Title = ERItem!Title
                PREEItem.Abbreviation = ERItem!Abbreviation
                PREEItem.ItemType = ERItem!ItemType
                PREEItem.Active = 1
                PREEItem.UseEmployer = 1
                PREEItem.Basis = ERItem!Basis
                PREEItem.AmtPct = ERItem!AmtPct
                PREEItem.Save (Equate.RecAdd)
                EECount = EECount + 1
            End If
            If PREmployee.GetNext = False Then Exit Do
        Loop
    End If

    x = ERItem!Title & " has been added to " & EECount & " employees"
    MsgBox x, vbInformation

End Sub


