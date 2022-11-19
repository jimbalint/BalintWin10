VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEmpForm 
   Caption         =   "Employee Maintenance"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   10785
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   615
      Left            =   7800
      TabIndex        =   111
      Top             =   9960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   3480
      TabIndex        =   110
      Top             =   9960
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8700
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   15346
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MAIN (F2)"
      TabPicture(0)   =   "frmEmpForm.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtSSN"
      Tab(0).Control(1)=   "tdbtxtComment"
      Tab(0).Control(2)=   "txtAltName"
      Tab(0).Control(3)=   "txtCheckComment"
      Tab(0).Control(4)=   "cmbWkcCat"
      Tab(0).Control(5)=   "cmbDept"
      Tab(0).Control(6)=   "cmb1099"
      Tab(0).Control(7)=   "cmbEICType"
      Tab(0).Control(8)=   "chkUseDeptWkc"
      Tab(0).Control(9)=   "chkInactive"
      Tab(0).Control(10)=   "chkStatutory"
      Tab(0).Control(11)=   "cmbState"
      Tab(0).Control(12)=   "tdbnumZipCode"
      Tab(0).Control(13)=   "chkUseAltName"
      Tab(0).Control(14)=   "txtAddress1"
      Tab(0).Control(15)=   "lngEmployeeNumber"
      Tab(0).Control(16)=   "txtFirstName"
      Tab(0).Control(17)=   "txtLastName"
      Tab(0).Control(18)=   "txtAddress2"
      Tab(0).Control(19)=   "txtCity"
      Tab(0).Control(20)=   "txtMI"
      Tab(0).Control(21)=   "Label22"
      Tab(0).Control(22)=   "Label18"
      Tab(0).Control(23)=   "Label17"
      Tab(0).Control(24)=   "Label16"
      Tab(0).Control(25)=   "Label9"
      Tab(0).Control(26)=   "Label1"
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "PAY INFORMATION (F3)"
      TabPicture(1)   =   "frmEmpForm.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblEEDefaultJob"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "curSalaryAmt"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "curHourlyAmt"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkNoStateTax"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chkNoFedTax"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkNoSSTax"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkSalaried"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkNoCityTax"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkNoFedUnemp"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkNoStateUnemp"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkSWTMarried"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkFWTMarried"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkNoMedTax"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Frame1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Frame2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Frame3"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Frame4"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmbPPY"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chkCourtAdd"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmbEEDfltCity"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmbCourtCWT"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "cmbEEDefaultJob"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "cmbPRW4"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "chkW4_2Jobs"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "tdbnumW4Dependents"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "tdbnumW4OtherDependents"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "tdbnumW4OtherIncome"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "tdbnumW4Deductions"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "tdbnumW4ExtraWH"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).ControlCount=   35
      TabCaption(2)   =   "DATES AND OTHER INFORMATION (F4)"
      TabPicture(2)   =   "frmEmpForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(5)=   "Label4"
      Tab(2).Control(6)=   "dteDateLastRecall"
      Tab(2).Control(7)=   "dteDateLastLayoff"
      Tab(2).Control(8)=   "dteDateLastPaid"
      Tab(2).Control(9)=   "DteDateofBirth"
      Tab(2).Control(10)=   "dteDateLastReview"
      Tab(2).Control(11)=   "dteDateLastRaise"
      Tab(2).Control(12)=   "DteDateHired"
      Tab(2).Control(13)=   "dteDateTerminated"
      Tab(2).Control(14)=   "cmbTermReason"
      Tab(2).Control(15)=   "lngWorkCompNo"
      Tab(2).Control(16)=   "cmbSex"
      Tab(2).Control(17)=   "cmbShiftCode"
      Tab(2).Control(18)=   "cmbEducationLevel"
      Tab(2).Control(19)=   "cmbMaritalStatus"
      Tab(2).Control(20)=   "cmbRaceCode"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "OTHER EARNINGS AND DEDUCTIONS (F5)"
      TabPicture(3)   =   "frmEmpForm.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label10"
      Tab(3).Control(1)=   "lblOEDEDTitle"
      Tab(3).Control(2)=   "lblW212"
      Tab(3).Control(3)=   "lblW214"
      Tab(3).Control(4)=   "Line1"
      Tab(3).Control(5)=   "lblRateDiff"
      Tab(3).Control(6)=   "Label21"
      Tab(3).Control(7)=   "fgOEDED"
      Tab(3).Control(8)=   "chkActive"
      Tab(3).Control(9)=   "chkUseEmpDef"
      Tab(3).Control(10)=   "fraItmBasis"
      Tab(3).Control(11)=   "chkTips"
      Tab(3).Control(12)=   "chkNotNet"
      Tab(3).Control(13)=   "tdbnumAmtPct"
      Tab(3).Control(14)=   "tdbnumMaxAmt"
      Tab(3).Control(15)=   "cmdOEDEDAdd"
      Tab(3).Control(16)=   "cmdOEDEDDelete"
      Tab(3).Control(17)=   "chkOEDEDNoSSTax"
      Tab(3).Control(18)=   "chkOEDEDNoMedTax"
      Tab(3).Control(19)=   "chkOEDEDNoFWTTax"
      Tab(3).Control(20)=   "chkOEDEDNoSWTTax"
      Tab(3).Control(21)=   "chkOEDEDNoCWTTax"
      Tab(3).Control(22)=   "chkOEDEDNoFUNTax"
      Tab(3).Control(23)=   "chkOEDEDNoSUNTax"
      Tab(3).Control(24)=   "chkPension"
      Tab(3).Control(25)=   "cmbW2Box12"
      Tab(3).Control(26)=   "cmbW2Box14"
      Tab(3).Control(27)=   "tdbtxtItemComment"
      Tab(3).Control(28)=   "chkDirDepRpt"
      Tab(3).Control(29)=   "chkSickPay"
      Tab(3).Control(30)=   "cmdBasis"
      Tab(3).Control(31)=   "cmdItemUpdate"
      Tab(3).Control(32)=   "cmbRateDiff"
      Tab(3).Control(33)=   "cmbOECity"
      Tab(3).ControlCount=   34
      TabCaption(4)   =   "DIRECT DEPOSIT (F6)"
      TabPicture(4)   =   "frmEmpForm.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtAccount"
      Tab(4).Control(1)=   "txtABA"
      Tab(4).Control(2)=   "fgDirDep"
      Tab(4).Control(3)=   "txtBankName"
      Tab(4).Control(4)=   "chkDirDepActive"
      Tab(4).Control(5)=   "fraType"
      Tab(4).Control(6)=   "fraBasis"
      Tab(4).Control(7)=   "tdbnumDDAmount"
      Tab(4).Control(8)=   "cmdDirDepAdd"
      Tab(4).Control(9)=   "cmdDirDepDelete"
      Tab(4).Control(10)=   "cmdDDUpdate"
      Tab(4).ControlCount=   11
      Begin TDBNumber6Ctl.TDBNumber tdbnumW4ExtraWH 
         Height          =   375
         Left            =   7800
         TabIndex        =   154
         Top             =   4920
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":008C
         Caption         =   "frmEmpForm.frx":00AC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0122
         Keys            =   "frmEmpForm.frx":0140
         Spin            =   "frmEmpForm.frx":018A
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumW4Deductions 
         Height          =   375
         Left            =   7800
         TabIndex        =   153
         Top             =   4440
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":01B2
         Caption         =   "frmEmpForm.frx":01D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":023A
         Keys            =   "frmEmpForm.frx":0258
         Spin            =   "frmEmpForm.frx":02A2
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumW4OtherIncome 
         Height          =   375
         Left            =   7800
         TabIndex        =   152
         Top             =   3960
         Width           =   4215
         _Version        =   65536
         _ExtentX        =   7435
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":02CA
         Caption         =   "frmEmpForm.frx":02EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0356
         Keys            =   "frmEmpForm.frx":0374
         Spin            =   "frmEmpForm.frx":03BE
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumW4OtherDependents 
         Height          =   375
         Left            =   3000
         TabIndex        =   151
         Top             =   4920
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":03E6
         Caption         =   "frmEmpForm.frx":0406
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0484
         Keys            =   "frmEmpForm.frx":04A2
         Spin            =   "frmEmpForm.frx":04EC
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumW4Dependents 
         Height          =   375
         Left            =   3000
         TabIndex        =   150
         Top             =   4440
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":0514
         Caption         =   "frmEmpForm.frx":0534
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":05A6
         Keys            =   "frmEmpForm.frx":05C4
         Spin            =   "frmEmpForm.frx":060E
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.CheckBox chkW4_2Jobs 
         Caption         =   "Two Jobs Total"
         Height          =   375
         Left            =   360
         TabIndex        =   149
         Top             =   4440
         Width           =   2295
      End
      Begin VB.ComboBox cmbPRW4 
         Height          =   360
         Left            =   3000
         TabIndex        =   148
         Text            =   "cmbPRW4"
         Top             =   3960
         Width           =   4455
      End
      Begin MSMask.MaskEdBox txtSSN 
         Height          =   375
         Left            =   -70440
         TabIndex        =   40
         Top             =   2880
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   11
         Mask            =   "###-##-####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbOECity 
         Height          =   360
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   7920
         Width           =   4695
      End
      Begin VB.ComboBox cmbRateDiff 
         Height          =   360
         Left            =   -66960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton cmdItemUpdate 
         Caption         =   "UPDATE"
         Height          =   495
         Left            =   -72360
         TabIndex        =   2
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdDDUpdate 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   -71400
         TabIndex        =   142
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton cmdBasis 
         Caption         =   "BASIS FOR DEDUCTION"
         Height          =   375
         Left            =   -68520
         TabIndex        =   13
         Top             =   4800
         Width           =   2895
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
         Height          =   255
         Left            =   -64560
         TabIndex        =   25
         Top             =   5880
         Width           =   2295
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
         Left            =   -66600
         TabIndex        =   23
         Top             =   7320
         Width           =   1815
      End
      Begin TDBText6Ctl.TDBText tdbtxtItemComment 
         Height          =   375
         Left            =   -68520
         TabIndex        =   11
         Top             =   3360
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":0636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":069A
         Key             =   "frmEmpForm.frx":06B8
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
         Height          =   1215
         Left            =   -73440
         TabIndex        =   50
         Top             =   5640
         Width           =   8895
         _Version        =   65536
         _ExtentX        =   15690
         _ExtentY        =   2143
         Caption         =   "frmEmpForm.frx":06FC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0760
         Key             =   "frmEmpForm.frx":077E
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
      Begin VB.ComboBox cmbEEDefaultJob 
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   8040
         Width           =   6135
      End
      Begin TDBText6Ctl.TDBText txtAltName 
         Height          =   375
         Left            =   -74640
         TabIndex        =   47
         Top             =   4560
         Width           =   9615
         _Version        =   65536
         _ExtentX        =   16960
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":07C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0828
         Key             =   "frmEmpForm.frx":0846
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
      Begin TDBText6Ctl.TDBText txtCheckComment 
         Height          =   375
         Left            =   -74640
         TabIndex        =   49
         Top             =   5040
         Width           =   10815
         _Version        =   65536
         _ExtentX        =   19076
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":088A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":08FA
         Key             =   "frmEmpForm.frx":0918
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
      Begin VB.ComboBox cmbWkcCat 
         Height          =   360
         Left            =   -68280
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3360
         Width           =   5655
      End
      Begin VB.ComboBox cmbW2Box14 
         Height          =   360
         Left            =   -64440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   7440
         Width           =   2175
      End
      Begin VB.ComboBox cmbW2Box12 
         Height          =   360
         Left            =   -64440
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   6600
         Width           =   2175
      End
      Begin VB.ComboBox cmbCourtCWT 
         Height          =   360
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   7200
         Width           =   5775
      End
      Begin VB.ComboBox cmbEEDfltCity 
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   7200
         Width           =   6135
      End
      Begin VB.ComboBox cmbDept 
         Height          =   360
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3360
         Width           =   3015
      End
      Begin VB.CheckBox chkCourtAdd 
         Caption         =   "Additional"
         Height          =   255
         Left            =   9840
         TabIndex        =   104
         Top             =   6720
         Width           =   1335
      End
      Begin VB.ComboBox cmb1099 
         Height          =   360
         Left            =   -66360
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3840
         Width           =   2175
      End
      Begin VB.ComboBox cmbEICType 
         Height          =   360
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CheckBox chkUseDeptWkc 
         Caption         =   "Use Department Category"
         Height          =   255
         Left            =   -66000
         TabIndex        =   41
         Top             =   3000
         Width           =   2775
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "  Inactive"
         Height          =   375
         Left            =   -73560
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkStatutory 
         Caption         =   "Statutory Employee"
         Height          =   255
         Left            =   -70440
         TabIndex        =   45
         Top             =   3960
         Width           =   2055
      End
      Begin VB.ComboBox cmbPPY 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbState 
         Height          =   360
         Left            =   -63600
         TabIndex        =   38
         Text            =   "cmbState"
         Top             =   2400
         Width           =   975
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumZipCode 
         Height          =   375
         Left            =   -74760
         TabIndex        =   39
         Top             =   2880
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":095C
         Caption         =   "frmEmpForm.frx":097C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":09E2
         Keys            =   "frmEmpForm.frx":0A00
         Spin            =   "frmEmpForm.frx":0A4A
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
         ValueVT         =   1949630469
         Value           =   0
         MaxValueVT      =   7602181
         MinValueVT      =   5
      End
      Begin VB.CheckBox chkUseAltName 
         Caption         =   "Use Alt Name for Docs."
         Height          =   375
         Left            =   -64920
         TabIndex        =   48
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Frame Frame4 
         Caption         =   " SWT ADDITIONAL AMOUNT "
         Height          =   735
         Left            =   6960
         TabIndex        =   134
         Top             =   5880
         Width           =   5055
         Begin TDBNumber6Ctl.TDBNumber tdbnumSWTExtraAmount 
            Height          =   375
            Left            =   3120
            TabIndex        =   103
            Top             =   240
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   661
            Calculator      =   "frmEmpForm.frx":0A72
            Caption         =   "frmEmpForm.frx":0A92
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmEmpForm.frx":0AF6
            Keys            =   "frmEmpForm.frx":0B14
            Spin            =   "frmEmpForm.frx":0B5E
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
         Begin VB.OptionButton optSWTAddPercent 
            Caption         =   "Percentage"
            Height          =   255
            Left            =   1440
            TabIndex        =   102
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optSWTAddAmount 
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " SWT BASIS "
         Height          =   735
         Left            =   1680
         TabIndex        =   133
         Top             =   5880
         Width           =   4935
         Begin TDBNumber6Ctl.TDBNumber tdbnumSWTAmount 
            Height          =   375
            Left            =   3480
            TabIndex        =   100
            Top             =   240
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   661
            Calculator      =   "frmEmpForm.frx":0B86
            Caption         =   "frmEmpForm.frx":0BA6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmEmpForm.frx":0C0A
            Keys            =   "frmEmpForm.frx":0C28
            Spin            =   "frmEmpForm.frx":0C72
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.OptionButton optSWTPercent 
            Caption         =   "Percentage"
            Height          =   255
            Left            =   1920
            TabIndex        =   99
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optSWTExemptions 
            Caption         =   "Exemptions"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " FWT ADDITIONAL AMOUNT "
         Height          =   735
         Left            =   6960
         TabIndex        =   132
         Top             =   3000
         Width           =   5055
         Begin TDBNumber6Ctl.TDBNumber tdbnumFWTExtraAmount 
            Height          =   375
            Left            =   3120
            TabIndex        =   96
            Top             =   240
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   661
            Calculator      =   "frmEmpForm.frx":0C9A
            Caption         =   "frmEmpForm.frx":0CBA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmEmpForm.frx":0D1E
            Keys            =   "frmEmpForm.frx":0D3C
            Spin            =   "frmEmpForm.frx":0D86
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
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin VB.OptionButton optFWTAddAmount 
            Caption         =   "Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optFWTAddPercent 
            Caption         =   "Percentage"
            Height          =   255
            Left            =   1440
            TabIndex        =   95
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " FWT BASIS "
         Height          =   735
         Left            =   1680
         TabIndex        =   131
         Top             =   3000
         Width           =   4935
         Begin TDBNumber6Ctl.TDBNumber tdbnumFWTAmount 
            Height          =   375
            Left            =   3480
            TabIndex        =   93
            Top             =   240
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   661
            Calculator      =   "frmEmpForm.frx":0DAE
            Caption         =   "frmEmpForm.frx":0DCE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmEmpForm.frx":0E32
            Keys            =   "frmEmpForm.frx":0E50
            Spin            =   "frmEmpForm.frx":0E9A
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
         Begin VB.OptionButton optFWTPercent 
            Caption         =   "Percentage"
            Height          =   255
            Left            =   1920
            TabIndex        =   92
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optFWTExemptions 
            Caption         =   "Exemptions"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdDirDepDelete 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   -70680
         TabIndex        =   53
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdDirDepAdd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   -72480
         TabIndex        =   52
         Top             =   5040
         Width           =   1335
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumDDAmount 
         Height          =   375
         Left            =   -67320
         TabIndex        =   63
         Top             =   4680
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":0EC2
         Caption         =   "frmEmpForm.frx":0EE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":0F44
         Keys            =   "frmEmpForm.frx":0F62
         Spin            =   "frmEmpForm.frx":0FAC
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
         Left            =   -64560
         TabIndex        =   24
         Top             =   5520
         Width           =   2055
      End
      Begin VB.CheckBox chkOEDEDNoSUNTax 
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
         Left            =   -66600
         TabIndex        =   20
         Top             =   5880
         Width           =   1695
      End
      Begin VB.CheckBox chkOEDEDNoFUNTax 
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
         Left            =   -66600
         TabIndex        =   19
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CheckBox chkOEDEDNoCWTTax 
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
         Left            =   -68520
         TabIndex        =   18
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CheckBox chkOEDEDNoSWTTax 
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
         Left            =   -68520
         TabIndex        =   17
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CheckBox chkOEDEDNoFWTTax 
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
         Left            =   -68520
         TabIndex        =   16
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CheckBox chkOEDEDNoMedTax 
         Caption         =   "No MED Tax"
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
         Left            =   -68520
         TabIndex        =   15
         Top             =   5880
         Width           =   1815
      End
      Begin VB.CheckBox chkOEDEDNoSSTax 
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
         Left            =   -68520
         TabIndex        =   14
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CommandButton cmdOEDEDDelete 
         Caption         =   "DELETE"
         Height          =   495
         Left            =   -70560
         TabIndex        =   3
         Top             =   7800
         Width           =   1335
      End
      Begin VB.CommandButton cmdOEDEDAdd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   -74160
         TabIndex        =   1
         Top             =   7800
         Width           =   1335
      End
      Begin TDBNumber6Ctl.TDBNumber tdbnumMaxAmt 
         Height          =   375
         Left            =   -65520
         TabIndex        =   6
         Top             =   1680
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":0FD4
         Caption         =   "frmEmpForm.frx":0FF4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":105E
         Keys            =   "frmEmpForm.frx":107C
         Spin            =   "frmEmpForm.frx":10C6
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
      Begin TDBNumber6Ctl.TDBNumber tdbnumAmtPct 
         Height          =   375
         Left            =   -68640
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":10EE
         Caption         =   "frmEmpForm.frx":110E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1170
         Keys            =   "frmEmpForm.frx":118E
         Spin            =   "frmEmpForm.frx":11D8
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
      Begin VB.Frame fraBasis 
         Caption         =   "Basis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67440
         TabIndex        =   124
         Top             =   3720
         Width           =   4095
         Begin VB.OptionButton optDDNet 
            Caption         =   "Net"
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
            Left            =   3000
            TabIndex        =   62
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optDDPercent 
            Caption         =   "Percent"
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
            Left            =   1560
            TabIndex        =   61
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optDDAmount 
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraType 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67440
         TabIndex        =   123
         Top             =   2760
         Width           =   2895
         Begin VB.OptionButton optSavings 
            Caption         =   "Savings"
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
            Left            =   1560
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optChecking 
            Caption         =   "Checking"
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
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkDirDepActive 
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
         Height          =   255
         Left            =   -67440
         TabIndex        =   54
         Top             =   960
         Width           =   975
      End
      Begin TDBText6Ctl.TDBText txtBankName 
         Height          =   375
         Left            =   -67440
         TabIndex        =   55
         Top             =   1320
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":1200
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1268
         Key             =   "frmEmpForm.frx":1286
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
      Begin VB.CheckBox chkNotNet 
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
         Left            =   -66600
         TabIndex        =   21
         Top             =   6600
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
         Left            =   -66600
         TabIndex        =   22
         Top             =   6960
         Width           =   1215
      End
      Begin VB.Frame fraItmBasis 
         Caption         =   "Basis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -67440
         TabIndex        =   122
         Top             =   2160
         Width           =   4215
         Begin VB.OptionButton optHrly 
            Caption         =   "Hourly"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2760
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optPct 
            Caption         =   "Percent"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1440
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optAmt 
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkUseEmpDef 
         Caption         =   "Use Employer Definition"
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
         Left            =   -68520
         TabIndex        =   10
         Top             =   3000
         Width           =   2655
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
         Height          =   255
         Left            =   -68520
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid fgDirDep 
         Height          =   3255
         Left            =   -72840
         TabIndex        =   51
         Top             =   1320
         Width           =   3735
         _cx             =   6588
         _cy             =   5741
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
      Begin VSFlex8Ctl.VSFlexGrid fgOEDED 
         Height          =   6615
         Left            =   -74520
         TabIndex        =   0
         Top             =   960
         Width           =   5295
         _cx             =   9340
         _cy             =   11668
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
      Begin VB.ComboBox cmbRaceCode 
         Height          =   360
         Left            =   -69120
         TabIndex        =   74
         Top             =   3060
         Width           =   615
      End
      Begin VB.ComboBox cmbMaritalStatus 
         Height          =   360
         Left            =   -65640
         TabIndex        =   75
         Top             =   3180
         Width           =   615
      End
      Begin VB.ComboBox cmbEducationLevel 
         Height          =   360
         Left            =   -72870
         TabIndex        =   76
         Top             =   3660
         Width           =   615
      End
      Begin VB.ComboBox cmbShiftCode 
         Height          =   360
         Left            =   -69120
         TabIndex        =   77
         Top             =   3660
         Width           =   615
      End
      Begin VB.ComboBox cmbSex 
         Height          =   360
         Left            =   -72840
         TabIndex        =   73
         Top             =   3060
         Width           =   615
      End
      Begin TDBNumber6Ctl.TDBNumber lngWorkCompNo 
         Height          =   375
         Left            =   -67440
         TabIndex        =   78
         Top             =   3780
         Width           =   2535
         _Version        =   65536
         _ExtentX        =   4471
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":12CA
         Caption         =   "frmEmpForm.frx":12EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":135A
         Keys            =   "frmEmpForm.frx":1378
         Spin            =   "frmEmpForm.frx":13C2
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
         MaxValueVT      =   6356997
         MinValueVT      =   7602181
      End
      Begin VB.ComboBox cmbTermReason 
         Height          =   360
         Left            =   -65760
         TabIndex        =   72
         Top             =   2460
         Width           =   735
      End
      Begin TDBDate6Ctl.TDBDate dteDateTerminated 
         Height          =   375
         Left            =   -70800
         TabIndex        =   71
         Top             =   2460
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":13EA
         Caption         =   "frmEmpForm.frx":14EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":155E
         Keys            =   "frmEmpForm.frx":157C
         Spin            =   "frmEmpForm.frx":15DA
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.CheckBox chkNoMedTax 
         Caption         =   "  No Med Tax"
         Height          =   375
         Left            =   3840
         TabIndex        =   87
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkFWTMarried 
         Caption         =   " Married"
         Height          =   375
         Left            =   360
         TabIndex        =   90
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkSWTMarried 
         Caption         =   " Married"
         Height          =   375
         Left            =   360
         TabIndex        =   97
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CheckBox chkNoStateUnemp 
         Caption         =   "  No State Unemp"
         Height          =   375
         Left            =   9480
         TabIndex        =   89
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CheckBox chkNoFedUnemp 
         Caption         =   "  No Fed Unemp"
         Height          =   375
         Left            =   9480
         TabIndex        =   86
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkNoCityTax 
         Caption         =   "  No City Tax"
         Height          =   375
         Left            =   7560
         TabIndex        =   85
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkSalaried 
         Caption         =   "  Salaried"
         Height          =   375
         Left            =   1800
         TabIndex        =   79
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkNoSSTax 
         Caption         =   "  No SS Tax"
         Height          =   375
         Left            =   3840
         TabIndex        =   83
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkNoFedTax 
         Caption         =   "  No Fed Tax"
         Height          =   375
         Left            =   5880
         TabIndex        =   84
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkNoStateTax 
         Caption         =   "  No State Tax"
         Height          =   375
         Left            =   5880
         TabIndex        =   88
         Top             =   2040
         Width           =   1575
      End
      Begin TDBText6Ctl.TDBText txtAddress1 
         Height          =   375
         Left            =   -74760
         TabIndex        =   35
         Top             =   1920
         Width           =   12015
         _Version        =   65536
         _ExtentX        =   21193
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":1602
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1668
         Key             =   "frmEmpForm.frx":1686
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
      Begin TDBNumber6Ctl.TDBNumber lngEmployeeNumber 
         Height          =   375
         Left            =   -74760
         TabIndex        =   31
         Top             =   1440
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":16CA
         Caption         =   "frmEmpForm.frx":16EA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1754
         Keys            =   "frmEmpForm.frx":1772
         Spin            =   "frmEmpForm.frx":17BC
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
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText txtFirstName 
         Height          =   375
         Left            =   -72120
         TabIndex        =   32
         Top             =   1440
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":17E4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":184E
         Key             =   "frmEmpForm.frx":186C
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
         Text            =   "First Name"
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
      Begin TDBText6Ctl.TDBText txtLastName 
         Height          =   375
         Left            =   -66600
         TabIndex        =   34
         Top             =   1440
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":18B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1918
         Key             =   "frmEmpForm.frx":1936
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
         Text            =   "Last Name"
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
         TabIndex        =   36
         Top             =   2400
         Width           =   6015
         _Version        =   65536
         _ExtentX        =   10610
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":197A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":19E0
         Key             =   "frmEmpForm.frx":19FE
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
         Left            =   -68520
         TabIndex        =   37
         Top             =   2400
         Width           =   3975
         _Version        =   65536
         _ExtentX        =   7011
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":1A42
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1AA0
         Key             =   "frmEmpForm.frx":1ABE
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
      Begin TDBText6Ctl.TDBText txtMI 
         Height          =   375
         Left            =   -67680
         TabIndex        =   33
         Top             =   1440
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":1B02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1B5A
         Key             =   "frmEmpForm.frx":1B78
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
         AlignHorizontal =   2
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   3
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
      Begin TDBNumber6Ctl.TDBNumber curHourlyAmt 
         Height          =   375
         Left            =   7920
         TabIndex        =   81
         Top             =   960
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":1BBC
         Caption         =   "frmEmpForm.frx":1BDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1C4C
         Keys            =   "frmEmpForm.frx":1C6A
         Spin            =   "frmEmpForm.frx":1CB4
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
         Format          =   "$ ###,###.##;($ ###,###.##)"
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
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBText6Ctl.TDBText lngDepartmentNumber 
         Height          =   375
         Left            =   -66840
         TabIndex        =   113
         Top             =   3120
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":1CDC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1D44
         Key             =   "frmEmpForm.frx":1D62
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
      Begin TDBNumber6Ctl.TDBNumber curSalaryAmt 
         Height          =   375
         Left            =   3960
         TabIndex        =   80
         Top             =   960
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   661
         Calculator      =   "frmEmpForm.frx":1DA6
         Caption         =   "frmEmpForm.frx":1DC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":1E36
         Keys            =   "frmEmpForm.frx":1E54
         Spin            =   "frmEmpForm.frx":1E9E
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
         Format          =   "$ ###,###.##;($ ###,###.##)"
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
         ValueVT         =   82575361
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate DteDateHired 
         DataSource      =   "premployee"
         Height          =   375
         Left            =   -74640
         TabIndex        =   64
         Top             =   1020
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":1EC6
         Caption         =   "frmEmpForm.frx":1FC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2032
         Keys            =   "frmEmpForm.frx":2050
         Spin            =   "frmEmpForm.frx":20AE
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
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   11
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   2
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   2
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteDateLastRaise 
         Height          =   375
         Left            =   -70800
         TabIndex        =   68
         Top             =   1740
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":20D6
         Caption         =   "frmEmpForm.frx":21D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":224A
         Keys            =   "frmEmpForm.frx":2268
         Spin            =   "frmEmpForm.frx":22C6
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteDateLastReview 
         Height          =   375
         Left            =   -74640
         TabIndex        =   67
         Top             =   1740
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":22EE
         Caption         =   "frmEmpForm.frx":23EE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2464
         Keys            =   "frmEmpForm.frx":2482
         Spin            =   "frmEmpForm.frx":24E0
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
         MinDate         =   2
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate DteDateofBirth 
         Height          =   375
         Left            =   -70800
         TabIndex        =   65
         Top             =   1020
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":2508
         Caption         =   "frmEmpForm.frx":2608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2678
         Keys            =   "frmEmpForm.frx":2696
         Spin            =   "frmEmpForm.frx":26F4
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
         EditMode        =   1
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
         MinDate         =   2
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteDateLastPaid 
         Height          =   375
         Left            =   -67440
         TabIndex        =   66
         Top             =   1020
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":271C
         Caption         =   "frmEmpForm.frx":281C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":288E
         Keys            =   "frmEmpForm.frx":28AC
         Spin            =   "frmEmpForm.frx":290A
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteDateLastLayoff 
         Height          =   375
         Left            =   -67440
         TabIndex        =   69
         Top             =   1740
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":2932
         Caption         =   "frmEmpForm.frx":2A32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2AA8
         Keys            =   "frmEmpForm.frx":2AC6
         Spin            =   "frmEmpForm.frx":2B24
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate dteDateLastRecall 
         Height          =   375
         Left            =   -74640
         TabIndex        =   70
         Top             =   2460
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   661
         Calendar        =   "frmEmpForm.frx":2B4C
         Caption         =   "frmEmpForm.frx":2C4C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2CC2
         Keys            =   "frmEmpForm.frx":2CE0
         Spin            =   "frmEmpForm.frx":2D3E
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2010382337
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText txtABA 
         Height          =   375
         Left            =   -67440
         TabIndex        =   56
         Top             =   1800
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":2D66
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2DC2
         Key             =   "frmEmpForm.frx":2DE0
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
      Begin TDBText6Ctl.TDBText txtAccount 
         Height          =   375
         Left            =   -67440
         TabIndex        =   57
         Top             =   2280
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   661
         Caption         =   "frmEmpForm.frx":2E24
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmEmpForm.frx":2E8E
         Key             =   "frmEmpForm.frx":2EAC
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
      Begin VB.Label Label13 
         Caption         =   "REVISED W4 FORM"
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
         Left            =   360
         TabIndex        =   147
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label22 
         Caption         =   "S S N:"
         Height          =   255
         Left            =   -71520
         TabIndex        =   146
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label21 
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
         Left            =   -68640
         TabIndex        =   145
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label lblRateDiff 
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
         Height          =   255
         Left            =   -68760
         TabIndex        =   143
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lblEEDefaultJob 
         Caption         =   "Default Job:"
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
         Left            =   360
         TabIndex        =   141
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "Courtesy City Withholding"
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
         Left            =   6840
         TabIndex        =   140
         Top             =   6720
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "1099 Employee:"
         Height          =   255
         Left            =   -67800
         TabIndex        =   139
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "EIC Type:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   138
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Workers Comp Category:"
         Height          =   255
         Left            =   -68400
         TabIndex        =   137
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Pays Per Year:"
         Height          =   255
         Left            =   840
         TabIndex        =   136
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Default City Tax / Default State"
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
         Left            =   360
         TabIndex        =   135
         Top             =   6840
         Width           =   3615
      End
      Begin VB.Label Label14 
         Caption         =   "STATE WITHHOLDING TAX"
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
         Left            =   360
         TabIndex        =   130
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Label Label11 
         Caption         =   "FEDERAL WITHHOLDING TAX"
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
         Left            =   360
         TabIndex        =   129
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Line Line1 
         X1              =   -68760
         X2              =   -62160
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblW214 
         Caption         =   "W2 Box 14 Code"
         Height          =   255
         Left            =   -64440
         TabIndex        =   128
         Top             =   7080
         Width           =   1695
      End
      Begin VB.Label lblW212 
         Caption         =   "W2 Box 12 Code"
         Height          =   255
         Left            =   -64440
         TabIndex        =   127
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label lblOEDEDTitle 
         Caption         =   "OE DED TITLE"
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
         Left            =   -67080
         TabIndex        =   126
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Item Display"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67320
         TabIndex        =   121
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Department:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Sex:"
         Height          =   375
         Left            =   -74640
         TabIndex        =   119
         Top             =   3180
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Shift Code:"
         Height          =   375
         Left            =   -70800
         TabIndex        =   118
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Education Level:"
         Height          =   375
         Left            =   -74640
         TabIndex        =   109
         Top             =   3780
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Marital Status:"
         Height          =   375
         Left            =   -67440
         TabIndex        =   117
         Top             =   3180
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Race:"
         Height          =   375
         Left            =   -70800
         TabIndex        =   116
         Top             =   3180
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Term Reason:"
         Height          =   375
         Left            =   -67440
         TabIndex        =   115
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "State:"
         Height          =   255
         Left            =   -64320
         TabIndex        =   114
         Top             =   2520
         Width           =   615
      End
   End
   Begin VB.Label Label20 
      Caption         =   "c"
      Height          =   375
      Left            =   240
      TabIndex        =   144
      Top             =   10320
      Width           =   375
   End
   Begin VB.Label Label12 
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
      Height          =   495
      Left            =   9600
      TabIndex        =   125
      Top             =   9960
      Width           =   2895
   End
   Begin VB.Label txtEmployeeDisplay 
      Alignment       =   2  'Center
      Caption         =   "Employee Display"
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
      Left            =   1950
      TabIndex        =   112
      Top             =   600
      Width           =   9735
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
      Left            =   1950
      TabIndex        =   108
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmEmpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DeptID As Long

Dim rsOEDED As ADODB.Recordset
Dim rsDirDep As ADODB.Recordset
Dim rsERItem As New ADODB.Recordset

Dim I, J, K As Long
Dim ActiveDrop, TypeDrop As String

Dim AddFlag As Boolean
Dim SkipSave As Boolean
Dim InitFlag As Boolean

Dim SaveFlag As Boolean

Dim OrigEmpNum As Long

Private Sub cmbPRW4_Click()
    FWTDisplay
End Sub

Private Sub cmdDDUpdate_Click()
    DirDepSave
End Sub

Private Sub cmdItemUpdate_Click()
    ItemSave
End Sub


Private Sub Form_Load()
    
    InitFlag = True
    
    ' *********
    AddFlag = False
    SaveFlag = False
    ' *********
    
    If Not PREmployee.GetByID(SelID) Then
        MsgBox "Employee ID NF: " & SelID, vbExclamation
        End
    End If
    
    If Not PRW4.GetBySQL("SELECT * FROM PRW4 WHERE PRW4.EmployeeID = " & PREmployee.EmployeeID) Then
        PRW4.Clear
        PRW4.EmployeeID = PREmployee.EmployeeID
        PRW4.Save (Equate.RecAdd)
    End If
    
    ' store the original employee number
    OrigEmpNum = PREmployee.EmployeeNumber
    
    ' PREmployee.GetBySQL ("SELECT * FROM PREmployee WHERE PREmployee.EmployeeNumber = 1102")
    DeptID = PREmployee.DepartmentID
    ' Populate Combo Boxes
    Me.cmbMaritalStatus.AddItem "M"
    Me.cmbMaritalStatus.AddItem "S"
    Me.cmbMaritalStatus.AddItem "D"
    Me.cmbSex.AddItem "M"
    Me.cmbSex.AddItem "F"
    Me.cmbEducationLevel.AddItem 1
    Me.cmbEducationLevel.AddItem 2
    Me.cmbEducationLevel.AddItem 3
    Me.cmbEducationLevel.AddItem 4
    Me.cmbEducationLevel.AddItem 5
    Me.cmbEducationLevel.AddItem 6
    Me.cmbEducationLevel.AddItem 7
    
    ' set tdbText parameters
    tdbTextSet Me.txtFirstName
    tdbTextSet Me.txtMI
    tdbTextSet Me.txtLastName
    tdbTextSet Me.txtAltName
    tdbTextSet Me.txtAddress1
    tdbTextSet Me.txtAddress2
    tdbTextSet Me.txtCity
    tdbTextSet Me.txtBankName
    tdbTextSet Me.txtABA
    tdbTextSet Me.txtAccount
    tdbTextSet Me.txtCheckComment

    ' comment
    SQLString = "SELECT * FROM Notes WHERE NoteType = " & Equate.NoteTypeEE & _
                "AND DateTm = 0 " & _
                "AND RelatedID = " & PREmployee.EmployeeID
    If Notes.GetBySQL(SQLString) = False Then
        Notes.Clear
        Notes.NoteType = Equate.NoteTypeEE
        Notes.RelatedID = PREmployee.EmployeeID
        Notes.Save (Equate.RecAdd)
    End If
    Me.tdbtxtComment.text = Notes.Notation
    
    tdbTextSet Me.tdbtxtItemComment
    Me.tdbtxtItemComment.MaxLength = 50

    ' set tdbNumber parameters
    tdbIntegerSet Me.lngEmployeeNumber
    ' tdbIntegerSet Me.lngFWTExemptions
    ' tdbIntegerSet Me.lngFWTPercent
    ' tdbIntegerSet Me.lngSWTExemptions
    ' tdbIntegerSet Me.lngSWTPercent
    tdbIntegerSet Me.tdbnumW4Dependents
    tdbIntegerSet Me.tdbnumW4OtherDependents
    
    ' set tdbAmount parameters - two decimal places
    tdbAmountSet Me.curSalaryAmt
    tdbAmountSet Me.curHourlyAmt
    
'    tdbAmountSet Me.curFWTExtraAmt
'    tdbAmountSet Me.curSWTExtraAmt
    
    tdbAmountSet Me.tdbnumAmtPct
    tdbAmountSet Me.tdbnumMaxAmt
    tdbAmountSet Me.tdbnumDDAmount
    
    tdbAmountSet Me.tdbnumW4OtherIncome
    tdbAmountSet Me.tdbnumW4Deductions
    tdbAmountSet Me.tdbnumW4ExtraWH
    
    ' set tdbDate parameters
    tdbDateSet Me.DteDateHired, PREmployee.DateHired
    tdbDateSet Me.dteDateLastLayoff, PREmployee.DateLastLayoff
    tdbDateSet Me.dteDateLastPaid, PREmployee.DateLastPaid
    tdbDateSet Me.dteDateLastRaise, PREmployee.DateLastRaise
    tdbDateSet Me.dteDateLastRecall, PREmployee.DateLastRecall
    tdbDateSet Me.dteDateLastReview, PREmployee.DateLastReview
    tdbDateSet Me.DteDateofBirth, PREmployee.DateOfBirth
    tdbDateSet Me.dteDateTerminated, PREmployee.DateTerminated
    
    ' Populate Main Employee screen tab variables from file
    Me.lblCompanyName.Caption = PRCompany.Name
    Me.txtEmployeeDisplay = PREmployee.EmployeeNumber & " " & PREmployee.LFName
    
    Me.lngEmployeeNumber.Format = "########0"
    Me.lngEmployeeNumber.DisplayFormat = "########0"
    Me.lngEmployeeNumber.HighlightText = True
    Me.lngEmployeeNumber.Key.Clear = ""
    Me.lngEmployeeNumber.MinValue = 0
    Me.lngEmployeeNumber.MaxValue = 999999999
    Me.lngEmployeeNumber = PREmployee.EmployeeNumber
    
    Me.txtFirstName.text = PREmployee.FirstName
    Me.txtMI.text = PREmployee.MidInit
    Me.txtLastName.text = PREmployee.LastName
    
    Me.txtAltName = PREmployee.AltName
    Me.txtCheckComment = PREmployee.CheckComment
    
    Me.chkUseAltName = PREmployee.UseAltName
    Me.txtAddress1 = PREmployee.Address1
    Me.txtAddress2 = PREmployee.Address2
    Me.txtCity = PREmployee.City
    Me.cmbState = PREmployee.State
    
    ' Me.txtSSN.HighlightText = dbiHighlightField
    ' Me.txtSSN.Value = Format(PREmployee.SSN, "000000000")
    Me.txtSSN.text = Format(PREmployee.SSN, "000-00-0000")
    
    Me.chkStatutory = PREmployee.Statutory
       
    tdbIntegerSet Me.tdbnumZipCode
    Me.tdbnumZipCode.Format = "00000"
    Me.tdbnumZipCode.DisplayFormat = ""
    ZipString = Format(PREmployee.ZipCode, "00000")
    Me.tdbnumZipCode = Mid(PREmployee.ZipCode, 1, 5)
       
    ' Populate Pay Parameter Employee screen tab variables from file
    Me.curSalaryAmt = PREmployee.SalaryAmount
    Me.curHourlyAmt = PREmployee.HourlyAmount
    Me.chkInactive.Value = PREmployee.Inactive
    Me.chkSalaried.Value = PREmployee.Salaried
    Me.chkNoSSTax.Value = PREmployee.NoSSTax
    Me.chkNoMedTax.Value = PREmployee.NoMedTax
    Me.chkNoFedTax.Value = PREmployee.NoFedTax
    Me.chkNoStateTax.Value = PREmployee.NoStateTax
    Me.chkNoCityTax.Value = PREmployee.NoCityTax
    Me.chkNoFedUnemp.Value = PREmployee.NoFedUnemp
    Me.chkNoStateUnemp.Value = PREmployee.NoStateUnemp
    
    Me.chkFWTMarried.Value = PREmployee.FWTMarried
    
    Me.chkCourtAdd = PREmployee.CourtesyAdd
    
    ' FWT Basis
    If PREmployee.FWTBasis = PREquate.BasisExemptions Then
        Me.optFWTExemptions = True
        Me.tdbnumFWTAmount.Format = "##0"
        Me.tdbnumFWTAmount.DisplayFormat = ""
    Else
        Me.optFWTPercent = True
        Me.tdbnumFWTAmount.Format = "##0.00 %"
        Me.tdbnumFWTAmount.DisplayFormat = ""
    End If
    Me.tdbnumFWTAmount = PREmployee.FWTAmount
        
    ' extra FWT
    If PREmployee.FWTExtraBasis = PREquate.BasisPercent Then
        Me.optFWTAddPercent = True
        Me.tdbnumFWTExtraAmount.Format = "##0.00 %"
        Me.tdbnumFWTExtraAmount.DisplayFormat = ""
    Else
        Me.optFWTAddAmount = True
        Me.tdbnumFWTExtraAmount.Format = "$ ##,##0.00"
        Me.tdbnumFWTExtraAmount.DisplayFormat = ""
    End If
    Me.tdbnumFWTExtraAmount = PREmployee.FWTExtraAmount
    
    Me.chkSWTMarried = PREmployee.SWTMarried
    
    ' SWT Basis
    If PREmployee.SWTBasis = PREquate.BasisExemptions Then
        Me.optSWTExemptions = True
        Me.tdbnumSWTAmount.Format = "##0"
        Me.tdbnumSWTAmount.DisplayFormat = ""
    Else
        Me.optSWTPercent = True
        Me.tdbnumSWTAmount.Format = "##0.00 %"
        Me.tdbnumSWTAmount.DisplayFormat = ""
    End If
    Me.tdbnumSWTAmount = PREmployee.SWTAmount
        
    ' extra SWT
    If PREmployee.SWTExtraBasis = PREquate.BasisPercent Then
        Me.optSWTAddPercent = True
        Me.tdbnumSWTExtraAmount.Format = "##0.00 %"
        Me.tdbnumSWTExtraAmount.DisplayFormat = ""
    Else
        Me.optSWTAddAmount = True
        Me.tdbnumSWTExtraAmount.Format = "$ ##,##0.00"
        Me.tdbnumSWTExtraAmount.DisplayFormat = ""
    End If
    Me.tdbnumSWTExtraAmount = PREmployee.SWTExtraAmount
    
    ' revised W4
    Me.chkW4_2Jobs = PRW4.TwoJobs
    Me.tdbnumW4Dependents = PRW4.Dependents
    Me.tdbnumW4OtherDependents = PRW4.DependentsOther
    Me.tdbnumW4OtherIncome = PRW4.OtherIncome
    Me.tdbnumW4Deductions = PRW4.Deductions
    Me.tdbnumW4ExtraWH = PRW4.ExtraWH
    
    ' Populate state dropdown box
    PRState.GetBySQL ("SELECT * FROM PRState order by PRState.StateAbbrev")
    Do
        Me.cmbState.AddItem PRState.StateAbbrev
        If Not PRState.GetNext Then
           Exit Do
        End If
    Loop

'    ' Populate city dropdown box
'    PRCity.GetBySQL ("SELECT * FROM PRCity order by PRCity.CityNumber")
'    Do
'        ' Me.cmbDfltCity.AddItem PRCity.CityNumber
'        If Not PRCity.GetNext Then
'           Exit Do
'        End If
'    Loop
    
'     Me.cmbworkcompnum = PREmployee.WorkCompNum
    On Error Resume Next
    If PREmployee.DateHired <> 0 Then Me.DteDateHired = PREmployee.DateHired
    If PREmployee.DateLastPaid <> 0 Then Me.dteDateLastPaid = PREmployee.DateLastPaid
    If PREmployee.DateLastRaise <> 0 Then Me.dteDateLastRaise = PREmployee.DateLastRaise
    If PREmployee.DateLastRecall <> 0 Then Me.dteDateLastRecall = PREmployee.DateLastRecall
    If PREmployee.DateLastReview <> 0 Then Me.dteDateLastReview = PREmployee.DateLastReview
    If PREmployee.DateLastLayoff <> 0 Then Me.dteDateLastLayoff = PREmployee.DateLastLayoff
    If PREmployee.DateHired <> 0 Then Me.DteDateHired = PREmployee.DateHired
    If PREmployee.DateTerminated <> 0 Then Me.dteDateTerminated = PREmployee.DateTerminated
    If PREmployee.DateOfBirth <> 0 Then Me.DteDateofBirth = PREmployee.DateOfBirth
    If PREmployee.TermReason <> 0 Then Me.cmbTermReason = PREmployee.TermReason
    If PREmployee.Sex <> "" Then Me.cmbSex = PREmployee.Sex
    If PREmployee.RaceCode <> 0 Then Me.cmbRaceCode = PREmployee.RaceCode
    If PREmployee.MaritalStatus <> "" Then Me.cmbMaritalStatus = PREmployee.MaritalStatus
    If PREmployee.EducationLevel <> 0 Then Me.cmbEducationLevel = PREmployee.EducationLevel
    If PREmployee.ShiftCode <> 0 Then Me.cmbShiftCode = PREmployee.ShiftCode
    On Error GoTo 0
    
    DropDownInit
    FWTDisplay
    GridInit
    
    ' populate the default Job combo?
    Me.cmbEEDefaultJob.Visible = False
    Me.lblEEDefaultJob.Visible = False
    If TableExists("JCJob", cn) = True Then
        With Me.cmbEEDefaultJob
            SQLString = "SELECT * FROM JCJob ORDER BY FullName"
            If JCJob.GetBySQL(SQLString) Then
                Me.lblEEDefaultJob.Visible = True
                .Visible = True
                .AddItem "NONE"
                .ItemData(.NewIndex) = 0
                Do
                    .AddItem JCJob.FullName
                    .ItemData(.NewIndex) = JCJob.JobID
                    If JCJob.GetNext = False Then Exit Do
                Loop
            End If
            If .ListCount > 0 Then
                .ListIndex = 0
                For I = 0 To .ListCount - 1
                    If .ItemData(I) = PREmployee.DefaultJobID Then
                        .ListIndex = I
                        Exit For
                    End If
                Next I
            End If
        End With
    End If
    
    ' start tab - zero based
    Me.SSTab1.Tab = 0
    
    ' Federal type
    FWTDisplay
    
    InitFlag = False

    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True

End Sub

Private Sub FWTDisplay()
    Dim booStandard As Boolean
    Dim booW4 As Boolean
    If Me.cmbPRW4.ListIndex = PREquate.PRW4Standard Then
        booStandard = True
        booW4 = False
    Else
        booStandard = False
        booW4 = True
    End If
    With Me
        .chkW4_2Jobs.Visible = booW4
        .tdbnumW4Dependents.Visible = booW4
        .tdbnumW4OtherDependents.Visible = booW4
        .tdbnumW4OtherIncome.Visible = booW4
        .tdbnumW4Deductions.Visible = booW4
        .tdbnumW4ExtraWH.Visible = booW4
        .chkFWTMarried.Visible = booStandard
        .Frame1.Visible = booStandard
        .Frame2.Visible = booStandard
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
        Case vbKeyF2: OpenTab 0
        Case vbKeyF3: OpenTab 1
        Case vbKeyF4: OpenTab 2
        Case vbKeyF5: OpenTab 3
        Case vbKeyF6: OpenTab 4
    End Select
    
End Sub

Private Sub OpenTab(ByVal TabNum As Byte)
    Me.SSTab1.Tab = TabNum
End Sub


Private Sub ItemDisplay()

 ' MsgBox "/" & AddFlag & "/" & SaveFlag
    
    If AddFlag Then Exit Sub
    If SaveFlag = True Then Exit Sub
    If rsOEDED.RecordCount = 0 Then Exit Sub
    
    ' get the associated employer item
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemID = " & rsOEDED!EmployerItemID
    rsInit SQLString, cn, rsERItem
    If rsERItem.BOF And rsERItem.EOF Then
        MsgBox "Employer Item NF: " & rsOEDED!EmployerItemID, vbExclamation
        End
    End If
    
    ' always take these from the employee item
    Me.chkActive = rsOEDED!Active
    
    If IsNull(rsOEDED!AmtPct) Then
        Me.tdbnumAmtPct = 0
    Else
        Me.tdbnumAmtPct = rsOEDED!AmtPct
    End If
    
    If IsNull(rsOEDED!MaxAmount) Then
        Me.tdbnumMaxAmt = 0
    Else
        Me.tdbnumMaxAmt = rsOEDED!MaxAmount
    End If
    
    If rsOEDED!Basis = PREquate.BasisAmount Then
        Me.optAmt = True
        Me.tdbnumAmtPct.Caption = "Amount"
    ElseIf rsOEDED!Basis = PREquate.BasisHourly Then
        Me.optHrly = True
        Me.tdbnumAmtPct.Caption = "Rate"
    Else
        Me.optPct = True
        Me.tdbnumAmtPct.Caption = "Percent"
    End If
    
    ' always take the title from the employer
    Me.lblOEDEDTitle = rsERItem!Title
    
    Me.chkUseEmpDef = rsOEDED!UseEmployer
    
    ' hide rate difference
    '    for deductions
    '    for ER = N/A
    If rsOEDED!ItemType = PREquate.ItemTypeDED Or nNull(rsERItem!RateDifference) = 0 Then
        Me.lblRateDiff.Visible = False
        Me.cmbRateDiff.Visible = False
    Else
        Me.lblRateDiff.Visible = True
        Me.cmbRateDiff.Visible = True
    End If
    
    ' must be hourly if ER has rate difference setup
    If nNull(rsERItem!RateDifference) = 0 Then
        Me.fraItmBasis.Enabled = True
        Me.optAmt.Enabled = True
        Me.optPct.Enabled = True
        Me.optHrly.Enabled = True
    Else
        Me.fraItmBasis.Enabled = False
        Me.optAmt.Enabled = False
        Me.optPct.Enabled = False
        Me.optHrly.Enabled = False
    End If
    
    If rsOEDED!UseEmployer = 1 Then
        
        Me.chkOEDEDNoSSTax = nNull(rsERItem!NoSSTax)
        Me.chkOEDEDNoMedTax = nNull(rsERItem!NoMedTax)
        Me.chkOEDEDNoFWTTax = nNull(rsERItem!NoFWTTax)
        Me.chkOEDEDNoSWTTax = nNull(rsERItem!NoSWTTax)
        Me.chkOEDEDNoCWTTax = nNull(rsERItem!NoCWTTax)
        Me.chkOEDEDNoFUNTax = nNull(rsERItem!NoFUNTax)
        Me.chkOEDEDNoSUNTax = nNull(rsERItem!NoSUNTax)
        Me.chkTips = nNull(rsERItem!Tips)
        Me.chkNotNet = nNull(rsERItem!NotInNet)
        Me.chkDirDepRpt = nNull(rsERItem!DirDepRpt)
        Me.chkPension = nNull(rsERItem!Pension)
        Me.chkSickPay = nNull(rsERItem!SickPay)
    
        cmbPoint Me.cmbW2Box12, rsERItem!W2Box12Code
        cmbPoint Me.cmbW2Box14, rsERItem!W2Box14Code
    
        With Me.cmbRateDiff
            Select Case rsERItem!RateDifference
                Case PREquate.BasisAmount:      .ListIndex = 1
                Case PREquate.BasisPercent:     .ListIndex = 2
                Case Else:                      .ListIndex = 0
            End Select
        End With
    
        ' lock in the add'l amt/pct if rate diff from employer
        If nNull(rsERItem!RateDifference) <> 0 Then
            Me.tdbnumAmtPct.Enabled = False
            Me.tdbnumAmtPct = rsERItem!AmtPct
            Me.cmbRateDiff.Enabled = False
            Me.fraItmBasis.Enabled = False
            Me.optAmt.Enabled = False
            Me.optPct.Enabled = False
            Me.optHrly.Enabled = False
            Me.optHrly = True
        Else
            Me.tdbnumAmtPct.Enabled = True
            Me.tdbnumAmtPct = rsOEDED!AmtPct    ' ***
            Me.cmbRateDiff.Enabled = True
            Me.fraItmBasis.Enabled = True
            Me.optAmt.Enabled = True
            Me.optPct.Enabled = True
            Me.optHrly.Enabled = True
            Me.optHrly = False
        End If
    
        Me.cmbOECity.ListIndex = -1
        If Not IsNull(rsERItem!CityID) Then
            If rsERItem!CityID <> 0 Then
                cmbPoint Me.cmbOECity, rsERItem!CityID
            End If
        End If
    
    Else
        
        Me.chkOEDEDNoSSTax = nNull(rsOEDED!NoSSTax)
        Me.chkOEDEDNoMedTax = nNull(rsOEDED!NoMedTax)
        Me.chkOEDEDNoFWTTax = nNull(rsOEDED!NoFWTTax)
        Me.chkOEDEDNoSWTTax = nNull(rsOEDED!NoSWTTax)
        Me.chkOEDEDNoCWTTax = nNull(rsOEDED!NoCWTTax)
        Me.chkOEDEDNoFUNTax = nNull(rsOEDED!NoFUNTax)
        Me.chkOEDEDNoSUNTax = nNull(rsOEDED!NoSUNTax)
        Me.chkTips = nNull(rsOEDED!Tips)
        Me.chkNotNet = nNull(rsOEDED!NotInNet)
        Me.chkDirDepRpt = nNull(rsOEDED!DirDepRpt)
        Me.chkPension = nNull(rsOEDED!Pension)
        Me.chkSickPay = nNull(rsOEDED!SickPay)
    
        cmbPoint Me.cmbW2Box12, rsOEDED!W2Box12Code
        cmbPoint Me.cmbW2Box14, rsOEDED!W2Box14Code
        
        With Me.cmbRateDiff
            Select Case rsOEDED!RateDifference
                Case PREquate.BasisAmount:      .ListIndex = 1
                Case PREquate.BasisPercent:     .ListIndex = 2
                Case Else:                      .ListIndex = 0
            End Select
        End With
    
        Me.cmbOECity.ListIndex = -1
        If Not IsNull(rsOEDED!CityID) Then
            If rsOEDED!CityID <> 0 Then
                cmbPoint Me.cmbOECity, rsOEDED!CityID
            End If
        End If
    
    End If

    Me.tdbtxtItemComment = rsOEDED!Comment & ""

End Sub


Private Sub chkUseEmpDef_Click()
        
    If chkUseEmpDef = False Then
        
        Me.chkOEDEDNoSSTax.Enabled = True
        Me.chkOEDEDNoMedTax.Enabled = True
        Me.chkOEDEDNoFWTTax.Enabled = True
        Me.chkOEDEDNoSWTTax.Enabled = True
        Me.chkOEDEDNoCWTTax.Enabled = True
        Me.chkOEDEDNoFUNTax.Enabled = True
        Me.chkOEDEDNoSUNTax.Enabled = True
    
        Me.chkTips.Enabled = True
        Me.chkNotNet.Enabled = True
        Me.chkDirDepRpt.Enabled = True
        Me.chkPension.Enabled = True
        Me.chkSickPay.Enabled = True
    
        Me.lblW212.Enabled = True
        Me.lblW214.Enabled = True
        
        Me.cmbW2Box12.Enabled = True
        Me.cmbW2Box14.Enabled = True
        
        Me.cmdBasis.Enabled = True
        
        Me.tdbnumAmtPct.Enabled = True
        Me.cmbRateDiff.Enabled = True
    
        Me.cmbOECity.Enabled = True
    
    Else
        
        ' get the info from the employer item
        SQLString = "SELECT * FROM PRItem WHERE ItemID = " & rsOEDED!EmployerItemID
        rsInit SQLString, cn, rsERItem
        
        If rsERItem.RecordCount = 0 Then
            MsgBox "Employer Item Err: ", vbExclamation
            End
        End If
        
        Me.chkOEDEDNoSSTax = nNull(rsERItem!NoSSTax)
        Me.chkOEDEDNoMedTax = nNull(rsERItem!NoMedTax)
        Me.chkOEDEDNoFWTTax = nNull(rsERItem!NoFWTTax)
        Me.chkOEDEDNoSWTTax = nNull(rsERItem!NoSWTTax)
        Me.chkOEDEDNoCWTTax = nNull(rsERItem!NoCWTTax)
        Me.chkOEDEDNoFUNTax = nNull(rsERItem!NoFUNTax)
        Me.chkOEDEDNoSUNTax = nNull(rsERItem!NoSUNTax)
        Me.chkTips = nNull(rsERItem!Tips)
        Me.chkNotNet = nNull(rsERItem!NotInNet)
        Me.chkDirDepRpt = nNull(rsERItem!DirDepRpt)
        Me.chkPension = nNull(rsERItem!Pension)
        Me.chkSickPay = nNull(rsERItem!SickPay)
        
        cmbPoint Me.cmbW2Box12, rsERItem!W2Box12Code
        cmbPoint Me.cmbW2Box14, rsERItem!W2Box14Code
        
        Me.chkOEDEDNoSSTax.Enabled = False
        Me.chkOEDEDNoMedTax.Enabled = False
        Me.chkOEDEDNoFWTTax.Enabled = False
        Me.chkOEDEDNoSWTTax.Enabled = False
        Me.chkOEDEDNoCWTTax.Enabled = False
        Me.chkOEDEDNoFUNTax.Enabled = False
        Me.chkOEDEDNoSUNTax.Enabled = False
        
        Me.chkTips.Enabled = False
        Me.chkNotNet.Enabled = False
        Me.chkDirDepRpt.Enabled = False
        Me.chkPension.Enabled = False
        Me.chkSickPay.Enabled = False
    
        Me.lblW212.Enabled = False
        Me.lblW214.Enabled = False
    
        Me.cmbW2Box12.Enabled = False
        Me.cmbW2Box14.Enabled = False
    
        ' deduction basis
        Me.cmdBasis.Enabled = False
        
        If nNull(rsERItem!RateDifference) <> 0 Then
            Me.tdbnumAmtPct = rsERItem!AmtPct
            Me.tdbnumAmtPct.Enabled = False
        End If
        
        Me.cmbRateDiff.Enabled = False
        With Me.cmbRateDiff
            Select Case rsERItem!RateDifference
                Case PREquate.BasisAmount:      .ListIndex = 1
                Case PREquate.BasisPercent:     .ListIndex = 2
                Case Else:                      .ListIndex = 0
            End Select
        End With
    
        Me.cmbOECity.ListIndex = -1
        Me.cmbOECity.Enabled = False
        If Not IsNull(rsERItem!CityID) Then
            cmbPoint Me.cmbOECity, rsERItem!CityID
        End If
    
    End If
End Sub

Private Sub cmdCancel_Click()

'    If MsgBox("All changes will be lost - OK to exit?", vbQuestion + vbYesNo, "Employee Maintenance") = vbNo Then
'        Exit Sub
'    End If

    Unload Me
    
End Sub

Private Sub cmdOEDEDAdd_Click()

Dim ItId As Long

    AddFlag = True
    
    frmOEDEDAdd.Init
    
    If frmOEDEDAdd.rs.RecordCount = 0 Then
        MsgBox "There are no available items left for this employee", vbInformation, "Employee Add Item"
        frmOEDEDAdd.rs.Close
        Unload frmOEDEDAdd
        Exit Sub
    End If

    frmOEDEDAdd.Show vbModal

    Unload frmOEDEDAdd

    ' canceled
    If TaskID = 0 Then Exit Sub

    ItId = TaskID
    
    If ItId = -1 Then Exit Sub ' cancel selected

    ' make sure this item not in disconnect record set
    ' --> added this session
    SQLString = "EmployerItemID = " & ItId
    rsOEDED.Find SQLString, 0, adSearchForward, 1
    If Not rsOEDED.EOF Then
        MsgBox "Item already selected for this employee!", vbExclamation
        Exit Sub
    End If
    
    ' use the PRItem class for the employer item
    If Not PRItem.GetByID(ItId) Then
        MsgBox "Employer Item error! " & ItId, vbExclamation
        End
    End If
    
    ' add it to the disconnected RecordSet
    rsOEDED.AddNew
    
    rsOEDED!EmployeeID = PREmployee.EmployeeID
    rsOEDED!Title = PRItem.Title
    rsOEDED!Abbreviation = PRItem.Abbreviation
    rsOEDED!Active = 1
    rsOEDED!ItemType = PRItem.ItemType
    rsOEDED!NoSSTax = PRItem.NoSSTax
    rsOEDED!NoMedTax = PRItem.NoMedTax
    rsOEDED!NoFWTTax = PRItem.NoFWTTax
    rsOEDED!NoSWTTax = PRItem.NoSWTTax
    rsOEDED!NoCWTTax = PRItem.NoCWTTax
    rsOEDED!NoFUNTax = PRItem.NoFUNTax
    rsOEDED!NoSUNTax = PRItem.NoSUNTax
    rsOEDED!Basis = PRItem.Basis
    rsOEDED!Tips = PRItem.Tips
    rsOEDED!NotInNet = PRItem.NotInNet
    rsOEDED!SDNumber = PRItem.SDNumber
    rsOEDED!EmployerItemID = PRItem.ItemID
    rsOEDED!Pension = PRItem.Pension
    rsOEDED!SickPay = PRItem.SickPay
    
    cmbPoint Me.cmbW2Box12, PRItem.W2Box12Code
    cmbPoint Me.cmbW2Box14, PRItem.W2Box14Code
    
    rsOEDED!UseEmployer = 1
    
    rsOEDED!MatchPct = 0
    rsOEDED!MaxPct = 0
    rsOEDED!MaxAmount = 0
    rsOEDED!AmtPct = 0
    
    rsOEDED.Update
    
    Unload frmOEDEDAdd

    AddFlag = False

    ItemDisplay

End Sub

Private Sub cmdOEDEDDelete_Click()

    ' trap if added on this session
    If IsNull(rsOEDED!ItemID) Then Exit Sub
    If rsOEDED!ItemID = 0 Then Exit Sub

    ' can not delete if any history exists !!!
    If rsOEDED!ItemType = PREquate.ItemTypeOE Then
        SQLString = "SELECT * FROM PRDist WHERE PRDist.EmployeeID = " & PREmployee.EmployeeID & _
                    " AND PRDist.ItemID = " & rsOEDED!ItemID
        If PRDist.GetBySQL(SQLString) Then
            MsgBox "Earning detail exists - delete not allowed!", vbExclamation
            Exit Sub
        End If
    Else
        SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.EmployeeID = " & PREmployee.EmployeeID & _
                    " AND PRItemHist.ItemID = " & rsOEDED!ItemID
    
        If PRItemHist.GetBySQL(SQLString) Then
            MsgBox "Deduction detail exists - delete not allowed!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' ok to proceed ???
    If MsgBox("OK to delete: " & rsOEDED!Title & "?", vbYesNo + vbQuestion, "Delete OE/Deduction") = vbNo Then
        Exit Sub
    End If
    
    ' mark for deletion
    rsOEDED!Title = "* DELETED *"
    rsOEDED!SDNumber = 99
    rsOEDED.Update

End Sub

Private Sub cmdSave_Click()
    
Dim Hrs As New ADODB.Recordset
    
    SaveFlag = True
    
    ' data verifications
    
    ' can't change the employee number
    ' if history exists
    If Me.lngEmployeeNumber <> OrigEmpNum Then
        SQLString = "SELECT * FROM PRHist WHERE PRHist.EmployeeID = " & SelID
        rsInit SQLString, cn, Hrs
        If Hrs.RecordCount > 0 Then
            MsgBox "Employee # change not allowed if history exists!", vbExclamation
            Exit Sub
        End If
    End If
    
    ' department dropdown - blank is OK
    With Me.cmbDept
        PREmployee.DepartmentID = .ItemData(.ListIndex)
    End With
    
    ' save data
    ' Main Screen Tab
    PREmployee.EmployeeNumber = Me.lngEmployeeNumber
    PREmployee.LastName = Me.txtLastName.text
    PREmployee.FirstName = Me.txtFirstName.text
    PREmployee.MidInit = Me.txtMI
    PREmployee.AltName = Me.txtAltName
    PREmployee.UseAltName = Me.chkUseAltName
    PREmployee.CheckComment = Me.txtCheckComment
    PREmployee.Address1 = Me.txtAddress1
    PREmployee.Address2 = Me.txtAddress2
    PREmployee.City = Me.txtCity
    PREmployee.State = Me.cmbState.text
    
    If IsNull(Me.tdbnumZipCode) Then
        PREmployee.ZipCode = 0
    Else
        PREmployee.ZipCode = Me.tdbnumZipCode
    End If
    
    PREmployee.SSN = CLng(Replace(Me.txtSSN, "-", ""))
    
    PREmployee.x1099Employee = Me.cmb1099.ListIndex
    
    ' Pay Parameters Screen Tab
    PREmployee.SalaryAmount = Me.curSalaryAmt
    PREmployee.HourlyAmount = Me.curHourlyAmt
    PREmployee.Inactive = Me.chkInactive
    PREmployee.Salaried = Me.chkSalaried
    PREmployee.NoSSTax = Me.chkNoSSTax
    PREmployee.NoMedTax = Me.chkNoMedTax
    PREmployee.NoFedTax = Me.chkNoFedTax
    PREmployee.NoStateTax = Me.chkNoStateTax
    PREmployee.NoCityTax = Me.chkNoCityTax
    PREmployee.NoFedUnemp = Me.chkNoFedUnemp
    PREmployee.NoStateUnemp = Me.chkNoStateUnemp
    
    ' comment
    SQLString = "SELECT * FROM Notes WHERE NoteType = " & Equate.NoteTypeEE & _
                "AND DateTm = 0 " & _
                "AND RelatedID = " & PREmployee.EmployeeID
    If Notes.GetBySQL(SQLString) = False Then
        Notes.Clear
        Notes.NoteType = Equate.NoteTypeEE
        Notes.RelatedID = PREmployee.EmployeeID
        Notes.Save (Equate.RecAdd)
    End If
    Notes.Notation = Me.tdbtxtComment
    Notes.Save (Equate.RecPut)
    
    ' *** FWT ***
    PREmployee.FWTMarried = Me.chkFWTMarried
        
    If Me.optFWTExemptions Then
        PREmployee.FWTBasis = PREquate.BasisExemptions
    Else
        PREmployee.FWTBasis = PREquate.BasisPercent
    End If
    PREmployee.FWTAmount = Me.tdbnumFWTAmount
    
    If Me.optFWTAddAmount Then
        PREmployee.FWTExtraBasis = PREquate.BasisAmount
    Else
        PREmployee.FWTExtraBasis = PREquate.BasisPercent
    End If
    PREmployee.FWTExtraAmount = Me.tdbnumFWTExtraAmount
    
    ' *** SWT ***
    PREmployee.SWTMarried = Me.chkSWTMarried
    
    If Me.optSWTExemptions Then
        PREmployee.SWTBasis = PREquate.BasisExemptions
    Else
        PREmployee.SWTBasis = PREquate.BasisPercent
    End If
    PREmployee.SWTAmount = Me.tdbnumSWTAmount
    
    If Me.optSWTAddAmount Then
        PREmployee.SWTExtraBasis = PREquate.BasisAmount
    Else
        PREmployee.SWTExtraBasis = PREquate.BasisPercent
    End If
    PREmployee.SWTExtraAmount = Me.tdbnumSWTExtraAmount
    
    With Me.cmbEEDfltCity
        PREmployee.DefaultCityID = .ItemData(.ListIndex)
    End With
        
    With Me.cmbCourtCWT
        PREmployee.CourtesyCityID = .ItemData(.ListIndex)
    End With
        
    PREmployee.CourtesyAdd = Me.chkCourtAdd
        
    ' Dates and Other Screen Tab
    If Me.cmbTermReason <> "" Then
        PREmployee.TermReason = Me.cmbTermReason
    End If
    
    If Me.cmbSex <> "" Then
        PREmployee.Sex = Me.cmbSex
    End If
    
    If Me.cmbRaceCode <> "" Then
        PREmployee.RaceCode = Me.cmbRaceCode
    End If
    
    If Me.cmbMaritalStatus <> "" Then
        PREmployee.MaritalStatus = Me.cmbMaritalStatus
    End If
    
    If Me.cmbEducationLevel <> "" Then
        PREmployee.EducationLevel = Me.cmbEducationLevel
    End If
    
    If Me.cmbShiftCode <> "" Then
        PREmployee.ShiftCode = Me.cmbShiftCode
    End If

    '  checking for null date values and setting them to zero
    If Me.DteDateHired.ValueIsNull Then
       PREmployee.DateHired = 0
    Else
       PREmployee.DateHired = Me.DteDateHired.Value
    End If
    
    If Me.dteDateLastPaid.ValueIsNull Then
        PREmployee.DateLastPaid = 0
    Else
        PREmployee.DateLastPaid = Me.dteDateLastPaid
    End If
    
    If Me.dteDateLastRaise.ValueIsNull Then
        PREmployee.DateLastRaise = 0
    Else
        PREmployee.DateLastRaise = Me.dteDateLastRaise
    End If
        
    If Me.dteDateLastRecall.ValueIsNull Then
        PREmployee.DateLastRecall = 0
    Else
        PREmployee.DateLastRecall = Me.dteDateLastRecall
    End If
    
    If Me.dteDateLastReview.ValueIsNull Then
        PREmployee.DateLastReview = 0
    Else
        PREmployee.DateLastReview = Me.dteDateLastReview
    End If
    
    If Me.dteDateLastLayoff.ValueIsNull Then
        PREmployee.DateLastLayoff = 0
    Else
        PREmployee.DateLastLayoff = Me.dteDateLastLayoff
    End If
    
    If Me.dteDateTerminated.ValueIsNull Then
        PREmployee.DateTerminated = 0
    Else
        PREmployee.DateTerminated = Me.dteDateTerminated
    End If
    
    If Me.DteDateofBirth.ValueIsNull Then
        PREmployee.DateOfBirth = 0
    Else
        PREmployee.DateOfBirth = Me.DteDateofBirth
    End If

    ' pays per year
    PREmployee.PaysPerYear = Me.cmbPPY.text

    ' EIC filing type
    PREmployee.EICType = Me.cmbEICType.ListIndex

    ' statutory employee
    PREmployee.Statutory = Me.chkStatutory

    ' Wkc Cat
    If Me.chkUseDeptWkc Then        ' use the department cat
        PREmployee.WkcUseDept = 1
        PREmployee.WkcCat = 0
    Else                            ' specific to the employee
        PREmployee.WkcUseDept = 0
        With Me.cmbWkcCat
            PREmployee.WkcCat = .ItemData(.ListIndex)
        End With
    End If

    With Me.cmbEEDefaultJob
        If .Visible = True Then
            PREmployee.DefaultJobID = .ItemData(.ListIndex)
        End If
    End With

    PREmployee.Save (Equate.RecPut)

    ' check for OE/DED deletes and Dir Deposit deletes
    If Not (rsOEDED.EOF And rsOEDED.BOF) Then
        If rsOEDED.RecordCount > 0 And rsOEDED!SDNumber <> 99 Then
            ItemSave
        End If
    End If
    
    ' save Dir Dep info currently selected
    If Not (rsDirDep.EOF And rsDirDep.BOF) Then
        If rsDirDep.RecordCount > 0 Then DirDepSave
    End If
    
    ' OE / DED deletes - deduction basis
    SkipSave = True     ' don't bother with screen to record set updates here ...
    If rsOEDED.RecordCount > 0 Then
        rsOEDED.MoveFirst
        Do
            
            If rsOEDED!SDNumber = 99 Then
                
                ' delete the deduction basis record
                If rsOEDED!ItemType = PREquate.ItemTypeDED And rsOEDED!UseEmployer = 1 Then
                    SQLString = "DELETE * FROM PRGlobal WHERE " & _
                                " TypeCode = " & PREquate.GlobalTypeDeductBasis & _
                                " AND UserID = " & PRCompany.CompanyID & _
                                " AND Description = '" & rsOEDED!EmployerItemID & "'" & _
                                " AND Var1 = '" & PREmployee.EmployeeID & "'"
                    cnDes.Execute SQLString
                End If
            
                rsOEDED.Delete
            
            End If
            
            rsOEDED.MoveNext
            If rsOEDED.EOF Then Exit Do
        Loop
    End If
    
    ' direct deposit deletes
    If rsDirDep.RecordCount > 0 Then
        rsDirDep.MoveFirst
        Do
            If rsDirDep!DirDepBasis = 99 Then rsDirDep.Delete
            rsDirDep.MoveNext
            If rsDirDep.EOF Then Exit Do
        Loop
    End If
    
    ' revised W4
    PRW4.FilingType = Me.cmbPRW4.ListIndex
    PRW4.TwoJobs = Me.chkW4_2Jobs
    PRW4.Dependents = Me.tdbnumW4Dependents
    PRW4.DependentsOther = Me.tdbnumW4OtherDependents
    PRW4.OtherIncome = Me.tdbnumW4OtherIncome
    PRW4.Deductions = Me.tdbnumW4Deductions
    PRW4.ExtraWH = Me.tdbnumW4ExtraWH
    PRW4.Save (Equate.RecPut)
    
    SkipSave = False

    ' save the disconnected record sets
    rsSave rsOEDED, cn
    rsSave rsDirDep, cn
    
    Unload Me

End Sub

Private Sub fgOEDED_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    'ItemDisplay
End Sub

Private Sub fgOEDED_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    ' save it to disconn record set
    'ItemSave
End Sub
Private Sub fgOEDED_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    ItemDisplay
End Sub

Private Sub fgOEDED_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    ItemSave
End Sub

Private Sub cmdDirDepAdd_Click()

Dim NetCount As Integer

    NetCount = 0
    If rsDirDep.RecordCount > 0 Then
'        rsDirDep.MoveFirst
'        Do
'            If rsDirDep!DirDepBasis = PREquate.DirDepBasisNet Then
'                NetCount = NetCount + 1
'            End If
'            rsDirDep.MoveNext
'        Loop Until rsDirDep.EOF
    
        ' @@@@@@@@@@@@@@@@@
        rsDirDep.Update
    
    Else
        
        ' hide the fields if no records
        With Me
            .chkDirDepActive.Enabled = True
            .txtBankName.Enabled = True
            .txtABA.Enabled = True
            .txtAccount.Enabled = True
            .fraType.Enabled = True
            .optChecking.Enabled = True
            .optSavings.Enabled = True
            .fraBasis.Enabled = True
            .optDDAmount.Enabled = True
            .optDDPercent.Enabled = True
            .optDDNet.Enabled = True
            .tdbnumDDAmount.Enabled = True
        End With
    
    End If

    rsDirDep.AddNew
    
    rsDirDep!Active = 1
    
    rsDirDep!ItemType = PREquate.ItemTypeDirDepDed
    rsDirDep!DirDepType = PREquate.DirDepTypeChecking
    rsDirDep!DirDepBank = ""
    rsDirDep!DirDepABA = ""
    rsDirDep!DirDepAccount = ""
    
    If rsDirDep.RecordCount = 0 Then
        rsDirDep!DirDepBasis = PREquate.DirDepBasisNet
    ElseIf NetCount = 0 Then
        rsDirDep!DirDepBasis = PREquate.DirDepBasisNet
    Else
        rsDirDep!DirDepBasis = PREquate.DirDepBasisNet
    End If
    
    rsDirDep!DirDepAmtPct = 0
    rsDirDep.Update
    
    DirDepDisplay

End Sub

Private Sub cmdDirDepDelete_Click()

    ' trap if added on this session
    If IsNull(rsDirDep!ItemID) Then Exit Sub
    If rsDirDep!ItemID = 0 Then Exit Sub

    ' deny if history exists
    SQLString = "SELECT * FROM PRItemHist WHERE PRItemHist.EmployeeID = " & PREmployee.EmployeeID & _
                " AND PRItemHist.ItemID = " & rsDirDep!ItemID
    If PRItemHist.GetBySQL(SQLString) Then
        MsgBox "Can't delete - detail exists", vbExclamation
        Exit Sub
    End If

    If MsgBox("OK to delete " & Trim(rsDirDep!DirDepBank) & " ?", vbQuestion + vbYesNo, "Delete Direct Deposit Record") = vbNo Then
        Exit Sub
    End If

    rsDirDep!DirDepBank = "* DELETED *"
    rsDirDep!DirDepBasis = 99
    rsDirDep.Update

End Sub

Private Sub optAmt_Click()
    If Me.optAmt = True Then Me.tdbnumAmtPct.Caption = "Amount"
End Sub

Private Sub optFWTAddAmount_Click()
    Me.tdbnumFWTExtraAmount.Format = "$ ##,##0.00"
    Me.tdbnumFWTExtraAmount.DisplayFormat = ""
    Me.tdbnumFWTExtraAmount = 0
End Sub

Private Sub optFWTAddPercent_Click()
    Me.tdbnumFWTExtraAmount.Format = "##0.00 %"
    Me.tdbnumFWTExtraAmount.DisplayFormat = ""
    Me.tdbnumFWTExtraAmount = 0
End Sub

Private Sub optFWTExemptions_Click()
    Me.tdbnumFWTAmount.Format = "##0"
    Me.tdbnumFWTAmount.DisplayFormat = ""
    Me.tdbnumFWTAmount = 0
End Sub

Private Sub optFWTPercent_Click()
    Me.tdbnumFWTAmount.Format = "##0.00 %"
    Me.tdbnumFWTAmount.DisplayFormat = ""
    Me.tdbnumFWTAmount = 0
End Sub


Private Sub optswtAddAmount_Click()
    Me.tdbnumSWTExtraAmount.Format = "$ ##,##0.00"
    Me.tdbnumSWTExtraAmount.DisplayFormat = ""
    Me.tdbnumSWTExtraAmount = 0
End Sub

Private Sub optswtAddPercent_Click()
    Me.tdbnumSWTExtraAmount.Format = "##0.00 %"
    Me.tdbnumSWTExtraAmount.DisplayFormat = ""
    Me.tdbnumSWTExtraAmount = 0
End Sub

Private Sub optswtExemptions_Click()
    Me.tdbnumSWTAmount.Format = "##0"
    Me.tdbnumSWTAmount.DisplayFormat = ""
    Me.tdbnumSWTAmount = 0
End Sub

Private Sub optswtPercent_Click()
    Me.tdbnumSWTAmount.Format = "##0.00 %"
    Me.tdbnumSWTAmount.DisplayFormat = ""
    Me.tdbnumSWTAmount = 0
End Sub

Private Sub optHrly_Click()
    If Me.optHrly = True Then Me.tdbnumAmtPct.Caption = "Rate"
End Sub

Private Sub optPct_Click()
    If Me.optPct = True Then Me.tdbnumAmtPct.Caption = "Percent"
End Sub

Private Sub optDDAmount_Click()
    If Me.optDDAmount = True Then
        Me.tdbnumDDAmount.Caption = "Amount"
        Me.tdbnumDDAmount.Enabled = True
    End If
End Sub

Private Sub optDDNet_Click()
    If Me.optDDNet = True Then
        Me.tdbnumDDAmount.Enabled = False
        Me.tdbnumDDAmount = 0
    End If
End Sub

Private Sub optDDPercent_Click()
    If Me.optDDPercent = True Then
        Me.tdbnumDDAmount.Caption = "Percent"
        Me.tdbnumDDAmount.Enabled = True
    End If
End Sub

Private Sub DropDownInit()
    
Dim RecFlag As Boolean
Dim x As String
    
    ' ******************************************************************************
    ' rate difference drop down
    With Me.cmbRateDiff
        .AddItem "N / A"
        .AddItem "Amount"
        .AddItem "Percent"
    End With
    
    ' ******************************************************************************
    ' Department Drop Down
    With Me.cmbDept
        
        .AddItem "NONE"
        .ItemData(.NewIndex) = 0
        SQLString = "SELECT * FROM PRDepartment ORDER BY Name"
        If PRDepartment.GetBySQL(SQLString) Then
            Do
                .AddItem Trim(PRDepartment.Name) & " " & PRDepartment.DepartmentNumber
                .ItemData(.NewIndex) = PRDepartment.DepartmentID
                If Not PRDepartment.GetNext Then Exit Do
            Loop
        End If
        
    End With
    cmbPoint Me.cmbDept, PREmployee.DepartmentID
    
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
    ' **********************************************************************************************
    ' Workers Comp Drop Down
    
    With Me.cmbWkcCat
        
        .AddItem "NONE"
        .ItemData(.NewIndex) = 0
    
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeWkcCat & _
                    " ORDER BY Description"

        If PRGlobal.GetBySQL(SQLString) Then
            Do
                .AddItem PRGlobal.Description & "   " & Format(PRGlobal.Percent, "##0.0000") & " %"
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If Not PRGlobal.GetNext Then Exit Do
            Loop
        End If
    
    End With
    
    ' initialize the selection
    If PREmployee.WkcUseDept Then
        If PRDepartment.GetByID(PREmployee.DepartmentID) Then   ' take from the department record
            Me.chkUseDeptWkc = 1
            cmbPoint Me.cmbWkcCat, PRDepartment.WkcCat
        Else    ' dept NF
            Me.chkUseDeptWkc = 0
            PREmployee.WkcUseDept = 0
            PREmployee.WkcCat = 0
            Me.cmbWkcCat.ListIndex = 0
        End If
    Else            ' take from the employee record
        Me.chkUseDeptWkc = 0
        cmbPoint Me.cmbWkcCat, PREmployee.WkcCat
    End If
    
    ' **********************************************************************************************
    ' Employee Default City / State init
    cmbCityInit Me.cmbEEDfltCity, PREmployee.DefaultCityID
    cmbCityInit Me.cmbCourtCWT, PREmployee.CourtesyCityID
    cmbCityInit Me.cmbOECity, 0
    
    ' init the PPY dropdown
    ' regular Windows Combo
    ' style 2 - user can not type in field
    cmbPPYSet Me.cmbPPY, PREmployee.PaysPerYear
    
    ' init the EIC drop down
    ' regular Windows Combo
    ' style 2 - user can not type in field
    ' cmb selection will correspone to the correct field assignment
    Me.cmbEICType.AddItem "NONE"                    ' = 0
    Me.cmbEICType.AddItem "SINGLE"                  ' = 1
    Me.cmbEICType.AddItem "MARRIED WITH"            ' = 2
    Me.cmbEICType.AddItem "MARRIED WITHOUT"         ' = 3
    Me.cmbEICType.ListIndex = PREmployee.EICType
    
    ' 1099 drop down
    ' regular Windows Combo
    ' style 2 - user can not type in field
    ' cmb selection will correspone to the correct field assignment
    Me.cmb1099.AddItem "N / A"          ' = 0
    Me.cmb1099.AddItem "1099 Regular"   ' = 1
    Me.cmb1099.AddItem "1099 Inc"       ' = 2
    Me.cmb1099.ListIndex = PREmployee.x1099Employee
    
    Me.cmbPRW4.AddItem "Standard"
    Me.cmbPRW4.AddItem "Single/Married Separate"
    Me.cmbPRW4.AddItem "Married Joint/Widow(er)"
    Me.cmbPRW4.AddItem "Head of Household"
    Me.cmbPRW4.ListIndex = PRW4.FilingType
    
End Sub

Private Sub GridInit()

    ' ******************************************************************************
    ' OE / DED Grid
    ' open as disconnected for "All or Nothing" save
    DisConn = True
    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = " & PREmployee.EmployeeID & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " OR PRITem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeSDTax & ")" & _
                " ORDER BY PRItem.ItemType, PRItem.ItemID"
    
    rsInit SQLString, cn, rsOEDED
    
    ' fill in the titles from the employer record
    If Not (rsOEDED.BOF And rsOEDED.EOF) Then
        rsOEDED.MoveFirst
        Do
            If PRItem.GetByID(rsOEDED!EmployerItemID) Then
                rsOEDED!Title = PRItem.Title
                rsOEDED.Update
            End If
            rsOEDED.MoveNext
            If rsOEDED.EOF Then Exit Do
        Loop
    End If
    
    If rsOEDED.RecordCount > 0 Then rsOEDED.MoveFirst
    
    SetGrid rsOEDED, Me.fgOEDED
    
    ' grid formatting
    Me.fgOEDED.ColHidden(0) = True  ' hide the ItemID column
    Me.fgOEDED.ColHidden(1) = True  ' hide the EmployeeID column
    Me.fgOEDED.ColHidden(3) = True  ' hide the Abbrev column
    
    Me.fgOEDED.TextMatrix(0, 2) = "Item Name"
    Me.fgOEDED.ColWidth(2) = 3000
    
    Me.fgOEDED.TextMatrix(0, 4) = "Item Type"
    Me.fgOEDED.ColWidth(4) = 1400
    
    Me.fgOEDED.TextMatrix(0, 5) = "Active"
    Me.fgOEDED.ColWidth(5) = 900
    
    Me.fgOEDED.SelectionMode = flexSelectionByRow
    Me.fgOEDED.ScrollBars = flexScrollBarVertical
    Me.fgOEDED.Editable = flexEDNone
    
    ' make a drop down string to translate active = 0 / 1 to No / Yes
    ActiveDrop = "|#0;No|#1;Yes"
    fgOEDED.ColComboList(5) = ActiveDrop
    
    ' make a drop down for item type - only two of them!!!
    TypeDrop = "|#" & PREquate.ItemTypeOE & ";Other Earning|#" & PREquate.ItemTypeDED & ";Deduction"
    TypeDrop = Trim(TypeDrop) & "|#" & PREquate.ItemTypeSDTax & ";SD Tax"
    fgOEDED.ColComboList(4) = TypeDrop
    
    If rsOEDED.RecordCount > 0 Then
        rsOEDED.MoveFirst
        ItemDisplay
    End If
    
    Me.fgOEDED.AllowSelection = False
    
    ' ******************************************************************************
    ' DirDep Grid
    ' open as disconnected for "All or Nothing" save
    DisConn = True
    
    SQLString = "SELECT DirDepBank, DirDepType, Active, DirDepABA, DirDepAccount, DirDepBasis, " & _
                " DirDepAmtPct, ItemID, ItemType, EmployeeID FROM PRItem " & _
                " WHERE PRItem.EmployeeID = " & PREmployee.EmployeeID & _
                " AND PRItem.ItemType = " & PREquate.ItemTypeDirDepDed & _
                " ORDER BY PRItem.Active DESC, PRItem.DirDepType, PRItem.ItemID"
    
    rsInit SQLString, cn, rsDirDep
    
    If rsDirDep.RecordCount > 0 Then
        rsDirDep.MoveFirst
    Else
        ' hide the fields if no records
        With Me
            .chkDirDepActive.Enabled = False
            .txtBankName.Enabled = False
            .txtABA.Enabled = False
            .txtAccount.Enabled = False
            .fraType.Enabled = False
            .optChecking.Enabled = False
            .optSavings.Enabled = False
            .fraBasis.Enabled = False
            .optDDAmount.Enabled = False
            .optDDPercent.Enabled = False
            .optDDNet.Enabled = False
            .tdbnumDDAmount.Enabled = False
        End With
    End If
    
    SetGrid rsDirDep, Me.fgDirDep
    
    ' grid formatting
    Me.fgDirDep.TextMatrix(0, 0) = "Bank Name"
    Me.fgDirDep.ColWidth(0) = 2500
    
    Me.fgDirDep.TextMatrix(0, 1) = "Type"
    Me.fgDirDep.ColWidth(1) = 1300
    
    Me.fgDirDep.TextMatrix(0, 2) = "Active"
    Me.fgDirDep.ColWidth(2) = 700
    
    Me.fgDirDep.SelectionMode = flexSelectionByRow
    Me.fgDirDep.ScrollBars = flexScrollBarVertical
    Me.fgDirDep.Editable = flexEDNone
    
    ' make a drop down string to translate active = 0 / 1 to No / Yes
    ActiveDrop = "|#0;No|#1;Yes"
    fgDirDep.ColComboList(2) = ActiveDrop
    
    ' Checking/Savings display
    TypeDrop = "|#" & PREquate.DirDepTypeChecking & ";Checking|#" & PREquate.DirDepTypeSavings & ";Savings"
    fgDirDep.ColComboList(1) = TypeDrop

    If rsDirDep.RecordCount > 0 Then
        rsDirDep.MoveFirst
        DirDepDisplay
    End If

    Me.fgDirDep.AllowSelection = False

End Sub

Private Sub ItemSave()

    If SkipSave = True Then Exit Sub

    ' always take it from the screen
    ' if from employer was picked - the screen was updated with it
    
    If rsOEDED.RecordCount = 0 Then Exit Sub
    
    rsOEDED!Active = Me.chkActive
    rsOEDED!AmtPct = Me.tdbnumAmtPct
    rsOEDED!MaxAmount = Me.tdbnumMaxAmt
    If Me.optAmt = True Then
        rsOEDED!Basis = PREquate.BasisAmount
    ElseIf Me.optPct = True Then
        rsOEDED!Basis = PREquate.BasisPercent
    Else
        rsOEDED!Basis = PREquate.BasisHourly
    End If
    
    rsOEDED!UseEmployer = Me.chkUseEmpDef
    
    rsOEDED!NoSSTax = Me.chkOEDEDNoSSTax
    rsOEDED!NoMedTax = Me.chkOEDEDNoMedTax
    rsOEDED!NoFWTTax = Me.chkOEDEDNoFWTTax
    rsOEDED!NoSWTTax = Me.chkOEDEDNoSWTTax
    rsOEDED!NoCWTTax = Me.chkOEDEDNoCWTTax
    rsOEDED!NoFUNTax = Me.chkOEDEDNoFUNTax
    rsOEDED!NoSUNTax = Me.chkOEDEDNoSUNTax
    rsOEDED!NotInNet = Me.chkNotNet
    rsOEDED!DirDepRpt = Me.chkDirDepRpt
    rsOEDED!Tips = Me.chkTips
    rsOEDED!Pension = Me.chkPension
    rsOEDED!SickPay = Me.chkSickPay
    
    With Me.cmbW2Box12
        If .ListIndex >= 0 Then
            rsOEDED!W2Box12Code = .ItemData(.ListIndex)
        Else
            rsOEDED!W2Box12Code = 0
        End If
    End With
    
    With Me.cmbW2Box14
        If .ListIndex >= 0 Then
            rsOEDED!W2Box14Code = .ItemData(.ListIndex)
        Else
            rsOEDED!W2Box14Code = 0
        End If
    End With
    
    rsOEDED!Comment = Trim(Me.tdbtxtItemComment)
    
    With Me.cmbRateDiff
        Select Case .ListIndex
            Case 0:     rsOEDED!RateDifference = 0
            Case 1:     rsOEDED!RateDifference = PREquate.BasisAmount
            Case 2:     rsOEDED!RateDifference = PREquate.BasisPercent
        End Select
    End With
    
    rsOEDED!CityID = 0
    With Me.cmbOECity
        If .ListIndex > 0 Then
            rsOEDED!CityID = .ItemData(.ListIndex)
        End If
    End With
    
    rsOEDED.Update

End Sub

Private Sub txtBankName_Change()
    If InitFlag = True Then Exit Sub
    If rsDirDep.RecordCount = 0 Then Exit Sub
    rsDirDep!DirDepBank = Trim(Me.txtBankName)
End Sub

Private Sub optChecking_Click()
    If Me.optChecking = True Then
        rsDirDep!DirDepType = PREquate.DirDepTypeChecking
    Else
        rsDirDep!DirDepType = PREquate.DirDepTypeSavings
    End If
End Sub
Private Sub optSavings_Click()
    If Me.optChecking = True Then
        rsDirDep!DirDepType = PREquate.DirDepTypeChecking
    Else
        rsDirDep!DirDepType = PREquate.DirDepTypeSavings
    End If
End Sub

Private Sub DirDepDisplay()
    
    If rsDirDep.RecordCount = 0 Then Exit Sub
    If IsNull(rsDirDep!Active) Then Exit Sub
    
    Me.chkDirDepActive = rsDirDep!Active
    Me.txtBankName = rsDirDep!DirDepBank & ""
    Me.txtABA = rsDirDep!DirDepABA & ""
    
    Me.txtAccount = rsDirDep!DirDepAccount
    
    If IsNull(rsDirDep!DirDepAmtPct) Then
        Me.tdbnumDDAmount = 0
    Else
        Me.tdbnumDDAmount = rsDirDep!DirDepAmtPct
    End If
    
    If rsDirDep!DirDepType = PREquate.DirDepTypeChecking Then
        Me.optChecking = True
    Else
        Me.optSavings = True
    End If
    
    If rsDirDep!DirDepBasis = PREquate.DirDepBasisAmt Then
        Me.optDDAmount = True
        Me.tdbnumDDAmount.Caption = "Amount"
    ElseIf rsDirDep!DirDepBasis = PREquate.DirDepBasisPct Then
        Me.optDDPercent = True
        Me.tdbnumDDAmount.Caption = "Percent"
    Else
        Me.optDDNet = True
        Me.tdbnumDDAmount.Caption = "Amount"
        Me.tdbnumDDAmount = 0
        Me.tdbnumDDAmount.Enabled = False
    End If
    
End Sub

Private Sub DirDepSave()

    If SkipSave = True Then Exit Sub
    If rsDirDep.RecordCount = 0 Then Exit Sub

    ' dont do if pending delete
    If rsDirDep!DirDepBasis = 99 Then Exit Sub

    rsDirDep!EmployeeID = SelID
    rsDirDep!Active = Me.chkDirDepActive
    rsDirDep!DirDepBank = Trim(Me.txtBankName)
    rsDirDep!DirDepABA = Trim(Me.txtABA)
    
    rsDirDep!DirDepAccount = Trim(Me.txtAccount)
    If Me.optChecking = True Then
        rsDirDep!DirDepType = PREquate.DirDepTypeChecking
    Else
        rsDirDep!DirDepType = PREquate.DirDepTypeSavings
    End If
    If Me.optDDAmount = True Then
        rsDirDep!DirDepBasis = PREquate.DirDepBasisAmt
    ElseIf Me.optDDPercent = True Then
        rsDirDep!DirDepBasis = PREquate.DirDepBasisPct
    Else
        rsDirDep!DirDepBasis = PREquate.DirDepBasisNet
    End If
    rsDirDep!DirDepAmtPct = Me.tdbnumDDAmount
    rsDirDep.Update
    
End Sub

Private Sub fgDirDep_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)

    DirDepDisplay

End Sub

Private Sub fgDirDep_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)

    ' save it to disconn record set
    DirDepSave

End Sub

Private Sub chkUseDeptWkc_Click()

    If Me.chkUseDeptWkc Then    ' use the department cat

        ' get the department selected on the screen
        J = Me.cmbDept.ItemData(Me.cmbDept.ListIndex)
        
        If J = 0 Then   ' no dept selected - cant do it
            Me.chkUseDeptWkc = 0
            Exit Sub
        ElseIf PRDepartment.GetByID(J) Then   ' take from the department record
            With Me.cmbWkcCat
                .Enabled = False
                .ListIndex = 0
                For I = 0 To .ListCount - 1
                    If .ItemData(I) = PRDepartment.WkcCat Then
                        .ListIndex = I
                        Exit For
                    End If
                Next I
            End With
        Else    ' dept NF
            Me.chkUseDeptWkc = 0
            PREmployee.WkcUseDept = 0
            PREmployee.WkcCat = 0
            Me.cmbWkcCat.Enabled = True
        End If
    Else                        ' per the employee
        Me.cmbWkcCat.Enabled = True
    End If
    Me.cmbWkcCat.Refresh

End Sub

Private Sub cmbCityInit(ByRef cmb As ComboBox, CtyID As Long)
    
    With cmb
        
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
        
    cmbPoint cmb, CtyID
        
End Sub

Private Sub cmdBasis_Click()
    If rsOEDED!ItemType <> PREquate.ItemTypeDED Then Exit Sub
    frmDeductBasis.EmployeeID = PREmployee.EmployeeID
    frmDeductBasis.ItemID = rsOEDED!EmployerItemID
    frmDeductBasis.Show vbModal
End Sub


Private Sub txtSSN_GotFocus()
    Me.txtSSN.SelStart = 0
    Me.txtSSN.SelLength = 11
End Sub
