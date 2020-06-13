VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMainMenu 
   Caption         =   "Balint Windows Accounting"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808000&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   600
      Picture         =   "frmMainMenu.frx":030A
      ScaleHeight     =   675
      ScaleWidth      =   2595
      TabIndex        =   77
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   12360
      TabIndex        =   0
      Top             =   10080
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "FILE"
      TabPicture(0)   =   "frmMainMenu.frx":1028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label17"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label18"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdFIUserMt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFiPSSWD"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFICopy"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdFiNew"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdFIOpen"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdSDGLImport"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdSDGLHImport"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdSDPRImport"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkHideGL"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkHidePR"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdDelete"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkHideJC"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSDPRHImport"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdSDGLFFImport"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdNewADO"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdJimBo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdBackUp"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "GENERAL LEDGER"
      TabPicture(1)   =   "frmMainMenu.frx":1342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGLMtStatements"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdGLDataEntry"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdFreeFormat"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "PAYROLL"
      TabPicture(2)   =   "frmMainMenu.frx":135E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdPREntry"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdPR1099"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "JOB COST"
      TabPicture(3)   =   "frmMainMenu.frx":137A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdJCMaint"
      Tab(3).Control(1)=   "cmdJCWageRpt"
      Tab(3).Control(2)=   "cmdJCJobMaint"
      Tab(3).Control(3)=   "cmdTimeSheetEntry"
      Tab(3).Control(4)=   "cmdJCTSReport"
      Tab(3).Control(5)=   "cmdPWMaint"
      Tab(3).Control(6)=   "cmdQBTaxPay"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "INVOICING"
      TabPicture(4)   =   "frmMainMenu.frx":1396
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdKPInvProcess"
      Tab(4).Control(1)=   "cmdKPInvStockMaint"
      Tab(4).Control(2)=   "cmdKPInvGlobalMaint"
      Tab(4).Control(3)=   "cmdInvCustMsg"
      Tab(4).Control(4)=   "cmdInvQBJob"
      Tab(4).Control(5)=   "cmdInvGlobal"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "1099 Processing"
      TabPicture(5)   =   "frmMainMenu.frx":13B2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdPA_Report"
      Tab(5).Control(1)=   "cmdPA_Import"
      Tab(5).Control(2)=   "cmdPA_Print"
      Tab(5).Control(3)=   "cmdPA_PayerMaint"
      Tab(5).Control(4)=   "cmdPA_Payee"
      Tab(5).ControlCount=   5
      Begin VB.CommandButton cmdBackUp 
         Caption         =   "BACKUP / RESTORE"
         Height          =   615
         Left            =   960
         TabIndex        =   128
         Top             =   6600
         Width           =   2535
      End
      Begin VB.CommandButton cmdJimBo 
         Caption         =   "SQL"
         Height          =   495
         Left            =   11280
         TabIndex        =   127
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewADO 
         Caption         =   "DATABASE CONVERT"
         Height          =   615
         Left            =   8040
         TabIndex        =   126
         Top             =   6840
         Width           =   2535
      End
      Begin VB.CommandButton cmdPA_Report 
         Caption         =   "REPORT"
         Height          =   855
         Left            =   -74280
         TabIndex        =   122
         Top             =   6240
         Width           =   2415
      End
      Begin VB.CommandButton cmdPA_Import 
         Caption         =   "IMPORT FROM SUPERDOS"
         Height          =   855
         Left            =   -74280
         TabIndex        =   121
         Top             =   4920
         Width           =   2415
      End
      Begin VB.CommandButton cmdPA_Print 
         Caption         =   "ENTRY / 1099 PRINT"
         Height          =   855
         Left            =   -74280
         TabIndex        =   120
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CommandButton cmdPA_PayerMaint 
         Caption         =   "PAYER MAINTENANCE/1096"
         Height          =   855
         Left            =   -74280
         TabIndex        =   119
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdPA_Payee 
         Caption         =   "PAYEE MAINTENANCE"
         Height          =   855
         Left            =   -74280
         TabIndex        =   118
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdInvGlobal 
         Caption         =   "GLOBAL QB SETTINGS"
         Height          =   615
         Left            =   -74040
         TabIndex        =   117
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton cmdInvQBJob 
         Caption         =   "QB JOB UPDATE"
         Height          =   615
         Left            =   -74040
         TabIndex        =   116
         Top             =   5520
         Width           =   2415
      End
      Begin VB.CommandButton cmdPR1099 
         Caption         =   "1099 PROCESSING"
         Height          =   615
         Left            =   -69600
         TabIndex        =   115
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton cmdInvCustMsg 
         Caption         =   "CUSTOMER MESSAGES"
         Height          =   615
         Left            =   -74040
         TabIndex        =   112
         Top             =   3840
         Width           =   2415
      End
      Begin VB.CommandButton cmdKPInvGlobalMaint 
         Caption         =   "SYSTEM SETTINGS"
         Height          =   615
         Left            =   -74040
         TabIndex        =   111
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CommandButton cmdKPInvStockMaint 
         Caption         =   "STOCK MAINTENANCE"
         Height          =   615
         Left            =   -74040
         TabIndex        =   110
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton cmdKPInvProcess 
         Caption         =   "INVOICE PROCESSING"
         Height          =   615
         Left            =   -74040
         TabIndex        =   109
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdFreeFormat 
         Caption         =   "FREE FORMAT STATEMENTS"
         Height          =   1215
         Left            =   -74640
         Picture         =   "frmMainMenu.frx":13CE
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmdSDGLFFImport 
         Caption         =   "GL FREE FORMAT IMPORT"
         Height          =   615
         Left            =   8040
         TabIndex        =   102
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CommandButton cmdQBTaxPay 
         Caption         =   "TAX PAYMENT TO QB"
         Height          =   735
         Left            =   -64560
         TabIndex        =   101
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdPWMaint 
         Caption         =   "PREVAILING WAGE MAINTENANCE"
         Height          =   735
         Left            =   -74400
         TabIndex        =   100
         Top             =   4920
         Width           =   2535
      End
      Begin VB.CommandButton cmdSDPRHImport 
         Caption         =   "PR HISTORY IMPORT"
         Height          =   615
         Left            =   8040
         TabIndex        =   97
         Top             =   5880
         Width           =   2535
      End
      Begin VB.CommandButton cmdJCTSReport 
         Caption         =   "TIME SHEET REPORT"
         Height          =   735
         Left            =   -74400
         TabIndex        =   95
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdTimeSheetEntry 
         Caption         =   "TIME SHEET ENTRY"
         Height          =   735
         Left            =   -74400
         TabIndex        =   93
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdJCJobMaint 
         Caption         =   "JOB MAINTENANCE"
         Height          =   735
         Left            =   -74400
         TabIndex        =   92
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox chkHideJC 
         Caption         =   "Hide JC Menu"
         Height          =   255
         Left            =   5280
         TabIndex        =   91
         Top             =   8520
         Width           =   1815
      End
      Begin VB.CommandButton cmdJCWageRpt 
         Caption         =   "WAGE BY JOB REPORT"
         Height          =   735
         Left            =   -74400
         TabIndex        =   90
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CommandButton cmdJCMaint 
         Caption         =   "CUSTOMER / JOB MAINTENANCE"
         Height          =   735
         Left            =   -64560
         TabIndex        =   89
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   615
         Left            =   960
         TabIndex        =   82
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CheckBox chkHidePR 
         Caption         =   "Hide PR Menu"
         Height          =   255
         Left            =   3120
         TabIndex        =   79
         Top             =   8520
         Width           =   1815
      End
      Begin VB.CheckBox chkHideGL 
         Caption         =   "Hide GL Menu"
         Height          =   255
         Left            =   840
         TabIndex        =   78
         Top             =   8520
         Width           =   1815
      End
      Begin VB.CommandButton cmdSDPRImport 
         Caption         =   "PR CLIENT IMPORT"
         Height          =   615
         Left            =   8040
         TabIndex        =   75
         Top             =   5160
         Width           =   2535
      End
      Begin VB.CommandButton cmdSDGLHImport 
         Caption         =   "GL HISTORY IMPORT"
         Height          =   615
         Left            =   8040
         TabIndex        =   74
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton cmdSDGLImport 
         Caption         =   "GL CLIENT IMPORT"
         Height          =   615
         Left            =   8040
         TabIndex        =   73
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdPREntry 
         Caption         =   "D A T A   E N T R Y"
         Height          =   975
         Left            =   -74160
         Picture         =   "frmMainMenu.frx":16D8
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   840
         Width           =   6495
      End
      Begin VB.Frame Frame7 
         Height          =   6495
         Left            =   -64920
         TabIndex        =   54
         Top             =   720
         Width           =   3015
         Begin VB.CommandButton cmdAcctImport 
            Caption         =   "IMPORT CHART OF ACCOUNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   108
            Top             =   5640
            Width           =   2535
         End
         Begin VB.CommandButton cmdAccountChange 
            Caption         =   "ACCOUNT # CHANGE"
            Height          =   615
            Left            =   240
            TabIndex        =   107
            Top             =   4920
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlClrAmts 
            Caption         =   "C&LEAR AMOUNTS AND UPDATE"
            Height          =   615
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlFiscYrClose 
            Caption         =   "&FISCAL YEAR CLOSING"
            Height          =   615
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlDelAccts 
            Caption         =   "DELETE ACCO&UNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   58
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlCopyBRBud 
            Caption         =   "C&OPY BRANCH/BUDGET"
            Height          =   615
            Left            =   240
            TabIndex        =   57
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlMultDivAccts 
            Caption         =   "&MULTIPLY/DIVIDE ACCOUNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   56
            Top             =   2760
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLUtlFileCopy 
            Caption         =   "FILE COP&Y"
            Height          =   615
            Left            =   240
            TabIndex        =   55
            Top             =   4200
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "U  T  I  L  I  T  Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   61
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5055
         Left            =   -72120
         TabIndex        =   46
         Top             =   720
         Width           =   3015
         Begin VB.CommandButton cmdGLRptDEJrnl 
            Caption         =   "DATA &ENTRY JOURNAL"
            Height          =   615
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLRptDetlGL 
            Caption         =   "DETAIL &GENERAL LEDGER"
            Height          =   615
            Left            =   240
            TabIndex        =   51
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLRptChofAccts 
            Caption         =   "&CHART OF ACCOUNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   50
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLRptPRGLAccts 
            Caption         =   "&PRINT GL ACCOUNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   49
            Top             =   2760
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLRptPrDescFile 
            Caption         =   "P&RINT DESCRIPTION FILE"
            Height          =   615
            Left            =   240
            TabIndex        =   48
            Top             =   4200
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLRptTrBal 
            Caption         =   "TRIAL &BALANCE"
            Height          =   615
            Left            =   240
            TabIndex        =   47
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   "R  E  P  O  R  T  S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   53
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6375
         Left            =   -68520
         TabIndex        =   40
         Top             =   720
         Width           =   3015
         Begin VB.CommandButton cmdFFColumn 
            Caption         =   "COLUMN DEFINITIONS"
            Height          =   735
            Left            =   240
            TabIndex        =   105
            Top             =   4200
            Width           =   2535
         End
         Begin VB.CommandButton cmdFFSched 
            Caption         =   "ACCOUNT SCHEDULES"
            Height          =   615
            Left            =   240
            TabIndex        =   104
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLMaintComp 
            Caption         =   "&COMPANY"
            Height          =   615
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLMtAcctsAmts 
            Caption         =   "&ACCOUNTS/AMOUNTS"
            Height          =   615
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   2535
         End
         Begin VB.CommandButton cmdGLMtJrnlSrc 
            Caption         =   "&JOURNAL SOURCE"
            Height          =   615
            Left            =   240
            TabIndex        =   42
            Top             =   2040
            Width           =   2535
         End
         Begin VB.CommandButton cmdMtDesc 
            Caption         =   "&DESCRIPTIONS"
            Height          =   615
            Left            =   240
            TabIndex        =   41
            Top             =   5520
            Width           =   2535
         End
         Begin VB.Label Label19 
            Caption         =   "F R E E   F O R M A T"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   103
            Top             =   3000
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "M A I N T E N A N C E"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdFIOpen 
         Caption         =   "&OPEN"
         DragIcon        =   "frmMainMenu.frx":19E2
         Height          =   615
         Left            =   960
         TabIndex        =   35
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdFiNew 
         Caption         =   "&NEW"
         Height          =   615
         Left            =   960
         TabIndex        =   34
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CommandButton cmdFICopy 
         Caption         =   "&COPY"
         Height          =   615
         Left            =   960
         TabIndex        =   33
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CommandButton cmdFiPSSWD 
         Caption         =   "&PASSWORD"
         Height          =   615
         Left            =   960
         TabIndex        =   32
         Top             =   3360
         Width           =   2535
      End
      Begin VB.CommandButton cmdFIUserMt 
         Caption         =   "&USER MAINTENANCE"
         Height          =   615
         Left            =   960
         TabIndex        =   31
         Top             =   5520
         Width           =   2535
      End
      Begin VB.CommandButton cmdGLDataEntry 
         Caption         =   "&DA&TA ENTRY"
         Height          =   1215
         Left            =   -74640
         Picture         =   "frmMainMenu.frx":1CEC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdGLMtStatements 
         Caption         =   "&STATEMENTS"
         Height          =   1215
         Left            =   -74640
         Picture         =   "frmMainMenu.frx":1FF6
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Height          =   6855
         Left            =   -74760
         TabIndex        =   10
         Top             =   2040
         Width           =   7575
         Begin VB.CommandButton cmdFUTA940 
            Caption         =   "FEDERAL 940"
            Height          =   615
            Left            =   2760
            TabIndex        =   125
            Top             =   6240
            Width           =   2175
         End
         Begin VB.CommandButton cmdDptDist 
            Caption         =   "DEPT DISTRIBUTION"
            Height          =   615
            Left            =   240
            TabIndex        =   114
            Top             =   4800
            Width           =   2175
         End
         Begin VB.CommandButton cmdItemListing 
            Caption         =   "ITEM LISTING"
            Height          =   615
            Left            =   5160
            TabIndex        =   113
            Top             =   6120
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRErnDed 
            Caption         =   "EARNG AND DEDUCTION SUMM"
            Height          =   615
            Left            =   2760
            TabIndex        =   99
            Top             =   5520
            Width           =   2175
         End
         Begin VB.CommandButton cmdW2 
            Caption         =   "W2 PROCESSING"
            Height          =   615
            Left            =   5160
            TabIndex        =   94
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton cdEarnSumm 
            Caption         =   "EARNINGS SUMMARY"
            Height          =   615
            Left            =   2760
            TabIndex        =   84
            Top             =   4800
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRepQPChkRecon 
            Caption         =   "CHECK RECONCI&LIATION"
            Height          =   615
            Left            =   2760
            TabIndex        =   62
            Top             =   4080
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRepPPPItmDtl 
            Caption         =   "ITE&M DETAIL"
            Height          =   615
            Left            =   240
            TabIndex        =   23
            Top             =   4080
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMiscListsLabels 
            Caption         =   "LISTS AND LA&BELS"
            Height          =   615
            Left            =   5160
            TabIndex        =   22
            Top             =   5400
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptQPCityTax 
            Caption         =   "&CITY TAX REPORT"
            Height          =   615
            Left            =   2760
            TabIndex        =   21
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptPPPDirDep 
            Caption         =   "&DI&RECT DEPOSIT REPORT"
            Height          =   615
            Left            =   240
            TabIndex        =   20
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMiscNewHire 
            Caption         =   "&NEW HIRE REPORT"
            Height          =   615
            Left            =   5160
            TabIndex        =   19
            Top             =   4680
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptQPOBUC 
            Caption         =   "&OHIO BUC"
            Height          =   615
            Left            =   2760
            TabIndex        =   18
            Top             =   2640
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptPPPEntryFrm 
            Caption         =   "&ENTRY FORM"
            Height          =   615
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMiscCityRate 
            Caption         =   "CIT&Y RATE LIST"
            Height          =   615
            Left            =   5160
            TabIndex        =   16
            Top             =   3960
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptQPFed941 
            Caption         =   "&FEDERAL 941"
            Height          =   615
            Left            =   2760
            TabIndex        =   15
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptPPPDepList 
            Caption         =   "EMPLOYER &DE&POSIT LISTING"
            Height          =   615
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptQPQtrRpts 
            Caption         =   "&QUARTERLY REPORTS"
            Height          =   615
            Left            =   2760
            TabIndex        =   13
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptPPPChkReg 
            Caption         =   "C&HECK REGISTER"
            Height          =   615
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRRptYECityTax 
            Caption         =   "Y/E CITY &TAX REPORT"
            Height          =   615
            Left            =   5160
            TabIndex        =   11
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label141 
            Caption         =   "Miscellaneous"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   28
            Top             =   3480
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Y/E Processing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   27
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Qtrly Processing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label11 
            Caption         =   "Pay Period Processing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label10 
            Caption         =   "R  E  P  O  R  T  S"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2950
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   -63960
         TabIndex        =   9
         Top             =   720
         Width           =   2655
         Begin VB.CommandButton cmdPRPurge 
            Caption         =   "HISTORY PURGE"
            Height          =   615
            Left            =   240
            TabIndex        =   124
            Top             =   4320
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRQBRegister 
            Caption         =   "QUICK BOOKS REGISTRATION"
            Height          =   615
            Left            =   240
            TabIndex        =   98
            Top             =   3600
            Width           =   2175
         End
         Begin VB.CommandButton cmdWkComp 
            Caption         =   "WRK COMP ASSIGN"
            Height          =   615
            Left            =   240
            TabIndex        =   80
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRUtlTaxWageSweep 
            Caption         =   "TAXABLE &WAGE SWEEP"
            Height          =   615
            Left            =   240
            TabIndex        =   69
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRUtlUpdtGL 
            Caption         =   "&UPDATE TO GENERAL LEDGER"
            Height          =   615
            Left            =   240
            TabIndex        =   68
            Top             =   1485
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRUtlCityAssign 
            Caption         =   "CHANGE CITY/DEPT"
            Height          =   615
            Left            =   240
            TabIndex        =   67
            Top             =   810
            Width           =   2175
         End
         Begin VB.Label Label14 
            Caption         =   "U T I L I T Y"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   70
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6975
         Left            =   -66960
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         Begin VB.CommandButton cmdPRCountyMaint 
            Caption         =   "COUNTY"
            Height          =   615
            Left            =   240
            TabIndex        =   96
            Top             =   5400
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtGloSysSet 
            Caption         =   "&SYSTEM SETTINGS"
            Height          =   615
            Left            =   240
            TabIndex        =   66
            Top             =   3960
            Width           =   2175
         End
         Begin VB.CommandButton cmdPrMtGloCity 
            Caption         =   "C&ITY"
            Height          =   615
            Left            =   240
            TabIndex        =   65
            Top             =   4680
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtGloState 
            Caption         =   "ST&ATE"
            Height          =   615
            Left            =   240
            TabIndex        =   64
            Top             =   6120
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtComp 
            Caption         =   "&COMPANY"
            Height          =   615
            Left            =   240
            TabIndex        =   7
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtGLUpdate 
            Caption         =   "&GENERAL LEDGER UPDATE"
            Height          =   615
            Left            =   240
            TabIndex        =   6
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtDept 
            Caption         =   "&DEPARTMENT"
            Height          =   615
            Left            =   240
            TabIndex        =   5
            Top             =   2160
            Width           =   2175
         End
         Begin VB.CommandButton cmdPRMtEmployee 
            Caption         =   "&EMPLOYEE"
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Global"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   63
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "M A I N T E N A N C E"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Label Label18 
         Caption         =   "DELETE the currently opened client"
         Height          =   495
         Left            =   3840
         TabIndex        =   83
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Payroll Import from SuperDOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8040
         TabIndex        =   81
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "General Ledger Import from SuperDOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         TabIndex        =   72
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Open Existing Client"
         Height          =   375
         Left            =   3840
         TabIndex        =   39
         Top             =   1380
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Create a New BLANK File"
         Height          =   375
         Left            =   3840
         TabIndex        =   38
         Top             =   2100
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Make a Copy of the Current File"
         Height          =   495
         Left            =   3840
         TabIndex        =   37
         Top             =   2700
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Set or Change Database Password of the Current File"
         Height          =   495
         Left            =   3840
         TabIndex        =   36
         Top             =   3400
         Width           =   2895
      End
   End
   Begin VB.Label lblBalintFolder 
      Caption         =   "Balint Folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   7920
      TabIndex        =   123
      Top             =   10440
      Width           =   4335
   End
   Begin VB.Label lblVersion 
      Caption         =   "New ADO 6/11/20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   5640
      TabIndex        =   88
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label lblFileName 
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   5640
      TabIndex        =   87
      Top             =   10080
      Width           =   6375
   End
   Begin VB.Label lblPRCompanyID 
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   480
      TabIndex        =   86
      Top             =   10440
      Width           =   4935
   End
   Begin VB.Label lblGLCompanyID 
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   480
      TabIndex        =   85
      Top             =   10080
      Width           =   4815
   End
   Begin VB.Label Label16 
      Caption         =   "Balint Windows Accounting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   3360
      TabIndex        =   76
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim x As String
Dim dbPassword, DriveLetter As String
Dim TaskID As Long
Dim LoadFlag As Boolean
Public dbName As String




Private Sub cmdJimBo_Click()
    Dim frmj As New frmJimBo
    frmj.lblHdr.Caption = GLCompany.Name
    frmj.Show
End Sub

Private Sub cmdNewADO_Click()
    RunADO_Conversion (BalintFolder)
End Sub

Private Sub Form_Load()
    
    ' *** for testing - dumb ass !!!
    DriveLetter = "C:"

    DriveLetter = Mid(App.Path, 1, 2)

    ' tab hides ? - see if PRGlobal exists
    Me.SSTab1.TabVisible(0) = True      ' File
    Me.SSTab1.TabVisible(1) = True      ' GL
    Me.SSTab1.TabVisible(2) = True      ' Payroll
    Me.SSTab1.TabVisible(3) = True      ' Job Cost
    Me.SSTab1.TabVisible(4) = False     ' KP Invoicing
    Me.SSTab1.TabVisible(5) = False     ' 1099 processing
    
    LoadFlag = True
    
    If TableExists("PRGlobal", cnDes) = True Then

        SQLString = "SELECT * FROM PRGlobal WHERE Description = 'MenuHidden' AND UserID = " & GLUser.ID
        rsInit SQLString, cnDes, rs
        If rs.RecordCount > 0 Then
            
            rs.MoveFirst

            If rs!Var1 = "1" Then      ' GL
                Me.SSTab1.TabVisible(1) = False
                Me.chkHideGL = 1
            End If
            
            If rs!Var2 = "1" Then      ' PR
                Me.SSTab1.TabVisible(2) = False
                Me.chkHidePR = 1
            End If
            
            If rs!Var3 = "1" Then      ' JC
                Me.SSTab1.TabVisible(3) = False
                Me.chkHideJC = 1
            End If
            
            rs.Close
            
        End If
    
        ' KP Invoice
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                    " AND Description = 'KPInvoice' AND Var2 = '" & GLCompany.ID & "'" & _
                    " AND Var1 = 'Yes'"
        If PRGlobal.GetBySQL(SQLString) = True Then
            Me.SSTab1.TabVisible(4) = True
        End If
    
    Else
        
        Me.chkHideGL = 0
        Me.chkHideGL.Enabled = False
        
        Me.chkHidePR = 1
        Me.chkHidePR.Enabled = False
        Me.SSTab1.TabVisible(2) = False
        
        Me.chkHideJC = 1
        Me.chkHideJC.Enabled = False
        Me.SSTab1.TabVisible(3) = False
    
        ' don't show JC if no PR files
    
    End If

    ' enable 1099 processing
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                " AND Description = '1099' " & _
                " AND Var1 = 'Yes'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Me.SSTab1.TabVisible(5) = True
    End If

    ' set focus based on open tab
    Me.Show
    
    ' from fresh login
    If Command() = "" Then
        OpenTab = 1         ' dflt to GL
        If Me.chkHideGL = 0 And Me.chkHidePR = 0 Then
            OpenTab = 0
        ElseIf Me.chkHideGL = 1 Then
            OpenTab = 2 ' dflt to PR
        End If
    End If
        
    If Me.SSTab1.TabVisible(OpenTab) Then
        Me.SSTab1.Tab = OpenTab
    Else
        Me.SSTab1.Tab = 0
    End If
    
    ' set the focus on the first button of the tab
    ' don't want EXIT to be in focus
    Me.cmdFIOpen.SetFocus
    If Me.SSTab1.Tab = 1 Then Me.cmdGLDataEntry.SetFocus
    If Me.SSTab1.Tab = 2 Then Me.cmdPREntry.SetFocus
    If Me.SSTab1.Tab = 3 Then Me.cmdJCMaint.SetFocus
        
    ShowCompanyID

    LoadFlag = False

End Sub
Private Sub chkHideGL_Click()
    
    If Me.chkHideGL = 1 Then
        Me.chkHidePR = 0
        Me.SSTab1.TabVisible(1) = False
        Me.SSTab1.Tab = 2
    Else
        Me.SSTab1.TabVisible(1) = True
    End If
    
    MenuHideUpdate
    
End Sub
Private Sub chkHidePR_Click()
    
    If Me.chkHidePR = 1 Then
        Me.chkHideGL = 0
        Me.SSTab1.TabVisible(2) = False
        Me.SSTab1.Tab = 1
    Else
        Me.SSTab1.TabVisible(2) = True
    End If
    
    MenuHideUpdate
    
End Sub

Private Sub MenuHideUpdate()
    
    If LoadFlag = True Then Exit Sub
    
    ' save the menu hide options
    ' clear it first
    x = "MenuHidden"
    SQLString = "DELETE * FROM PRGlobal WHERE Description = 'MenuHidden' AND UserID = " & GLUser.ID
    cnDes.Execute SQLString
    
    ' save it?
    SQLString = "SELECT * FROM PRGlobal"
    rsInit SQLString, cnDes, rs
    rs.AddNew
    rs!Description = "MenuHidden"
    rs!UserID = GLUser.ID
    If Me.chkHideGL = 1 Then
        rs!Var1 = "1"
    Else
        rs!Var1 = "0"
    End If
    If Me.chkHidePR = 1 Then
        rs!Var2 = "1"
    Else
        rs!Var2 = "0"
    End If
    If Me.chkHideJC = 1 Then
        rs!Var3 = "1"
    Else
        rs!Var3 = "0"
    End If
    rs.Update

End Sub

Private Sub cmdFiPSSWD_Click()
    
    On Error Resume Next
    cn.Close
    On Error GoTo 0
    
    ' RE-open the company database
    Dim gfnm As String
    gfnm = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    
     ' open the company database
    If BalintFolder = "" Then
        gfnm = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    Else
        gfnm = Replace(BalintFolder, "^", " ") & "\Data\" & mdbName(GLCompany.FileName)
    End If
    
    If NewADO Then
        gfnm = Replace(gfnm, ".mdb", ".accdb")
    Else
        gfnm = Replace(gfnm, ".accdb", ".mdb")
    End If
    
    frmSetDBPassword.lblCompanyName = Me.lblCompanyName
    frmSetDBPassword.lblFileName = gfnm
    frmSetDBPassword.Show vbModal
    
    dbName = gfnm
    CNOpen gfnm, frmSetDBPassword.tdbNewPassword
    dbPwd = frmSetDBPassword.tdbNewPassword
    

End Sub


Private Sub cmdFIUserMt_Click()
    NewCall "GLMaint", "User"
End Sub

Private Sub cmdFICopy_Click()
'    If BalintFolder <> "" Then
'        MsgBox "Run this from a standard network location!", vbExclamation
'        Exit Sub
'    End If
    NewCall "GLUtil", "GLFileCopy"
End Sub
Private Sub cmdGLUtlFileCopy_Click()
    NewCall "GLUtil", "GLFileCopy"
End Sub

Private Sub cmdSDGLHistImport_Click()
    NewCall "GLUtil", "HistImport"
End Sub

Private Sub cmdSDGLImport_Click()
    If BalintFolder <> "" Then
        MsgBox "Run this from a standard network location!", vbExclamation
        Exit Sub
    End If
    NewCall "GLUtil", "Import"
End Sub

Private Sub cmdSDPRImport_Click()
    If BalintFolder <> "" Then
        MsgBox "Run this from a standard network location!", vbExclamation
        Exit Sub
    End If
    NewCall "GLUtil", "PRImport"
End Sub

Private Sub cmdFiNew_Click()
'    If BalintFolder <> "" Then
'        MsgBox "Run this from a standard network location!", vbExclamation
'        Exit Sub
'    End If
    NewCall "GLUtil", "NewFile"
End Sub
Private Sub cmdPREntry_Click()
    NewCall "PREntry", ""
End Sub

Private Sub cmdFIOpen_Click()
    frmCompanyList.Show vbModal
    ShowCompanyID
    ' PR conversion fix
    On Error Resume Next
    cn.Execute "update PREmployee set CheckComment = """" where CheckComment = ""0"""
    On Error GoTo 0
End Sub

Private Sub cmdGLDataEntry_Click()
    NewCall "GLEntryADO", "Entry"
End Sub

Private Sub cmdGLMaintComp_Click()
    NewCall "GLMaint", "Company"
End Sub

Private Sub cmdGLMtAcctsAmts_Click()
    NewCall "GLMaint", "Account"
End Sub

Private Sub cmdGLMtJrnlSrc_Click()
    NewCall "GLMaint", "Journal"
End Sub
Private Sub cmdFFSched_Click()
    NewCall "GLMaint", "FFSchedule"
End Sub
Private Sub cmdFFColumn_Click()
    NewCall "GLMaint", "FFColumn"
End Sub
Private Sub cmdGLMtStatements_Click()
    NewCall "GLPrint", "Statement"
End Sub
Private Sub cmdFreeFormat_Click()
    NewCall "GLPrint", "FreeFormat"
End Sub
Private Sub cmdGLRptChofAccts_Click()
    NewCall "GLPrint", "ChartOfAccounts"
End Sub

Private Sub cmdGLRptDEJrnl_Click()
    NewCall "GLPrint", "GLHistJnl"
End Sub

Private Sub cmdGLRptDetlGL_Click()
    NewCall "GLPrint", "DetailGL"
End Sub

Private Sub cmdGLRptPrDescFile_Click()
    NewCall "GLPrint", "PrintDesc"
End Sub

Private Sub cmdGLRptPRGLAccts_Click()
    NewCall "GLPrint", "PrintGLAccount"
End Sub

Private Sub cmdGLRptTrBal_Click()
    NewCall "GLPrint", "TrialBal"
End Sub

Private Sub cmdGLUtlClrAmts_Click()
    NewCall "GLUtil", "ClearGLAmount"
End Sub

Private Sub cmdGLUtlCopyBRBud_Click()
    NewCall "GLUtil", "CopyBB"
End Sub

Private Sub cmdGLUtlDelAccts_Click()
    NewCall "GLUtil", "DeleteAccts"
End Sub

Private Sub cmdGLUtlFiscYrClose_Click()
    NewCall "GLUtil", "YearEnd"
End Sub

Private Sub cmdGLUtlMultDivAccts_Click()
    NewCall "GLUtil", "GLMultDiv"
End Sub

Private Sub cmdMtDesc_Click()
    NewCall "GLMAINT", "Descriptions"
End Sub

Private Sub cmdPRMiscCityRate_Click()
    NewCall "PRReport", "CityList"
End Sub

Private Sub cmdPRMiscListsLabels_Click()
    NewCall "PRReport", "EEList"
End Sub
Private Sub cmdItemListing_Click()
    NewCall "PRReport", "ItemListing"
End Sub

Private Sub cmdPRMiscNewHire_Click()
    NewCall "PRReport", "NewHire"
End Sub

Private Sub cmdPRMtComp_Click()
    NewCall "PRMaint", "Employer"
End Sub

Private Sub cmdPRMtDept_Click()
    NewCall "PRMaint", "Department"
End Sub

Private Sub cmdPRMtEmployee_Click()
    NewCall "PRMaint", "Employee"
End Sub

Private Sub cmdPrMtGloCity_Click()
    NewCall "PRMaint", "City"
End Sub

Private Sub cmdPRMtGloState_Click()
    NewCall "PrMaint", "State"
End Sub

Private Sub cmdPRMtGloSysSet_Click()
    NewCall "PRGlobMaint", ""
End Sub

Private Sub cmdPRMtGLUpdate_Click()
    NewCall "PRMaint", "GLUpd"
End Sub

Private Sub cmdPRRepPPPItmDtl_Click()
    NewCall "PRGReps", "ItemDetail"
End Sub
Private Sub cmdDptDist_Click()
    NewCall "PRReport", "DptDist"
End Sub

Private Sub cmdPRRepQPChkRecon_Click()
    NewCall "PRReport", "CheckRecon"
End Sub

Private Sub cmdPRRptPPPChkReg_Click()
    NewCall "PRReport", "CheckReg"
End Sub

Private Sub cmdPRRptPPPDepList_Click()
    NewCall "PRReport", "Deposit"
End Sub

Private Sub cmdPRRptPPPDirDep_Click()
    NewCall "PRReport", "DirDep"
End Sub

Private Sub cmdPRRptPPPEntryFrm_Click()
    NewCall "PRReport", "EntryForm"
End Sub

Private Sub cmdPRRptQPCityTax_Click()
    NewCall "PRReport", "CityTax"
End Sub

Private Sub cmdPRRptQPFed941_Click()
    NewCall "PRGReps", "Form941"
End Sub

Private Sub cmdPRRptQPOBUC_Click()
    NewCall "PRReport", "OHBUC"
End Sub

Private Sub cmdPRRptQPQtrRpts_Click()
    NewCall "PRReport", "QtrRpts"
End Sub

Private Sub cmdPRRptYECityTax_Click()
    NewCall "PRReport", "YECityTax"
End Sub

Private Sub cmdPRUtlCityAssign_Click()
    NewCall "PRMaint", "AssignCity"
End Sub

Private Sub cmdPRUtlTaxWageSweep_Click()
    NewCall "PRMaint", "TaxSweep"
End Sub

Private Sub cmdPRUtlUpdtGL_Click()
    NewCall "PRReport", "GLUpdate"
End Sub
Private Sub cmdPA_Payee_Click()
    NewCall "Win99", "Payee"
End Sub
Private Sub cmdPA_PayerMaint_Click()
    NewCall "Win99", "Payer"
End Sub

Private Sub cmdPA_Print_Click()
    NewCall "Win99", "Print"
End Sub
Private Sub cmdPA_Import_Click()
    NewCall "Win99", "SDImport"
End Sub
Private Sub cmdPA_Report_Click()
    NewCall "Win99", "Report"
End Sub


Private Sub CmdExit_Click()
    End
End Sub
Private Sub cmdSDGLHImport_Click()
    If BalintFolder <> "" Then
        MsgBox "Run this from a standard network location!", vbExclamation
        Exit Sub
    End If
    NewCall "GLUtil", "HistImport"
End Sub
Private Sub cmdSDGLFFImport_Click()
    If BalintFolder <> "" Then
        MsgBox "Run this from a standard network location!", vbExclamation
        Exit Sub
    End If
    NewCall "GLUtil", "FFImport"
End Sub
Private Sub cdEarnSumm_Click()
    NewCall "PRGReps", "EarnSummary"
End Sub
Private Sub cmdJCMaint_Click()
    NewCall "PRMaint", "JCList"
End Sub
Private Sub cmdJCJobMaint_Click()
    NewCall "PRMaint", "JCJobMaint"
End Sub
Private Sub cmdTimeSheetEntry_Click()
    NewCall "PRMaint", "TimeSheet"
End Sub
Private Sub cmdW2_Click()
    NewCall "PRGReps", "W2"
End Sub
Private Sub cmdPR1099_Click()
    NewCall "PRReport", "1099"
End Sub

Private Sub cmdJCWageRpt_Click()
    NewCall "PRReport", "WAGEBYJOB"
End Sub
Private Sub cmdJCTSReport_Click()
    NewCall "PRMaint", "TSPRINT"
End Sub
Private Sub cmdPRCountyMaint_Click()
    NewCall "PRMaint", "COUNTY"
End Sub
Private Sub cmdSDPRHImport_Click()
    If BalintFolder <> "" Then
        MsgBox "Run this from a standard network location!", vbExclamation
        Exit Sub
    End If
    NewCall "PRMaint", "HISTIMPORT"
End Sub
Private Sub cmdPRQBRegister_Click()
    NewCall "PRMaint", "QBRegister"
End Sub
Private Sub cmdPRErnDed_Click()
    NewCall "PRReport", "ErnDed"
End Sub
Private Sub cmdPWMaint_Click()
    NewCall "PRMaint", "PWMaint"
End Sub
Private Sub cmdQBTaxPay_Click()
    NewCall "PRQBFunc", "TaxPay"
End Sub
Private Sub cmdAccountChange_Click()
    NewCall "GLMaint", "AccountChange"
End Sub
Private Sub cmdAcctImport_Click()
    NewCall "GLUtil", "AcctImport"
End Sub
Private Sub cmdKPInvProcess_Click()
    NewCall "KPInvoice", "Process"
End Sub
Private Sub cmdKPInvGlobalMaint_Click()
    NewCall "KPInvoice", "Global"
End Sub
Private Sub cmdInvCustMsg_Click()
    NewCall "KPInvoice", "CustMsg"
End Sub

Private Sub cmdKPInvStockMaint_Click()
    NewCall "KPInvoice", "StockMaint"
End Sub
Private Sub cmdInvQBJob_Click()
    NewCall "KPInvoice", "QBJob"
End Sub
Private Sub cmdInvGlobal_Click()
    NewCall "KPInvoice", "GlobalQB"
End Sub
Private Sub cmdPRPurge_Click()
    NewCall "PRMaint", "PURGE"
End Sub

Private Sub cmdFUTA940_Click()
    NewCall "PRReport", "FUTA940"
End Sub

Private Sub chkHideJC_Click()
    If Me.chkHideJC = 0 Then
        Me.SSTab1.TabVisible(3) = True
    Else
        Me.SSTab1.TabVisible(3) = False
    End If
    MenuHideUpdate
End Sub


Private Sub NewCall(ByVal ModuleName As String, ByVal ProgName As String)

    If BalintFolder = "" Then
        x = DriveLetter & "\Balint\" & ModuleName & ".exe" & _
            " ProgName=" & ProgName & _
            " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
            " UserID=" & UserID & _
            " BackName=" & DriveLetter & "\Balint\GLMenu.exe" & _
            " MenuName=GLMenu.exe"
    Else
        x = BalintFolder & "\" & ModuleName & ".exe" & _
            " ProgName=" & ProgName & _
            " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
            " UserID=" & UserID & _
            " BackName=" & BalintFolder & "\GLMenu.exe" & _
            " BalintFolder=" & BalintFolder & _
            " MenuName=GLMenu.exe"
        
        ' folder redirect - EXE's reside on local drive
        x = "C:\Balint\" & ModuleName & ".exe" & _
            " ProgName=" & ProgName & _
            " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
            " UserID=" & UserID & _
            " BackName=" & "c:\Balint\GLMenu.exe" & _
            " BalintFolder=" & BalintFolder & _
            " MenuName=GLMenu.exe"
    End If
    
    ' database password if required
    If dbPwd <> "" Then
       x = x & " dbPWD=" & dbPwd
    End If
        
    ' TaskID = Shell(x, vbMaximizedFocus)
    
    cnDes.Close
    Set cnDes = Nothing
    TaskID = Shell(x, vbNormalFocus)
'    AppActivate TaskID

    Unload Me
    End

End Sub


Private Sub cmdPRMaintComp_Click()
    NewCall "PRMaint", "Employer"
End Sub

Public Function TableExists(ByVal TableName As String, _
                            ByRef adoConn As ADODB.Connection) _
                            As Boolean

Dim cm As ADODB.Command
Dim frs As ADODB.Recordset
Dim FldFlag As Boolean
Dim FString As String
                         
    ' see if the field is already in the Table
    Set frs = New ADODB.Recordset
    frs.CursorLocation = adUseClient
    frs.CursorType = adOpenStatic
    frs.LockType = adLockBatchOptimistic
    Set frs = adoConn.OpenSchema(adSchemaColumns)
           
    TableExists = False
           
    Do Until frs.EOF = True
                  
        If frs!Table_Name = TableName Then
            TableExists = True
            Exit Do
        End If
        
       frs.MoveNext
   
   Loop

End Function

Private Sub cmdWkComp_Click()
    If MsgBox("OK to set ALL employees to Dept Worker's Comp Code?", vbQuestion + vbYesNo, "Windows PR") = vbNo Then
        Exit Sub
    End If
    SQLString = "SELECT * FROM PREmployee"
    If Not PREmployee.GetBySQL(SQLString) Then Exit Sub
    Do
        PREmployee.WkcUseDept = 1
        PREmployee.Save (Equate.RecPut)
        If Not PREmployee.GetNext Then Exit Do
    Loop
    MsgBox "All employees are set to Dept Worker's Comp Category", vbInformation, "Windows PR"
End Sub

Private Sub cmdDelete_Click()
    
Dim FName As String
Dim cID As Long
    
'    If BalintFolder <> "" Then
'        MsgBox "Run this from a standard network location!", vbExclamation
'        Exit Sub
'    End If
    
    If GLUser.LastCompany = 0 Then
        MsgBox "No client opened!", vbInformation
        Exit Sub
    End If
    
    If MsgBox("OK to DELETE " & GLCompany.Name, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' 2015-10-31 - work for balint folder also
    If BalintFolder = "" Then
        FName = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    Else
        
        ' 2015-11-21 0 use the UNC
        Dim LastBS As Integer
        LastBS = InStrRev(GLCompany.FileName, "\")
        FName = BalintFolder & "\Data\" & Mid(GLCompany.FileName, LastBS + 1)
    
    End If
    
    If MsgBox(String(40, "X") & vbCr & "ALL INFORMATION FOR " & GLCompany.Name & vbCr & _
              FName & vbCr & _
              "WILL BE DELETED!!!" & vbCr & _
              "OK TO CONTINUE???" & vbCr & String(40, "X"), vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    cn.Close
    
    ' delete the file
    Kill FName
    
    cID = GLUser.LastCompany
    
    ' delete the company records
    SQLString = "DELETE * FROM GLCompany WHERE ID = " & cID
    cnDes.Execute SQLString
    
    SQLString = "DELETE * FROM PRCompany WHERE GLCompanyID = " & cID
    cnDes.Execute SQLString
 
    GLUser.LastCompany = 0
    GLUser.LastPRCompany = 0
    GLUser.Save (Equate.RecPut)
 
    MsgBox "ALL information for " & GLCompany.Name & " has been deleted" & vbCr & _
           "Open and existing client to continue", vbInformation
    
    GLCompany.Clear
    Me.lblCompanyName = "No Company Loaded"
 
End Sub

Private Sub ShowCompanyID()
    Me.lblGLCompanyID = "GL CompanyID: " & GLCompany.ID
    Me.lblFileName = GLCompany.FileName
    ' Me.lblFileName = dbName
    Me.lblCompanyName = GLCompany.Name
    
    Me.lblPRCompanyID = ""
    If TableExists("PRCompany", cnDes) = True Then
        On Error Resume Next
        SQLString = "SELECT * FROM PRCompany WHERE GLCompanyID = " & GLCompany.ID
        If PRCompany.GetBySQL(SQLString) Then
            Me.lblPRCompanyID = "PR CompanyID: " & PRCompany.CompanyID
        Else
            Me.lblPRCompanyID = ""
        End If
        On Error GoTo 0
    End If

    If BalintFolder = "" Then
        Me.lblBalintFolder = "BalintFolder=Dflt"
    Else
        Me.lblBalintFolder = "BalintFolder=" & BalintFolder
    End If

End Sub


