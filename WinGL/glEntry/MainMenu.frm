VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form MainMenu 
   Caption         =   " GENERAL LEDGER DATA ENTRY BATCH LIST"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   Icon            =   "MainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&COPY"
      Height          =   495
      Left            =   8640
      TabIndex        =   26
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CheckBox chkAcctDesc 
      Caption         =   "&Include Acct Desc"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin TrueDBGrid80.TDBGrid BatchList 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "BATCH"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "UPDATED"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "FISCAL PD"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DEBITS"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "CREDITS"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Currency"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "RECORDS"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "JNL SRC"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "USER"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2752"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2752"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2752"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2752"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2752"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2752"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2752"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   3
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   8421504
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.namedParent=40"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HC0C0C0&"
      _StyleDefs(21)  =   ":id=9,.fgcolor=&H80000008&,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(22)  =   ":id=9,.strikethrough=0,.charset=0"
      _StyleDefs(23)  =   ":id=9,.fontname=MS Sans Serif"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&H80000005&"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(71)  =   "Named:id=33:Normal"
      _StyleDefs(72)  =   ":id=33,.parent=0"
      _StyleDefs(73)  =   "Named:id=34:Heading"
      _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   ":id=34,.wraptext=-1"
      _StyleDefs(76)  =   "Named:id=35:Footing"
      _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=36:Selected"
      _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=37:Caption"
      _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(82)  =   "Named:id=38:HighlightRow"
      _StyleDefs(83)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(84)  =   "Named:id=39:EvenRow"
      _StyleDefs(85)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(86)  =   "Named:id=40:OddRow"
      _StyleDefs(87)  =   ":id=40,.parent=33"
      _StyleDefs(88)  =   "Named:id=41:RecordSelector"
      _StyleDefs(89)  =   ":id=41,.parent=34"
      _StyleDefs(90)  =   "Named:id=42:FilterBar"
      _StyleDefs(91)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DE&LETE"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "ENTRY &JOURNAL"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdDataEntry 
      Caption         =   "&DATA ENTRY"
      Height          =   495
      Left            =   3528
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdEditBatch 
      Caption         =   "&EDIT BATCH"
      Height          =   495
      Left            =   1824
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddBatch 
      Caption         =   "&ADD NEW BATCH"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label15c 
      Caption         =   "F8"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9120
      TabIndex        =   28
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "COPY the selected batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   8640
      TabIndex        =   27
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Click on column header to change sort"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label12 
      Caption         =   "F7"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Delete ALL entries for the selected batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   6960
      TabIndex        =   23
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Print the journal for the selected batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   5280
      TabIndex        =   22
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Add entries to the selected batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   3480
      TabIndex        =   21
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Edit batch info if no records exist"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Start a new Data Entry Batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "F6"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "F4"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "F3"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "F2"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label lblUser 
      Caption         =   "User Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "USER:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "FILE:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblCityStateZip 
      Caption         =   "City/State/Zip"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TaskID As Double
Dim CompanyID As Long
Dim xdb As New XArrayDB
Dim SortBy(7) As Byte
Dim i, j, k As Long
Dim HeadString(7) As String
Dim SortType(7) As Variant
Dim x As String
Dim rsADO As New ADODB.Recordset

Private Sub Form_Load()
    
Dim i, j As Long
Dim x As String
    
    On Error GoTo glErr
    lblUser = use.Name
    
    If BalintFolder = "" Then
        FileName = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
        lblFileName = Mid(FileName, 3, Len(FileName) - 2)
    Else
        FileName = BalintFolder & "\Data\" & mdbName(com.FileName)
        lblFileName = FileName
    End If
    
    lblCompanyName = com.Name
    lblAddress = com.address1
    lblCityStateZip = com.city
    CompanyID = com.ID
    If Not com.city = "" Then lblCityStateZip = lblCityStateZip & " " & com.state
    If com.zipcode > 0 Then lblCityStateZip = lblCityStateZip & " " & CStr(com.zipcode)
    
'    Set xdb = xFactory.GetBatch(FileName, "SELECT * FROM glBatch ORDER BY -FiscalYear, -Period, JournalSource")
    
    Set xdb = xFactory.GetBatch(FileName, "SELECT * FROM glBatch ORDER BY -BatchNumber")

'    Set xdb = xFactory.GetBatch(FileName, "SELECT * FROM glBatch")
    
'    Set xdb = xFactory.GetBatch(FileName, "SELECT * FROM glBatch")
    
    ' column headers / types
    HeadString(0) = "Batch#"
    SortType(0) = XTYPE_LONG
    
    HeadString(1) = "Updated"
    SortType(1) = XTYPE_DATE
    
    HeadString(2) = "Fiscal Pd"
    SortType(2) = XTYPE_STRING
    
    HeadString(3) = "Debits"
    SortType(3) = XTYPE_CURRENCY
    
    HeadString(4) = "Credits"
    SortType(4) = XTYPE_CURRENCY
    
    HeadString(5) = "Records"
    SortType(5) = XTYPE_LONG
    
    HeadString(6) = "Journal Source"
    SortType(6) = XTYPE_STRING
    
    HeadString(7) = "User"
    SortType(7) = XTYPE_STRING

    ' set sort by parameters - start with batch number descending
    SortBy(0) = 1
    BatchList.Columns(0).Caption = UCase(HeadString(0) & "-")
    BatchList.Columns(0).Font.Bold = True

    ' initiate other columns
    For i = 1 To 7
        SortBy(i) = 0
        BatchList.Columns(i).Caption = HeadString(i)
    Next i

    BatchList.Columns(0).Width = 800
    BatchList.Columns(1).Width = 1200
    BatchList.Columns(2).Width = 1000
    BatchList.Columns(3).Width = 1600
    BatchList.Columns(4).Width = 1600
    BatchList.Columns(5).Width = 1000
    BatchList.Columns(6).Width = 1000
    BatchList.Columns(7).Width = 1300
    
    BatchList.Columns(3).Alignment = dbgRight
    BatchList.Columns(4).Alignment = dbgRight
    BatchList.Columns(3).NumberFormat = "###,###,##0.00"
    BatchList.Columns(4).NumberFormat = "###,###,##0.00"
    
    BatchList.Columns(5).Alignment = dbgRight
    BatchList.Columns(6).Alignment = dbgRight
    
    SetUserNames
    
    BatchList.Font.Size = 10
    BatchList.OddRowStyle.Font.Size = 10
    BatchList.EvenRowStyle.Font.Size = 10
    BatchList.AlternatingRowStyle = True
    
    ' BatchList.AlternatingRowStyle = False
    
    BatchList.AllowColSelect = False
    
    ' JS column adjust - show if budget - js > 100
    j = xdb.UpperBound(1)
    For i = 1 To j
        If CLng(xdb(i, 6)) > 100 Then
           xdb(i, 6) = "BUDG " & CLng(xdb(i, 6)) - 100
        End If
    Next i
    
    Set BatchList.Array = xdb
    
    frmProgress.Hide
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub BatchList_HeadClick(ByVal ColIndex As Integer)

Dim SortOrder As Byte

    j = xdb.UpperBound(1)

    For i = 0 To 7
        
        If i = ColIndex Then
           
           If SortBy(i) = 0 Or SortBy(i) = 2 Then
              ' descending
              SortBy(i) = 1
              x = "-"
              SortOrder = XORDER_DESCEND
           ElseIf SortBy(i) = 1 Then
              ' ascending
              SortBy(i) = 2
              x = "+"
              SortOrder = XORDER_ASCEND
           End If
        
           BatchList.Columns(i).Font.Bold = True
           BatchList.Columns(i).Caption = UCase(HeadString(i) & x)
        
        Else          ' not sorted by
        
           BatchList.Columns(i).Caption = HeadString(i)
           BatchList.Columns(i).Font.Bold = False
        
        End If
    
    Next i
        
    ' sort it
    xdb.QuickSort 1, j, ColIndex, SortOrder, SortType(ColIndex)

    Set BatchList.Array = xdb
    BatchList.ReBind
    BatchList.Refresh
    BatchList.col = 0
    BatchList.Row = 0
    BatchList.SetFocus

End Sub

Private Sub cmdAddBatch_Click()
    OnAdd (False)
    If Response = False Then Exit Sub
'    cmdDataEntry_Click
End Sub

Private Sub OnAdd(ByVal BatchCpy As Boolean)

Dim BatchNum As Long
Dim BatchFrom As Long

    ' store the batch copying from if necessary
    If BatchCpy Then
        BatchFrom = xdb.Value(BatchList.Bookmark, 0)
    End If
    
'    On Error GoTo glErr
'    com.lastbatch = com.lastbatch + 1
'    BatchNum = com.lastbatch
'    com.PutRecord curCompany
    
    bat.AddBatch BatchNum, com.curFiscalYear, com.curPeriod, FileName
    
    xdb.InsertRows 1
    BatchList.Bookmark = 1
    xdb.Value(BatchList.Bookmark, 0) = CStr(BatchNum)
    xdb.Value(BatchList.Bookmark, 1) = ShowDate(bat.Updated)
    xdb.Value(BatchList.Bookmark, 2) = CStr(bat.fiscalYear) & "-" & CStr(bat.period)
    xdb.Value(BatchList.Bookmark, 3) = ShowValue(bat.debits)
    xdb.Value(BatchList.Bookmark, 4) = ShowValue(bat.credits)
    xdb.Value(BatchList.Bookmark, 5) = CStr(bat.nRecords)
    xdb.Value(BatchList.Bookmark, 6) = CStr(bat.JournalSource)
    
    BatchList.ReBind
    
    If BatchCpy = False Then
        BatchForm.BatchNumber = BatchNum
        BatchForm.Init
        BatchForm.Show vbModal
        Unload BatchForm
    Else
        BatchCopy.BatchNumberC = BatchNum
        BatchCopy.BatchFrom = BatchFrom
        BatchCopy.Init
        BatchCopy.Show vbModal
        Unload BatchCopy
    End If
    
'    If BatchForm.userOK = True Then
    If Response Then
        
        If BatchCpy = False Then
            bat.GetBatch BatchForm.BatchNumber, FileName
        Else
            bat.GetBatch BatchCopy.BatchNumberC, FileName
        End If
        
        xdb.Value(BatchList.Bookmark, 1) = ShowDate(bat.Updated)
        xdb.Value(BatchList.Bookmark, 2) = CStr(bat.fiscalYear) & "-" & CStr(bat.period)
        xdb.Value(BatchList.Bookmark, 3) = ShowValue(bat.debits)
        xdb.Value(BatchList.Bookmark, 4) = ShowValue(bat.credits)
        xdb.Value(BatchList.Bookmark, 5) = CStr(bat.nRecords)
        
        If bat.JournalSource > 100 Then
           xdb(BatchList.Bookmark, 6) = "BUDG " & CStr(bat.JournalSource - 100)
        Else
           xdb.Value(BatchList.Bookmark, 6) = CStr(bat.JournalSource)
        End If
        
        i = bat.updateUser
        use.GetRecord (i)
        xdb.Value(BatchList.Bookmark, 7) = use.Name
        
        BatchList.RefetchRow
        OnDataEntry
    
    Else
        xdb.DeleteRows 1
        bat.DeleteRecord BatchNum, FileName
        BatchList.ReBind
    End If
        
    BatchList.SetFocus
    Exit Sub
glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdCopy_Click()
Dim BMark
Dim BatchNum As Long
Dim i As Long
Dim db As DAO.Database
Dim Fnm As String
Dim x As String

'    On Error GoTo glErr
    Dim bat As New rBatch
    
    BMark = BatchList.Bookmark
    
    If IsNull(BMark) Then Exit Sub
    
    BatchNum = CLng(xdb.Value(BMark, 0))
    
    i = MsgBox("Are you SURE you want to copy this Batch # " & BatchNum, _
        vbQuestion + vbYesNo + vbDefaultButton2, "Windows GL Entry")
    
    If i = vbNo Then Exit Sub
    
    OnAdd (True)
    
End Sub

Private Sub cmdDataEntry_Click()
    OnDataEntry
End Sub

Private Sub OnDataEntry()
    
Dim BMark
    
    If IsNull(BatchList.Bookmark) Then Exit Sub
    
    frmProgress.Show
    
'    On Error GoTo glErr
    Dim bat As New rBatch
    
    BMark = BatchList.Bookmark
    
    DataEntry.ID = xdb.Value(BatchList.Bookmark, 0)
    DataEntry.Show vbModal
    Unload DataEntry


'    bat.GetSQL "select * from glBatch where BatchNumber=" & BatchList.SelectedItem.Text, FileName
'    BatchList.SelectedItem.SubItems(1) = Format(bat(1).Updated, "mm/dd/yy")
'    BatchList.SelectedItem.SubItems(2) = CStr(bat(1).FiscalYear) & "-" & CStr(bat(1).Period)
'    BatchList.SelectedItem.SubItems(3) = gl.ShowValue(bat(1).Debits)
'    BatchList.SelectedItem.SubItems(4) = gl.ShowValue(bat(1).Credits)
'    BatchList.SelectedItem.SubItems(5) = CStr(bat(1).nRecords)
    
    bat.GetBatch xdb.Value(BatchList.Bookmark, 0), FileName
    
    xdb.Value(BatchList.Bookmark, 1) = bat.Updated
    xdb.Value(BatchList.Bookmark, 2) = bat.fiscalYear & "-" & bat.period
    xdb.Value(BatchList.Bookmark, 3) = bat.debits
    xdb.Value(BatchList.Bookmark, 4) = bat.credits
    xdb.Value(BatchList.Bookmark, 5) = bat.nRecords
    
    BatchList.ReBind
    BatchList.SetFocus
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdDelete_Click()

Dim BMark
Dim BatchNum As Long
Dim i As Long
Dim db As DAO.Database
Dim Fnm As String
Dim x As String

'    On Error GoTo glErr
    Dim bat As New rBatch
    
    BMark = BatchList.Bookmark
    
    If IsNull(BMark) Then Exit Sub
    
    BatchNum = CLng(xdb.Value(BMark, 0))
    
    i = MsgBox("Are you SURE you want to delete this Batch # " & BatchNum, _
        vbCritical + vbYesNo + vbDefaultButton2, "Windows GL Entry")
    
    If i = vbNo Then Exit Sub
    
    If BalintFolder = "" Then
        Fnm = Mid(App.Path, 1, 2) & Mid(FileName, 3, Len(FileName) - 2)
    Else
        Fnm = BalintFolder & "\Data\" & mdbName(FileName)
    End If
    
    ' store the period yyyymm from the batch record
    ' handle bogus batches
    If Right(xdb(BMark, 2), 2) = "00" Then
        i = 0
    Else
        i = Mid(xdb(BMark, 2), 1, 4) * 100 + Mid(xdb(BMark, 2), 6, 2)
    End If
 
    ' delete from GLHistory
    x = "DELETE * FROM GLHistory WHERE BatchNumber = " & BatchNum

    Set db = OpenDatabase(Name:=Fnm, _
                          Options:=False, _
                          ReadOnly:=False, _
                          Connect:=";pwd=" & Password)

    db.Execute (x)

    db.Close
    Set db = Nothing

    bat.DeleteRecord BatchNum, Fnm
    
    If i <> 0 Then
        If BalintFolder = "" Then
            x = "\Balint\GLUtil.exe" & _
                " ProgName=ClearGLAmount " & _
                " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
                " UserID=" & curUser & _
                " BackName=" & DriveLetter & "\Balint\GLEntry.exe" & _
                " Period=" & i
        Else
            x = "c:\Balint\GLUtil.exe" & _
                " ProgName=ClearGLAmount " & _
                " SysFile=" & BalintFolder & "\Data\GLSystem.mdb" & _
                " UserID=" & curUser & _
                " BackName=" & "c:\Balint\GLEntry.exe" & _
                " Period=" & i & _
                " BalintFolder=" & BalintFolder
        End If
    Else
        If BalintFolder = "" Then
            x = "\Balint\GLEntry.exe" & _
                " ProgName=GLEntry " & _
                " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
                " UserID=" & curUser & _
                " BackName=" & DriveLetter & "\Balint\GLEntry.exe"
        Else
            x = "c:\Balint\GLEntry.exe" & _
                " ProgName=GLEntry " & _
                " SysFile=" & BalintFolder & "\Data\GLSystem.mdb" & _
                " UserID=" & curUser & _
                " BackName=" & "c:\Balint\GLEntry.exe" & _
                " BalintFolder=" & BalintFolder
        End If
    End If

    ' database password if required
    If Password <> "" Then
       x = x & " dbPWd=" & Password
    End If
        
    If Not TestMode Then TaskID = Shell(x, vbMaximizedFocus)
    
    Unload Me
    End
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)

End Sub

Private Sub cmdEditBatch_Click()
    OnEdit
End Sub

Private Sub OnEdit()
    
    If IsNull(BatchList.Bookmark) Then Exit Sub
    
    If xdb.Value(BatchList.Bookmark, 5) <> 0 Then
       MsgBox "Batch edit not allowed if history records exist!", vbExclamation + vbOKOnly, "GL Data Entry"
       Exit Sub
    End If
    
    BatchForm.BatchNumber = xdb.Value(BatchList.Bookmark, 0)
    BatchForm.Init
    BatchForm.Show vbModal
    If BatchForm.userOK = True Then
        bat.GetBatch BatchForm.BatchNumber, FileName
        xdb.Value(BatchList.Bookmark, 1) = ShowDate(bat.Updated)
        xdb.Value(BatchList.Bookmark, 2) = CStr(bat.fiscalYear) & "-" & CStr(bat.period)
        xdb.Value(BatchList.Bookmark, 3) = ShowValue(bat.debits)
        xdb.Value(BatchList.Bookmark, 4) = ShowValue(bat.credits)
        xdb.Value(BatchList.Bookmark, 5) = CStr(bat.nRecords)
        xdb.Value(BatchList.Bookmark, 6) = CStr(bat.JournalSource)
        BatchList.ReBind
        BatchList.RefetchRow
    End If
    Unload BatchForm
    BatchList.SetFocus
End Sub

Private Sub CmdExit_Click()
    
Dim x As String
    
    If Not TestMode Then
        If BalintFolder = "" Then
            BackName = "\Balint\GLMenu.exe"
        Else
            BackName = "c:\Balint\GLMenu.exe"
        End If
    End If
    
    If BackName <> "" Then
       
       x = BackName & " UserID=" & curUser & " OpenTab=1 "
       If BalintFolder <> "" Then
           x = x & "BalintFolder=" & Replace(BalintFolder, "^", " ")
       End If
       If Password <> "" Then
          x = x & " dbPwd=" & Password
       End If
       TaskID = Shell(x, vbMaximizedFocus)
    
    End If
    
    Unload Me
    End
    
End Sub


Private Sub Sort(ByRef cc As Collection)

    Dim i, j, N, Temp, x() As Integer
    N = cc.Count
    ReDim x(1 To N)
    For i = N To 1 Step -1
        x(i) = CInt(cc(i))
        cc.Remove i
    Next i
    For i = 1 To N - 1
        For j = i + 1 To N
            If x(i) > x(j) Then
                Temp = x(i)
                x(i) = x(j)
                x(j) = Temp
            End If
        Next j
    Next i
    For i = 1 To N
        cc.Add CStr(x(i))
    Next i
End Sub

Private Sub cmdPrint_Click()
    If IsNull(BatchList.Bookmark) Then Exit Sub
    OnPrint
    BatchList.SetFocus
End Sub

Private Sub OnPrint()
    
Dim x As String
    
    If BalintFolder = "" Then
        
        x = Mid(App.Path, 1, 3) & "Balint\GLPrint.exe" & _
            " SysFile=\Balint\Data\GLSystem.mdb" & _
            " UserID=" & curUser & _
            " BackName=\Balint\GLEntry.exe" & _
            " Batch=" & xdb.Value(BatchList.Bookmark, 0) & _
            " ProgName=GLHistJnl " & _
            " AcctDesc=" & Me.chkAcctDesc
    Else
        
        ' balint folder used - EXE is on C:\Balint
        x = "c:\Balint\GLPrint.exe" & _
        " SysFile=" & BalintFolder & "\Data\GLSystem.mdb" & _
        " UserID=" & curUser & _
        " BackName=c:\Balint\GLEntry.exe" & _
        " Batch=" & xdb.Value(BatchList.Bookmark, 0) & _
        " ProgName=GLHistJnl " & _
        " AcctDesc=" & Me.chkAcctDesc & _
        " BalintFolder=" & BalintFolder
    
    End If
     
     If Password <> "" Then
        x = x & " dbPwd=" & Password
     End If
     
     TaskID = Shell(x, vbMaximizedFocus)
     Unload Me
     End
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2: OnAdd (False)
        Case vbKeyF3: OnEdit
        Case vbKeyF4: OnDataEntry
        Case vbKeyF6: OnPrint
        Case vbKeyF7: cmdDelete_Click
        Case vbKeyF8: cmdCopy_Click
    End Select
End Sub



Private Sub SetUserNames()

Dim i, j, k As Long

    use.OpenDB
    
    j = xdb.UpperBound(1)
    For i = 1 To j
        k = CLng(xdb.Value(i, 7))
        If use.FindByID(k) Then
           xdb.Value(i, 7) = use.Name
        Else
           xdb.Value(i, 7) = "Not Found"
        End If
    Next i
    BatchList.ReBind
    
    use.CloseDB
End Sub

Private Sub Form_Terminate()
    If BackName <> "" Then TaskID = Shell(BackName, vbMaximizedFocus)
    Unload Me
    End
End Sub

