VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form MainMenu 
   Caption         =   " GENERAL LEDGER DATA ENTRY BATCH LIST"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "MainMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10215
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
      Width           =   9495
      _ExtentX        =   16748
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.Label Label15 
      Caption         =   "ADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   7200
      Width           =   975
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
Dim xDB As New XArrayDB
Dim SortBy(7) As Byte
Dim I, J, K As Long
Dim HeadString(7) As String
Dim SortType(7) As Variant
Dim x As String
Dim ndx As Integer
Dim FileName As String

Private Sub Form_Load()
    
Dim I, J As Long
Dim x As String
    
    On Error GoTo glErr
    lblUser = GLUser.Name

    FileName = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    lblFileName = Mid(FileName, 3, Len(FileName) - 2)
    
    lblCompanyName = GLCompany.Name
    lblAddress = GLCompany.Address1
    lblCityStateZip = GLCompany.City & ", " & GLCompany.State & "  " & GLCompany.ZipCode
    CompanyID = GLCompany.ID
        
    ' populate the batch grid
    SQLString = " SELECT * FROM GLBatch ORDER BY ID DESC"
    If GLBatch.GetByString(SQLString) Then
        xDB.ReDim 1, GLBatch.RecCt, 0, 7
        ndx = 0
        Do
            ndx = ndx + 1
            xDB.Value(ndx, 0) = CStr(GLBatch.BatchNumber)
            xDB.Value(ndx, 1) = ShowDate(GLBatch.Updated)
            xDB.Value(ndx, 2) = CStr(GLBatch.FiscalYear) & "-" & Format(GLBatch.Period, "00")
            xDB.Value(ndx, 3) = ShowValue(GLBatch.Debits)
            xDB.Value(ndx, 4) = ShowValue(GLBatch.Credits)
            xDB.Value(ndx, 5) = CStr(GLBatch.Records)
            xDB.Value(ndx, 6) = CStr(GLBatch.JournalSource)
            xDB.Value(ndx, 7) = CStr(GLBatch.UpdateUser)
            If GLBatch.GetNext = False Then Exit Do
        Loop
    Else
        xDB.ReDim 1, 0, 0, 7
    End If

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
    For I = 1 To 7
        SortBy(I) = 0
        BatchList.Columns(I).Caption = HeadString(I)
    Next I

    BatchList.Columns(0).Width = 800
    BatchList.Columns(1).Width = 1000
    BatchList.Columns(2).Width = 1000
    BatchList.Columns(3).Width = 1600
    BatchList.Columns(4).Width = 1600
    BatchList.Columns(5).Width = 1000
    BatchList.Columns(6).Width = 1000
    BatchList.Columns(7).Width = 1000
    
    BatchList.Columns(3).Alignment = dbgRight
    BatchList.Columns(4).Alignment = dbgRight
    BatchList.Columns(3).NumberFormat = "###,###,##0.00"
    BatchList.Columns(4).NumberFormat = "###,###,##0.00"
    
    BatchList.Columns(5).Alignment = dbgRight
    BatchList.Columns(6).Alignment = dbgRight
    
    SetUserNames
    
    BatchList.AlternatingRowStyle = True
    BatchList.AllowColSelect = False
    
    ' JS column adjust - show if budget - js > 100
    J = xDB.UpperBound(1)
    For I = 1 To J
        If CLng(xDB(I, 6)) > 100 Then
           xDB(I, 6) = "BUDG " & CLng(xDB(I, 6)) - 100
        End If
    Next I
    
    Set BatchList.Array = xDB
    
    frmProgress.Hide
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub BatchList_HeadClick(ByVal ColIndex As Integer)

Dim SortOrder As Byte

    J = xDB.UpperBound(1)

    For I = 0 To 7
        
        If I = ColIndex Then
           
           If SortBy(I) = 0 Or SortBy(I) = 2 Then
              ' descending
              SortBy(I) = 1
              x = "-"
              SortOrder = XORDER_DESCEND
           ElseIf SortBy(I) = 1 Then
              ' ascending
              SortBy(I) = 2
              x = "+"
              SortOrder = XORDER_ASCEND
           End If
        
           BatchList.Columns(I).Font.Bold = True
           BatchList.Columns(I).Caption = UCase(HeadString(I) & x)
        
        Else          ' not sorted by
        
           BatchList.Columns(I).Caption = HeadString(I)
           BatchList.Columns(I).Font.Bold = False
        
        End If
    
    Next I
        
    ' sort it
    xDB.QuickSort 1, J, ColIndex, SortOrder, SortType(ColIndex)

    Set BatchList.Array = xDB
    BatchList.ReBind
    BatchList.Refresh
    BatchList.Col = 0
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
        BatchFrom = xDB.Value(BatchList.Bookmark, 0)
    End If
    
    GLBatch.AddBatch GLCompany.CurFiscalYear, GLCompany.CurPeriod
    
    ' reget the batch - save closes it
    If GLBatch.GetBatch(GLBatch.BatchNumber) = False Then
        MsgBox "GL Batch error! " & GLBatch.BatchNumber, vbExclamation
        GoBack
    End If
    
    GLBatch.FiscalYear = GLCompany.CurFiscalYear
    GLBatch.Period = GLCompany.CurPeriod
    GLBatch.Updated = Now
    GLBatch.Save (Equate.RecPut)
    
    ' reget the batch - save closes it
    If GLBatch.GetBatch(GLBatch.BatchNumber) = False Then
        MsgBox "GL Batch error! " & GLBatch.BatchNumber, vbExclamation
        GoBack
    End If
    
    xDB.InsertRows 1
    BatchList.Bookmark = 1
    xDB.Value(BatchList.Bookmark, 0) = CStr(GLBatch.BatchNumber)
    xDB.Value(BatchList.Bookmark, 1) = ShowDate(GLBatch.Updated)
    xDB.Value(BatchList.Bookmark, 2) = CStr(GLBatch.FiscalYear) & "-" & CStr(GLBatch.Period)
    xDB.Value(BatchList.Bookmark, 3) = ShowValue(GLBatch.Debits)
    xDB.Value(BatchList.Bookmark, 4) = ShowValue(GLBatch.Credits)
    xDB.Value(BatchList.Bookmark, 5) = CStr(GLBatch.RecCt)
    xDB.Value(BatchList.Bookmark, 6) = CStr(GLBatch.JournalSource)
    
    BatchList.ReBind
    
    If BatchCpy = False Then
        BatchForm.BatchNumber = GLBatch.BatchNumber
        BatchForm.Init
        BatchForm.Show vbModal
        Unload BatchForm
    Else
        BatchCopy.BatchNumberC = GLBatch.BatchNumber
        BatchCopy.BatchFrom = BatchFrom
        BatchCopy.Init
        BatchCopy.Show vbModal
        Unload BatchCopy
    End If
    
'    If BatchForm.userOK = True Then
    If Response Then
        
        If BatchCpy = False Then
            GLBatch.GetBatch (BatchForm.BatchNumber)
        Else
            GLBatch.GetBatch (BatchCopy.BatchNumberC)
        End If
        
        xDB.Value(BatchList.Bookmark, 1) = ShowDate(GLBatch.Updated)
        xDB.Value(BatchList.Bookmark, 2) = CStr(GLBatch.FiscalYear) & "-" & CStr(GLBatch.Period)
        xDB.Value(BatchList.Bookmark, 3) = ShowValue(GLBatch.Debits)
        xDB.Value(BatchList.Bookmark, 4) = ShowValue(GLBatch.Credits)
        xDB.Value(BatchList.Bookmark, 5) = CStr(GLBatch.RecCt)
        
        If GLBatch.JournalSource > 100 Then
           xDB(BatchList.Bookmark, 6) = "BUDG " & CStr(GLBatch.JournalSource - 100)
        Else
           xDB.Value(BatchList.Bookmark, 6) = CStr(GLBatch.JournalSource)
        End If
        
'        I = GLBatch.UpdateUser
'        use.GetRecord (I)
        xDB.Value(BatchList.Bookmark, 7) = GLUser.Name
        
        BatchList.RefetchRow
        OnDataEntry
    
    Else
        xDB.DeleteRows 1
        GLBatch.DeleteBatch BatchNum
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
Dim I As Long
'Dim db As DAO.Database
Dim Fnm As String
Dim x As String

'    On Error GoTo glErr
'    Dim bat As New rBatch
    
    BMark = BatchList.Bookmark
    
    If IsNull(BMark) Then Exit Sub
    
    BatchNum = CLng(xDB.Value(BMark, 0))
    
    I = MsgBox("Are you SURE you want to copy this Batch # " & BatchNum, _
        vbQuestion + vbYesNo + vbDefaultButton2, "Windows GL Entry")
    
    If I = vbNo Then Exit Sub
    
    OnAdd (True)
    
End Sub

Private Sub cmdDataEntry_Click()
    OnDataEntry
End Sub

Private Sub OnDataEntry()
    
Dim BMark
    
    If IsNull(BatchList.Bookmark) Then Exit Sub
    
    frmProgress.Show
    
    BMark = BatchList.Bookmark
    
    DataEntry.ID = xDB.Value(BatchList.Bookmark, 0)
    DataEntry.Show vbModal
    Unload DataEntry


'    glbatch.GetSQL "select * from glBatch where BatchNumber=" & BatchList.SelectedItem.Text, FileName
'    BatchList.SelectedItem.SubItems(1) = Format(bat(1).Updated, "mm/dd/yy")
'    BatchList.SelectedItem.SubItems(2) = CStr(bat(1).FiscalYear) & "-" & CStr(bat(1).Period)
'    BatchList.SelectedItem.SubItems(3) = gl.ShowValue(bat(1).Debits)
'    BatchList.SelectedItem.SubItems(4) = gl.ShowValue(bat(1).Credits)
'    BatchList.SelectedItem.SubItems(5) = CStr(bat(1).nRecords)
    
    GLBatch.GetBatch (xDB.Value(BatchList.Bookmark, 0))
    
    xDB.Value(BatchList.Bookmark, 1) = GLBatch.Updated
    xDB.Value(BatchList.Bookmark, 2) = GLBatch.FiscalYear & "-" & GLBatch.Period
    xDB.Value(BatchList.Bookmark, 3) = GLBatch.Debits
    xDB.Value(BatchList.Bookmark, 4) = GLBatch.Credits
    xDB.Value(BatchList.Bookmark, 5) = GLBatch.Records
    
    BatchList.ReBind
    BatchList.SetFocus
    
    Exit Sub

glErr:
    MsgBox Error(Err.Number)
End Sub

Private Sub cmdDelete_Click()

Dim BMark
Dim BatchNum As Long
Dim I As Long
'Dim db As DAO.Database
Dim Fnm As String
Dim x As String
Dim DriveLetter As String

'    On Error GoTo glErr
'    Dim bat As New rBatch
    
    BMark = BatchList.Bookmark
    
    If IsNull(BMark) Then Exit Sub
    
    BatchNum = CLng(xDB.Value(BMark, 0))
    
    I = MsgBox("Are you SURE you want to delete this Batch # " & BatchNum, _
        vbCritical + vbYesNo + vbDefaultButton2, "Windows GL Entry")
    
    If I = vbNo Then Exit Sub
    
    Fnm = Mid(App.Path, 1, 2) & Mid(FileName, 3, Len(FileName) - 2)
    
    ' store the period yyyymm from the batch record
' 107   I = Mid(xDB(BMark, 2), 1, 4) * 100 + Mid(xDB(BMark, 2), 6, 2)
 
    ' delete from GLHistory
    GLBatch.DeleteBatch (BatchNum)
    SQLString = " DELETE * FROM GLHistory WHERE BatchNumber = " & BatchNum
    cn.Execute SQLString
    
'    x = "\Balint\GLUtil.exe" & _
'        " ProgName=ClearGLAmount " & _
'        " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
'        " UserID=" & curUser & _
'        " BackName=" & DriveLetter & "\Balint\GLEntry.exe" & _
'        " Period=" & I


    DriveLetter = Left(BalintFolder, 1)
    x = "\Balint\GLUtil.exe" & _
        " ProgName=ClearGLAmount " & _
        " SysFile=" & DriveLetter & "\Balint\Data\GLSystem.mdb" & _
        " UserID=" & GLUser.ID & _
        " MenuName=" & MenuName & _
        " BackName=" & DriveLetter & "\Balint\GLEntry.exe" & _
        " Period=" & I
    
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
    
    If xDB.Value(BatchList.Bookmark, 5) <> 0 Then
       MsgBox "Batch edit not allowed if history records exist!", vbExclamation + vbOKOnly, "GL Data Entry"
       Exit Sub
    End If
    
    BatchForm.BatchNumber = xDB.Value(BatchList.Bookmark, 0)
    BatchForm.Init
    BatchForm.Show vbModal
    If BatchForm.userOK = True Then
        GLBatch.GetBatch BatchForm.BatchNumber
        xDB.Value(BatchList.Bookmark, 1) = ShowDate(GLBatch.Updated)
        xDB.Value(BatchList.Bookmark, 2) = CStr(GLBatch.FiscalYear) & "-" & CStr(GLBatch.Period)
        xDB.Value(BatchList.Bookmark, 3) = ShowValue(GLBatch.Debits)
        xDB.Value(BatchList.Bookmark, 4) = ShowValue(GLBatch.Credits)
        xDB.Value(BatchList.Bookmark, 5) = CStr(GLBatch.RecCt)
        xDB.Value(BatchList.Bookmark, 6) = CStr(GLBatch.JournalSource)
        BatchList.ReBind
        BatchList.RefetchRow
    End If
    Unload BatchForm
    BatchList.SetFocus
End Sub

Private Sub CmdExit_Click()
    
Dim x As String
    
'    If TestMode Then
'        End
'    End If
'
'    MenuName = "GLMenu.exe"
'    If BalintFolder = "" Then
'        BackName = "\Balint\" & MenuName
'    Else
'        BackName = BalintFolder & "\" & MenuName
'    End If
    
    GoBack
    
'    If BackName <> "" Then
'       x = BackName & " UserID=" & GLUser.ID
'       If Password <> "" Then
'          x = x & " dbPwd=" & Password
'       End If
'       TaskID = Shell(x, vbMaximizedFocus)
'    End If
'    Unload Me
'    End
    
    
End Sub


Private Sub Sort(ByRef cc As Collection)

    Dim I, J, n, Temp, x() As Integer
    n = cc.Count
    ReDim x(1 To n)
    For I = n To 1 Step -1
        x(I) = CInt(cc(I))
        cc.Remove I
    Next I
    For I = 1 To n - 1
        For J = I + 1 To n
            If x(I) > x(J) Then
                Temp = x(I)
                x(I) = x(J)
                x(J) = Temp
            End If
        Next J
    Next I
    For I = 1 To n
        cc.Add CStr(x(I))
    Next I
End Sub

Private Sub cmdPrint_Click()
    If IsNull(BatchList.Bookmark) Then Exit Sub
    OnPrint
    BatchList.SetFocus
End Sub

Private Sub OnPrint()
    
Dim x As String
    
'            " SysFile=\Balint\Data\GLSystem.mdb"
'            " SysFile=\Balint\Data\GLSystem.mdb"
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 3) & "Balint\GLPrint.exe" & _
            " UserID=" & GLUser.ID & _
            " BackName=\Balint\GLEntryADO.exe" & _
            " Batch=" & xDB.Value(BatchList.Bookmark, 0) & _
            " MenuName=" & MenuName & _
            " ProgName=GLHistJnl " & _
            " AcctDesc=" & Me.chkAcctDesc
    Else
        x = BalintFolder & "\GLPrint.exe" & _
            " UserID=" & GLUser.ID & _
            " BackName=\Balint\GLEntryADO.exe" & _
            " Batch=" & xDB.Value(BatchList.Bookmark, 0) & _
            " ProgName=GLHistJnl " & _
            " MenuName=" & MenuName & _
            " BalintFolder=" & BalintFolder & _
            " AcctDesc=" & Me.chkAcctDesc
    End If
    
    x = "c:\Balint\GLPrint.exe" & _
        " UserID=" & GLUser.ID & _
        " BackName=c:\Balint\GLEntryADO.exe" & _
        " Batch=" & xDB.Value(BatchList.Bookmark, 0) & _
        " ProgName=GLHistJnl " & _
        " MenuName=" & MenuName & _
        " BalintFolder=" & BalintFolder & _
        " AcctDesc=" & Me.chkAcctDesc
    
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

Dim I, J, K As Long

    Dim uid As Integer
    uid = GLUser.ID

    J = xDB.UpperBound(1)
    For I = 1 To J
        K = CLng(xDB.Value(I, 7))
        If GLUser.GetByID(K) = True Then
            xDB.Value(I, 7) = GLUser.Name
        Else
            xDB.Value(I, 7) = "Not Found"
        End If
    Next I
    
    ' re-get the user
    GLUser.GetByID (uid)
    
'    use.OpenDB
'
'    j = xDB.UpperBound(1)
'    For I = 1 To j
'        k = CLng(xDB.Value(I, 7))
'        If use.FindByID(k) Then
'           xDB.Value(I, 7) = use.Name
'        Else
'           xDB.Value(I, 7) = "Not Found"
'        End If
'    Next I
'    BatchList.ReBind
'
'    use.CloseDB

End Sub

Private Sub Form_Terminate()
    If BackName <> "" Then TaskID = Shell(BackName, vbMaximizedFocus)
    Unload Me
    End
End Sub

Public Function GetBatch(ByVal DataFile As String, ByVal SQL As String) As XArrayDB
    
    Dim xDB As New XArrayDB
    Dim n, ndx As Long
    On Error GoTo glErr
    xDB.ReDim 1, 0, 0, 7
    
    SQLString = " SELECT * FROM GLBatch "
    If GLBatch.GetByString(SQLString) = False Then
        Exit Function
    End If
    
    n = GLBatch.RecCt
    xDB.ReDim 1, n, 0, 7
    
    Do
        xDB.Value(ndx, 0) = CStr(GLBatch.BatchNumber)
        xDB.Value(ndx, 1) = ShowDate(GLBatch.Updated)
        xDB.Value(ndx, 2) = CStr(GLBatch.FiscalYear) & "-" & Format(GLBatch.Period, "00")
        xDB.Value(ndx, 3) = ShowValue(GLBatch.Debits)
        xDB.Value(ndx, 4) = ShowValue(GLBatch.Credits)
        xDB.Value(ndx, 5) = CStr(GLBatch.Records)
        xDB.Value(ndx, 6) = CStr(GLBatch.JournalSource)
        xDB.Value(ndx, 7) = CStr(GLBatch.UpdateUser)
        If GLBatch.GetNext = False Then Exit Do
    Loop
    
glErr:
    Set GetBatch = xDB
    
    
'    Dim xDB As New XArrayDB
'    Dim n, ndx As Long
'    On Error GoTo glErr
'    xDB.ReDim 1, 0, 0, 7
'
'    Set db = OpenDatabase(Name:=DataFile, _
'                          Options:=False, _
'                          ReadOnly:=False, _
'                          Connect:=";pwd=" & Password)
'
'    Set rs = db.OpenRecordset(SQL)
'    rs.MoveLast
'    n = rs.RecordCount
'    xDB.ReDim 1, n, 0, 7
'    rs.MoveFirst
'    For ndx = 1 To n
'        xDB.Value(ndx, 0) = CStr(rs!BatchNumber)
'        xDB.Value(ndx, 1) = ShowDate(rs!Updated)
'        xDB.Value(ndx, 2) = CStr(rs!FiscalYear) & "-" & Format(rs!Period, "00")
'        xDB.Value(ndx, 3) = ShowValue(rs!Debits)
'        xDB.Value(ndx, 4) = ShowValue(rs!Credits)
'        xDB.Value(ndx, 5) = CStr(rs!Records)
'        xDB.Value(ndx, 6) = CStr(rs!JournalSource)
'        xDB.Value(ndx, 7) = CStr(rs!UpdateUser)
'        rs.MoveNext
'    Next ndx
'    rs.Close
'glErr:
'    Set GetBatch = xDB

End Function


