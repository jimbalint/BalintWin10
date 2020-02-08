VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form DataEntry 
   Caption         =   "  DATA ENTRY TABLE"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   Icon            =   "DataEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDupeRef 
      Caption         =   "Duplicate R&eference"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox chkNoDecimal 
      Caption         =   "No Deci&mal"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkCapLock 
      Caption         =   "&Caps Lock"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin TrueDBGrid80.TDBDropDown dropAccount 
      Height          =   1695
      Left            =   1080
      TabIndex        =   10
      Top             =   4440
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2990
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
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
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   -1  'True
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   12632256
      ValueTranslate  =   0   'False
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
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
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H808080&"
      _StyleDefs(21)  =   ":id=9,.fgcolor=&HFFFFFF&"
      _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(45)  =   "Named:id=33:Normal"
      _StyleDefs(46)  =   ":id=33,.parent=0"
      _StyleDefs(47)  =   "Named:id=34:Heading"
      _StyleDefs(48)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   ":id=34,.wraptext=-1"
      _StyleDefs(50)  =   "Named:id=35:Footing"
      _StyleDefs(51)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   "Named:id=36:Selected"
      _StyleDefs(53)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=37:Caption"
      _StyleDefs(55)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(56)  =   "Named:id=38:HighlightRow"
      _StyleDefs(57)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(58)  =   "Named:id=39:EvenRow"
      _StyleDefs(59)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(60)  =   "Named:id=40:OddRow"
      _StyleDefs(61)  =   ":id=40,.parent=33"
      _StyleDefs(62)  =   "Named:id=41:RecordSelector"
      _StyleDefs(63)  =   ":id=41,.parent=34"
      _StyleDefs(64)  =   "Named:id=42:FilterBar"
      _StyleDefs(65)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   7320
      Width           =   975
   End
   Begin TrueDBGrid80.TDBGrid EntryLog 
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6165
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ACCOUNT"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ACCT DESC."
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "REFERENCE"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DESCRIPTION"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "AMOUNT"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Fixed"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      AnimateWindowDirection=   2
      DeadAreaBackColor=   8421504
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=111,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Arial"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Arial"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HC0C0C0&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HC0C0C0&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   7320
      Width           =   975
   End
   Begin VB.CheckBox AutoIncrement 
      Caption         =   " Auto Increment &Reference"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "ADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   7920
      TabIndex        =   30
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label txtHashTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "HASH TOTAL"
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label txtCredits 
      Alignment       =   1  'Right Justify
      Caption         =   "CREDITS"
      Height          =   255
      Left            =   5040
      TabIndex        =   28
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblBudget 
      Alignment       =   2  'Center
      Caption         =   "Budget Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   9240
      TabIndex        =   27
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "F7"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   7920
      Width           =   375
   End
   Begin VB.Label txtCompany 
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
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   15
      Left            =   2040
      TabIndex        =   24
      Top             =   480
      Width           =   135
   End
   Begin VB.Label txtFileName 
      Caption         =   "FileName"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label txtJournal 
      Caption         =   "Journal Name"
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "JOURNAL:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "FILE:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "F6"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "F5"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "F4"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   7920
      Width           =   255
   End
   Begin VB.Label txtBalance 
      Alignment       =   1  'Right Justify
      Caption         =   "BALANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label txtDebits 
      Alignment       =   1  'Right Justify
      Caption         =   "DEBITS"
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label lblUpdated 
      Caption         =   "Update User and Date"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1200
      Width           =   4935
   End
   Begin VB.Label lblCreated 
      Caption         =   "Created User and Date"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label txtRecords 
      Alignment       =   1  'Right Justify
      Caption         =   "RECORDS IN BATCH"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblBatchNumber 
      Caption         =   "BATCH NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ID As Long
Public userOK As Boolean

Private BatchNumber, CreateUser, UpdateUser As Long
Private LastAccount, nRecords, nAutoNum As Long
Private createDate, updateDate As Date
Private FiscalYear As Integer
Private Period As Byte
Private Debits, Credits As Currency

Private HashTotal As Long

Dim Balance As Currency
Dim xDB As New XArrayDB
Dim AcctArray As New XArrayDB
Dim Flg As Boolean
Dim AcctFind As String
Dim x As String
Dim LastDes As String

Dim SetToLastAc As Boolean
Dim AddNext As Boolean
Dim RecordCount As Integer
Dim LastNum As Long
Dim LastRef As String

Dim HRecs As Long
Dim Btch As Long
Dim I, j, k As Integer

Private Sub Form_Load()
        
    SetToLastAc = True            ' set new line to last acct# by defaul
    AddNext = True                ' auto add next line
        
    LastAccount = 0
    userOK = False
    GLBatch.GetBatch (ID)
    
    BatchNumber = GLBatch.BatchNumber
    
    If GLBatch.BatchNumber = 0 Or IsNull(GLBatch.BatchNumber) Then
       frmProgress.Hide
       MsgBox "Batch number missing !!!", vbCritical + vbOKOnly, "GL Data Entry"
       End
    End If
    
    CreateUser = GLBatch.CreateUser
    createDate = GLBatch.Created
    updateDate = GLBatch.Updated
    UpdateUser = GLBatch.UpdateUser
    nRecords = GLBatch.RecCt
    
    Period = GLBatch.Period
    
    If GLBatch.Period = 0 Or IsNull(GLBatch.Period) Then
       MsgBox "Period number missing !!!", vbCritical + vbOKOnly, "GL Data Entry"
       End
    End If
    
    FiscalYear = GLBatch.FiscalYear
    If GLBatch.FiscalYear = 0 Or IsNull(GLBatch.FiscalYear) Then
       MsgBox "Fiscal Year missing !!!", vbCritical + vbOKOnly, "GL Data Entry"
       End
    End If
    
    lblBatchNumber = "Batch # " & BatchNumber & ":" & GLCompany.MonthName(Period, FiscalYear)
    lblCreated = "Created by " & GLUser.Name & " on " & CStr(createDate)
    lblUpdated = "Record is OPEN (Not Updated)"
    
    ' load the account drop down array with posting accts - type = "0"
    ' GLAccount.OpenRS
    If GLAccount.GetAcctsByType("0") = False Then
        MsgBox "No type 0 accounts found!", vbExclamation
        GoBack
    End If
    
    AcctArray.ReDim 1, GLAccount.RecCt, 0, 1
    I = 0
    Do
        I = I + 1
        AcctArray(I, 0) = GLAccount.Account
        AcctArray(I, 1) = GLAccount.Description
        If GLAccount.GetNext = False Then Exit Do
    Loop
    
    dropAccount.Array = AcctArray
    
    I = 0
    If GLHistory.GetByString("SELECT * FROM GLHistory WHERE BatchNumber = " & BatchNumber & " ORDER BY ID") = True Then
        xDB.ReDim 1, GLHistory.Records, 0, 6
        Do
            I = I + 1
            
            If GLAccount.GetAccount(GLHistory.Account) = True Then
                x = GLAccount.Description
            Else
                x = "GL Acct not found!"
            End If
            
            xDB(I, 0) = GLHistory.Account
            xDB(I, 1) = x
            xDB(I, 2) = GLHistory.Reference
            xDB(I, 3) = GLHistory.Description
            xDB(I, 4) = GLHistory.Amount
            If GLHistory.GetNext = False Then Exit Do
        Loop
    Else
        xDB.ReDim 0, 0, 0, 6
    End If
    
    ' Set xDB = xFactory.GetHistory(FileName, BatchNumber)
    
    Set EntryLog.Array = xDB
    EntryLog.Columns(0).Width = 1200    ' ACCT
    EntryLog.Columns(1).Width = 3000    ' ACCT DESC
    EntryLog.Columns(2).Width = 1400    ' REFERENCE
    EntryLog.Columns(3).Width = 2400    ' HIST. DESC
    EntryLog.Columns(4).Width = 2000    ' AMOUNT
    
    EntryLog.Columns(4).Alignment = dbgRight
    EntryLog.Columns(4).NumberFormat = "###,###,##0.00"
    
    EntryLog.AlternatingRowStyle = True
    
    Dim ndx, temp As Long
    nAutoNum = 0
    ShowBalance
    For ndx = 1 To nRecords
        If IsNumeric(xDB.Value(ndx, 2)) Then
            temp = CLng(xDB.Value(ndx, 2))
            If temp > nAutoNum Then nAutoNum = temp
        End If
    Next ndx
    
    ' GLCompany.GetRecord (curCompany)
    
    txtCompany = GLCompany.Name
    ' txtFileName = FileName
    
    If GLBatch.JournalSource < 100 Then
        GLJournal.GetData GLBatch.JournalSource
    Else
        GLJournal.GetData GLBatch.JournalSource - 100
    End If
    
    txtJournal = CStr(GLJournal.JournalSource) & "-" & GLJournal.JournalName
    
    EntryLog.Columns(0).DropDown = dropAccount

    ' budget entry ???
    If GLBatch.JournalSource > 100 Then
       Me.lblBudget = "Budget Entry"
    Else
       Me.lblBudget = ""
    End If
    
    frmProgress.Hide
    
    OnAdd   ' start with a blank line for entry

End Sub

Private Sub OnAdd()
    
    LastNum = 0
    
    EntryLog.Update

    If xDB.UpperBound(1) = 0 Then
        
        xDB.ReDim 1, 1, xDB.LowerBound(2), xDB.UpperBound(2)
        EntryLog.Bookmark = 1
        EntryLog.ReBind
    
    Else
        
        If Not IsNull(EntryLog.Bookmark) Then
           If IsNumeric(xDB.Value(EntryLog.Bookmark, 2)) Then
              If CLng(xDB.Value(EntryLog.Bookmark, 2)) <> 0 Then
                 LastNum = CLng(xDB.Value(EntryLog.Bookmark, 2))
                 LastRef = xDB.Value(EntryLog.Bookmark, 2)
              End If
           End If
        Else

        End If

'        xdb.InsertRows 1
'        EntryLog.Bookmark = 1
'        EntryLog.ReBind
    
        xDB.AppendRows 1
        
        EntryLog.Bookmark = xDB.UpperBound(1)
        EntryLog.ReBind
    
    End If
    
    ' date/time stamp for GLHistory.PostDate
    If Not IsNull(EntryLog.Bookmark) Then
       xDB.Value(EntryLog.Bookmark, 6) = Now()
    End If
    
    If SetToLastAc = True Then
        xDB.Value(EntryLog.Bookmark, 0) = CStr(LastAccount)
        If xDB(EntryLog.Bookmark, 0) <> 0 Then AcctDescription
    End If
    
    If AutoIncrement.Value = 1 Then
       If LastNum <> 0 Then xDB.Value(EntryLog.Bookmark, 2) = CStr(LastNum + 1)
       EntryLog.ReBind
    End If
    
    If chkDupeRef And Not IsNull(LastRef) Then
        xDB.Value(EntryLog.Bookmark, 2) = LastRef
    End If
    
    ' stuff the last description
    If Not IsNull(LastDes) Then
       xDB.Value(EntryLog.Bookmark, 3) = LastDes
    End If
    
    xDB.Value(EntryLog.Bookmark, 5) = "0"
    
    EntryLog.ReBind
    
    If SetToLastAc = True Then
        EntryLog.Col = 2
        EntryLog.Col = 0
    Else
        EntryLog.Col = 0
    End If
    
    ShowBalance

    ' don't add current line to hash total yet
    HashTotal = (HashTotal - CLng(xDB.Value(EntryLog.Bookmark, 0))) Mod 10 ^ 9
    txtHashTotal = "HASH Total " & Format(HashTotal, "########0")
    

' suzy
'    EntryLog.SetFocus

End Sub

Private Sub AutoIncrement_Click()
    EntryLog.SetFocus
    If AutoIncrement Then chkDupeRef = False
End Sub

Private Sub chkDupeRef_Click()
    If chkDupeRef Then AutoIncrement = False
End Sub

Private Sub cmdAdd_Click()
    OnAdd

    EntryLog.SetFocus

'    EntryLog.Row = 1
'    EntryLog.Bookmark = 1
'    EntryLog.Refresh

End Sub

Private Sub cmdDelete_Click()
    OnDelete
End Sub

Private Sub CmdExit_Click()
    
Dim TID As Double
Dim tm As Single
Dim Resp As Integer
    
    Resp = MsgBox("Would you like to SAVE before EXITING ?", _
           vbQuestion + vbYesNoCancel + vbDefaultButton1, "GL Data Entry")
       
    If Resp = vbCancel Then Exit Sub
     
    If Resp = vbYes Then
       
       OnSave
    
       ' call to update program
       If BalintFolder = "" Then
            x = "\Balint\GLUtil.exe " & _
                "SysFile=\Balint\Data\GLSystem.mdb " & _
                "UserID=" & GLUser.ID & " " & _
                "BackName=\Balint\GLEntryADO.exe " & _
                "ProgName=UpdateB " & _
                "Batch=" & BatchNumber
       Else
            x = "c:\Balint\GLUtil.exe " & _
               "SysFile=" & BalintFolder & "\Data\GLSystem.mdb " & _
               "UserID=" & GLUser.ID & " " & _
               "BackName=" & "c:\Balint\GLEntryADO.exe " & _
               "ProgName=UpdateB " & _
               "Batch=" & BatchNumber & _
               " BalintFolder=" & BalintFolder
       End If
       
       If Password <> "" Then
          x = x & " dbPwd=" & Password
       End If
           
       If Not TestMode Then TID = Shell(x, vbMaximizedFocus)
       Unload Me
       End
    
    End If
    
    EntryLog.Update

'    xFactory.PutHistory FileName, xdb, EntryLog.Bookmark
    Me.Hide
End Sub

Private Sub ShowBalance()
    Dim ndx As Long
    Dim temp As Currency
    nRecords = xDB.UpperBound(1)
    Debits = 0
    Credits = 0
    HashTotal = 0
    For ndx = 1 To nRecords
        temp = CCur(xDB.Value(ndx, 4))
        If temp > 0 Then
            Debits = Debits + temp
        Else
            Credits = Credits + temp
        End If
        HashTotal = (CLng(xDB.Value(ndx, 0)) + HashTotal) Mod 10 ^ 9
    Next ndx
    txtRecords = CStr(nRecords) & " Records"
    txtDebits = "Debits = " & ShowValue(Debits)
    txtCredits = "Credits = " & ShowValue(Credits)
    txtHashTotal = "Hash Total = " & Format(HashTotal, "########0")
    Balance = Round(Debits + Credits, 2)
    If Balance = 0# Then
        txtBalance = ""
    Else
        txtBalance = "BALANCE = " & ShowValue(Balance)
    End If
End Sub

Private Sub cmdPrint_Click()
    ' OnPrint
End Sub

Private Sub OnEdit()
End Sub

Private Sub OnDelete()
    
    x = "Are you SURE you want to delete: " & vbCrLf & _
        xDB.Value(EntryLog.Bookmark, 0) & " for: " & _
        Format(xDB.Value(EntryLog.Bookmark, 4), "Currency")
      
    If MsgBox(x, vbQuestion + vbYesNo + vbDefaultButton2, "GL Data Entry") = vbNo Then Exit Sub
    
    Dim ndx As Long
    ndx = EntryLog.Bookmark
    If ndx < 0 Then Exit Sub
    xDB.DeleteRows (ndx)
    If ndx > xDB.UpperBound(1) Then
        EntryLog.Bookmark = xDB.UpperBound(1)
    End If
      
    EntryLog.ReBind
    ShowBalance
    EntryLog.Bookmark = 1
    EntryLog.SetFocus

End Sub

Private Sub cmdSave_Click()
    OnSave
End Sub

Private Sub EntryLog_AfterColEdit(ByVal ColIndex As Integer)
    
Dim x As String
Dim Y As String

Dim I As Integer
Dim j As Integer
    
Dim l As Double
    
Dim dc As Double
    
Dim NegVal As Boolean
Dim DecEntered As Boolean
    
    ' account number entered
    If ColIndex = 0 Then
        EntryLog.Update
        
        If IsNumeric(xDB.Value(EntryLog.Bookmark, 0)) Then
            
            LastAccount = CLng(xDB.Value(EntryLog.Bookmark, 0))
            
            AcctDescription
            
            ' make sure the account exists
            If AcctFind = "NOT FOUND" Then
               MsgBox "Account not found !!!", vbExclamation + vbOKOnly, "GL Data Entry"
               Exit Sub
            End If
                 
            If Mid(AcctFind, 1, 10) = "Wrong Type" Then
               MsgBox AcctFind, vbExclamation + vbOKOnly, "GL Data Entry"
            End If
        
        Else
            AcctDescription
            MsgBox "Invalid Account Number has been entered!!!", vbExclamation, "GL Data Entry"
            Exit Sub
        End If
    
    End If
    
    If ColIndex = 2 Then
        EntryLog.Update
        LastRef = xDB.Value(EntryLog.Bookmark, 2)
    End If

    If ColIndex = 3 Then    ' save the last description
       EntryLog.Update
       LastDes = xDB.Value(EntryLog.Bookmark, 3)
    End If
    
    If ColIndex = 4 Then     ' update balance sums
        
        EntryLog.Update
        
        If Not IsNumeric(xDB.Value(EntryLog.Bookmark, 4)) Then
           Beep
           EntryLog.Col = 4
           Exit Sub
        End If
        
        ' decimal point not required
        ' handle if decimal is entered
        If chkNoDecimal Then
           
           x = xDB.Value(EntryLog.Bookmark, 4)
           Y = ""
           j = Len(x)
           NegVal = False
           DecEntered = False
           For I = 1 To j
              If Mid(x, I, 1) = "-" Then
                 NegVal = True
              Else
                 If Mid(x, I, 1) = "." Then
                    DecEntered = True
                 End If
                 Y = Trim(Y) & Mid(x, I, 1)
              End If
           Next I
           
           ' put in the decimal if not entered
           If Not DecEntered Then
              If Len(Y) = 1 Then
                 Y = "0.0" & Y
              ElseIf Len(Y) = 2 Then
                 Y = "0." & Y
              Else
                 Y = Mid(Y, 1, Len(Y) - 2) & "." & Mid(Y, Len(Y) - 1, 2)
              End If
           End If
           
           dc = CDbl(Y)
           If NegVal Then dc = -dc
           
           xDB.Value(EntryLog.Bookmark, 4) = Format(dc, "###,###,##0.00")
           
        End If
        
        If CLng(xDB.Value(EntryLog.Bookmark, 0)) = -1 And _
           CCur(xDB.Value(EntryLog.Bookmark, 4)) <> 0 Then
           MsgBox "Amount MUST be ZERO for Memo Entry !", vbExclamation + vbOKOnly, "GL Data Entry"
           Exit Sub
        End If
        
        EntryLog.ReBind
        
        ShowBalance
        
'        If AddNext = True And EntryLog.Bookmark = 1 Then OnAdd
        
        If AddNext = True And EntryLog.Bookmark = xDB.UpperBound(1) Then OnAdd
    
    End If

End Sub

Private Sub EntryLog_KeyPress(KeyAscii As Integer)
    ' Convert key to upper case
    If chkCapLock.Value = 1 Then
       KeyAscii = Asc(UCase(Chr$(KeyAscii)))
    End If

End Sub


Private Sub EntryLog_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If LastCol = 0 Then
        If EntryLog.Col = 1 Then EntryLog.Col = 2
    End If
    If LastCol = 2 Then
        If EntryLog.Col = 1 Then EntryLog.Col = 0
    End If
    If AcctFind = "NOT FOUND" Then    ' force back to acct # if not found
       EntryLog.Col = 0
    End If
End Sub

Private Sub AcctDescription()
            
    If LastAccount = -1 Then
        AcctFind = "Memo Entry"
    Else
        
        If GLAccount.GetAccount(LastAccount) = False Then
            AcctFind = "NOT FOUND"
        Else
            AcctFind = GLAccount.Description
        End If
        
        If GLAccount.AcctType <> "0" And AcctFind <> "NOT FOUND" Then ' must be a type 0
            AcctFind = "Wrong Type: " & x
        End If
                    
    End If
            
    xDB.Value(EntryLog.Bookmark, 1) = AcctFind
    EntryLog.RefetchRow
    EntryLog.Col = 2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        ' Case vbKeyF6: OnPrint
        Case vbKeyF7: OnSave
        Case vbKeyF4: OnAdd
        Case vbKeyF5: OnDelete
    End Select
End Sub

Private Sub OnSave()
    
Dim NumAcct As Long
Dim rs As New ADODB.Recordset
Dim r As Integer
Dim TID As Double
    
    If Balance <> 0 Then
       
       r = MsgBox("Entries are not in balance !!!" & vbCrLf & _
                  "Would you like to save the data as unbalanced ? ", _
                  vbExclamation + vbYesNo + vbDefaultButton2, _
                  "GL Data Entry")
              
       If r = vbYes Then
              
          r = MsgBox("Entries are not in balance !!!" & vbCrLf & _
                     "Are you SURE you would like to save the data as unbalanced ? ", _
                     vbExclamation + vbYesNo + vbDefaultButton2, _
                     "GL Data Entry")
          If r = vbNo Then Exit Sub
       
       Else
          
          Exit Sub
       
       End If
              
    End If
    
    Dim ndx, xrecs As Integer
    
    ' remove the batch entries first
    SQLString = " DELETE * FROM GLHistory WHERE BatchNumber = " & GLBatch.BatchNumber
    cn.Execute SQLString
    
    SQLString = " SELECT * FROM GLHistory "
    rsInit SQLString, cn, rs
    
    ' loop thru array and add entries
    xrecs = xDB.UpperBound(1)
    HRecs = 0
    For ndx = 1 To xrecs
    
        NumAcct = CLng(xDB.Value(ndx, 0))
        
        If (NumAcct > 0 Or NumAcct = -1) = False Then GoTo NextHist
        If IsEmpty(xDB.Value(ndx, 4)) Then GoTo NextHist
        
        HRecs = HRecs + 1
        rs.AddNew
        xDB.Value(ndx, 5) = CStr(rs!ID)
        rs!BatchNumber = GLBatch.BatchNumber
        
        rs!SourceCode = 0       ' ???

        If GLBatch.JournalSource >= 100 Then
           rs!HisType = "B"
        Else
           rs!HisType = "A"
        End If
        
        rs!UpdateFlag = False
        
        If IsEmpty(xDB.Value(ndx, 2)) Then xDB.Value(ndx, 2) = " "
        If IsEmpty(xDB.Value(ndx, 3)) Then xDB.Value(ndx, 3) = " "
        
        rs!Account = CLng(xDB.Value(ndx, 0))
        rs!Reference = Left(xDB.Value(ndx, 2), 20)
        rs!Description = Left(xDB.Value(ndx, 3), 20)
        rs!Amount = CCur(xDB.Value(ndx, 4))
        rs!PostDate = CDate(xDB.Value(ndx, 6))
        rs!FiscalYear = FiscalYear
        rs!Period = Period
        rs!JournalSource = GLBatch.JournalSource
        rs!SourceCode = 0   ' ????
        rs.Update
        ' PutHistory = PutHistory + 1
        
NextHist:
    Next ndx
    
    If ndx = 0 Then
        MsgBox "No Records Saved", vbInformation + vbOKOnly, "GL Data Entry"
    Else
        MsgBox CStr(HRecs) & " Records Saved", vbInformation + vbOKOnly, "GL Data Entry"
    End If
    
    ' GLBatch.nRecords = ndx
    Btch = GLBatch.BatchNumber
    GLBatch.Records = HRecs
    GLBatch.Debits = Debits
    GLBatch.Credits = Credits
    GLBatch.Save (Equate.RecPut)
    
    ' reget the batch
    If GLBatch.GetBatch(Btch) = False Then
        MsgBox "GL Batch error: " & Btch, vbExclamation
        GoBack
    End If
    
    EntryLog.SetFocus
End Sub

Private Sub OnPrint()
    ' ReviewReport.BatchNumber = BatchNumber
    ' ReviewReport.Show vbModal
End Sub







