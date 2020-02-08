VERSION 5.00
Object = "{82392BA0-C18D-11D2-B0EA-00A024695830}#1.0#0"; "ticaldr8.ocx"
Begin VB.Form FormDate 
   Caption         =   "Form3"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3975
   ScaleWidth      =   5565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin TDBCalendar6Ctl.TDBCalendar TDBCalendar1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _Version        =   65536
      _ExtentX        =   8916
      _ExtentY        =   4471
      ShowContextMenu =   -1  'True
      Appearance      =   1
      AutoSize        =   0   'False
      BorderStyle     =   1
      BackColor       =   -2147483643
      StartOfMonth    =   0
      EmptyRows       =   0
      Enabled         =   -1  'True
      FirstMonth      =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineColors0     =   -2147483632
      LineStyles0     =   0
      LineColors1     =   -2147483632
      LineStyles1     =   0
      LineColors2     =   -2147483632
      LineStyles2     =   0
      LineColors3     =   -2147483632
      LineStyles3     =   0
      LineColors4     =   -2147483632
      LineStyles4     =   0
      LineColors5     =   -2147483632
      LineStyles5     =   0
      LineColors6     =   -2147483632
      LineStyles6     =   2
      MarginBottom    =   0
      MarginTitle     =   0
      MarginTop       =   0
      MarginLeft      =   0
      MarginRight     =   0
      MarginWidth     =   0
      MarginHeight    =   0
      MaxDate         =   5373484
      MinDate         =   1757585
      MousePointer    =   0
      YearType        =   0
      MonthRows       =   1
      MonthCols       =   1
      MultiSelect     =   0
      NavOrientation  =   2
      ScrollRate      =   1
      ScrollTipAlign  =   0
      SelEdgeWidth    =   8
      SelectStyle     =   0
      SelectWhat      =   0
      ShowMenu        =   -1  'True
      ShowNavigator   =   3
      ShowScrollTip   =   -1  'True
      ShowTrailing    =   -1  'True
      StartOfWeek     =   1
      Templates       =   0
      TipInterval     =   500
      TitleHeight     =   0
      TitleFormat     =   "mmmm yyy"
      ValueIsNull     =   0   'False
      Value           =   2455266
      OverrideTipText =   ""
      TopDate         =   2455257
      AttribStyles    =   "FormDate.frx":0000
      StyleSets       =   "FormDate.frx":00C0
      CtrlType        =   8
      CtrlValue       =   "CtrlStyle"
      DayType         =   8
      DayValue        =   "DayStyle"
      TitleType       =   8
      TitleValue      =   "TitleStyle"
      WeekType        =   8
      WeekValue       =   "WeekStyle"
      TrailType       =   8
      TrailValue      =   "TrailAttrib"
      SelType         =   8
      SelValue        =   "SelAttrib"
      WeekRests0      =   0
      WeekReflect0    =   0
      WeekCaption0    =   "Sun"
      WeekAttrib0Type =   8
      WeekAttrib0Value=   "SunAttrib"
      WeekRests1      =   0
      WeekReflect1    =   0
      WeekCaption1    =   "Mon"
      WeekAttrib1Type =   1
      WeekRests2      =   0
      WeekReflect2    =   0
      WeekCaption2    =   "Tue"
      WeekAttrib2Type =   1
      WeekRests3      =   0
      WeekReflect3    =   0
      WeekCaption3    =   "Wed"
      WeekAttrib3Type =   1
      WeekRests4      =   0
      WeekReflect4    =   0
      WeekCaption4    =   "Thu"
      WeekAttrib4Type =   1
      WeekRests5      =   0
      WeekReflect5    =   0
      WeekCaption5    =   "Fri"
      WeekAttrib5Type =   1
      WeekRests6      =   0
      WeekReflect6    =   0
      WeekCaption6    =   "Sat"
      WeekAttrib6Type =   8
      WeekAttrib6Value=   "SatAttrib"
      HolidayStyles   =   "FormDate.frx":0224
      UserStyles      =   ""
      Key             =   "FormDate.frx":0240
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
End
Attribute VB_Name = "FormDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SelDate As Date

Private Sub Command1_Click()
    SelDate = Me.TDBCalendar1
    Me.Hide
End Sub
