VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGLPrint 
   Caption         =   "GL Print File Setup"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   Icon            =   "GLPrintAll.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbHorzNudge 
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   7920
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   661
      Calculator      =   "GLPrintAll.frx":030A
      Caption         =   "GLPrintAll.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "GLPrintAll.frx":039C
      Keys            =   "GLPrintAll.frx":03BA
      Spin            =   "GLPrintAll.frx":0404
      AlignHorizontal =   0
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
   Begin VB.CheckBox chkTextOutput 
      Caption         =   "Text File Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   23
      Top             =   7320
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdlTextOutput 
      Left            =   240
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkAcctDesc 
      Caption         =   "&Include Account Description?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CommandButton cmdLookHi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9240
      Picture         =   "GLPrintAll.frx":042C
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cmdLookLow 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      Picture         =   "GLPrintAll.frx":0736
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4440
      Width           =   375
   End
   Begin VB.CheckBox chkBudget 
      Caption         =   "B&udget"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   6600
      Width           =   1335
   End
   Begin VB.ComboBox cmbJournalSource 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3780
      TabIndex        =   20
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "O&ther Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7155
      TabIndex        =   27
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtHiCons 
      Height          =   285
      Left            =   7320
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtLoCons 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtHiBranch 
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtLoBranch 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   5010
      Width           =   1095
   End
   Begin VB.TextBox txtHiAccount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   15
      Top             =   4410
      Width           =   1815
   End
   Begin VB.TextBox txtLoAccount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Top             =   4410
      Width           =   1815
   End
   Begin VB.ComboBox cmbStartPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Frame fraType4 
      Height          =   735
      Left            =   3960
      TabIndex        =   35
      Top             =   3360
      Width           =   6255
      Begin VB.OptionButton optIncomeStatement 
         Caption         =   "I&ncome Stmt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optBalanceSheet 
         Caption         =   "Ba&lance Sheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optBoth 
         Caption         =   "Print Bo&th"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraType3 
      Height          =   735
      Left            =   240
      TabIndex        =   34
      Top             =   3360
      Width           =   3495
      Begin VB.OptionButton optComparative 
         Caption         =   "Co&mparative"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optRegular 
         Caption         =   "&Regular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame fraType2 
      Height          =   735
      Left            =   6720
      TabIndex        =   33
      Top             =   2280
      Width           =   3375
      Begin VB.OptionButton optSchedules 
         Caption         =   "Sc&hedules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optStatements 
         Caption         =   "&Statements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame fraType1 
      Height          =   735
      Left            =   240
      TabIndex        =   32
      Top             =   2280
      Width           =   6255
      Begin VB.OptionButton optBudget 
         Caption         =   "Bud&get"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optConsolidated 
         Caption         =   "&Consolidated"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optBranch 
         Caption         =   "&Branch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "&Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4620
      TabIndex        =   26
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2085
      TabIndex        =   25
      Top             =   8520
      Width           =   1215
   End
   Begin VB.ComboBox cmbEndPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.ComboBox cmbFiscalYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2700
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblProgName 
      Alignment       =   2  'Center
      Caption         =   "Program Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   45
      Top             =   720
      Width           =   9735
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   44
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label lblJournalSource 
      Alignment       =   2  'Center
      Caption         =   "Journal Source:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   43
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lblHiCons 
      Caption         =   "Hi Consolidated Account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   42
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblLoCons 
      Caption         =   "Low Consolidated Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   41
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblHiBranch 
      Caption         =   "Hi Branch:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   40
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblLoBranch 
      Caption         =   "Low Branch:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   39
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblHiAccount 
      Caption         =   "Hi Account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   38
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblLoAccount 
      Caption         =   "Low Account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label lblStartPd 
      Caption         =   "Start Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblEndPeriod 
      Caption         =   "End Period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   31
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblFY 
      Caption         =   "Fiscal Year:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1260
      TabIndex        =   30
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmGLPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EndYMs(11) As Long
Dim StartYMs(11) As Long
Dim jj As Integer
Dim ll As Long

Public bytFrame1 As Byte
Public bytFrame2 As Byte
Public bytFrame3 As Byte
Public bytFrame4 As Byte

' need to add to GLPrint !!!
Public FiscalYear As Long
Public StartPD As Integer
Public EndPd As Integer
Public JournalSource As Integer
Public AllJnlOption As Boolean

Dim flg As Boolean
Private Sub Form_Load()
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim RetVal As Boolean
    
   ' assign the GLPrint values
'   GLPrint.Clear
'   If GLPrint.GetData("SYSTEM") = False Then
'
'      GLPrint.User = "SYSTEM"
'      GLPrint.RegBraCon = Equate.Regular     ' 1=regular  2=branch  3=consolidated  4=budget  AB%
'      GLPrint.StaSch = Equate.Stmt           ' 1=statment  2=schedule
'      GLPrint.RegCmp = Equate.Regular        ' 1=regular  2=comparative
'      GLPrint.PrintBIB = Equate.PrtBoth      ' 1=bal sht  2=inc stmt  3=both
'      GLPrint.BeginDate = 0
'      GLPrint.EndDate = 0
'      GLPrint.LowAccount = 1
'      GLPrint.HiAccount = 999999999
'      GLPrint.LowBranchAcct = 1
'      GLPrint.HiBranchAcct = 99
'      GLPrint.LowConsAcct = 1
'      GLPrint.HiConsAcct = 99
'      GLPrint.UseMathRec = True
'
'      GLPrint.Save (Equate.RecAdd)
'
'   End If
   
   GLPrint.GetData User, flg
   
   ' new GLPrint record was created for the user
   ' load defaults from the GLCompany file
   If flg = True Then
      GLPrint.LowAccount = 1
      GLPrint.HiAccount = 999999999
      GLPrint.LowBranchAcct = GLCompany.LowBranch
      GLPrint.HiBranchAcct = GLCompany.HiBranch
      GLPrint.LowConsAcct = GLCompany.LowConsolidated
      GLPrint.HiConsAcct = GLCompany.HiConsolidated
      GLPrint.Save (Equate.RecPut)
   End If
   
   Me.lblCompName = GLCompany.Name
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' set the buttons from glprint
   If GLPrint.RegBraCon = Equate.Regular Then Me.optNormal = True
   If GLPrint.RegBraCon = Equate.Branch Then Me.optBranch = True
   If GLPrint.RegBraCon = Equate.Consol Then Me.optConsolidated = True
   If GLPrint.RegBraCon = Equate.Budget Then Me.optBudget = True
   
   If GLPrint.StaSch = Equate.Stmt Then Me.optStatements = True
   If GLPrint.StaSch = Equate.Sched Then Me.optSchedules = True
   
   If GLPrint.RegCmp = Equate.Regular Then Me.optRegular = True
   If GLPrint.RegCmp = Equate.Comp Then Me.optComparative = True
   
   If GLPrint.PrintBIB = Equate.PrtBSOnly Then Me.optBalanceSheet = True
   If GLPrint.PrintBIB = Equate.PrtISOnly Then Me.optIncomeStatement = True
   If GLPrint.PrintBIB = Equate.PrtBoth Then Me.optBoth = True
   
   If GLPrint.LowAccount = 0 Then GLPrint.LowAccount = 1
   If GLPrint.HiAccount = 0 Then GLPrint.HiAccount = 999999999
   
   ' not a branch client
   If GLCompany.SubDigits = 0 Then
      Me.optConsolidated.Enabled = False
      Me.optBranch.Enabled = False
      Me.optNormal = True
      Me.txtLoBranch.Enabled = False
      Me.txtHiBranch.Enabled = False
      Me.txtLoCons.Enabled = False
      Me.txtHiCons.Enabled = False
   End If
   
   ' budget for spreadsheet tool only???
   Me.optBudget.Visible = False
   
   Me.txtLoAccount = GLPrint.LowAccount
   Me.txtHiAccount = GLPrint.HiAccount
   
   Me.txtLoBranch = GLPrint.LowBranchAcct
   Me.txtHiBranch = GLPrint.HiBranchAcct
   
   Me.txtLoCons = GLPrint.LowConsAcct
   Me.txtHiCons = GLPrint.HiConsAcct
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   GLPrint.ReportDate = 20030630

'   GLPrint.Output = ""
'   GLPrint.Copies = 1
'   GLPrint.User = "JIM"

   rs.Source = "Select DISTINCT FiscalYear from GLAmount order by FiscalYear Desc"
   
   Set rs.ActiveConnection = cn
        
   rs.Open
        
   If rs.EOF = True And rs.BOF = True Then
      MsgBox "No amount data ???"
      End
   End If

   ll = 0
   jj = 0
   Do Until rs.EOF = True
      cmbFiscalYear.AddItem rs.Fields("FiscalYear")
      If rs!FiscalYear = GLPrint.FiscalYear Then jj = ll
      rs.MoveNext
      ll = ll + 1
   Loop
   cmbFiscalYear.ListIndex = jj
   
   Set rs = Nothing
   
   EndPeriodSet (CInt(cmbFiscalYear))

   If GLPrint.BeginDate = 0 Then
      Me.cmbStartPeriod.ListIndex = 0
   Else
      For jj = 0 To 11
         If StartYMs(jj) = GLPrint.BeginDate Then
            Me.cmbStartPeriod.ListIndex = jj
         End If
      Next jj
   End If
   
   If GLPrint.EndDate = 0 Then
      Me.cmbEndPeriod.ListIndex = 0
   Else
      For jj = 0 To 11
         If EndYMs(jj) = GLPrint.EndDate Then
            Me.cmbEndPeriod.ListIndex = jj
         End If
      Next jj
   End If
      
   ' init the journal source combo
   If ProgName = "GLHISTJNL" Then
      cmbJournalSource.AddItem "All Jnls"
   End If
   
   For ll = 1 To 10
       If GLJournal.GetData(CLng(ll)) = True Then
          cmbJournalSource.AddItem ll & " - " & GLJournal.JournalName
       Else
          cmbJournalSource.AddItem ll
       End If
   Next ll
   cmbJournalSource.ListIndex = 0
   
   ' only show the budget checkbox for de jnls
   Me.chkBudget.Visible = False
   
   ' only show include acct desc if de jnls
   Me.chkAcctDesc.Visible = False
   
   ' set the screen according to the calling program
   Select Case ProgName
   
      Case "STATEMENT"
         Me.lblProgName = "Statement Print"
         Me.cmbJournalSource.Enabled = False
   
      Case "CHARTOFACCOUNTS"
         Me.lblProgName = "Chart Of Accounts"
         Me.optBranch.Enabled = False
         Me.optBudget.Enabled = False
         Me.cmdOptions.Enabled = False
         DisableFrame2
         DisableFrame3
         DisableFrame4
         DisableJS
         DisableBranchRange
         DisableConsRange
         DisableFY
         DisablePdRange
      
      Case "PRINTDESC"
         Me.lblProgName = "Print Description File"
         Me.optBranch.Enabled = False
         Me.optBudget.Enabled = False
         Me.cmdOptions.Enabled = False
         DisableFrame1
         DisableFrame2
         DisableFrame3
         DisableFrame4
         DisableJS
         DisableBranchRange
         DisableConsRange
         DisableFY
         DisablePdRange
         
         Me.lblLoAccount = "Low Desc Number:"
         Me.lblHiAccount = "Hi Desc Number:"
   
      Case "PRINTGLACCOUNT"
         
         Me.lblProgName = "Print GL Account File"
         DisableFrame1
         DisableFrame2
         DisableFrame3
         DisableFrame4
         DisableJS
'         DisableFY
'         DisablePdRange
         DisableConsRange
   
      Case "GLHISTJNL"
         
         Me.chkAcctDesc.Visible = True
         Me.chkBudget.Visible = True
         Me.AllJnlOption = True
         Me.lblProgName = "Data Entry Journal"
         Me.cmdOptions.Enabled = False
         DisableFrame1
         DisableFrame2
         DisableFrame3
         DisableFrame4
         DisableAccountRange
         DisableBranchRange
         DisableConsRange
         
      Case "DETAILGL"
         Me.lblProgName = "Detail General Ledger"
         
         DisableFrame2
         DisableFrame3
         DisableFrame4
         DisableJS
         
         Me.optBudget.Enabled = False
         
      Case "TRIALBAL"
         Me.lblCompName = GLCompany.Name
         Me.lblProgName = "Trial Balance"
         Me.cmbStartPeriod.Enabled = False
         Me.fraType1.Visible = False
         Me.fraType2.Visible = False
         Me.fraType3.Visible = False
         Me.fraType4.Visible = False
         Me.txtLoBranch.Enabled = False
         Me.txtHiBranch.Enabled = False
         Me.txtLoCons.Enabled = False
         Me.txtHiCons.Enabled = False
         Me.cmbJournalSource.Enabled = False
   
   End Select

    ' get tab - Horz Nudge
    SetNudge Me.tdbHorzNudge
    PRGlobal.OpenRS
    SQLString = "SELECT * FROM PRGlobal WHERE Description = 'GLTab' " & _
                " AND UserID = " & GLUser.ID
    If PRGlobal.GetBySQL(SQLString) = True Then
        Me.tdbHorzNudge.Value = PRGlobal.Var1
    End If

    Me.KeyPreview = True

End Sub

Private Sub cmbFiscalYear_Click()
    EndPeriodSet (CInt(cmbFiscalYear))
End Sub

Private Sub cmdColumns_Click()
    frmGLColumn.Show vbModal
End Sub

Private Sub cmbStartPeriod_LostFocus()
     
     cmbEndPeriod.ListIndex = cmbStartPeriod.ListIndex
     cmbEndPeriod.Refresh

End Sub

Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub cmdLookHi_Click()
    frmAcctLookup.Show vbModal
    Me.txtHiAccount = frmAcctLookup.SelAcct
    Me.txtLoBranch.SetFocus
End Sub

Private Sub cmdLookLow_Click()
    frmAcctLookup.Show vbModal
    Me.txtLoAccount = frmAcctLookup.SelAcct
    Me.txtHiAccount.SetFocus
End Sub

Private Sub cmdOK_Click()
     
Dim EMsg As String
Dim Response As Integer
     
    ' save the nudge setting
    SQLString = "SELECT * FROM PRGlobal WHERE Description = 'GLTab' " & _
                " AND UserID = " & GLUser.ID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.UserID = GLUser.ID
        PRGlobal.Description = "GLTab"
        PRGlobal.Save (Equate.RecAdd)
    End If
    PRGlobal.Var1 = Me.tdbHorzNudge.Value
    PRGlobal.Save (Equate.RecPut)
    TabValue = Me.tdbHorzNudge.Value
    
    ' output to text file?
    TextFileName = ""
    TextChannel = 0
    
    If chkTextOutput Then

        cdlTextOutput.CancelError = True
        
        ' set to current
        cdlTextOutput.Flags = cdlCFBoth Or cdlCFEffects
        cdlTextOutput.Filter = "Comma Separated Values|*.csv"
        cdlTextOutput.FileName = GLUser.Logon & ".csv"
        cdlTextOutput.DialogTitle = "Select a file for Text Export"
        cdlTextOutput.CancelError = True
        cdlTextOutput.InitDir = "\Balint\Data"

        ' call the file dialog
        On Error Resume Next
        cdlTextOutput.ShowOpen
        
        If Err.Number = 0 Then

            ' assign
            TextFileName = cdlTextOutput.FileName
            TextChannel = FreeFile

            Do
                
                On Error Resume Next
                Open TextFileName For Output As #TextChannel
                
                If Err.Number <> 0 Then
                    
                    ErrMsg = "Error Opening: " & TextFileName & vbCr & vbCr & _
                        " " & Err.Number & " " & Err.Description
                        
                    Response = MsgBox(ErrMsg, vbRetryCancel + vbExclamation, "File Open Error")
                    If Response <> vbRetry Then
                        TextChannel = 0
                        TextFileName = ""
                        Exit Do
                    End If
                    
                Else
                    Exit Do
                End If
            
            Loop

        End If

        On Error GoTo 0

    End If
    On Error GoTo OkErr
     
    EMsg = "01"
     
    GLPrint.RegBraCon = Equate.Regular
    If optNormal Then GLPrint.RegBraCon = Equate.Regular
    If optBranch Then GLPrint.RegBraCon = Equate.Branch
    If optConsolidated Then GLPrint.RegBraCon = Equate.Consol
    If optBudget Then GLPrint.RegBraCon = Equate.Budget
    
    EMsg = "02"
    
    GLPrint.StaSch = Equate.Stmt
    If optStatements Then GLPrint.StaSch = Equate.Stmt
    If optSchedules Then GLPrint.StaSch = Equate.Sched
    
    EMsg = "03"
    
    GLPrint.RegCmp = Equate.NonComp
    If optRegular Then GLPrint.RegCmp = Equate.NonComp
    If optComparative Then GLPrint.RegCmp = Equate.Comp
    
    EMsg = "04"
    
    GLPrint.PrintBIB = Equate.PrtBoth
    If optBoth Then GLPrint.PrintBIB = Equate.PrtBoth
    If optBalanceSheet Then GLPrint.PrintBIB = Equate.PrtBSOnly
    If optIncomeStatement Then GLPrint.PrintBIB = Equate.PrtISOnly
    
    EMsg = "05"
    
    GLPrint.FiscalYear = Me.cmbFiscalYear
    If Me.cmbStartPeriod.Visible = True Then
       GLPrint.BeginDate = StartYMs(cmbStartPeriod.ListIndex)
       GLPrint.EndDate = EndYMs(cmbEndPeriod.ListIndex)
    
       If GLPrint.EndDate < GLPrint.BeginDate Then
          MsgBox "Ending period is before the starting period !!!", vbExclamation + vbOKOnly, "Windows GL"
          Me.cmbStartPeriod.SetFocus
          Exit Sub
       End If
    
    End If
    
    EMsg = "06"
    
    GLPrint.LowAccount = Me.txtLoAccount
    GLPrint.HiAccount = Me.txtHiAccount
    
    EMsg = "07"
    
    GLPrint.LowBranchAcct = Me.txtLoBranch
    GLPrint.HiBranchAcct = Me.txtHiBranch
    
    EMsg = "08"
    
    GLPrint.LowConsAcct = Me.txtLoCons
    GLPrint.HiConsAcct = Me.txtHiCons
    
    EMsg = "09"
    
    GLPrint.PrtZeroBal = frmGLPrint2.chkPrintZeroBal
    
    GLPrint.Save (Equate.RecPut)
    
    EMsg = "10"
    
    FiscalYear = Me.cmbFiscalYear
    StartPD = cmbStartPeriod.ListIndex + 1
    EndPd = cmbEndPeriod.ListIndex + 1
    
    EMsg = "11"
    
    If cmbJournalSource.Visible = False Then
       JournalSource = 0
    ElseIf cmbJournalSource = "All Jnls" Then
       JournalSource = 0
    Else
       JournalSource = CInt(Mid(cmbJournalSource, 1, 2))
    End If
    
    Response = True
    
    On Error GoTo 0
    
    Select Case ProgName
       
       Case "STATEMENT"
          GLStatement
       
       Case "CHARTOFACCOUNTS"
          
          If Me.optNormal Then
             X = "Reg"
          Else
             X = "Cons"
          End If
          ChartOfAccts Me.txtLoAccount, _
                       Me.txtHiAccount, _
                       X, _
                       GLCompany.SubDigits
      
       Case "PRINTDESC"
          PrintDesc Me.txtLoAccount, Me.txtHiAccount
      
       Case "PRINTGLACCOUNT"
          PrintGLAccount Me.FiscalYear, _
                         Me.StartPD, _
                         Me.EndPd, _
                         Me.txtLoAccount, _
                         Me.txtHiAccount, _
                         Me.txtLoCons, _
                         Me.txtHiCons, _
                         Me.txtLoBranch, _
                         Me.txtHiBranch, _
                         GLCompany.SubDigits
       
       Case "GLHISTJNL"
          
          If chkBudget And Me.JournalSource <> 0 Then Me.JournalSource = Me.JournalSource + 100
          
          GLHistJnl Me.FiscalYear, _
                    Me.StartPD, _
                    Me.EndPd, _
                    Me.JournalSource, _
                    0, _
                    Me.chkAcctDesc
       
       Case "DETAILGL"
          
          If Me.optRegular Then X = "Reg"
          If Me.optBranch Then X = "Bra"
          If Me.optConsolidated Then X = "Cons"
          
          DetailGL X, _
                   Me.FiscalYear, _
                   Me.StartPD, _
                   Me.EndPd, _
                   Me.txtLoAccount, _
                   Me.txtHiAccount, _
                   Me.txtLoCons, _
                   Me.txtHiCons, _
                   Me.txtLoBranch, _
                   Me.txtHiBranch, _
                   GLCompany.SubDigits, _
                   GLPrint.SepPage, _
                   GLPrint.PrtZeroBal, _
                   CompanyID
       
       Case "TRIALBAL"
          GLTrialBal
    
    End Select
    
    If TextChannel <> 0 Then Close #TextChannel

    If Response Then
       Prvw.vsp.EndDoc
       Prvw.Show vbModal
    End If
    
    GoBack

OkErr:
    MsgBox "Error " & Err.Number & " " & Err.Description & vbCrLf & _
           "Form err # " & EMsg, vbExclamation + vbOKOnly, "Windows GL"
    On Error GoTo 0
    Unload Me

End Sub

Private Sub cmdOptions_Click()
   frmGLPrint2.Show vbModal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub EndPeriodSet(ByVal FY As Integer)
    
'    Dim i As Integer
'    Dim v As Variant
'
'    cmbEndPeriod.Clear
'    cmbStartPeriod.Clear
'
'    If GLCompany.FirstPeriod = 1 Then
'       v = DateSerial(FY, GLCompany.FirstPeriod, 1)
'    Else
'       v = DateSerial(FY - 1, GLCompany.FirstPeriod, 1)
'    End If
'
'    cmbEndPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
'    cmbStartPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
'    EndYMs(0) = Year(v) * 100 + Month(v)
'    StartYMs(0) = Year(v) * 100 + Month(v)
'
'    For i = 1 To 11
'        v = DateSerial(Year(v), Month(v) + 1, 1)
'        cmbEndPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
'        cmbStartPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
'        EndYMs(i) = Year(v) * 100 + Month(v)
'        StartYMs(i) = Year(v) * 100 + Month(v)
'    Next i
'
'    cmbEndPeriod.ListIndex = 0
'    cmbStartPeriod.ListIndex = 0
    
    
    
    
    Dim I As Integer
    Dim v As Variant
    
    cmbEndPeriod.Clear
    cmbStartPeriod.Clear
      
    If GLCompany.FirstPeriod = 1 Then
       v = DateSerial(FY, GLCompany.FirstPeriod, 1)
    Else
       v = DateSerial(FY - 1, GLCompany.FirstPeriod, 1)
    End If

    cmbEndPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    cmbStartPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    EndYMs(0) = Year(v) * 100 + Month(v)
    StartYMs(0) = Year(v) * 100 + Month(v)
    
    For I = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbEndPeriod.AddItem "Pd. #:" & I + 1 & " - " & Format(v, "mmmm-yyyy")
        cmbStartPeriod.AddItem "Pd. #:" & I + 1 & " - " & Format(v, "mmmm-yyyy")
        EndYMs(I) = Year(v) * 100 + Month(v)
        StartYMs(I) = Year(v) * 100 + Month(v)
    Next I
    
    cmbEndPeriod.ListIndex = 0
    cmbStartPeriod.ListIndex = 0
    
End Sub

Private Sub txtLoAccount_GotFocus()
   txtLoAccount.SelStart = 0
   txtLoAccount.SelLength = Len(txtLoAccount.Text)
End Sub
Private Sub txtHiAccount_GotFocus()
   txtHiAccount.SelStart = 0
   txtHiAccount.SelLength = Len(txtHiAccount.Text)
End Sub
Private Sub txtLoBranch_GotFocus()
   txtLoBranch.SelStart = 0
   txtLoBranch.SelLength = Len(txtLoBranch.Text)
End Sub
Private Sub txtHiBranch_GotFocus()
   txtHiBranch.SelStart = 0
   txtHiBranch.SelLength = Len(txtHiBranch.Text)
End Sub
Private Sub txtLoCons_GotFocus()
   txtLoCons.SelStart = 0
   txtLoCons.SelLength = Len(txtLoCons.Text)
End Sub
Private Sub txtHiCOns_GotFocus()
   txtHiCons.SelStart = 0
   txtHiCons.SelLength = Len(txtHiCons.Text)
End Sub


Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub DisableFrame1()

    Me.fraType1.Enabled = False
    Me.optNormal.Enabled = False
    Me.optBranch.Enabled = False
    Me.optConsolidated.Enabled = False
    Me.optBudget.Enabled = False

End Sub

Private Sub DisableFrame2()

    Me.fraType2.Enabled = False
    Me.optStatements.Enabled = False
    Me.optSchedules.Enabled = False

End Sub

Private Sub DisableFrame3()

    Me.fraType3.Enabled = False
    Me.optRegular.Enabled = False
    Me.optComparative.Enabled = False

End Sub

Private Sub DisableFrame4()

    Me.fraType4.Enabled = False
    Me.optBoth.Enabled = False
    Me.optBalanceSheet.Enabled = False
    Me.optIncomeStatement.Enabled = False

End Sub

Private Sub DisableAccountRange()

    Me.txtLoAccount.Enabled = False
    Me.txtHiAccount.Enabled = False
    Me.lblLoAccount.Enabled = False
    Me.lblHiAccount.Enabled = False
    
End Sub

Private Sub DisableBranchRange()

    Me.txtLoBranch.Enabled = False
    Me.txtHiBranch.Enabled = False
    Me.lblLoBranch.Enabled = False
    Me.lblHiBranch.Enabled = False
    
End Sub

Private Sub DisableConsRange()

    Me.txtLoCons.Enabled = False
    Me.txtHiCons.Enabled = False
    Me.lblLoCons.Enabled = False
    Me.lblHiCons.Enabled = False
    
End Sub

Private Sub DisableJS()

    Me.cmbJournalSource.Visible = False
    Me.lblJournalSource.Enabled = False
    Me.chkBudget.Visible = False

End Sub

Private Sub DisablePdRange()

    Me.cmbStartPeriod.Visible = False
    Me.cmbEndPeriod.Visible = False
    Me.lblStartPd.Enabled = False
    Me.lblEndPeriod.Enabled = False

End Sub

Private Sub DisableFY()

    Me.lblFY.Enabled = False
    Me.cmbFiscalYear.Visible = False

End Sub

Private Sub SetNudge(ByRef tdbNum As TDBNumber)
    tdbIntegerSet tdbNum
    With tdbNum
        .Spin = dbiShowAlways
        .MinValue = -255
        .MaxValue = 255
    End With
End Sub

Private Sub tdbIntegerSet(ByRef tdbAmt As TDBNumber)

    tdbAmt.Format = "##,###,##0;(##,###,##0)"
    tdbAmt.DisplayFormat = "##,###,##0;(##,###,##0);0"
    tdbAmt.HighlightText = True
    tdbAmt.Key.Clear = ""
    tdbAmt.MinValue = -99999999
    tdbAmt.MaxValue = 99999999
    tdbAmt.Value = 0

End Sub

