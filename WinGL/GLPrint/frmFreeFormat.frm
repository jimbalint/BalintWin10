VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFreeFormat 
   Caption         =   "Free Format Statements"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFreeFormat.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbLoBranch 
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      Top             =   6120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":030A
      Caption         =   "frmFreeFormat.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":038E
      Keys            =   "frmFreeFormat.frx":03AC
      Spin            =   "frmFreeFormat.frx":03F6
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
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   13200
      Picture         =   "frmFreeFormat.frx":041E
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5160
      Width           =   495
   End
   Begin VB.CommandButton cmdLoAcct 
      Height          =   375
      Left            =   13080
      Picture         =   "frmFreeFormat.frx":0728
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4320
      Width           =   495
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoAcct 
      Height          =   375
      Left            =   10560
      TabIndex        =   24
      Top             =   4320
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":0A32
      Caption         =   "frmFreeFormat.frx":0A52
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":0AB6
      Keys            =   "frmFreeFormat.frx":0AD4
      Spin            =   "frmFreeFormat.frx":0B1E
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
   Begin VB.ComboBox cmbBIB 
      Height          =   360
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3480
      Width           =   3255
   End
   Begin VB.ComboBox cmbRC 
      Height          =   360
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3000
      Width           =   3255
   End
   Begin VB.ComboBox cmbSS 
      Height          =   360
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2520
      Width           =   3255
   End
   Begin VB.ComboBox cmbNBC 
      Height          =   360
      Left            =   10560
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2040
      Width           =   3255
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DEL"
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   7440
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlTextOutput 
      Left            =   10800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkTextOutput 
      Caption         =   "Text File Output"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton cmdOtherOptions 
      Caption         =   "Other Options"
      Height          =   495
      Left            =   12000
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&CLEAR ALL"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "SELECT A&LL"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   9600
      TabIndex        =   9
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4935
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   9855
      _cx             =   17383
      _cy             =   8705
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
   Begin VB.ComboBox cmbEndPeriod 
      Height          =   360
      Left            =   7080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   3495
   End
   Begin VB.ComboBox cmbStartPeriod 
      Height          =   360
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
   End
   Begin VB.ComboBox cmbFiscalYear 
      Height          =   360
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiAcct 
      Height          =   375
      Left            =   10560
      TabIndex        =   25
      Top             =   5160
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":0B46
      Caption         =   "frmFreeFormat.frx":0B66
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":0BCA
      Keys            =   "frmFreeFormat.frx":0BE8
      Spin            =   "frmFreeFormat.frx":0C32
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
   Begin TDBNumber6Ctl.TDBNumber tdbHiBranch 
      Height          =   375
      Left            =   10320
      TabIndex        =   29
      Top             =   6600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":0C5A
      Caption         =   "frmFreeFormat.frx":0C7A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":0CDE
      Keys            =   "frmFreeFormat.frx":0CFC
      Spin            =   "frmFreeFormat.frx":0D46
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
   Begin TDBNumber6Ctl.TDBNumber tdbLoCons 
      Height          =   375
      Left            =   12360
      TabIndex        =   30
      Top             =   6120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":0D6E
      Caption         =   "frmFreeFormat.frx":0D8E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":0DF2
      Keys            =   "frmFreeFormat.frx":0E10
      Spin            =   "frmFreeFormat.frx":0E5A
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
   Begin TDBNumber6Ctl.TDBNumber tdbHiCons 
      Height          =   375
      Left            =   12360
      TabIndex        =   31
      Top             =   6600
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmFreeFormat.frx":0E82
      Caption         =   "frmFreeFormat.frx":0EA2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFreeFormat.frx":0F06
      Keys            =   "frmFreeFormat.frx":0F24
      Spin            =   "frmFreeFormat.frx":0F6E
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
   Begin VB.Label Label7 
      Caption         =   "Low/Hi Consol"
      Height          =   255
      Left            =   12360
      TabIndex        =   23
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Low/Hi Branch"
      Height          =   255
      Left            =   10320
      TabIndex        =   22
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Hi Account #"
      Height          =   255
      Left            =   10560
      TabIndex        =   21
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Low Account #"
      Height          =   255
      Left            =   10560
      TabIndex        =   20
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "End Period:"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Start Period:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fiscal Year:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13455
   End
End
Attribute VB_Name = "frmFreeFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X, Y, Z As String
Dim I, J, K As Long
Dim boo As Boolean

Dim SchedDrop As String
Dim ColDrop As String

Dim ColNum As Long
Dim LoadFlag As Boolean
Dim Flag As Boolean

Dim rsCol As New ADODB.Recordset
Dim rsFY As New ADODB.Recordset

Dim v As Variant
Dim EndYMs(11) As Long
Dim StartYMs(11) As Long

Dim ErrMsg As String
Dim rw, LastRw As Long
Dim FFPrintID As Long

Private Sub Form_Load()

    ' ---------------------------------
    ' -- PRGlobal storage of FF setup
    '
    ' GLCompanyID       UserID
    ' PrintID           Var1        GLPrint Rec ID#
    ' Order #           Byte1
    '
    ' Select flag       Byte2
    ' FFCol             Var2
    ' FFSched           Var3
    ' Description       Var4
    '
    ' all other print parameters stored in GLPrint
    '
    ' GLPrint.User = FF### - ### = Order #
    '
    ' ----------------------------------
    
    LoadFlag = True
    
    ' *****************************************************************************************
    ' jb 10/27/10 - clear out bogus PRGlobal Records
    Dim rsPRG As New ADODB.Recordset
    rsPRG.CursorLocation = adUseClient
    rsPRG.Fields.Append "PRGlobalID", adDouble
    rsPRG.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSetup & _
                " AND UserID = " & GLCompany.ID & _
                " ORDER BY Byte1"
    If PRGlobal.GetBySQL(SQLString) = True Then
        If GLPrint.GetByID(PRGlobal.Var1) = False Then
            rsPRG.AddNew
            rsPRG!PRGlobalID = PRGlobal.GlobalID
            rsPRG.Update
        End If
        If rsPRG.RecordCount > 0 Then
            rsPRG.MoveFirst
            Do
                SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & rsPRG!PRGlobalID
                cnDes.Execute SQLString
                rsPRG.MoveNext
            Loop Until rsPRG.EOF
        End If
    End If
    ' *****************************************************************************************
    
    ' recordset for the main grid
    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    rsCol.Open , , adOpenDynamic, adLockOptimistic
    
    AddCol "", "Col0", 300
    AddCol "Select", "", 800, flexDTBoolean
    AddCol "Description", "Desc", 2600
    AddCol "FFCol", "", 2600
    AddCol "FFSched", "", 2600
    AddCol "GlobalID", "", 0
    AddCol "PrintID", "", 0
    
    ' setup the grid
    With fg
        
        ColNum = 0
                
        .Rows = 1
        .Cols = rsCol.RecordCount
        
        .FixedRows = 1
        .FixedCols = 1
        
        .ExplorerBar = flexExMoveRows
        .AllowBigSelection = False
        .Editable = flexEDKbdMouse
            
        I = 0
        rsCol.MoveFirst
        Do
            .TextMatrix(0, I) = rsCol!Title
            .ColWidth(I) = rsCol!Width
            .ColData(I) = rsCol!Abbrev
            If rsCol!DataType <> 0 Then
                .ColDataType(I) = rsCol!DataType
            End If
            If rsCol!Format <> 0 Then
                .ColFormat(I) = rsCol!Format
            End If
            I = I + 1
            rsCol.MoveNext
        Loop Until rsCol.EOF
    
        .AllowSelection = False
        .AllowBigSelection = False
    
    End With
    
    Init
    
'''''    If rs.RecordCount > 0 Then
'''''        rs.MoveFirst
'''''        SetOptions
'''''    End If
    
    With fg
        If .Rows > 1 Then
            ' .TextMatrix(1, GetCol("Col0")) = "*"
            .Row = 1
            SetOptions
        End If
    End With
    
    LoadFlag = False
    
    Me.KeyPreview = True

End Sub

Private Sub CmdExit_Click()
    
    SaveData
    SaveDateRanges
    GoBack

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub
Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If LoadFlag = True Then Exit Sub
    
    With fg
        If .Rows = 1 Then Exit Sub
'        If OldRow <> 0 Then
'            .TextMatrix(OldRow, GetCol("Col0")) = " "
'        End If
'        If NewRow <> 0 Then
'            .TextMatrix(NewRow, GetCol("Col0")) = "*"
'        End If
    End With
    
    SetOptions

End Sub
Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If LoadFlag = True Then Exit Sub
    SaveData
End Sub
Private Sub fg_BeforeMoveRow(ByVal Row As Long, Position As Long)
    SaveData
    LoadFlag = True
End Sub

Private Sub fg_AfterMoveRow(ByVal Row As Long, Position As Long)
    SetSeq
End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    ' must be acct order if standard columns
    With fg
        If Col = GetCol("FFSched") And GridValue(.TextMatrix(Row, GetCol("FFCol"))) = 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With fg
        If Col = GetCol("FFCol") And GridValue(.TextMatrix(Row, GetCol("FFCol"))) = 0 Then
            .TextMatrix(Row, GetCol("FFSched")) = "0"
        End If
    End With
    SetOptions
End Sub

Private Sub SetOptions()
        
'    boo = False
'    If IsNull(rs!PrintID) Then
'        boo = True
'    ElseIf rs!PrintID = 0 Then
'        boo = True
'    End If
'    If boo = False Then
'        If GLPrint.GetByID(rs!PrintID) = False Then boo = True
'    End If
'    If boo Then
'        GLPrint.Clear
'        GLPrint.Save (Equate.RecAdd)
'        rs!PrintID = GLPrint.ID
'        rs.Update
'    End If
    
    
    With fg
    
        ' on new add .....
        If GridValue(TxtMatrix("GlobalID")) = "0" Then Exit Sub
    
        If TxtMatrix("PrintID") = "0" Then
            MsgBox "PrintID Error", vbExclamation
            GoBack
        End If
        
        If GLPrint.GetByID(GridValue(TxtMatrix("PrintID"))) = False Then
                        
'            GLPrint.Clear
'            GLPrint.ID = GridValue(TxtMatrix("PrintID"))
'            GLPrint.User = "Default"
'            GLPrint.Save (Equate.RecAdd)
            
            MsgBox "PrintID Error " & TxtMatrix("PrintID"), vbExclamation
            GoBack
        
        End If
    
        If .Rows = 1 Then Exit Sub
        
        ' standard columns
        '       not using free format options
        Me.tdbLoBranch = GLPrint.LowBranchAcct
        Me.tdbHiBranch = GLPrint.HiBranchAcct
        Me.tdbLoCons = GLPrint.LowConsAcct
        Me.tdbHiCons = GLPrint.HiConsAcct
    
        If TxtMatrix("FFCol") = "0" Then
            ' standard columns
            .TextMatrix(.Row, GetCol("FFSched")) = "0"
        
            cmbPoint Me.cmbNBC, GLPrint.RegBraCon
            cmbPoint Me.cmbSS, GLPrint.StaSch
            cmbPoint Me.cmbRC, GLPrint.RegCmp
            cmbPoint Me.cmbBIB, GLPrint.PrintBIB
            
            Me.tdbLoAcct.Enabled = True
            Me.tdbLoAcct = nNull(GLPrint.LowAccount)
            Me.tdbHiAcct.Enabled = True
            Me.tdbHiAcct = nNull(GLPrint.HiAccount)
        
            Me.tdbLoBranch.Enabled = True
            Me.tdbLoBranch = GLPrint.LowBranchAcct
            Me.tdbHiBranch.Enabled = True
            Me.tdbHiBranch = GLPrint.HiBranchAcct
            Me.tdbLoCons.Enabled = True
            Me.tdbLoCons = GLPrint.LowConsAcct
            Me.tdbHiCons.Enabled = True
            Me.tdbHiCons = GLPrint.HiConsAcct
        
        Else        ' free format
    
            cmbShutOff Me.cmbRC
            cmbShutOff Me.cmbBIB
        
            If GridValue(TxtMatrix("FFSched")) <> 0 Then         ' use schedule of rows
        
                cmbShutOff Me.cmbNBC
                cmbShutOff Me.cmbSS
                tdbShutOff Me.tdbLoAcct
                tdbShutOff Me.tdbHiAcct
                tdbShutOff Me.tdbLoBranch
                tdbShutOff Me.tdbHiBranch
                tdbShutOff Me.tdbLoCons
                tdbShutOff Me.tdbHiCons
            
            Else                            ' use acct num range w/ col schedule GLFG
            
                cmbPoint Me.cmbNBC, GLPrint.RegBraCon
                cmbPoint Me.cmbSS, GLPrint.StaSch
            
                Me.tdbLoAcct = GLPrint.LowAccount
                Me.tdbLoAcct.Enabled = True
                Me.tdbHiAcct = GLPrint.HiAccount
                Me.tdbHiAcct.Enabled = True
        
                tdbShutOff Me.tdbHiBranch
                Me.tdbLoBranch.Enabled = True
                Me.tdbLoBranch = GLPrint.LowBranchAcct
                Me.tdbLoCons.Enabled = True
                Me.tdbLoCons = GLPrint.LowConsAcct
                Me.tdbHiCons.Enabled = True
                Me.tdbHiCons = GLPrint.HiConsAcct
            
                With Me.cmbNBC
                    If .ItemData(.ListIndex) = Equate.Consol Then
                        Me.tdbHiBranch.Enabled = True
                        Me.tdbHiBranch = GLPrint.LowBranchAcct
                    End If
                End With
            
            End If
    
        End If
    
    End With

End Sub

Private Sub Init()

    GLPrint.OpenRS
    
    Me.lblCompanyName = GLCompany.Name

    frmProgress.MousePointer = vbHourglass
    frmProgress.lblMsg1 = GLCompany.Name & " Initializing ...."
    frmProgress.Show

    ' ----------- drop down inits -------------------------------

    ' FF Column dropdown
    ColDrop = "|#0;Standard"
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFColumn & _
                " ORDER BY Description"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            ColDrop = ColDrop & "|#" & PRGlobal.GlobalID & ";" & PRGlobal.Description
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If
    
    ' FF Schedule dropdowns
    SchedDrop = "|#0;Account Order"
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSched & _
                " AND UserID = " & GLCompany.ID & _
                " ORDER BY Description"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            SchedDrop = SchedDrop & "|#" & PRGlobal.GlobalID & ";" & PRGlobal.Description
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If

    ' GL Account drop
    
    I = 0

    frmProgress.lblMsg2 = "Now gathering History data ...."
    frmProgress.Refresh
    
    ' -------- fiscal year and period drop down ----------------
    
    SQLString = "SELECT DISTINCT FiscalYear from GLAmount ORDER BY FiscalYear DESC"
    rsInit SQLString, cn, rsFY
    If rsFY.RecordCount = 0 Then
        MsgBox "No amount data found!", vbExclamation
        GoBack
    End If
    rsFY.MoveFirst
    With Me.cmbFiscalYear
        Do
            .AddItem rsFY!FiscalYear
            rsFY.MoveNext
        Loop Until rsFY.EOF
        .ListIndex = 0
    End With
    
    EndPeriodSet CInt(Me.cmbFiscalYear)
    
    With Me.cmbNBC
        
        .AddItem "N/A"
        .ItemData(.NewIndex) = 0
        
        .AddItem "Normal"
        .ItemData(.NewIndex) = Equate.Regular
        
        .AddItem "Branch"
        .ItemData(.NewIndex) = Equate.Branch
        
        .AddItem "Consolidated"
        .ItemData(.NewIndex) = Equate.Consol
        
        .ListIndex = 0
    
    End With
    
    With Me.cmbSS
        
        .AddItem "N/A"
        .ItemData(.NewIndex) = 0
        
        .AddItem "Statements"
        .ItemData(.NewIndex) = Equate.Stmt
        
        .AddItem "Schedules"
        .ItemData(.NewIndex) = Equate.Sched
    
        .ListIndex = 0
    
    End With
    
    With Me.cmbRC
        
        .AddItem "N/A"
        .ItemData(.NewIndex) = 0
        
        .AddItem "Regular"
        .ItemData(.NewIndex) = Equate.NonComp
        
        .AddItem "Comparative"
        .ItemData(.NewIndex) = Equate.Comp
    
        .ListIndex = 0
    
    End With
    
    With Me.cmbBIB
        
        .AddItem "N/A"
        .ItemData(.NewIndex) = 0
        
        .AddItem "Print B/S & I/S"
        .ItemData(.NewIndex) = Equate.PrtBoth
        
        .AddItem "Print B/S Only"
        .ItemData(.NewIndex) = Equate.PrtBSOnly
        
        .AddItem "Print I/S Only"
        .ItemData(.NewIndex) = Equate.PrtISOnly
        
        .ListIndex = 0
    
    End With
    
    ' ---------- main grid setup ------------------------
    ' PRGlobal to store setups per client
    ' String1 thru String10
    ' Select / ColID / SchedID / LoAcct / HiAcct / Branch
    I = 1
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSetup & _
                " AND UserID = " & GLCompany.ID & _
                " ORDER BY Byte1"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            I = I + 1
            With fg
                .Rows = I
                .TextMatrix(.Rows - 1, GetCol("Select")) = PRGlobal.Byte2
                .TextMatrix(.Rows - 1, GetCol("Desc")) = PRGlobal.Var4
                .TextMatrix(.Rows - 1, GetCol("FFCol")) = PRGlobal.Var2
                .TextMatrix(.Rows - 1, GetCol("FFSched")) = PRGlobal.Var3
                .TextMatrix(.Rows - 1, GetCol("GlobalID")) = PRGlobal.GlobalID
                .TextMatrix(.Rows - 1, GetCol("PrintID")) = PRGlobal.Var1
            End With
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If
                
    With fg

        .ColAlignment(GetCol("Col0")) = flexAlignCenterCenter
        .ColAlignment(GetCol("Desc")) = flexAlignLeftCenter

        .ColComboList(GetCol("FFCol")) = ColDrop
        .ColComboList(GetCol("FFSched")) = SchedDrop

        .TextMatrix(0, GetCol("FFCol")) = "Column Set"
        .TextMatrix(0, GetCol("FFSched")) = "Acct Schedule Set"

    End With

    IntSet Me.tdbLoAcct
    IntSet Me.tdbHiAcct
    IntSet Me.tdbLoBranch
    IntSet Me.tdbHiBranch
    IntSet Me.tdbLoCons
    IntSet Me.tdbHiCons
    
    ' date range
    SQLString = "SELECT * FROM GLPrint WHERE ReportName = 'FreeForm'"
    If GLPrint.GetBySQL(SQLString) = True Then
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
        FFPrintID = GLPrint.ID
    Else
        Me.cmbFiscalYear.ListIndex = 0
        Me.cmbStartPeriod.ListIndex = 0
        Me.cmbEndPeriod.ListIndex = 0
        FFPrintID = 0
    End If
    
    frmProgress.MousePointer = vbArrow
    frmProgress.Hide

    LoadFlag = False

End Sub
Private Sub cmdAdd_Click()
    
Dim CurrRow, OldMaxRows As Long
Dim OldRow, NewRow As Long
    
    If fg.Rows > 1 Then SaveData
    LoadFlag = True
    
    With fg
    
        If .Rows = 255 Then
            MsgBox "Only 255 setups per client allowed!", vbExclamation
            LoadFlag = False
            Exit Sub
        End If
        
        
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSetup
        PRGlobal.UserID = GLCompany.ID
        
        GLPrint.OpenRS
        GLPrint.Clear
        GLPrint.RegBraCon = Equate.Regular
        GLPrint.StaSch = Equate.Stmt
        GLPrint.RegCmp = Equate.NonComp
        GLPrint.PrintBIB = Equate.PrtBoth
        GLPrint.LowAccount = 1
        GLPrint.HiAccount = 999999999
        GLPrint.UseMathRec = True
        GLPrint.WidePrint = True
        GLPrint.Save (Equate.RecAdd)
        
        PRGlobal.Var1 = GLPrint.ID
        
        PRGlobal.Save (Equate.RecAdd)
        
        ' insert row at current position
        SQLString = "" & vbTab & "0" & vbTab & _
                    "" & vbTab & _
                    "0" & vbTab & _
                    "0" & vbTab & _
                    PRGlobal.GlobalID & vbTab & _
                    GLPrint.ID
        
        If fg.Row <> 0 Then
            .AddItem SQLString, fg.Row + 1
            fg.Row = fg.Row + 1
        Else
            .AddItem SQLString, 1
            fg.Row = 1
        End If
            
        SetSeq      ' update the seq numbering
    
        If fg.Row = 0 Then fg.Row = 1
    
    End With
        
End Sub

Private Sub EndPeriodSet(ByVal FY As Integer)
    
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

Private Sub cmbFiscalYear_Click()
    EndPeriodSet (CInt(cmbFiscalYear))
End Sub

Private Sub SaveData()
    
    If LoadFlag = True Then Exit Sub
    
    If fg.Row = 0 Then Exit Sub
    
    If TxtMatrix("PrintID") = "0" Then
        MsgBox "PrintID error ...", vbExclamation
        GoBack
    End If
    
    If GLPrint.GetByID(GridValue(TxtMatrix("PrintID"))) = False Then
        MsgBox "GLPrint error " & GridValue(TxtMatrix("PrintID")), vbExclamation
        GoBack
    End If
    
    GLPrint.User = "FF" & Format(fg.Row, "000")
    GLPrint.RegBraCon = cmbValue(Me.cmbNBC)
    GLPrint.StaSch = cmbValue(Me.cmbSS)
    GLPrint.RegCmp = cmbValue(Me.cmbRC)
    GLPrint.PrintBIB = cmbValue(Me.cmbBIB)
    GLPrint.LowAccount = nNull(Me.tdbLoAcct.Value)
    GLPrint.HiAccount = nNull(Me.tdbHiAcct.Value)
    GLPrint.LowBranchAcct = nNull(Me.tdbLoBranch)
    GLPrint.HiBranchAcct = nNull(Me.tdbHiBranch)
    GLPrint.LowConsAcct = nNull(Me.tdbLoCons)
    GLPrint.HiConsAcct = nNull(Me.tdbHiCons)
    GLPrint.Save (Equate.RecPut)
    
    If TxtMatrix("GlobalID") = "0" Then
        MsgBox "PRGlobal Error", vbExclamation
        GoBack
    End If
    If PRGlobal.GetByID(GridValue(TxtMatrix("GlobalID"))) = False Then
        MsgBox "PRGlobal Error", vbExclamation
        GoBack
    End If

    With fg
        If GridValue(.TextMatrix(.Row, GetCol("Select"))) <> 0 Then
            PRGlobal.Byte2 = 1
        Else
            PRGlobal.Byte2 = 0
        End If
        PRGlobal.Var1 = .TextMatrix(.Row, GetCol("PrintID"))
        PRGlobal.Var2 = .TextMatrix(.Row, GetCol("FFCol"))
        PRGlobal.Var3 = .TextMatrix(.Row, GetCol("FFSched"))
        PRGlobal.Var4 = .TextMatrix(.Row, GetCol("Desc"))
    End With
        
    PRGlobal.Save (Equate.RecPut)
        
End Sub
Private Sub cmdClearAll_Click()
    SetSelect "0"
End Sub

Private Sub cmdSelectAll_Click()
    SetSelect "1"
End Sub

Private Sub SetSelect(ByVal str As String)
    
    SaveData
    LoadFlag = True
    
    With fg
        
        If .Rows = 1 Then
            LoadFlag = False
            Exit Sub
        End If
        
        rw = fg.Row
        If rw = 0 Then rw = 1
        
        I = 1
        Do
            
            .TextMatrix(I, GetCol("Select")) = str
            
            fg.Row = I
            If PRGlobal.GetByID(GridValue(TxtMatrix("GlobalID"))) = False Then
                MsgBox "GlobalID error " & TxtMatrix("GlobalID"), vbExclamation
                GoBack
            End If
            
            PRGlobal.Byte2 = str
            PRGlobal.Save (Equate.RecPut)
            
            I = I + 1
            If I > .Rows - 1 Then Exit Do
        Loop
    
        .Row = rw
    
    End With
    
    SetOptions
    LoadFlag = False

End Sub


Private Sub cmdPrint_Click()

Dim RecCount, PrintCount As Byte
Dim PrtFlag As Boolean

Dim FY, PD1, PD2 As Long

    If fg.Rows = 1 Then Exit Sub
    
    SaveData
    
    ' save to GLPrint w/ user = "FreeForm"
    ' set up variables for statement printing
    SaveDateRanges
    FY = GLPrint.FiscalYear
    PD1 = GLPrint.BeginDate
    PD2 = GLPrint.EndDate
    
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
    
    frmProgress.Show
    frmProgress.MousePointer = vbHourglass
    frmProgress.lblMsg1 = "Free Format: " & GLCompany.Name
    frmProgress.Refresh
    
    RecCount = 0
    PrintCount = 0
    
    PrtFlag = False
    LoadFlag = True
    
    I = 1
    Do
        With fg
        
            fg.Row = I
            SetOptions
            
            If GLPrint.GetByID(GridValue(TxtMatrix("PrintID"))) = False Then
                MsgBox "PrintID Error", vbExclamation
                GoBack
            End If
    
            If GridValue(TxtMatrix("Select")) <> 0 Then
                
                PrintCount = PrintCount + 1
                
                ' ------------ verify ranges -----------------------------------------------------------
                If GLPrint.LowAccount > GLPrint.HiAccount Then
                    MsgBox TxtMatrix("Desc") & vbCr & "Low Account # greater than Hi Account #" & vbCr & _
                           GLPrint.LowAccount & " " & GLPrint.HiAccount, vbExclamation
                    GoTo NxtI
                End If
                
                If GLPrint.LowBranchAcct > GLPrint.HiBranchAcct And GLPrint.HiBranchAcct <> 0 Then
                    MsgBox TxtMatrix("Desc") & vbCr & "Low Branch # greater than Hi Branch #" & vbCr & _
                           GLPrint.LowBranchAcct & " " & GLPrint.HiBranchAcct, vbExclamation
                    GoTo NxtI
                End If
                                
                If GLPrint.LowConsAcct > GLPrint.HiConsAcct Then
                    MsgBox TxtMatrix("Desc") & vbCr & "Low Cons # greater than Hi Cons #" & vbCr & _
                           GLPrint.LowConsAcct & " " & GLPrint.HiConsAcct, vbExclamation
                    GoTo NxtI
                End If
                
                ' ---------------------------------------------------------------------------------------
                
                ' take from "FreeForm" GLPrint record
                GLPrint.FiscalYear = FY
                GLPrint.BeginDate = PD1
                GLPrint.EndDate = PD2
                
                frmProgress.lblMsg1 = GLCompany.Name & vbCr & "Free Format: " & TxtMatrix("Desc")
                frmProgress.Refresh
                
                If GridValue(TxtMatrix("FFCol")) <> 0 Then
                    FreeFormatPrint GridValue(TxtMatrix("FFCol")), GridValue(TxtMatrix("FFSched")), PrintCount
                Else
                    GLStatement PrintCount
                End If
            
                PrtFlag = True
            
            End If
    
NxtI:
            I = I + 1
            If I > .Rows - 1 Then Exit Do
        
        End With
        
    Loop
    
    frmProgress.MousePointer = vbArrow
    frmProgress.Hide

    If PrtFlag = True Then
        Prvw.vsp.EndDoc
        PrvwReturn = True
        Prvw.Show vbModal
    Else
        MsgBox "No Free Format Statement has been selected!", vbInformation
    End If
    
    fg.Row = 1
    SetOptions
    LoadFlag = False

End Sub

Private Sub cmdOtherOptions_Click()
    If fg.Rows = 1 Or fg.Row = 0 Then Exit Sub
    frmGLPrint2.Show vbModal
    GLPrint.Save (Equate.RecPut)
End Sub

Private Sub cmbShutOff(ByRef cmb As ComboBox)

    With cmb
        If .ListCount > 0 Then .ListIndex = 0
        .Enabled = False
    End With

End Sub

Private Function cmbValue(ByRef cmb As ComboBox) As Long

    With cmb
        If IsNull(.ListIndex) Or .ListIndex = -1 Or .ListCount = 0 Then
            cmbValue = 0
        Else
            cmbValue = .ItemData(.ListIndex)
        End If
    End With

End Function

Private Function StringValue(ByVal str As String) As Long

    StringValue = 0
    If IsNull(str) Then Exit Function
    If str = "" Then Exit Function
    StringValue = CLng(str)

End Function
Private Sub cmdLoAcct_Click()
    frmAcctLookup.Show vbModal
    Me.tdbLoAcct = frmAcctLookup.SelAcct
    Me.tdbHiAcct.SetFocus
End Sub
Private Sub cmdHiAcct_Click()
    frmAcctLookup.Show vbModal
    Me.tdbHiAcct = frmAcctLookup.SelAcct
End Sub

Private Sub IntSet(ByRef tdb As TDBNumber)
    
    With tdb
        .Format = "########0"
        .DisplayFormat = .Format
        .HighlightText = True
        .MinValue = 0
        .MaxValue = 999999999
    End With

End Sub

Private Sub tdbShutOff(ByRef tdb As TDBNumber)

    With tdb
        .Value = 0
        .Enabled = False
    End With

End Sub
Private Sub cmdDelete_Click()

    LoadFlag = True
    
    With fg
    
        If .Rows = 1 Or .Row = 0 Then
            LoadFlag = False
            Exit Sub
        End If
        
        If MsgBox("OK to delete: " & TxtMatrix("Desc"), vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        
        If PRGlobal.GetByID(GridValue(TxtMatrix("GlobalID"))) = False Then
            MsgBox "PRGlobal error ..", vbExclamation
            GoBack
        End If
        
        SQLString = "DELETE * FROM GLPrint WHERE ID = " & TxtMatrix("PrintID")
        cn.Execute SQLString
        
        SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & TxtMatrix("GlobalID")
        cnDes.Execute SQLString
        
        ' update the grid
        .RemoveItem .Row
    
        SetSeq
    
    End With
    
End Sub

Private Sub AddCol(ByVal Title As String, _
                   ByVal Abbrev As String, _
                   ByVal Width As Long, _
                   Optional DType As Byte, _
                   Optional fmt As String)

    If Abbrev = "" Then Abbrev = Title
    
    rsCol.AddNew
    rsCol!Title = Mid(Title, 1, 30)
    rsCol!Abbrev = Mid(Abbrev, 1, 30)
    rsCol!Width = Width
    rsCol!Number = ColNum
    rsCol!DataType = DType
    rsCol!Format = fmt
    rsCol.Update
    
    ColNum = ColNum + 1

End Sub

Private Function GetCol(ByVal ColData As String) As Long

    SQLString = "Abbrev = '" & ColData & "'"
    rsCol.Find SQLString, 0, adSearchForward, 1
    If rsCol.EOF Then
        GetCol = -1
    Else
        GetCol = rsCol!Number
    End If

End Function

Private Function TxtMatrix(ByVal ColID As String) As String

    With fg
        TxtMatrix = .TextMatrix(.Row, GetCol(ColID))
    End With

End Function
Private Function GridValue(ByVal str As String) As Long

    GridValue = 0
    If IsNull(str) Then Exit Function
    If str = "" Then Exit Function
    If str = "0" Then Exit Function
    If IsNumeric(str) = False Then Exit Function
    GridValue = CLng(str)

End Function

Private Sub SetSeq()

    LoadFlag = True
    SaveData
    
    With fg
    
        If .Rows = 1 Then
            LoadFlag = False
            Exit Sub
        End If
        
        rw = .Row
        
        For I = 1 To .Rows - 1
                        
            .Row = I
            
            If PRGlobal.GetByID(GridValue(TxtMatrix("GlobalID"))) = False Then
                MsgBox "PRGlobal error", vbExclamation
                GoBack
            End If

            PRGlobal.Byte1 = I
            PRGlobal.Save (Equate.RecPut)
        
'            If GLPrint.GetByID(GridValue(TxtMatrix("PrintID"))) = False Then
'                MsgBox "GLPrint error", vbExclamation
'                GoBack
'            End If
'
'            GLPrint.User = "FF" & Format(I, "000")
'            GLPrint.Save (Equate.RecPut)
            
        Next I
        
        .Row = rw
        
    End With
        
    SetOptions
    LoadFlag = False

End Sub

Private Sub SaveDateRanges()

    If FFPrintID = 0 Then
        GLPrint.Clear
        GLPrint.ReportName = "FreeForm"
        GLPrint.Save (Equate.RecAdd)
        FFPrintID = GLPrint.ID
    End If
    
    If GLPrint.GetByID(FFPrintID) = False Then
        MsgBox "Date save error: " & FFPrintID, vbExclamation
        GoBack
    End If
    
    GLPrint.FiscalYear = Me.cmbFiscalYear
    GLPrint.BeginDate = StartYMs(cmbStartPeriod.ListIndex)
    GLPrint.EndDate = EndYMs(cmbEndPeriod.ListIndex)
    GLPrint.Save (Equate.RecPut)

End Sub
Private Sub cmbStartPeriod_Click()
    Me.cmbEndPeriod.ListIndex = Me.cmbStartPeriod.ListIndex
End Sub


