VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFreeFormat 
   Caption         =   "Free Format Statements"
   ClientHeight    =   9165
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
   ScaleHeight     =   9165
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber tdbLoBranch 
      Height          =   375
      Left            =   10320
      TabIndex        =   28
      Top             =   6960
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
      Left            =   13080
      Picture         =   "frmFreeFormat.frx":041E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton cmdLoAcct 
      Height          =   375
      Left            =   13080
      Picture         =   "frmFreeFormat.frx":0728
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   495
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoAcct 
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   5040
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
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4080
      Width           =   3255
   End
   Begin VB.ComboBox cmbRC 
      Height          =   360
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3480
      Width           =   3255
   End
   Begin VB.ComboBox cmbSS 
      Height          =   360
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ComboBox cmbNBC 
      Height          =   360
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DEL"
      Height          =   615
      Left            =   1560
      TabIndex        =   15
      Top             =   8400
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlTextOutput 
      Left            =   12720
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkTextOutput 
      Caption         =   "Text File Output"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton cmdOtherOptions 
      Caption         =   "Other Options"
      Height          =   495
      Left            =   12000
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&CLEAR ALL"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "SELECT A&LL"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   8400
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   9600
      TabIndex        =   9
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   7560
      TabIndex        =   8
      Top             =   8280
      Width           =   1695
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5895
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   9855
      _cx             =   17383
      _cy             =   10398
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
      Left            =   10440
      TabIndex        =   25
      Top             =   5880
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
      Top             =   7440
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
      Top             =   6960
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
      Top             =   7440
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
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Low/Hi Branch"
      Height          =   255
      Left            =   10320
      TabIndex        =   22
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Hi Account #"
      Height          =   255
      Left            =   10440
      TabIndex        =   21
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Low Account #"
      Height          =   255
      Left            =   10440
      TabIndex        =   20
      Top             =   4680
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

Dim LoadFlag As Boolean
Dim Flag As Boolean

Dim rs As New ADODB.Recordset
Dim rsFY As New ADODB.Recordset

Dim v As Variant
Dim EndYMs(11) As Long
Dim StartYMs(11) As Long

Dim ErrMsg As String
Dim rw, LastRw As Long


Private Sub cmdDelete_Click()

    If MsgBox("OK to delete: " & rs!Description, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    ' delete PRGlobal
    If rs!GlobalID > 0 Then
        SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & rs!GlobalID
        cnDes.Execute SQLString
    End If
    
    rs.Delete
    fg.Refresh

End Sub

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
    
    ' recordset for the main grid
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Select", adBoolean
    rs.Fields.Append "Description", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "FFCol", adDouble
    rs.Fields.Append "FFSched", adDouble
    rs.Fields.Append "GlobalID", adDouble
    rs.Fields.Append "PrintID", adDouble
    rs.Fields.Append "OrderNum", adDouble
    rs.Fields.Append "OrderNum2", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic

    Init
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        SetOptions
    End If
    
    Me.KeyPreview = True

End Sub

Private Sub CmdExit_Click()
    SaveData
    GoBack
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub
Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If LoadFlag = True Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    SetOptions
End Sub
Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If LoadFlag = True Then Exit Sub
    SaveData
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
    
    
    ' on new add .....
    If rs!PrintID = 0 Then Exit Sub
    
    If rs!PrintID = 0 Then
        MsgBox "PrintID Error", vbExclamation
        GoBack
    End If
    If GLPrint.GetByID(rs!PrintID) = False Then
        MsgBox "PrintID Error", vbExclamation
        GoBack
    End If
    
    ' standard columns
    '       not using free format options
    
    Me.tdbLoBranch = GLPrint.LowBranchAcct
    Me.tdbHiBranch = GLPrint.HiBranchAcct
    Me.tdbLoCons = GLPrint.LowConsAcct
    Me.tdbHiCons = GLPrint.HiConsAcct
    
    If rs!FFCol = 0 Then
        
        rs!FFSched = 0          ' acct # range
        
        cmbPoint Me.cmbNBC, GLPrint.RegBraCon
        cmbPoint Me.cmbSS, GLPrint.StaSch
        cmbPoint Me.cmbRC, GLPrint.RegCmp
        cmbPoint Me.cmbBIB, GLPrint.PrintBIB
        Me.tdbLoAcct = nNull(GLPrint.LowAccount)
        Me.tdbHiAcct = nNull(GLPrint.HiAccount)
        
        Me.tdbLoBranch.Enabled = True
        Me.tdbHiBranch.Enabled = True
        Me.tdbLoCons.Enabled = True
        Me.tdbHiCons.Enabled = True
        
    Else        ' free format
    
        
        cmbShutOff Me.cmbRC
        cmbShutOff Me.cmbBIB
        
        If rs!FFSched <> 0 Then         ' use schedule of rows
        
            cmbShutOff Me.cmbNBC
            cmbShutOff Me.cmbSS
        
            tdbShutOff Me.tdbLoAcct
            tdbShutOff Me.tdbHiAcct
            tdbShutOff Me.tdbLoBranch
        
        Else                            ' use acct num range w/ col schedule GLFG
            
            cmbPoint Me.cmbNBC, GLPrint.RegBraCon
            cmbPoint Me.cmbSS, GLPrint.StaSch
            
            Me.tdbLoAcct = GLPrint.LowAccount
            Me.tdbHiAcct = GLPrint.HiAccount
        
        End If
    
    End If
    
    If fg.Row = 1 Then
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
    
    End If
    
End Sub

Private Sub Init()

    LoadFlag = True

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
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSetup & _
                " AND UserID = " & GLCompany.ID & _
                " ORDER BY Byte1"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            rs.AddNew
            If PRGlobal.Byte2 = 1 Then
                rs!Select = True
            Else
                rs!Select = False
            End If
            rs!Description = Trim(Mid(PRGlobal.Var4, 1, 30))
            rs!FFCol = StringValue(PRGlobal.Var2)
            rs!FFSched = StringValue(PRGlobal.Var3)
            rs!GlobalID = PRGlobal.GlobalID
            rs!PrintID = StringValue(PRGlobal.Var1)
            rs!OrderNum = PRGlobal.Byte1
            rs.Update
            
            If PRGlobal.GetNext = False Then Exit Do
        
        Loop
    
    End If
    
    SetGrid rs, fg
    
    With fg

        For I = 0 To .Cols - 1
            .ColKey(I) = .TextMatrix(0, I)
        Next I

        .ColWidth(.ColIndex("Select")) = 800
        .ColWidth(.ColIndex("Description")) = 2700
        .ColWidth(.ColIndex("FFCol")) = 2700
        .ColWidth(.ColIndex("FFSched")) = 2700
        .ColWidth(.ColIndex("OrderNum")) = 0
        .ColWidth(.ColIndex("OrderNum2")) = 0

        .ColHidden(.ColIndex("GlobalID")) = True
        .ColHidden(.ColIndex("PrintID")) = True

        .ColComboList(.ColIndex("FFCol")) = ColDrop
        .ColComboList(.ColIndex("FFSched")) = SchedDrop

        .TextMatrix(0, .ColIndex("FFCol")) = "Column Set"
        .TextMatrix(0, .ColIndex("FFSched")) = "Acct Schedule Set"

    End With

    IntSet Me.tdbLoAcct
    IntSet Me.tdbHiAcct
    IntSet Me.tdbLoBranch
    IntSet Me.tdbHiBranch
    IntSet Me.tdbLoCons
    IntSet Me.tdbHiCons
    
    frmProgress.MousePointer = vbArrow
    frmProgress.Hide

    LoadFlag = False

End Sub
Private Sub cmdAdd_Click()
    If rs.RecordCount = 255 Then
        MsgBox "Only 255 setups per client allowed!", vbExclamation
        Exit Sub
    End If
    
    LoadFlag = True
    
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSetup
    PRGlobal.UserID = GLCompany.ID
    PRGlobal.Save (Equate.RecAdd)
    
    GLPrint.Clear
    GLPrint.UseMathRec = True
    GLPrint.WidePrint = True
    GLPrint.Save (Equate.RecAdd)
    
    rs.AddNew
    rs!GlobalID = PRGlobal.GlobalID
    rs!PrintID = GLPrint.ID
    rs.Update
    
    LoadFlag = False

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
    
    If rs!PrintID = 0 Then
        MsgBox "PrintID error ...", vbExclamation
        GoBack
    End If
    
    If GLPrint.GetByID(rs!PrintID) = False Then
        MsgBox "GLPrint error", vbExclamation
        GoBack
    End If
    
    GLPrint.User = "FF" & Format(fg.Row, "000")
    
    ' store the date range in the first record
    If fg.Row = 1 Then
        GLPrint.FiscalYear = Me.cmbFiscalYear
        GLPrint.BeginDate = StartYMs(cmbStartPeriod.ListIndex)
        GLPrint.EndDate = EndYMs(cmbEndPeriod.ListIndex)
    End If

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
    
    If rs!GlobalID = 0 Then
        MsgBox "PRGlobal Error", vbExclamation
        GoBack
    End If
    If PRGlobal.GetByID(rs!GlobalID) = False Then
        MsgBox "PRGlobal Error", vbExclamation
        GoBack
    End If

    ' PRGlobal.Byte1 = fg.Row
    If rs!Select = True Then
        PRGlobal.Byte2 = 1
    Else
        PRGlobal.Byte2 = 0
    End If
    
    PRGlobal.Var1 = nNull(rs!PrintID)
    PRGlobal.Var2 = nNull(rs!FFCol)
    PRGlobal.Var3 = nNull(rs!FFSched)
    PRGlobal.Var4 = rs!Description & ""
    PRGlobal.Save (Equate.RecPut)
    
    PRGlobal.Save (Equate.RecPut)
        
End Sub
Private Sub cmdClearAll_Click()

    If rs.RecordCount = 0 Then Exit Sub
    
    LoadFlag = True
    
    rw = fg.Row
    rs.MoveFirst
    Do
        rs!Select = False
        rs.Update
        
        If PRGlobal.GetByID(rs!GlobalID) Then
            PRGlobal.Byte2 = 0
            PRGlobal.Save (Equate.RecPut)
        End If
        
        rs.MoveNext
    Loop Until rs.EOF

    rs.MoveFirst
    SetOptions
    LoadFlag = False

End Sub

Private Sub cmdSelectAll_Click()

    If rs.RecordCount = 0 Then Exit Sub
    
    LoadFlag = True
    
    rw = fg.Row
    rs.MoveFirst
    Do
        rs!Select = True
        rs.Update
        
        If PRGlobal.GetByID(rs!GlobalID) Then
            PRGlobal.Byte2 = 1
            PRGlobal.Save (Equate.RecPut)
        End If
        
        rs.MoveNext
    Loop Until rs.EOF

    rs.MoveFirst
    SetOptions
    LoadFlag = False

End Sub

Private Sub cmdPrint_Click()

Dim RecCount, PrintCount As Byte
Dim PrtFlag As Boolean

Dim FY, PD1, PD2 As Long

    If rs.RecordCount = 0 Then Exit Sub
    
    SaveData
    
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
    
    rw = fg.Row
    RecCount = 0
    PrintCount = 0
    
    PrtFlag = False
    LoadFlag = True
    
    rs.MoveFirst
    Do
        
        If GLPrint.GetByID(rs!PrintID) = False Then
            MsgBox "PrintID Error", vbExclamation
            GoBack
        End If
        
        ' get the date range from the first line
        ' even if not selected
        RecCount = RecCount + 1
        If RecCount = 1 Then
            FY = GLPrint.FiscalYear
            PD1 = GLPrint.BeginDate
            PD2 = GLPrint.EndDate
        End If
        
        If rs!Select = True Then
            
            PrintCount = PrintCount + 1
            
            ' take from the first record
            GLPrint.FiscalYear = FY
            GLPrint.BeginDate = PD1
            GLPrint.EndDate = PD2
            
            frmProgress.lblMsg1 = GLCompany.Name & vbCr & "Free Format: " & rs!Description
            frmProgress.Refresh
            
            If rs!FFCol <> 0 Then
                FreeFormatPrint rs!FFCol, rs!FFSched, PrintCount
            Else
                GLStatement PrintCount
            End If
        
            PrtFlag = True
        
        End If
        rs.MoveNext
    Loop Until rs.EOF
    
    frmProgress.Hide

    If PrtFlag = True Then
        Prvw.vsp.EndDoc
        PrvwReturn = True
        Prvw.Show vbModal
    Else
        MsgBox "No Free Format Statement has been selected!", vbInformation
    End If
    
    rs.MoveFirst
    SetOptions
    LoadFlag = False

End Sub

Private Sub cmdOtherOptions_Click()
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
