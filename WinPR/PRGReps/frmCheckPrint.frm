VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCheckPrint 
   Caption         =   "Check Printing"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbAmtFilter 
      Height          =   390
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   6600
      Width           =   4815
   End
   Begin VB.CheckBox chkBillStub 
      Caption         =   "Print Billing Stub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   12
      Top             =   8760
      Width           =   2295
   End
   Begin VB.CheckBox chkRevOrder 
      Caption         =   "Print in reverse order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   11
      Top             =   8280
      Width           =   3015
   End
   Begin VB.ComboBox cmbCheckStyle 
      Height          =   390
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CheckBox chkNoRate 
      Caption         =   "Don't print rate on stub"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   10
      Top             =   7800
      Width           =   3135
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumHorzNudge 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   6240
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmCheckPrint.frx":0000
      Caption         =   "frmCheckPrint.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":0096
      Keys            =   "frmCheckPrint.frx":00B4
      Spin            =   "frmCheckPrint.frx":00FE
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
   Begin VB.CheckBox chkBottomPanel 
      Caption         =   "Print bottom stub panel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8400
      TabIndex        =   9
      Top             =   7320
      Width           =   3015
   End
   Begin VB.CheckBox chkSveChk 
      Caption         =   " Save Check Numbers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   3
      Top             =   7800
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin TDBText6Ctl.TDBText tdbtextMsg 
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Top             =   7200
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   609
      Caption         =   "frmCheckPrint.frx":0126
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":018A
      Key             =   "frmCheckPrint.frx":01A8
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
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6503
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3683
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   11355
      _cx             =   20029
      _cy             =   5530
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumStartNumber 
      Height          =   345
      Left            =   1200
      TabIndex        =   14
      Top             =   5520
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   609
      Calculator      =   "frmCheckPrint.frx":01EC
      Caption         =   "frmCheckPrint.frx":020C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":027C
      Keys            =   "frmCheckPrint.frx":029A
      Spin            =   "frmCheckPrint.frx":02E4
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
      ForeColor       =   -2147483635
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   8640
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumVertNudge 
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   6720
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmCheckPrint.frx":030C
      Caption         =   "frmCheckPrint.frx":032C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":039E
      Keys            =   "frmCheckPrint.frx":03BC
      Spin            =   "frmCheckPrint.frx":0406
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
   Begin TDBNumber6Ctl.TDBNumber tdbnumEndNumber 
      Height          =   345
      Left            =   1200
      TabIndex        =   22
      Top             =   6000
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   609
      Calculator      =   "frmCheckPrint.frx":042E
      Caption         =   "frmCheckPrint.frx":044E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":04BA
      Keys            =   "frmCheckPrint.frx":04D8
      Spin            =   "frmCheckPrint.frx":0522
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
      ForeColor       =   -2147483635
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Checks to print:"
      Height          =   615
      Left            =   480
      TabIndex        =   24
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Check Style:"
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "Counts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2400
      TabIndex        =   20
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblCountBottom 
      Alignment       =   2  'Center
      Caption         =   "Counts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   19
      Top             =   5160
      Width           =   4455
   End
   Begin VB.Label lblBatchNumber 
      Caption         =   "Batch Number:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   18
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblCheckDate 
      Caption         =   "Check Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   17
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblPEDate 
      Caption         =   "Period Ending Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmCheckPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BatchID As Long
Public trs As New ADODB.Recordset
Dim StartCheckNum, CheckCount As Long
Dim TotalNet As Currency
Dim cnPRCK As ADODB.Connection
Public CheckFileName As String
Dim GlobID As Long

Dim BlankStock As Boolean

Dim x As String
Dim i, j As Long
Dim LoadFlag As Boolean

Dim PRBilling As Boolean

Public rsTS As New ADODB.Recordset

Private Sub Form_Load()
    LoadFlag = True
    InitForm
    LoadFlag = False
    Me.KeyPreview = True
End Sub

Private Sub cmdOK_Click()

    HorzNudge = Me.tdbnumHorzNudge
    VertNudge = Me.tdbnumVertNudge

    With Me.cmbCheckStyle
        If .ItemData(.ListIndex) = PREquate.CheckTypeBlankStock Then
            SaveNudge User.ID, "CHECKPRINTBLANK"
        Else
            SaveNudge User.ID, "CHECKPRINT"
        End If
    End With
    
    ' save check printing options
    If PRGlobal.GetByID(GlobID) Then
        
        PRGlobal.Var1 = Me.chkBottomPanel
        
        With Me.cmbCheckStyle
            PRGlobal.Var2 = .ItemData(.ListIndex)
        End With
        
        PRGlobal.Var3 = Me.chkNoRate
        PRGlobal.Var4 = Me.chkRevOrder
        
        If Me.chkBillStub.Visible = True And Me.chkBillStub = 1 Then
            PRGlobal.Byte1 = 1
        Else
            PRGlobal.Byte1 = 0
        End If
        
        PRGlobal.Save (Equate.RecPut)
    
    End If
    
    BatchID = PRBatchID
    
    ' create the record set for billing?
    If Me.chkBillStub = 1 Then GetTimeSheet
    
    With Me.cmbCheckStyle
        CheckPrint .ItemData(.ListIndex)
    End With
    
    GoBack

End Sub

Private Sub GetTimeSheet()
        
Dim tFlag As Boolean
        
    On Error Resume Next
    rsTS.Close
    On Error GoTo 0
    rsTS.CursorLocation = adUseClient
    rsTS.Fields.Append "EmployeeID", adDouble
    rsTS.Fields.Append "WEDate", adDate
    rsTS.Fields.Append "JobID", adVarChar, 30, adFldIsNullable
    rsTS.Fields.Append "ItemID", adDouble
    rsTS.Fields.Append "Hours", adCurrency
    rsTS.Fields.Append "Rate", adCurrency
    rsTS.Fields.Append "Amount", adCurrency
    rsTS.Open , , adOpenDynamic, adLockOptimistic
        
    SQLString = "SELECT * FROM PRDist WHERE BatchID = " & PRBatch.BatchID
    If PRDist.GetBySQL(SQLString) = False Then Exit Sub
        
    Do
    
        If PRDist.Hours = 0 Then GoTo NxtTimeSheet
        If PRDist.JobID = 0 Then GoTo NxtTimeSheet
        If PRDist.BillingRate = 0 Then GoTo NxtTimeSheet
            
        tFlag = False
        
'       03/11/2011 - create a line on the billing stub for each PRDist entry
'        If rsTS.RecordCount > 0 Then
'            rsTS.MoveFirst
'            Do
'                If PRDist.EmployeeID = rsTS!EmployeeID _
'                   And PRDist.JobID = rsTS!JobID _
'                   And DistItem() = rsTS!ItemID Then
'
'                    tFlag = True
'                    Exit Do
'                End If
'                rsTS.MoveNext
'            Loop Until rsTS.EOF
'        End If
        
        If tFlag = False Then
            rsTS.AddNew
            rsTS!EmployeeID = PRDist.EmployeeID
            rsTS!JobID = PRDist.JobID
            rsTS!ItemID = DistItem()
            rsTS!Hours = 0
            rsTS!Rate = 0
            rsTS!Amount = 0
            rsTS.Update
        End If
        rsTS!Hours = rsTS!Hours + PRDist.Hours
        rsTS!Amount = rsTS!Amount + PRDist.Amount
        rsTS!Rate = PRDist.Rate
        rsTS.Update
    
NxtTimeSheet:
        If PRDist.GetNext = False Then Exit Do
    Loop

    If rsTS.RecordCount > 0 Then
        rsTS.Sort = "EmployeeID, JobID, WEDate, ItemID"
    End If

End Sub

Private Function DistItem() As Long
    If PRDist.ItemType = PREquate.ItemTypeRegPay Then
        DistItem = 99991
    ElseIf PRDist.ItemType = PREquate.ItemTypeOvtPay Then
        DistItem = 99992
    Else
        DistItem = PRDist.EmployerItemID
    End If
End Function

Private Sub cmdSelectAll_Click()
    trs.MoveFirst
    Do
        trs!PrintCheck = True
        trs.Update
        trs.MoveNext
        If trs.EOF Then Exit Do
    Loop
    trs.MoveFirst
End Sub
Private Sub cmdClearAll_Click()
    trs.MoveFirst
    Do
        trs!PrintCheck = False
        trs.Update
        trs.MoveNext
        If trs.EOF Then Exit Do
    Loop
    trs.MoveFirst
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF7: Renumber
    End Select
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub InitForm()
Dim StartChk As Long

    ' does a blank stock check setup exist?
    BlankStock = False
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePRCheckPrefix
    If PRGlobal.GetBySQL(SQLString) Then
        
        
        If BalintFolder = "" Then
            CheckFileName = "\Balint\Data\PRCK" & _
                            Trim(PRGlobal.Description) & _
                            Format(PRCompany.CompanyID, "000000") & ".MDB"
        Else
            CheckFileName = Replace(BalintFolder, "^", " ") & "\Data\PRCK" & _
                            Trim(PRGlobal.Description) & _
                            Format(PRCompany.CompanyID, "000000") & ".MDB"
        End If
        
        On Error Resume Next
        GetAttr (CheckFileName)
        If Err.Number = 0 Then
            BlankStock = True
        End If
        On Error GoTo 0
    End If
    
    ' populate the drop down box
    For i = 1 To 4
        
        ' blank stock option available
        If i = 1 And BlankStock = False Then GoTo NextI
        
        Select Case i
            Case 1
                x = "Blank Stock"
                j = PREquate.CheckTypeBlankStock
            Case 2
                x = "Pre Printed A"
                j = PREquate.CheckTypePrePrintedA
            Case 3
                x = "Pre Printed B"
                j = PREquate.CheckTypePrePrintedB
            Case 4
                x = "Pre Printed C"
                j = PREquate.CheckTypePrePrintedC
        End Select
        
        With Me.cmbCheckStyle
            .AddItem x
            .ItemData(.NewIndex) = j
        End With

NextI:
    Next i
    
    ' Print Options
    ' always default blank form if file exists
    SQLString = "SELECT * FROM PRGlobal WHERE Description = 'CheckPrintB' AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        
        ' print bottom panel ?
        Me.chkBottomPanel = PRGlobal.Var1 & ""
        
        ' no rate print option if pre-preprinted
        If PRGlobal.Var3 = "1" And BlankStock = False Then
            If PRGlobal.Var3 = "1" Then
                Me.chkNoRate = 1
            Else
                Me.chkNoRate = 0
            End If
        End If
        
        ' set the drop down combo to the proper check style
        Me.cmbCheckStyle.ListIndex = 0      ' dflt to blank stock or first pre-printed option
        If BlankStock = False Then
            With Me.cmbCheckStyle
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = PRGlobal.Var2 Then
                        .ListIndex = i
                        Exit For
                    End If
                Next i
            End With
        End If
    
        ' print in reverse order ?
        If PRGlobal.Var4 = "1" Then
            Me.chkRevOrder = 1
        Else
            Me.chkRevOrder = 0
        End If
    
    Else
        
        PRGlobal.Clear
        PRGlobal.Description = "CheckPrintB"
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Var1 = "0"
        PRGlobal.Var2 = "0"
        PRGlobal.Var2 = "3"
        PRGlobal.Var4 = "0"
        PRGlobal.Byte1 = 0
        PRGlobal.Save (Equate.RecAdd)
    
        Me.cmbCheckStyle.ListIndex = 0
    
    End If
    GlobID = PRGlobal.GlobalID

    ' show billing stub option
    Me.chkBillStub.Visible = False
    If PRGlobal.Byte1 = 1 Then
        Me.chkBillStub.Visible = True
        Me.chkBillStub = 1
    Else        ' use bill rate option set in time sheet entry
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeCompanyOption & _
                    " AND Description = 'PayrollBilling'" & _
                    " AND Var1 = 'Yes'" & _
                    " AND Var2 = '" & PRCompany.GLCompanyID & "'"
        If PRGlobal.GetBySQL(SQLString) = True Then
            Me.chkBillStub.Visible = True
            Me.chkBillStub = 1
        End If
    End If

    Me.lblCompanyName.Caption = PRCompany.Name
    
    ' temp record set to show checks for the batch
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "PrintCheck", adBoolean
    trs.Fields.Append "CheckNumber", adDouble
    trs.Fields.Append "OrigCheckNumber", adDouble
    trs.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    trs.Fields.Append "CheckAmount", adCurrency
    trs.Fields.Append "HistID", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    
    trs.Open , , adOpenDynamic, adLockOptimistic
    If Not PRBatch.GetByID(PRBatchID) Then
        MsgBox "Batch Not Found: " & BatchID, vbCritical
        End
    End If
    
    Me.lblBatchNumber = "Batch Number: " & PRBatch.BatchID
    Me.lblCheckDate = "Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy")
    Me.lblPEDate = "PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
    
    SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & PRBatch.BatchID & _
                " ORDER BY CheckNumber"
    
'    ' *** SEH ***
'    SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & PRBatch.BatchID & _
'                " ORDER BY HistID"
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History records found for Batch Number: " & PRBatch.BatchID, vbCritical
        End
    End If
    
    StartCheckNum = 0
    
'    Me.tdbnumStartNumber.Format = "########0"
'    Me.tdbnumStartNumber.DisplayFormat = ""
'    Me.tdbnumStartNumber.MinValue = 0
'    Me.tdbnumStartNumber.MaxValue = 999999999
'    Me.tdbtextMsg.MaxLength = 30
'    Me.tdbnumStartNumber = PRHist.CheckNumber
'    Me.TDBNumEndNumber = StartCheckNum + CheckCount - 1
    
    Do
        
        If Not PREmployee.GetByID(PRHist.EmployeeID) Then
            MsgBox "Employee Not Found: " & PRHist.EmployeeID, vbCritical
            End
        End If
        
        trs.AddNew
        trs!PrintCheck = True
        trs!CheckNumber = PRHist.CheckNumber
        trs!OrigCheckNumber = PRHist.CheckNumber
        trs!EmployeeID = PRHist.EmployeeID
        trs!EmployeeName = PREmployee.LFName
        trs!CheckAmount = PRHist.Net
        trs!HistID = PRHist.HistID
        trs.Update
        
        CheckCount = CheckCount + 1
        TotalNet = TotalNet + PRHist.Net
        
        If StartCheckNum = 0 And PRHist.CheckNumber <> 0 Then
            StartCheckNum = PRHist.CheckNumber
        End If
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    SetGrid trs, fg
    Me.lblCount = "Checks to Print: " & Format(CheckCount, "#,##0") & " For $" & Format(TotalNet, "##,###,##0.00")
    
    ' **** nudge setup ****
    SetNudge Me.tdbnumHorzNudge
    Me.tdbnumHorzNudge.ToolTipText = "MOVE TEXT TO THE RIGHT"
    SetNudge Me.tdbnumVertNudge
    Me.tdbnumVertNudge.ToolTipText = "MOVE TEXT DOWN"
    ChkNudgeSet
    ' **** nudge setup ****
    
    Me.tdbnumStartNumber.Format = "########0"
    Me.tdbnumStartNumber.DisplayFormat = ""
    Me.tdbnumStartNumber.MinValue = 0
    Me.tdbnumStartNumber.MaxValue = 999999999
    Me.tdbnumStartNumber = StartCheckNum
    
    Me.tdbnumEndNumber.Format = "########0"
    Me.tdbnumEndNumber.DisplayFormat = ""
    Me.tdbnumEndNumber.MinValue = 0
    Me.tdbnumEndNumber.MaxValue = 999999999
    Me.tdbnumEndNumber = StartCheckNum + CheckCount - 1
    Me.tdbnumEndNumber.Enabled = False
    
    Me.tdbtextMsg.MaxLength = 30
    Me.lblCountBottom = "Number of Checks to Print: " & CheckCount
    
    With Me.cmbAmtFilter
        .AddItem "ALL CHECKS"
        .AddItem "NON ZERO CHECK AMOUNT"
        .AddItem "ZERO CHECK AMOUNT(DIR DEP)"
        .ListIndex = 0
    End With
    
    ' get the check prefix
    ' disable selection for Scott Molders
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePRCheckPrefix
    If PRGlobal.GetBySQL(SQLString) Then
        x = Trim(PRGlobal.Description)
        If LCase(x) = "sct" Then
            Me.cmbCheckStyle.Enabled = False
        End If
    End If
    
End Sub
    
Private Sub ChkNudgeSet()
    With Me.cmbCheckStyle
        If .ItemData(.ListIndex) = PREquate.CheckTypeBlankStock Then
            GetNudge User.ID, "CHECKPRINTBLANK"
            Me.tdbnumHorzNudge = HorzNudge
            Me.tdbnumVertNudge = VertNudge
        Else
            GetNudge User.ID, "CHECKPRINT"
            Me.tdbnumHorzNudge = HorzNudge
            Me.tdbnumVertNudge = VertNudge
        End If
    End With
End Sub

Private Sub ChkOptBlank_Click()
    If LoadFlag = True Then Exit Sub
    ChkNudgeSet
End Sub

Private Sub ChkOptPP_Click()
    ChkNudgeSet
End Sub


Private Sub tdbnumStartNumber_LostFocus()
    
Dim Rw As Long
    
    Me.tdbnumEndNumber = Me.tdbnumStartNumber + CheckCount - 1
    CheckNum = Me.tdbnumStartNumber
    Rw = fg.Row
    ScreenUpdate
    fg.Row = Rw
    fg.Select Rw, 0
    fg.ShowCell Rw, 0

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Renumber()
    
Dim ChkNum As Long
    
    ' renumber the checks by order entered
    On Error Resume Next
    trs.Close
    On Error GoTo 0
    SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & PRBatch.BatchID & _
                " ORDER BY HistID"
    
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History records found for Batch Number: " & PRBatch.BatchID, vbCritical
        End
    End If
    
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "PrintCheck", adBoolean
    trs.Fields.Append "CheckNumber", adDouble
    trs.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    trs.Fields.Append "CheckAmount", adCurrency
    trs.Fields.Append "HistID", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    ChkNum = Me.tdbnumStartNumber
    StartCheckNum = Me.tdbnumStartNumber
    CheckCount = 0
    TotalNet = 0
    
    Do
        
        If Not PREmployee.GetByID(PRHist.EmployeeID) Then
            MsgBox "Employee Not Found: " & PRHist.EmployeeID, vbCritical
            End
        End If
        
        trs.AddNew
        trs!PrintCheck = True
        trs!CheckNumber = ChkNum
        trs!EmployeeID = PRHist.EmployeeID
        trs!EmployeeName = PREmployee.LFName
        trs!CheckAmount = PRHist.Net
        trs!HistID = PRHist.HistID
        trs.Update
        
        ChkNum = ChkNum + 1
        CheckCount = CheckCount + 1
        TotalNet = TotalNet + PRHist.Net
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    SetGrid trs, fg
    Me.lblCount = "Checks to Print: " & Format(CheckCount, "#,##0") & " For $" & Format(TotalNet, "##,###,##0.00")

End Sub

Private Sub ScreenUpdate()

Dim CNum As Long

    CNum = Me.tdbnumStartNumber
    CheckCount = 0
    TotalNet = 0
    trs.MoveFirst
    Do
        If trs!PrintCheck = True Then
            CheckCount = CheckCount + 1
            TotalNet = TotalNet + trs!CheckAmount
            trs!CheckNumber = CNum
            CNum = CNum + 1
        Else
            trs!CheckNumber = trs!OrigCheckNumber
        End If
        trs.Update
        trs.MoveNext
    Loop Until trs.EOF
    
    Me.lblCount = "Checks to Print: " & Format(CheckCount, "#,##0") & " For $" & Format(TotalNet, "##,###,##0.00")
    Me.lblCountBottom = "Number of Checks to Print: " & CheckCount
    Me.tdbnumEndNumber = Me.tdbnumStartNumber + CheckCount - 1
    Me.Refresh

End Sub
Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
Dim TopRow As Long
    
    TopRow = fg.TopRow
    ScreenUpdate
    fg.Row = Row
    fg.Select Row, 0
    fg.TopRow = TopRow
    
End Sub

