VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmCheckPrint 
   Caption         =   "Check Printing"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin TDBText6Ctl.TDBText tdbtextMsg 
      Height          =   350
      Left            =   2880
      TabIndex        =   17
      Top             =   7100
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   617
      Caption         =   "frmCheckPrint.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":0064
      Key             =   "frmCheckPrint.frx":0082
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
   Begin VB.Frame fraNudge 
      Caption         =   "Nudge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   4560
      TabIndex        =   14
      Top             =   5600
      Width           =   3975
      Begin TDBNumber6Ctl.TDBNumber tdbnumHorzNudge 
         Height          =   300
         Left            =   200
         TabIndex        =   15
         Top             =   240
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   529
         Calculator      =   "frmCheckPrint.frx":00C6
         Caption         =   "frmCheckPrint.frx":00E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCheckPrint.frx":016C
         Keys            =   "frmCheckPrint.frx":018A
         Spin            =   "frmCheckPrint.frx":01D4
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
      Begin TDBNumber6Ctl.TDBNumber tdbnumVertNudge 
         Height          =   300
         Left            =   195
         TabIndex        =   16
         Top             =   555
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   529
         Calculator      =   "frmCheckPrint.frx":01FC
         Caption         =   "frmCheckPrint.frx":021C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCheckPrint.frx":02A8
         Keys            =   "frmCheckPrint.frx":02C6
         Spin            =   "frmCheckPrint.frx":0310
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Check Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   920
      Left            =   1200
      TabIndex        =   11
      Top             =   5600
      Width           =   2775
      Begin VB.OptionButton ChkOptPP 
         Caption         =   "&Pre-Printed Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   13
         Top             =   550
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton ChkOptBlank 
         Caption         =   "&Blank Stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   280
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&Clear All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   9
      Top             =   1500
      Width           =   1095
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
      Height          =   375
      Left            =   3060
      TabIndex        =   8
      Top             =   1500
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3495
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   7335
      _cx             =   12938
      _cy             =   6165
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
      Height          =   350
      Left            =   2820
      TabIndex        =   3
      Top             =   6650
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   617
      Calculator      =   "frmCheckPrint.frx":0338
      Caption         =   "frmCheckPrint.frx":0358
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCheckPrint.frx":03C8
      Keys            =   "frmCheckPrint.frx":03E6
      Spin            =   "frmCheckPrint.frx":0430
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
      ValueVT         =   1245189
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
      Left            =   6960
      TabIndex        =   2
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      TabIndex        =   1
      Top             =   7800
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
      Left            =   1680
      TabIndex        =   10
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label lblBatchNumber 
      Caption         =   "Batch Number:"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   680
      Width           =   1455
   End
   Begin VB.Label lblCheckDate 
      Caption         =   "Check Date:"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   390
      Width           =   1215
   End
   Begin VB.Label lblPEDate 
      Caption         =   "Period Ending Date:"
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
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
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
Dim CheckCount As Long
Dim TotalNet As Currency


Private Sub Form_Load()
    InitForm
    Me.KeyPreview = True
End Sub


Private Sub cmdOK_Click()
    BatchID = PRBatchID
    CheckPrint
    GoBack
End Sub

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
    End Select
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub InitForm()

    Me.lblCompanyName.Caption = PRCompany.Name
    
    ' temp record set to show checks for the batch
    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "PrintCheck", adBoolean
    trs.Fields.Append "CheckNumber", adDouble
    trs.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    trs.Fields.Append "CheckAmount", adCurrency
    trs.Fields.Append "HistID", adDouble
    trs.Fields.Append "EmployeeID", adDouble
    
    trs.Open , , adOpenDynamic, adLockOptimistic
    SQLString = "SELECT * FROM PRBatch ORDER BY BatchID Desc"
    If Not PRBatch.GetByID(PRBatchID) Then
        MsgBox "Batch Not Found: " & BatchID, vbCritical
        End
    End If
    
    Me.lblBatchNumber = "Batch Number: " & PRBatch.BatchID
    Me.lblCheckDate = "Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yyyy")
    Me.lblPEDate = "PE Date: " & Format(PRBatch.PEDate, "mm/dd/yyyy")
    
    SQLString = "SELECT * FROM PRHist WHERE PRHist.BatchID = " & PRBatch.BatchID & _
                " ORDER BY CheckNumber"
    If Not PRHist.GetBySQL(SQLString) Then
        MsgBox "No History records found for Batch Number: " & PRBatch.BatchID, vbCritical
        End
    End If
    
    
    Do
        
        If Not PREmployee.GetByID(PRHist.EmployeeID) Then
            MsgBox "Employee Not Found: " & PRHist.EmployeeID, vbCritical
            End
        End If
        
        trs.AddNew
        trs!PrintCheck = True
        trs!CheckNumber = PRHist.CheckNumber
        trs!EmployeeID = PRHist.EmployeeID
        trs!EmployeeName = PREmployee.LFName
        trs!CheckAmount = PRHist.Net
        trs!HistID = PRHist.HistID
        trs.Update
        
        CheckCount = CheckCount + 1
        TotalNet = TotalNet + PRHist.Net
        
        If Not PRHist.GetNext Then Exit Do
    
    Loop

    SetGrid trs, fg
    Me.lblCount = "Checks to Print: " & Format(CheckCount, "#,##0") & " For $" & Format(TotalNet, "##,###,##0.00")
    
    Me.ChkOptBlank = False
    Me.ChkOptBlank.Enabled = False
    Me.ChkOptPP = True
    
    
    ' **** nudge setup ****
    SetNudge Me.tdbnumHorzNudge
    Me.tdbnumHorzNudge.ToolTipText = "MOVE TEXT TO THE RIGHT"
    SetNudge Me.tdbnumVertNudge
    Me.tdbnumVertNudge.ToolTipText = "MOVE TEXT DOWN"
    
    GetNudge User.ID, "OHBUC"
    Me.tdbnumHorzNudge = HorzNudge
    Me.tdbnumVertNudge = VertNudge
    
    ' **** nudge setup ****
    
    Me.tdbnumStartNumber.Format = "########0"
    Me.tdbnumStartNumber.DisplayFormat = ""
    Me.tdbnumStartNumber.MinValue = 0
    Me.tdbnumStartNumber.MaxValue = 999999999
    Me.tdbnumStartNumber = PRHist.CheckNumber
    Me.tdbtextMsg.MaxLength = 30
        
End Sub

