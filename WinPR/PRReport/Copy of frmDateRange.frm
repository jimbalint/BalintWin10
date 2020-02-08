VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDateRange 
   Caption         =   "Standard Date Ranges"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEndCheckPE 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
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
      Left            =   2160
      TabIndex        =   20
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox txtStartCheckPE 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
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
      Left            =   2160
      TabIndex        =   19
      Top             =   5520
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   " Use:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   910
      Left            =   5640
      TabIndex        =   16
      Top             =   4920
      Width           =   1935
      Begin VB.OptionButton optPEDate 
         Caption         =   " P/E Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optCheckDate 
         Caption         =   " Check Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   1695
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNoofMo 
      Height          =   360
      Left            =   3480
      TabIndex        =   3
      Top             =   4440
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   635
      Calculator      =   "frmDateRange.frx":0000
      Caption         =   "frmDateRange.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDateRange.frx":008C
      Keys            =   "frmDateRange.frx":00AA
      Spin            =   "frmDateRange.frx":00F4
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
      EditMode        =   1
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
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox TxtEndMon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   13
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   5003
      TabIndex        =   8
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   1163
      TabIndex        =   7
      Top             =   7440
      Width           =   2055
   End
   Begin TDBDate6Ctl.TDBDate TDBStartPEDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5880
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calendar        =   "frmDateRange.frx":011C
      Caption         =   "frmDateRange.frx":0234
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDateRange.frx":02A0
      Keys            =   "frmDateRange.frx":02BE
      Spin            =   "frmDateRange.frx":031C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   2.02345000269425E-316
      CenturyMode     =   0
   End
   Begin VB.ComboBox cmbStartMon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   2655
      Left            =   2280
      TabIndex        =   9
      Top             =   1080
      Width           =   5295
      _cx             =   9340
      _cy             =   4683
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
      ScrollBars      =   2
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
   Begin VB.OptionButton OptChkPeDate 
      Caption         =   " Check/P/E      Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton optMonths 
      Caption         =   " Months"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   1320
   End
   Begin VB.OptionButton optBatch 
      Caption         =   " Batch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate TDBEndPEDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   6720
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calendar        =   "frmDateRange.frx":0344
      Caption         =   "frmDateRange.frx":045C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDateRange.frx":04C8
      Keys            =   "frmDateRange.frx":04E6
      Spin            =   "frmDateRange.frx":0544
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "01/05/2009"
      ValidateMode    =   0
      ValueVT         =   2010382343
      Value           =   39818
      CenturyMode     =   0
   End
   Begin VB.Line Line9 
      X1              =   5520
      X2              =   5520
      Y1              =   3840
      Y2              =   7200
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2055
      Y1              =   7200
      Y2              =   7215
   End
   Begin VB.Line Line7 
      X1              =   360
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      X1              =   7800
      X2              =   7800
      Y1              =   960
      Y2              =   7200
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   360
      Y1              =   960
      Y2              =   7200
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   7800
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line3 
      X1              =   360
      X2              =   7800
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   5520
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblProgram 
      Alignment       =   2  'Center
      Caption         =   "Program Name"
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
      Height          =   255
      Left            =   203
      TabIndex        =   15
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label lblClient 
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
      Height          =   255
      Left            =   203
      TabIndex        =   14
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label3 
      Caption         =   "End Month"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "# of Mths"
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
      Left            =   2160
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Start Month"
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumMon As Long
Public AlphMon As String
Public EndAlphMon
Public rs As ADODB.Recordset
Public SYRMO As String
Public EYRMO As String
Dim YMStartDate As String
Dim YMEndDate As String
Dim EYMYear As Long
Dim EYMMonth As String
Dim AlphStartMon As String
Dim AlphEndMon As String
Dim StartYear As Long
Dim EndYear As Long
Public NumStartMo As Long
Public NumEndMo As Long
Dim GFocus As Boolean
Dim FirstBatch As Long
Dim FirstPEDate As Long
Dim FirstCheckDt As Long

Sub Form_Load()

Dim YrFour As Long
Dim LastYM As String
Dim CurrNumMon As Long
Dim NextNumMon As Long
Dim YM, LowYM, HiYM As Long

    Me.KeyPreview = True
    GFocus = False
    Startdate = 0
    EndDate = 0
    
    Me.lblClient.Caption = PRCompany.Name
    Me.optCheckDate = 1
    Me.optBatch = 1
    SQLString = "SELECT BatchID, PEDate, CheckDate,RecCount, USERID, yearmonth " & _
                " FROM PRBATCH ORDER BY YEARMONTH DESC, PEDATE DESC"
                
    rsInit SQLString, cn, rs
    
    If rs.RecordCount = 0 Then
        MsgBox "No Batch files found!", vbCritical
        End
    End If
    
    ' loop thru the temp rs to get the lowest and highest YM
    rs.MoveFirst
    Do
        If LowYM = 0 Or rs!YearMonth < LowYM Then LowYM = rs!YearMonth
        If rs!YearMonth > HiYM Then HiYM = rs!YearMonth
        rs.MoveNext
    Loop Until rs.EOF
    
    rs.MoveFirst
    
    SetGrid rs, fg
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    fg.ColFormat(1) = "mm/dd/yyyy"
    fg.ColWidth(1) = 1200
    fg.AllowSelection = False
    
    For YM = HiYM To LowYM Step -1
        If YM Mod 100 = 0 Then YM = YM - 88
        x = GetAlphMon(YM Mod 100)
        cmbStartMon.AddItem Int(YM / 100) & " " & x
    Next YM
    
    Me.TDBStartPEDate = Now()
    Me.TDBEndPEDate = Now()

    BatchNumbr = 0
    cmbStartMon.ListIndex = 0

    rs.MoveFirst
    fg.Row = 1

End Sub
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Public Sub cmbStartMon_Click()
    If TDBNoofMo > 0 Then
        AlphStartMon = Mid(cmbStartMon, 6, 3)
        StartYear = Mid(cmbStartMon, 1, 4)
        GetNumMon (AlphStartMon)
        ConvNumMon = ConvNumMon + TDBNoofMo
        If ConvNumMon = 13 Then
            ConvNumMon = 1
            GetAlphMon (ConvNumMon - 1)
            If TDBNoofMo <> 0 Then
                TxtEndMon = StartYear + 1 & " " & AlphMon
            End If
        Else
            ConvNumMon = ConvNumMon
            GetAlphMon (ConvNumMon - 1)
            If TDBNoofMo <> 0 Then
                TxtEndMon = StartYear & " " & AlphMon
            End If
        End If
    End If
    
End Sub

Private Sub cmbStartMon_LostFocus()

    AlphStartMon = Mid(cmbStartMon, 6, 3)
    If Me.cmbStartMon.Text = "" Or IsNull(Me.cmbStartMon) Then
        MsgBox "Start Date was not selected !!!", vbCritical, "Date Range"
    Else
        StartYear = Mid(cmbStartMon, 1, 4)
    End If

    GetNumMon (AlphStartMon)
    ConvNumMon = ConvNumMon + TDBNoofMo
    If ConvNumMon = 13 Then
        ConvNumMon = 1
        GetAlphMon (ConvNumMon - 1)
        If TDBNoofMo <> 0 Then
            TxtEndMon = StartYear + 1 & " " & AlphMon
        End If
    Else
        ConvNumMon = ConvNumMon
        GetAlphMon (ConvNumMon - 1)
        If TDBNoofMo <> 0 Then
            TxtEndMon = StartYear & " " & AlphMon
        End If
    End If
    
End Sub


Public Sub cmdOK_Click()

Dim MM As Byte
Dim SDtChk, EDtChk

    InitFlag = True
    If optCheckDate = True Then
        OptDate = "CHECK DATE"
    Else
        OptDate = "P/E DATE"
    End If
    If RangeType = PREquate.RangeTypeBatch Then
        BatchNumbr = rs!BatchID
        PEDate = Format(rs!PEDate, "mm/dd/yyyy")
        CheckDt = Format(rs!CheckDate, "mm/dd/yyyy")
        TxtDisplay = "Batch " & BatchNumbr
    ElseIf RangeType = PREquate.RangeTypeMonths Then
        MM = GetNumMon(Mid(Me.cmbStartMon.Text, 6, 3))
        If Me.cmbStartMon.Text = "" Or IsNull(Me.cmbStartMon) Then
            MsgBox "Start Date was not selected !!!", vbCritical, "Date Range"
            Exit Sub
        End If
        If Me.TxtEndMon.Text = "" Or IsNull(Me.TxtEndMon) Then
            MsgBox "End Date was not selected !!!", vbCritical, "Date Range"
            Exit Sub
        End If
        Startdate = DateSerial(Mid(Me.cmbStartMon.Text, 1, 4), MM, 1)
        EndDate = DateSerial(Year(Startdate), Month(Startdate) + Me.TDBNoofMo, 1) - 1
    ElseIf RangeType = PREquate.RangeTypePEDate Then
        If Me.TDBStartPEDate.Text = "" Or IsNull(Me.TDBStartPEDate) Then
            MsgBox "Start Date was not selected !!!", vbCritical, "Date Range"
            Exit Sub
        End If
        If Me.TDBEndPEDate.Text = "" Or IsNull(Me.TDBEndPEDate) Then
            MsgBox "End Date was not selected !!!", vbCritical, "Date Range"
            Exit Sub
        End If
        Startdate = TDBStartPEDate
        EndDate = TDBEndPEDate
        TxtDisplay = "Date Range: " & Startdate & " To: " & EndDate & " Using " & OptDate
    End If
    If Me.optCheckDate = 0 And Me.optPEDate = 0 Then
        MsgBox "Please select Check Date or P/E Date", vbCritical, "Date Range"
        Exit Sub
    End If

    ' Unload frmDateRange
    Me.Hide
    
'    Set rs = Nothing

End Sub

Private Sub fg_GotFocus()
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    fg.ColFormat(1) = "mm/dd/yyyy"
    fg.ColWidth(1) = 1200
    fg.AllowSelection = False

End Sub


Public Sub optbatch_Click()
    If optBatch = True Then
        fg.Enabled = True
        cmbStartMon.Enabled = False
        TxtEndMon.Enabled = False
        TDBNoofMo.Enabled = False
        TDBStartPEDate.Enabled = False
        TDBEndPEDate.Enabled = False
        optCheckDate.Enabled = False
        optPEDate.Enabled = False
        RangeType = PREquate.RangeTypeBatch

    End If
    
End Sub

Private Sub optCheckDate_Click()
    
    If OptChkPeDate = True Then
        txtStartCheckPE = "Start Check Date"
        txtEndCheckPE = "End Check Date"
    End If
End Sub

Private Sub optPEDate_Click()
    If OptChkPeDate = True Then
        txtStartCheckPE = "Start P/E Date"
        txtEndCheckPE = "End P/E Date"
    Else
        txtStartCheckPE = ""
        txtEndCheckPE = ""
    End If
End Sub

Private Sub OptChkPeDate_Click()
    If OptChkPeDate = True Then
        TDBStartPEDate.Enabled = True
        TDBEndPEDate.Enabled = True
        cmbStartMon.Enabled = False
        TxtEndMon.Enabled = False
        TDBNoofMo.Enabled = False
        fg.Enabled = False
        optPEDate.Enabled = True
        optCheckDate.Enabled = True
        RangeType = PREquate.RangeTypePEDate
        If optCheckDate Then
            txtStartCheckPE = "Start Check Date"
            txtEndCheckPE = "End Check Date"
        Else
            txtStartCheckPE = "Start P/E Date"
            txtEndCheckPE = "End Check Date"
        End If
        BatchNumbr = 0
    End If
        
End Sub

Public Sub OptMonths_Click()
    If optMonths = True Then
        cmbStartMon.Enabled = True
        TxtEndMon.Enabled = True
        TDBNoofMo.Enabled = True
        TDBStartPEDate.Enabled = False
        TDBEndPEDate.Enabled = False
        fg.Enabled = False
        optPEDate.Enabled = True
        optCheckDate.Enabled = True
        RangeType = PREquate.RangeTypeMonths
        txtStartCheckPE = " "
        txtEndCheckPE = " "
        BatchNumbr = 0
    End If
End Sub

Private Function GetAlphMon(ByVal NumMon As Integer)
    AlphMon = ""
    If NumMon = 1 Then GetAlphMon = "JAN"
    If NumMon = 2 Then GetAlphMon = "FEB"
    If NumMon = 3 Then GetAlphMon = "MAR"
    If NumMon = 4 Then GetAlphMon = "APR"
    If NumMon = 5 Then GetAlphMon = "MAY"
    If NumMon = 6 Then GetAlphMon = "JUN"
    If NumMon = 7 Then GetAlphMon = "JUL"
    If NumMon = 8 Then GetAlphMon = "AUG"
    If NumMon = 9 Then GetAlphMon = "SEP"
    If NumMon = 10 Then GetAlphMon = "OCT"
    If NumMon = 11 Then GetAlphMon = "NOV"
    If NumMon = 12 Then GetAlphMon = "DEC"
End Function



Private Sub TDBNoofMo_Change()

Dim Dt As Date
Dim Mth As Byte

    AlphStartMon = Mid(cmbStartMon, 6, 3)
    StartYear = Mid(cmbStartMon, 1, 4)
    Mth = GetNumMon(AlphStartMon)
    Dt = DateSerial(StartYear, Mth + Me.TDBNoofMo - 1, 1)
    
    Me.TxtEndMon = Year(Dt) & " " & GetAlphMon(Month(Dt))
    
'    NumStartMo = ConvNumMon
'    NumEndMo = NumStartMo + TDBNoofMo
'    ConvNumMon = NumEndMo
'    GetAlphMon (NumEndMo - 1)
'    AlphEndMon = AlphMon
'    If NumEndMo = 13 Then
'        NumEndMo = 1
'        AlphEndMon = "JAN"
'        If TDBNoofMo <> 0 Then
'            TxtEndMon = StartYear + 1 & " " & AlphEndMon
'        End If
'    Else
'        YMStartDate = NumStartMo & "/01/" & StartYear
'        Startdate = CDate(YMStartDate)
'        YMEndDate = NumEndMo & "/01/" & StartYear
'        EndDate = CDate(YMEndDate)
'        If TDBNoofMo <> 0 Then
'            TxtEndMon = StartYear & " " & AlphMon
'        End If
'    End If

End Sub

Private Sub TDBNoofMo_LostFocus()

'    AlphStartMon = Mid(cmbStartMon, 6, 3)
'    StartYear = Mid(cmbStartMon, 1, 4)
'    GetNumMon (AlphStartMon)
'    NumStartMo = ConvNumMon
'    NumEndMo = NumStartMo + TDBNoofMo
'    GetAlphMon (NumEndMo)
'    AlphEndMon = AlphMon
'    If NumEndMo = 13 Then
'        NumEndMo = 1
'        AlphEndMon = "JAN"
'        TxtEndMon = StartYear + 1 & " " & AlphEndMon
'    Else
'        YMStartDate = NumStartMo & "/01/" & StartYear
'        Startdate = CDate(YMStartDate)
'        YMEndDate = NumEndMo & "/01/" & StartYear
'        EndDate = CDate(YMEndDate)
'        If TDBNoofMo <> 0 Then
'            TxtEndMon = StartYear & " " & AlphMon
'        End If
'    End If
    
End Sub


Public Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
End Sub

