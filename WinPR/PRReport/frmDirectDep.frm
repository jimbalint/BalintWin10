VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmDirectDep 
   Caption         =   "Direct Deposit Report & File Create"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
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
   ScaleHeight     =   8865
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOutFile 
      Height          =   1935
      Left            =   1050
      TabIndex        =   12
      Top             =   5640
      Width           =   8415
      Begin VB.CheckBox chkBalFile 
         Caption         =   " Balanced File?"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin TDBDate6Ctl.TDBDate tdbEffDate 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   661
         Calendar        =   "frmDirectDep.frx":0000
         Caption         =   "frmDirectDep.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDirectDep.frx":0192
         Keys            =   "frmDirectDep.frx":01B0
         Spin            =   "frmDirectDep.frx":020E
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
         EditMode        =   0
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
         Text            =   "06/24/2009"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   39988
         CenturyMode     =   0
      End
      Begin TDBText6Ctl.TDBText tdbtxtFileName 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   661
         Caption         =   "frmDirectDep.frx":0236
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmDirectDep.frx":02A2
         Key             =   "frmDirectDep.frx":02C0
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
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
         Text            =   "Output File Name"
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
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   2415
      Left            =   2130
      TabIndex        =   0
      Top             =   1200
      Width           =   6375
      _cx             =   11245
      _cy             =   4260
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
   Begin VB.CheckBox chkOutputFile 
      Caption         =   " Output File?"
      Height          =   375
      Left            =   4343
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   8040
      TabIndex        =   8
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   360
      TabIndex        =   7
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3743
      TabIndex        =   9
      Top             =   3960
      Width           =   4095
      Begin VB.OptionButton optEmpName 
         Caption         =   "Employee &Name"
         Height          =   300
         Left            =   480
         TabIndex        =   2
         Top             =   620
         Width           =   2415
      End
      Begin VB.OptionButton optEmpNo 
         Caption         =   "&Employee Number"
         Height          =   300
         Left            =   480
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Select Batch(es) to report:"
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   720
      Width           =   3375
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
      Left            =   480
      TabIndex        =   10
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmDirectDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsDDBatch As New ADODB.Recordset
Public BatchHeader As String
Dim Flg As Boolean
Dim DirDepFolder, DirDepHeader As String

Private Sub Form_Load()
    
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDirDepFolder & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) Then
        DirDepFolder = PRGlobal.Var1
        DirDepHeader = PRGlobal.Var2
        BatchHeader = PRGlobal.Var3
    Else
        DirDepFolder = ""
        DirDepHeader = ""
        BatchHeader = ""
    End If
    
    Me.chkBalFile.Enabled = False
    
    Me.fraOutFile.Visible = False
    ' fields for temp record set of batches
    rsDDBatch.CursorLocation = adUseClient
    
    rsDDBatch.Fields.Append "Select", adBoolean
    rsDDBatch.Fields.Append "BatchNumber", adDouble
    rsDDBatch.Fields.Append "PEDate", adDate
    rsDDBatch.Fields.Append "CheckDate", adDate
    rsDDBatch.Fields.Append "RecordCount", adDouble
    
    rsDDBatch.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRBatch ORDER BY BatchID Desc"
    If Not PRBatch.GetBySQL(SQLString) Then
        MsgBox "No PR Batches found!", vbCritical
        End
    End If
    
    Do
    
        rsDDBatch.AddNew
        
        ' batch selected from Data Entry link
        If PRBatch.BatchID = PRBatchID Then
            rsDDBatch!Select = True
            Me.tdbEffDate = Int(PRBatch.CheckDate)
        End If
            
        rsDDBatch!BatchNumber = PRBatch.BatchID
        rsDDBatch!PEDate = Int(PRBatch.PEDate)
        rsDDBatch!CheckDate = Int(PRBatch.CheckDate)
        rsDDBatch!RecordCount = PRBatch.RecCount
        If PRBatch.BatchID = PRBatchID Then
            rsDDBatch!Select = True
        End If
        rsDDBatch.Update
        
        If Not PRBatch.GetNext Then Exit Do
    
    Loop
    
    SetGrid rsDDBatch, fg
    
    Me.lblCompanyName = PRCompany.Name
'    Me.tdbEffDate.Enabled = False
'    Me.tdbtxtFileName.Enabled = False
'    Me.chkBalFile.Enabled = False
    
    ' Me.tdbtxtFileName.Text = Mid(PRCompany.FileName, 1, Len(PRCompany.FileName) - 4) & Format(Now(), "yymmdd") & ".txt"
    
    Me.KeyPreview = True
    Me.tdbtxtFileName = ""
    Me.tdbEffDate = 0

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdDateRange_Click()
    frmDateRange.Show vbModal

    If BatchNumbr > 0 Then
        txtDisplay = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
    Else
        txtDisplay = "Date Range: " & StartDate & " - " & EndDate
    End If
    Me.Refresh
End Sub

Private Sub cmdOK_Click()
    
    ' must select a batch
    Flg = False
    rsDDBatch.MoveFirst
    Do
        If rsDDBatch!Select = True Then
            Flg = True
            Exit Do
        End If
        rsDDBatch.MoveNext
    Loop Until rsDDBatch.EOF
    rsDDBatch.MoveFirst
    If Flg = False Then
        MsgBox "You must select a batch to process!", vbExclamation, "Direct Deposit Processing"
        Exit Sub
    End If
    
    ' if creating a file
    If Me.chkOutputFile Then
        
        ' must select a file name if creating a file
        If Me.tdbtxtFileName.Text = "" Or IsNull(Me.tdbtxtFileName) Then
            MsgBox "You must select a file name!", vbExclamation, "Direct Deposit Processing"
            Exit Sub
        End If
        
        ' blank eff date?
        If Me.tdbEffDate.Value = 0 Or IsNull(Me.tdbEffDate.Value) Then
            MsgBox "You must select an effective date!", vbExclamation, "Direct Deposit Processing"
            Exit Sub
        End If
        
    End If
    
    Me.Hide
    
    ' SEH - option for NACHA header line
    SEH_Flag = True
    If PRCompany.CompanyID = 37 And InStr(LCase(PRCompany.Name), "south east") Then
        Select Case MsgBox("Include file header?" & vbCr & "Yes for OLD method" & vbCr & "No for NEW method", vbYesNoCancel, "Direct Deposit Report")
            Case vbYes
                SEH_Flag = True
            Case vbNo
                SEH_Flag = False
            Case vbCancel
                GoBack
        End Select
    End If
    
    DirectDepositRpt
    GoBack

End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub chkOutputFile_Click()
    
Dim CharPosn As Integer
Dim rw As Integer
Dim HiCheckDate As Date
    
    ' get the highest ck date selected
    HiCheckDate = 0
    rw = Me.fg.Row
    rsDDBatch.MoveFirst
    Do
        If rsDDBatch!Select = True Then
            If rsDDBatch!CheckDate > HiCheckDate Then
                HiCheckDate = rsDDBatch!CheckDate
            End If
        End If
        rsDDBatch.MoveNext
    Loop Until rsDDBatch.EOF
    fg.Row = rw
    fg.TopRow = rw

    If HiCheckDate = 0 And Me.chkOutputFile = 1 Then
        MsgBox "Select a batch to output!", vbInformation
        Me.chkOutputFile = 0
        Exit Sub
    End If

    If Me.chkOutputFile Then
        Me.fraOutFile.Visible = True
'        Me.tdbEffDate.Enabled = True
'        Me.tdbtxtFileName.Enabled = True
'        Me.chkBalFile.Enabled = True
    Else
        Me.fraOutFile.Visible = False
'        Me.tdbEffDate.Enabled = False
'        Me.tdbtxtFileName.Enabled = False
'        Me.chkBalFile.Enabled = False
    End If

    Me.chkBalFile = PRCompany.DirDepBalanced
    Me.tdbEffDate = HiCheckDate

    ' default file name
    If DirDepFolder <> "" Then
        
        CharPosn = InStr(1, DirDepFolder, "[]", vbTextCompare)
        If CharPosn <> 0 Then
            Me.tdbtxtFileName = Mid(DirDepFolder, 1, CharPosn - 1) & _
                                Format(HiCheckDate, "yymmdd") & _
                                Mid(DirDepFolder, CharPosn + 2, Len(DirDepFolder) - CharPosn + 2)
        Else
            Me.tdbtxtFileName = DirDepFolder
        End If
    
    Else
        
        Me.tdbtxtFileName = ""
    
    End If

End Sub


