VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmFFAddAcct 
   Caption         =   "Free Format Schedules - Add Accounts"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSignReverse 
      Caption         =   "Reverse Sign of selected accounts"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   3495
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoBranch 
      Height          =   375
      Left            =   8520
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Calculator      =   "frmFFAddAcct.frx":0000
      Caption         =   "frmFFAddAcct.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFFAddAcct.frx":008A
      Keys            =   "frmFFAddAcct.frx":00A8
      Spin            =   "frmFFAddAcct.frx":00F2
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   8640
      TabIndex        =   10
      Top             =   9000
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5655
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   12975
      _cx             =   22886
      _cy             =   9975
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
   Begin VB.ComboBox cmbHiAccount 
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2640
      Width           =   3855
   End
   Begin VB.ComboBox cmbLoAccount 
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2160
      Width           =   3855
   End
   Begin VB.CommandButton cmdBranchRange 
      Caption         =   "APPLY &BRANCH RANGE"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton cmdNumRange 
      Caption         =   "APPLY ACCT &NUMBER RANGE"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "&CLEAR ALL"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "SELECT &ALL"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   10920
      TabIndex        =   0
      Top             =   9000
      Width           =   1815
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiBranch 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   2640
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Calculator      =   "frmFFAddAcct.frx":011A
      Caption         =   "frmFFAddAcct.frx":013A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFFAddAcct.frx":01A2
      Keys            =   "frmFFAddAcct.frx":01C0
      Spin            =   "frmFFAddAcct.frx":020A
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
   Begin VB.Label lblMsg1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   9000
      Width           =   7575
   End
   Begin VB.Label lblFFName 
      Alignment       =   2  'Center
      Caption         =   "FFName"
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
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   12975
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
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "frmFFAddAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x, Y, Z As String
Dim I, J, K As Long
Public OK As Boolean
Dim LoadFlag As Boolean
Dim rw As Long

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.lblFFName = frmFFSchedule.FFName
    Me.lblMsg1 = ""
    
    LoadFlag = True
    
    Init

    LoadFlag = False
    
    Me.KeyPreview = True

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    OK = False
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    OK = True
    Me.Hide
End Sub

Private Sub Init()

    ' do branch accounts apply?
    If GLCompany.SubDigits = 0 Then
        Me.cmdBranchRange.Enabled = False
        Me.tdbLoBranch.Enabled = False
        Me.tdbHiBranch.Enabled = False
    End If
    
    If GLAccount.GetAllAccounts = False Then
        MsgBox "No GL Accounts found!", vbExclamation
        GoBack
    End If
    
    With Me.fg
    
        .Cols = 4
        .Rows = 1
        
        .TextMatrix(0, 0) = "Select"
        .TextMatrix(0, 1) = "Account#"
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "Description"
        
        .FixedRows = 1
        .FixedCols = 0
                
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 1000
        .ColWidth(3) = 8000
        
        .ColDataType(0) = flexDTBoolean
        
        .ColAlignment(2) = flexAlignCenterCenter
        
        .ExplorerBar = flexExSort
        
        frmProgress.Show
        frmProgress.lblMsg1 = GLCompany.Name
        frmProgress.MousePointer = vbHourglass
        I = 0
        Do
            I = I + 1
            If I Mod 50 = 1 Then
                frmProgress.lblMsg2 = "Loading Accounts: " & GLAccount.Account
                frmProgress.Refresh
            End If
            .Rows = .Rows + 1
            .TextMatrix(I, 0) = False
            .TextMatrix(I, 1) = GLAccount.Account
            .TextMatrix(I, 2) = GLAccount.AcctType
            .TextMatrix(I, 3) = GLAccount.FullDesc
            If GLAccount.GetNext = False Then Exit Do
        Loop
        frmProgress.MousePointer = vbArrow
        frmProgress.Hide
    
        .Editable = flexEDKbdMouse
    
    End With
    
    Me.cmbLoAccount.AddItem "Select Start Acct#"
    Me.cmbLoAccount.ListIndex = 0
    Me.cmbHiAccount.AddItem "Select End Acct #"
    Me.cmbHiAccount.ListIndex = 0
    
End Sub
Private Sub cmdClearAll_Click()
    For I = 1 To fg.Rows - 1
        fg.TextMatrix(I, 0) = False
    Next I
End Sub

Private Sub cmdSelectAll_Click()
    For I = 1 To fg.Rows - 1
        fg.TextMatrix(I, 0) = True
    Next I
End Sub

Private Sub cmbLoAccount_GotFocus()

    With Me.cmbLoAccount
        If .ListCount = 1 Then
            PopCombos
        End If
    End With

End Sub
Private Sub cmbHiAccount_GotFocus()

    With Me.cmbHiAccount
        If .ListCount = 1 Then
            PopCombos
        End If
    End With

End Sub

Private Sub PopCombos()

    If LoadFlag = True Then Exit Sub
    
    I = 0
    
    Me.MousePointer = vbHourglass
    
    GLAccount.GetAllAccounts
    Do
        
        I = I + 1
        If I Mod 20 = 0 Then
            Me.lblMsg1 = "Loading Accounts: " & GLAccount.Account
            Me.Refresh
        End If
        
        With Me.cmbLoAccount
            .AddItem GLAccount.Account & " " & GLAccount.FullDesc
            .ItemData(.NewIndex) = GLAccount.Account
        End With
        
        With Me.cmbHiAccount
            .AddItem GLAccount.Account & " " & GLAccount.FullDesc
            .ItemData(.NewIndex) = GLAccount.Account
        End With

        If GLAccount.GetNext = False Then Exit Do
    
    Loop

    Me.cmbLoAccount.ListIndex = 1
    Me.cmbHiAccount.ListIndex = 1

    Me.lblMsg1 = ""
    Me.MousePointer = vbArrow

End Sub
Private Sub cmdBranchRange_Click()

    If IsNull(Me.tdbLoBranch) Then Exit Sub
    If IsNull(Me.tdbHiBranch) Then Exit Sub
    If Me.tdbLoBranch < Me.tdbHiBranch Then Exit Sub

    With fg
        rw = .Row
        For I = 1 To .Rows - 1
            K = CLng(.TextMatrix(I, 1))
            J = K Mod 10 ^ GLCompany.SubDigits
            If J >= Me.tdbLoBranch And J <= Me.tdbHiBranch Then
                .TextMatrix(I, 0) = True
            End If
        Next I
    End With

End Sub

Private Sub cmdNumRange_Click()

Dim LoAcct, HiAcct As Long

    If Me.cmbLoAccount.ListIndex <= 0 Then Exit Sub
    If Me.cmbHiAccount.ListIndex <= 0 Then Exit Sub
    If Me.cmbHiAccount.ListIndex < Me.cmbLoAccount.ListIndex Then Exit Sub
    
    LoAcct = Me.cmbLoAccount.ItemData(Me.cmbLoAccount.ListIndex)
    HiAcct = Me.cmbHiAccount.ItemData(Me.cmbHiAccount.ListIndex)
    
    With fg
        rw = .Row
        For I = 1 To .Rows - 1
            K = CLng(.TextMatrix(I, 1))
            If K >= LoAcct And K <= HiAcct Then
                .TextMatrix(I, 0) = True
            End If
        Next I
        .Row = rw
    End With

End Sub


