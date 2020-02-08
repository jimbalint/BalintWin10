VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmYearEnd 
   Caption         =   "Year End Closing"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYearEnd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4575
      Left            =   900
      TabIndex        =   9
      Top             =   2400
      Width           =   10335
      _cx             =   18230
      _cy             =   8070
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
   Begin VB.CommandButton cmdRetLook 
      Height          =   375
      Left            =   10200
      Picture         =   "frmYearEnd.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6660
      TabIndex        =   4
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4140
      TabIndex        =   3
      Top             =   7320
      Width           =   1215
   End
   Begin TDBNumber6Ctl.TDBNumber tdbRetEarn 
      Height          =   375
      Left            =   7890
      TabIndex        =   1
      Top             =   960
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      Calculator      =   "frmYearEnd.frx":0614
      Caption         =   "frmYearEnd.frx":0634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYearEnd.frx":0698
      Keys            =   "frmYearEnd.frx":06B6
      Spin            =   "frmYearEnd.frx":0700
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
      MaxValue        =   999999999
      MinValue        =   0
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
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbFiscalYear 
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   960
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      Calculator      =   "frmYearEnd.frx":0728
      Caption         =   "frmYearEnd.frx":0748
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmYearEnd.frx":07AC
      Keys            =   "frmYearEnd.frx":07CA
      Spin            =   "frmYearEnd.frx":0814
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   900
      TabIndex        =   8
      Top             =   240
      Width           =   9375
   End
   Begin VB.Label Label4 
      Caption         =   "Account Numbers to be closed into the Retained Earnings/Capital Account:"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Label Label3 
      Caption         =   "Retained Earnings/Capital Account:"
      Height          =   255
      Left            =   4170
      TabIndex        =   6
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Next Fiscal Year:"
      Height          =   255
      Left            =   780
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmYearEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FY As Integer
Dim rs As New ADODB.Recordset
Dim AcctDrop As String
Dim i As Integer

Private Sub Form_Load()
        
    Me.lblCompanyName = GLCompany.Name
    
    Me.tdbFiscalYear.MaxValue = 2099
    Me.tdbFiscalYear.MinValue = 1980
    Me.tdbFiscalYear.HighlightText = True
    Me.tdbFiscalYear.Format = "####0"
    Me.tdbFiscalYear.DisplayFormat = "####0;;Null"
    
    i = Int(GLCompany.LastClose / 10 ^ 4) + 1
    If i <= 1980 Or i > 2099 Then i = Year(Now())
    
    Me.tdbFiscalYear.Text = i
    FY = Me.tdbFiscalYear.Text
    
    With Me.tdbRetEarn
        .HighlightText = True
        .Format = ""
        .Format = "########0"
        .DisplayFormat = ""
        .MinValue = 0
        .MaxValue = 999999999
        .Text = GLCompany.RetEarnAcct
    End With
    
    ' init the ado Record Set
    rs.CursorLocation = adUseClient
    rs.Fields.Append "GLDesc", adVarChar, 255, adFldIsNullable
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    ' 100 lines on the grid
    For i = 1 To 100
        rs.AddNew
        rs.Update
    Next i
    
    ' init the flex grid
    ' use type 0 (zero) only
    GLAccount.GetAllAccounts
    Do
        If GLAccount.AcctType = "0" Then
            AcctDrop = Trim(AcctDrop) & "|#" & GLAccount.Account & ";" & _
                       GLAccount.Account & " " & GLAccount.FullDesc
        End If
        If Not GLAccount.GetNext Then Exit Do
    Loop
    
    SetGrid rs, fg
    
    fg.ColComboList(0) = AcctDrop
    fg.ColWidth(0) = 10000
    fg.TextMatrix(0, 0) = "Click on a line to get the list of accounts"
    
    ' get Account and Amount (prev year) record sets
    GLAccount.OpenRS

    ' test first p rec
    If Not (GLAccount.GetAccount(GLCompany.FirstPAcct)) Then
       MsgBox "First P record NOT FOUND !! " & GLCompany.FirstPAcct, vbCritical + vbOKOnly, "Year End"
       Exit Sub
    End If
   
    If GLAccount.AcctType <> "P" Then
       MsgBox "First P record wrong type: " & _
              GLCompany.FirstPAcct & " " & GLAccount.AcctType, vbCritical + vbOKOnly, "Year End"
       Exit Sub
    End If
   
    ' test N record
    If Not (GLAccount.GetAccount(GLCompany.NetProfitAcct)) Then
       MsgBox "N record NOT FOUND !! " & GLCompany.NetProfitAcct, vbCritical + vbOKOnly, "Year End"
       Exit Sub
    End If
   
    If GLAccount.AcctType <> "N" Then
       MsgBox "N record wrong type: " & _
              GLCompany.NetProfitAcct & " " & GLAccount.AcctType, vbCritical + vbOKOnly, "Year End"
       Exit Sub
    End If
   
    If GLCompany.FirstPAcct <= GLCompany.NetProfitAcct Then
       MsgBox "First P account: " & GLCompany.FirstPAcct & _
              "must be greater then N account#: " & GLCompany.NetProfitAcct, _
              vbCritical + vbOKOnly, "Year End"
       Exit Sub
    End If
    
    ' default the retained earnings account
    Me.tdbRetEarn = GLCompany.RetEarnAcct
    
    ' capture keyboard before form controls do
    ' used for escape
    Me.KeyPreview = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: GoBack
    End Select
    
End Sub

Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()
    
Dim i As Integer
Dim tID As Variant
    
    If CInt(Me.tdbFiscalYear.Text) <> FY Then
       i = MsgBox("Are you SURE you want to close to " & Me.tdbFiscalYear.Text & " ???", _
                vbCritical + vbOKCancel + vbDefaultButton2, "Year End")
       If i = vbCancel Then GoBack
    End If

    i = MsgBox("Are you SURE you want to close Fiscal Year " & _
               Me.tdbFiscalYear - 1 & _
               " into Fiscal Year " & Me.tdbFiscalYear & " ?", vbExclamation + vbOKCancel + vbDefaultButton2, "Year End")
    If i = vbCancel Then GoBack

    Set uDB = YearEnd(frmYearEnd.tdbFiscalYear, _
                      frmYearEnd.tdbRetEarn, _
                      rs)

    frmResults.lblCompanyName = GLCompany.Name
    frmResults.lblMsg1 = "Year End Process"
    frmResults.lblMsg2 = ""
    frmResults.lblMsg3 = ""

    For i = 1 To uDB.UpperBound(1)
        frmResults.List1.AddItem uDB(i, 0)
    Next i
    frmResults.Show vbModal

    ' call to update program
    x = "\Balint\GLUtil.exe " & _
        "SysFile=\Balint\Data\GLSystem.mdb " & _
        "UserID=" & UserID & " " & _
        "BackName=\Balint\GLMenu.exe " & _
        "ProgName=UpdateB " & _
        "Batch=" & GLBatch.BatchNumber

    If dbPwd <> "" Then
       x = x & " dbPwd=" & dbPwd
    End If

    ' 2015-03-17
    If BalintFolder <> "" Then
        x = x & " BalintFolder=" & BalintFolder
    End If

    tID = Shell(x, vbMaximizedFocus)
    Unload Me
    End
     
    GoBack

End Sub

Private Sub cmdRetLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbRetEarn = frmAcctLookup.SelAcct
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

