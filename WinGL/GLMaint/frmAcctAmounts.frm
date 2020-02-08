VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAcctAmounts 
   Caption         =   "DISPLAY AMOUNT/BUDGET VALUES"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
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
   ScaleHeight     =   8940
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&NEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&PREVIOUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   8280
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4935
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   7695
      _cx             =   13573
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
   Begin VB.ComboBox cmbFiscalYear 
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.ComboBox cmbAcct 
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Fiscal Year:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Account:"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmAcctAmounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DisplayAccount As Long
Dim rs As New ADODB.Recordset
Dim rsGL As New ADODB.Recordset
Dim LoadFlag As Boolean
Dim I, J, K As Long
Dim x, Y, Z As String

Private Sub Form_Load()

    LoadFlag = True
    
    Me.lblCompanyName = GLCompany.Name
    
    ' init the combo's if necessary
    If Me.cmbFiscalYear.ListCount = 0 Then Init

    ' select most recent fiscal year
    Me.cmbFiscalYear.ListIndex = 0
    
    ' init to account to display
    With Me.cmbAcct
        .ListIndex = 1
        J = .ListCount
        For I = 0 To J - 1
            If .ItemData(I) = DisplayAccount Then
                .ListIndex = I
                Exit For
            End If
        Next I
    End With
    
    DisplayAmounts
    
    Me.KeyPreview = True
    
    LoadFlag = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
    
End Sub
Private Sub cmbFiscalYear_Click()
    If LoadFlag = True Then Exit Sub
    DisplayAmounts
End Sub
Private Sub cmbAcct_Click()
    If LoadFlag = True Then Exit Sub
    DisplayAmounts
End Sub
Private Sub cmdPrev_Click()
    With Me.cmbAcct
        If .ListIndex = 0 Then Exit Sub
        .ListIndex = .ListIndex - 1
        DisplayAmounts
    End With
End Sub
Private Sub cmdNext_Click()
    With Me.cmbAcct
        If .ListIndex = .ListCount - 1 Then Exit Sub
        .ListIndex = .ListIndex + 1
        DisplayAmounts
    End With
End Sub

Private Sub DisplayAmounts()
    
Dim TotalAmount, TotalBudget As Currency
    
    TotalAmount = 0
    TotalBudget = 0
    
    On Error Resume Next
    rsGL.Close
    Set rsGL = Nothing
    fg.DataMode = flexDMFree
    On Error GoTo 0
    rsGL.CursorLocation = adUseClient
    rsGL.Fields.Append "Period", adVarChar, 30, adFldIsNullable
    rsGL.Fields.Append "Amount", adCurrency
    rsGL.Fields.Append "Budget", adCurrency
    rsGL.Open , , adOpenDynamic, adLockOptimistic
    
    If GLAmount.Find(Me.cmbAcct.ItemData(Me.cmbAcct.ListIndex), Me.cmbFiscalYear) = False Then
        Exit Sub
    End If
    
    For I = 1 To GLCompany.NumberPds
        rsGL.AddNew
        x = PeriodName(Me.cmbFiscalYear, _
            I, _
            GLCompany.FirstPeriod, _
            GLCompany.NumberPds)
        rsGL!Period = Mid(x, 1, 30)
        If I = 1 Then
            rsGL!Amount = GLAmount.Amount01
            rsGL!Budget = GLAmount.Budget01
        ElseIf I = 2 Then
            rsGL!Amount = GLAmount.Amount02
            rsGL!Budget = GLAmount.Budget02
        ElseIf I = 3 Then
            rsGL!Amount = GLAmount.Amount03
            rsGL!Budget = GLAmount.Budget03
        ElseIf I = 4 Then
            rsGL!Amount = GLAmount.Amount04
            rsGL!Budget = GLAmount.Budget04
        ElseIf I = 5 Then
            rsGL!Amount = GLAmount.Amount05
            rsGL!Budget = GLAmount.Budget05
        ElseIf I = 6 Then
            rsGL!Amount = GLAmount.Amount06
            rsGL!Budget = GLAmount.Budget06
        ElseIf I = 7 Then
            rsGL!Amount = GLAmount.Amount07
            rsGL!Budget = GLAmount.Budget07
        ElseIf I = 8 Then
            rsGL!Amount = GLAmount.Amount08
            rsGL!Budget = GLAmount.Budget08
        ElseIf I = 9 Then
            rsGL!Amount = GLAmount.Amount09
            rsGL!Budget = GLAmount.Budget09
        ElseIf I = 10 Then
            rsGL!Amount = GLAmount.Amount10
            rsGL!Budget = GLAmount.Budget10
        ElseIf I = 11 Then
            rsGL!Amount = GLAmount.Amount11
            rsGL!Budget = GLAmount.Budget11
        ElseIf I = 12 Then
            rsGL!Amount = GLAmount.Amount12
            rsGL!Budget = GLAmount.Budget12
        ElseIf I = 13 Then
            rsGL!Amount = GLAmount.Amount13
            rsGL!Budget = GLAmount.Budget13
        End If
        
        rsGL.Update
    
    Next I
        
    rsGL.AddNew
    rsGL!Period = "TOTAL"
    rsGL!Amount = GLAmount.TotalAmount
    rsGL!Budget = GLAmount.TotalBudget
    rsGL.Update
        
    SetGrid rsGL, Me.fg
    
    With fg
    
        .ColWidth(0) = 3000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
    
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, 2) = True
        .Editable = flexEDNone
        .Select .Rows - 1, 0, .Rows - 1, 2
    
    End With

End Sub

Private Sub CmdExit_Click()
    Me.Hide
End Sub

Private Sub Init()

    ' populate the fiscal year combo
    SQLString = "SELECT DISTINCT FiscalYear FROM GLHistory ORDER BY FiscalYear DESC"
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then
        MsgBox "No GL History Exists!", vbExclamation
        Exit Sub
    End If
    
    rs.MoveFirst
    Do
        Me.cmbFiscalYear.AddItem rs!FiscalYear
        rs.MoveNext
    Loop Until rs.EOF
    
    ' populate the account combo
    If GLAccount.GetAllAccounts = False Then
        MsgBox "No GL Account records found!", vbExclamation
        Exit Sub
    End If
        
    Do
        If GLAccount.AcctType = "0" Then
            Me.cmbAcct.AddItem GLAccount.Account & " " & GLAccount.FullDesc
            Me.cmbAcct.ItemData(Me.cmbAcct.NewIndex) = GLAccount.Account
        End If
        If GLAccount.GetNext = False Then Exit Do
    Loop
    
    If Me.cmbAcct.ListCount = 0 Then
        MsgBox "No Type 0 GL Accounts found!", vbExclamation
        Exit Sub
    End If

    GLAmount.OpenRS

End Sub

