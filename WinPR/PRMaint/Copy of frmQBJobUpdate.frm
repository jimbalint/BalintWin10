VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQBJobUpdate 
   Caption         =   "Update to QB Jobs"
   ClientHeight    =   11295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   11295
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   11400
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoadPRBatch 
      Caption         =   "LOAD PR DATA"
      Height          =   615
      Left            =   7680
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fgBatch 
      Height          =   1335
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   5655
      _cx             =   9975
      _cy             =   2355
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
   Begin VSFlex8Ctl.VSFlexGrid fgS 
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      Top             =   5880
      Width           =   11775
      _cx             =   20770
      _cy             =   1296
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
   Begin VSFlex8Ctl.VSFlexGrid fgEE 
      Height          =   2055
      Left            =   1320
      TabIndex        =   5
      Top             =   2040
      Width           =   11775
      _cx             =   20770
      _cy             =   3625
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
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   10440
      Width           =   1575
   End
   Begin VB.CommandButton cmdQBAccounts 
      Caption         =   "GET QB ACCTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   10440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   10440
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fgER 
      Height          =   2895
      Left            =   1320
      TabIndex        =   6
      Top             =   6840
      Width           =   11775
      _cx             =   20770
      _cy             =   5106
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
   Begin VSFlex8Ctl.VSFlexGrid fgC 
      Height          =   1575
      Left            =   1320
      TabIndex        =   7
      Top             =   4200
      Width           =   11775
      _cx             =   20770
      _cy             =   2778
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
   Begin VB.Label Label5 
      Caption         =   "Batch(es) to update"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Employer Expenses"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Employee State Tax Expenses"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Employee City Tax Expenses"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Expenses:"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmQBJobUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsER As New ADODB.Recordset
Dim rsEE As New ADODB.Recordset
Dim rsC As New ADODB.Recordset
Dim rsS As New ADODB.Recordset
Dim rsPRBatch As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsQB As New ADODB.Recordset

Dim i, j, k As Long
Dim X, Y, z As String

Dim TotalGross As Currency
Dim QBDrop As String
Dim DrCol, CrCol As Byte

Dim FUNRate, SUNRate As Double

Private Sub Form_Load()

    ' **** TO DO ****
    ' multi state unemp rate
    ' state unemp rate historical

    Me.lblCompanyName = PRCompany.Name

    ' ------------------------------------------------------------------
    ' get the state unemployment rate from the company record
    ' **** multi state ****
    SUNRate = PRCompany.StateUnempPct

    ' ------------------------------------------------------------------

    ' ------------------------------------------------------------------
    ' QB Account dropdown string
    QBDrop = LoadQBDrop("SELECT * FROM QBAccount " & _
                "WHERE AccountType <> 'VENDOR'" & _
                " ORDER BY Name")
    If QBDrop = "" Then
        MsgBox "No QB Account data found!" & vbCr & _
               "Open QB and press the Get Accounts button", vbInformation
    End If
    ' ------------------------------------------------------------------

    ' -------------------------------------------------------------------
    ' setup the PRBatch Grid
    rsPRBatch.CursorLocation = adUseClient
    rsPRBatch.Fields.Append "Select", adBoolean
    rsPRBatch.Fields.Append "BatchID", adDouble
    rsPRBatch.Fields.Append "CheckDate", adDate
    rsPRBatch.Fields.Append "PEDate", adDate
    rsPRBatch.Fields.Append "Records", adDouble
    rsPRBatch.Open , , adOpenDynamic, adLockOptimistic
        
    SQLString = "SELECT * FROM PRBatch ORDER BY CheckDate DESC"
    If PRBatch.GetBySQL(SQLString) = False Then
        MsgBox "No PR data found!", vbExclamation
        GoBack
    End If
    Do
        rsPRBatch.AddNew
        rsPRBatch!Select = False
        rsPRBatch!BatchID = PRBatch.BatchID
        rsPRBatch!CheckDate = PRBatch.CheckDate
        rsPRBatch!PEDate = PRBatch.PEDate
        rsPRBatch!Records = PRBatch.RecCount
        rsPRBatch.Update
        If PRBatch.GetNext = False Then Exit Do
    Loop
    
    SetGrid rsPRBatch, Me.fgBatch
    
    With Me.fgBatch
        .SelectionMode = flexSelectionByRow
    End With
    ' -------------------------------------------------------------------

    ' -------------------------------------------------------------------
    ' Job RS
    rsJob.CursorLocation = adUseClient
    rsJob.Fields.Append "JobID", adDouble
    rsJob.Fields.Append "CityID", adDouble
    rsJob.Fields.Append "Gross", adCurrency
    rsJob.Open , , adOpenDynamic, adLockOptimistic
    ' -------------------------------------------------------------------

    DefineRS rsER, fgER
    DefineRS rsEE, fgEE
    DefineRS rsC, fgC
    DefineRS rsS, fgS

    FillRS

    fgEE.Row = 1
    fgEE.TopRow = 1
    
    fgER.Row = 1
    fgER.TopRow = 1

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub DefineRS(ByRef rs As ADODB.Recordset, fg As VSFlexGrid)

    rs.CursorLocation = adUseClient
    rs.Fields.Append "Type", adVarChar, 1, adFldIsNullable
    rs.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "DebitAcct", adVarChar, 50, adFldIsNullable
    rs.Fields.Append "DebitAmount", adCurrency
    rs.Fields.Append "CreditAcct", adVarChar, 50, adFldIsNullable
    rs.Fields.Append "CreditAmount", adCurrency
    rs.Fields.Append "ID", adDouble
    rs.Fields.Append "GlobalID", adDouble
    rs.Fields.Append "Gross", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic

    SetGrid rs, fg

    fg.ColWidth(0) = 0          ' Type
    fg.ColWidth(1) = 2200       ' Title
    fg.ColWidth(2) = 3000       ' Debit Acct
    fg.ColWidth(3) = 1500       ' Debit Amt
    fg.ColWidth(4) = 3000       ' Credit Acct
    fg.ColWidth(5) = 1500       ' Credit Amt
    fg.ColWidth(6) = 0          ' ID
    fg.ColWidth(7) = 0          ' GlobalID
    fg.ColWidth(8) = 0          ' Gross

    fg.ColComboList(2) = QBDrop
    fg.ColComboList(4) = QBDrop
    DrCol = 2
    CrCol = 4

    fg.Font.Size = 9

End Sub

Private Sub FillRS()
    
Dim FromGlobal As Boolean
Dim Glob(10) As String
    
    ' Employee Exp Grid
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobEE & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        Glob(1) = PRGlobal.Var1 & ""
        Glob(2) = PRGlobal.Var2 & ""
        Glob(3) = PRGlobal.Var3 & ""
        Glob(4) = PRGlobal.Var4 & ""
        Glob(5) = PRGlobal.Var5 & ""
        Glob(6) = PRGlobal.Var6 & ""
        Glob(7) = PRGlobal.Var7 & ""
        Glob(8) = PRGlobal.Var8 & ""
        Glob(9) = PRGlobal.Var9 & ""
        Glob(10) = PRGlobal.Var10 & ""
    Else
        For i = 1 To 10
            Glob(i) = ""
        Next i
    End If
    For i = 1 To 6
        rsEE.AddNew
        If i = 1 Then
            rsEE!Type = "D"
            rsEE!DebitAcct = Mid(Glob(i), 1, 50)
            rsEE!CreditAcct = ""
        Else
            rsEE!Type = "C"
            rsEE!DebitAcct = ""
            rsEE!CreditAcct = Mid(Glob(i), 1, 50)
        End If
        If i = 1 Then rsEE!Title = "GROSS PAY"
        If i = 2 Then rsEE!Title = "SS TAX"
        If i = 3 Then rsEE!Title = "MED TAX"
        If i = 4 Then rsEE!Title = "FWT TAX"
        If i = 5 Then rsEE!Title = "DEDUCTIONS"
        If i = 6 Then rsEE!Title = "NET PAY"
        rsEE.Update
    Next i
    
    ' Employer Exp Grid
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobER & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = True Then
        Glob(1) = PRGlobal.Var1 & ""
        Glob(2) = PRGlobal.Var2 & ""
        Glob(3) = PRGlobal.Var3 & ""
        Glob(4) = PRGlobal.Var4 & ""
        Glob(5) = PRGlobal.Var5 & ""
        Glob(6) = PRGlobal.Var6 & ""
        Glob(7) = PRGlobal.Var7 & ""
        Glob(8) = PRGlobal.Var8 & ""
        Glob(9) = PRGlobal.Var9 & ""
        Glob(10) = PRGlobal.Var10 & ""
    Else
        For i = 1 To 10
            Glob(i) = ""
        Next i
    End If
    For i = 1 To 10
        rsER.AddNew
        If i <= 5 Then
            rsER!Type = "D"
            rsER!DebitAcct = Mid(Glob(i), 1, 50)
            rsER!CreditAcct = ""
        Else
            rsER!Type = "C"
            rsER!DebitAcct = ""
            rsER!CreditAcct = Mid(Glob(i), 1, 50)
        End If
        If i = 1 Then rsER!Title = "SS TAX"
        If i = 2 Then rsER!Title = "MED TAX"
        If i = 3 Then rsER!Title = "SUTA TAX"
        If i = 4 Then rsER!Title = "FUTA TAX"
        If i = 5 Then rsER!Title = "WORK COMP"
        If i = 6 Then rsER!Title = "ACCR SS TAX"
        If i = 7 Then rsER!Title = "ACCR MED TAX"
        If i = 8 Then rsER!Title = "ACCR SUTA TAX"
        If i = 9 Then rsER!Title = "ACCR FUTA TAX"
        If i = 10 Then rsER!Title = "ACCR WORK COMP"
        rsER.Update
    Next i
    
End Sub

Private Sub cmdSave_Click()
    
    ' Employee Exp grid
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobEE & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeQBJobEE
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    i = 0
    rsEE.MoveFirst
    Do
        i = i + 1
        If i = 1 Then PRGlobal.Var1 = rsEE!DebitAcct
        If i = 2 Then PRGlobal.Var2 = rsEE!CreditAcct
        If i = 3 Then PRGlobal.Var3 = rsEE!CreditAcct
        If i = 4 Then PRGlobal.Var4 = rsEE!CreditAcct
        If i = 5 Then PRGlobal.Var5 = rsEE!CreditAcct
        If i = 6 Then PRGlobal.Var6 = rsEE!CreditAcct
        rsEE.MoveNext
        If rsEE.EOF Then Exit Do
    Loop
    PRGlobal.Save (Equate.RecPut)
    
    ' Employer Exp grid
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobER & _
                " AND UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeQBJobER
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    i = 0
    rsER.MoveFirst
    Do
        i = i + 1
        
        If i = 1 Then PRGlobal.Var1 = rsER!DebitAcct
        If i = 2 Then PRGlobal.Var2 = rsER!DebitAcct
        If i = 3 Then PRGlobal.Var3 = rsER!DebitAcct
        If i = 4 Then PRGlobal.Var4 = rsER!DebitAcct
        If i = 5 Then PRGlobal.Var5 = rsER!DebitAcct
        
        If i = 6 Then PRGlobal.Var6 = rsER!CreditAcct
        If i = 7 Then PRGlobal.Var7 = rsER!CreditAcct
        If i = 8 Then PRGlobal.Var8 = rsER!CreditAcct
        If i = 9 Then PRGlobal.Var9 = rsER!CreditAcct
        If i = 10 Then PRGlobal.Var10 = rsER!CreditAcct
        
        rsER.MoveNext
        If rsER.EOF Then Exit Do
    
    Loop
    PRGlobal.Save (Equate.RecPut)
    
    ' state grid
    If rsS.RecordCount > 0 Then
        rsS.MoveFirst
        Do
            If rsS!GlobalID = 0 Or IsNull(rsS!GlobalID) Then
                MsgBox "Global ID not set for state: " & rsS!ID, vbExclamation
                GoBack
            End If
            If PRGlobal.GetByID(rsS!GlobalID) = False Then
                MsgBox "Global ID not found for state: " & rsS!ID, vbExclamation
                GoBack
            End If
            PRGlobal.Var2 = rsS!CreditAcct
            PRGlobal.Save (Equate.RecPut)
            rsS.MoveNext
        Loop Until rsS.EOF
    End If
    
    ' city grid
    If rsC.RecordCount > 0 Then
        rsC.MoveFirst
        Do
            If rsC!GlobalID = 0 Or IsNull(rsC!GlobalID) Then
                MsgBox "Global ID not set for state: " & rsC!ID, vbExclamation
                GoBack
            End If
            If PRGlobal.GetByID(rsC!GlobalID) = False Then
                MsgBox "Global ID not found for state: " & rsC!ID, vbExclamation
                GoBack
            End If
            PRGlobal.Var2 = rsC!CreditAcct
            PRGlobal.Save (Equate.RecPut)
            rsC.MoveNext
        Loop Until rsC.EOF
    End If
    
End Sub
Private Sub cmdLoadPRBatch_Click()

    ' clear everything out
    ClearAmounts rsEE
    ClearAmounts rsER
    ClearAmounts rsC
    ClearAmounts rsS

    TotalGross = 0
    
    If rsJob.RecordCount > 0 Then
        rsJob.MoveFirst
        Do
            rsJob!Gross = 0
            rsJob.Update
            rsJob.MoveNext
        Loop Until rsJob.EOF
    End If

    rsPRBatch.MoveFirst
    Do
        If rsPRBatch!Select = False Then GoTo NxtBatch

        ' FED unemp rate for the year
        FUNRate = PRGlobal.GetAmount(PREquate.GlobalTypeFUNPct, Year(PRBatch.CheckDate))

        SQLString = "SELECT * FROM PRHist WHERE BatchID = " & rsPRBatch!BatchID
        If PRHist.GetBySQL(SQLString) = True Then
            Do
        
                i = 0
                rsEE.MoveFirst
                Do
                    i = i + 1
                    If i = 1 Then rsEE!DebitAmount = rsEE!DebitAmount + PRHist.Gross
                    If i = 2 Then rsEE!CreditAmount = rsEE!CreditAmount + PRHist.SSTax
                    If i = 3 Then rsEE!CreditAmount = rsEE!CreditAmount + PRHist.MedTax
                    If i = 4 Then rsEE!CreditAmount = rsEE!CreditAmount + PRHist.FWTTax
                    If i = 6 Then rsEE!CreditAmount = rsEE!CreditAmount + PRHist.Net + PRHist.DirectDeposit
                    rsEE.Update
                    rsEE.MoveNext
                Loop Until rsEE.EOF
                
                i = 0
                rsER.MoveFirst
                Do
                    i = i + 1
                    If i = 1 Then rsER!DebitAmount = rsER!DebitAmount + PRHist.SSTax
                    If i = 2 Then rsER!DebitAmount = rsER!DebitAmount + PRHist.MedTax
                    If i = 3 Then rsER!DebitAmount = rsER!DebitAmount + PRHist.SUNWage * SUNRate / 100
                    If i = 4 Then rsER!DebitAmount = rsER!DebitAmount + PRHist.FUNWage * FUNRate / 100
                    If i = 5 Then rsER!DebitAmount = rsER!DebitAmount + PRHist.WkcAmount
                    If i = 6 Then rsER!CreditAmount = rsER!CreditAmount + PRHist.SSTax
                    If i = 7 Then rsER!CreditAmount = rsER!CreditAmount + PRHist.MedTax
                    If i = 8 Then rsER!CreditAmount = rsER!CreditAmount + PRHist.SUNWage * SUNRate / 100
                    If i = 9 Then rsER!CreditAmount = rsER!CreditAmount + PRHist.FUNWage * FUNRate / 100
                    If i = 10 Then rsER!CreditAmount = rsER!CreditAmount + PRHist.WkcAmount
                    rsER.Update
                    rsER.MoveNext
                Loop Until rsER.EOF
        
                ' ----------------------------------------------------------------------------
                ' one PRHist record per state
                rsS.Find "ID = " & PRHist.StateID, 0, adSearchForward, 1
                If rsS.EOF Then
                    If PRState.GetByID(PRHist.StateID) = False Then
                        MsgBox "PR State not found: " & PRHist.StateID, vbExclamation
                        GoBack
                    End If
                    rsS.AddNew
                    rsS!ID = PRHist.StateID
                    rsS!Type = "C"
                    rsS!Title = PRState.StateAbbrev & " TAX"
                    rsS!DebitAcct = ""
                    rsS!DebitAmount = 0
                
                    ' QB account store in PRGlobal
                    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobState & _
                                " AND UserID = " & PRCompany.CompanyID & _
                                " AND Var1 = '" & PRHist.StateID & "'"
                    If PRGlobal.GetBySQL(SQLString) = False Then
                        PRGlobal.Clear
                        PRGlobal.TypeCode = PREquate.GlobalTypeQBJobState
                        PRGlobal.UserID = PRCompany.CompanyID
                        PRGlobal.Var1 = PRHist.StateID
                        PRGlobal.Var2 = ""
                        PRGlobal.Save (Equate.RecAdd)
                        rsS!CreditAcct = ""
                    Else
                        rsS!CreditAcct = PRGlobal.Var2
                    End If
                    rsS!GlobalID = PRGlobal.GlobalID
                Else
                    '
                End If
                rsS!CreditAmount = rsS!CreditAmount + PRHist.SWTTax
                rsS.Update
                
                ' ----------------------------------------------------------------------------

                If PRHist.GetNext = False Then Exit Do
                        
            Loop
        
            ' ----------------------------------------------------------------------------
            ' load PRDist to City rs/fg
            SQLString = "SELECT * FROM PRDist WHERE BatchID = " & rsPRBatch!BatchID
            If PRDist.GetBySQL(SQLString) Then
                Do
                    rsC.Find "ID = " & PRDist.CityID, 0, adSearchForward, 1
                    If rsC.EOF Then
                        If PRCity.GetByID(PRDist.CityID) = False Then
                            MsgBox "CityID not found: " & PRDist.CityID, vbExclamation
                            GoBack
                        End If
                        rsC.AddNew
                        rsC!Type = "C"
                        rsC!Title = PRCity.CityName & " TAX"
                        rsC!DebitAcct = ""
                        rsC!DebitAmount = 0
                        rsC!ID = PRDist.CityID
                        
                        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQBJobCity & _
                                    " AND UserID = " & PRCompany.CompanyID & _
                                    " AND Var1 = '" & PRDist.CityID & "'"
                        If PRGlobal.GetBySQL(SQLString) = False Then
                            PRGlobal.Clear
                            PRGlobal.TypeCode = PREquate.GlobalTypeQBJobCity
                            PRGlobal.UserID = PRCompany.CompanyID
                            PRGlobal.Var1 = PRDist.CityID
                            PRGlobal.Var2 = ""
                            PRGlobal.Save (Equate.RecAdd)
                            rsC!CreditAcct = ""
                        Else
                            rsC!CreditAcct = PRGlobal.Var2
                        End If
                        rsC!GlobalID = PRGlobal.GlobalID
                    Else
                        '
                    End If
                    rsC!CreditAmount = rsC!CreditAmount + PRDist.CityTax
                    rsC.Update
                    
                    ' update Gross by Job
                    rsJob.Find "JobID = " & PRDist.JobID, 0, adSearchForward, 1
                    If rsJob.EOF Then
                        rsJob.AddNew
                        rsJob!JobID = PRDist.JobID
                        rsJob!CityID = PRDist.CityID
                        rsJob!Gross = 0
                    End If
                    rsJob!Gross = rsJob!Gross + PRDist.Amount
                    rsJob.Update
                    
                    If PRDist.GetNext = False Then Exit Do
                
                Loop
            
            End If
            ' ----------------------------------------------------------------------------
                            
            ' ----------------------------------------------------------------------------
            ' update deductions to the EE expenses
            rsEE.Find "Title = 'DEDUCTIONS'", 0, adSearchForward, 1
            If rsEE.EOF Then
                MsgBox "EE deduction cat not found", vbExclamation
                GoBack
            End If
            SQLString = "SELECT * FROM PRItemHist WHERE BatchID = " & rsPRBatch!BatchID & _
                        " AND ItemType = " & PREquate.ItemTypeDED
            If PRItemHist.GetBySQL(SQLString) = True Then
                Do
                    rsEE!CreditAmount = rsEE!CreditAmount + PRItemHist.Amount
                    If PRItemHist.GetNext = False Then Exit Do
                Loop
                rsEE.Update
            End If
            ' ----------------------------------------------------------------------------
        
        End If

NxtBatch:
        rsPRBatch.MoveNext
    Loop Until rsPRBatch.EOF
    
    rsPRBatch.MoveFirst
    fgBatch.Row = 1
    fgBatch.TopRow = 1

End Sub

Private Sub cmdPrint_Click()

Dim Debits(2), Credits(2) As Currency

    CalcQB


End Sub

Private Sub CalcQB()

    ' create rsQB each time
    On Error Resume Next
    rsQB.Close
    On Error GoTo 0
    rsQB.CursorLocation = adUseClient
    rsQB.Fields.Append "Title", adVarChar, 20, adFldIsNullable
    rsQB.Fields.Append "Name", adVarChar, 50, adFldIsNullable
    rsQB.Fields.Append "DebitAmount", adCurrency
    rsQB.Fields.Append "CreditAmount", adCurrency
    rsQB.Open , , adOpenDynamic, adLockOptimistic

    ' postings per company
    rsEE.MoveFirst
    


End Sub

Private Sub ClearAmounts(ByRef rs As ADODB.Recordset)
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    Do
        rs!DebitAmount = 0
        rs!CreditAmount = 0
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
End Sub

Private Sub fgEE_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If rsEE!Type = "D" And Col <> DrCol Then Cancel = True
    If rsEE!Type = "C" And Col <> CrCol Then Cancel = True
End Sub
Private Sub fgER_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If rsER!Type = "D" And Col <> DrCol Then Cancel = True
    If rsER!Type = "C" And Col <> CrCol Then Cancel = True
End Sub
Private Sub fgBatch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
Private Sub fgS_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> CrCol Then Cancel = True
End Sub
Private Sub fgC_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> CrCol Then Cancel = True
End Sub



