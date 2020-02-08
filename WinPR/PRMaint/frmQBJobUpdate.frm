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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7695
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   12735
      _cx             =   22463
      _cy             =   13573
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   11400
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoadPRBatch 
      Caption         =   "LOAD PR DATA"
      Height          =   615
      Left            =   7680
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fgBatch 
      Height          =   1335
      Left            =   1320
      TabIndex        =   5
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
   Begin VB.Label Label5 
      Caption         =   "Batch(es) to update"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   840
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

Dim ItemDrop, TypeDrop, JobDrop, CatDrop, QBDrop As String

Dim FUNRate, SUNRate As Double
Dim rsPRBatch As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim i, j, k As Long
Dim X, Y, z As String


Private Sub Form_Load()

    If TableExists("QBUpdate", cn) = False Then
        QBUpdateCreate
    End If
    SQLString = "DELETE * FROM QBUpdate"
    cn.Execute SQLString
    
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
    
    ' ------------------------------------------------------------------
    ' Other drop downs
    CatDrop = "|#1;EE WAGE|#2;EE TAX|#3;EE DED|#4;ER TAX"
    JobDrop = "|#0;NO|#1;YES"
    
    For i = 1 To 20
        TypeDrop = Trim(TypeDrop) & "|#" & i & ";"
        If i = 1 Then TypeDrop = Trim(TypeDrop) & "REG PAY"
        If i = 2 Then TypeDrop = Trim(TypeDrop) & "OVT PAY"
        If i = 3 Then TypeDrop = Trim(TypeDrop) & "OTHR EARNG"
        If i = 4 Then TypeDrop = Trim(TypeDrop) & "SS TAX"
        If i = 5 Then TypeDrop = Trim(TypeDrop) & "MED TAX"
        If i = 6 Then TypeDrop = Trim(TypeDrop) & "FWT TAX"
        If i = 7 Then TypeDrop = Trim(TypeDrop) & "SWT TAX"
        If i = 8 Then TypeDrop = Trim(TypeDrop) & "CWT TAX"
        If i = 9 Then TypeDrop = Trim(TypeDrop) & "DEDUCTION"
        If i = 10 Then TypeDrop = Trim(TypeDrop) & "NET PAY"
        If i = 11 Then TypeDrop = Trim(TypeDrop) & "SS TAX"
        If i = 12 Then TypeDrop = Trim(TypeDrop) & "MED TAX"
        If i = 13 Then TypeDrop = Trim(TypeDrop) & "FUTA TAX"
        If i = 14 Then TypeDrop = Trim(TypeDrop) & "SUTA TAX"
        If i = 15 Then TypeDrop = Trim(TypeDrop) & "WORK COMP"
        If i = 16 Then TypeDrop = Trim(TypeDrop) & "SS TAX"
        If i = 17 Then TypeDrop = Trim(TypeDrop) & "MED TAX"
        If i = 18 Then TypeDrop = Trim(TypeDrop) & "FUTA TAX"
        If i = 19 Then TypeDrop = Trim(TypeDrop) & "SUTA TAX"
        If i = 20 Then TypeDrop = Trim(TypeDrop) & "WORK COMP"
    Next i
    
    ' employer items
    ItemDrop = "|#999999;==="
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 " & _
                " AND (ItemType = " & PREquate.ItemTypeOE & _
                " OR ItemType = " & PREquate.ItemTypeDED & ")" & _
                " ORDER BY ItemType, ItemID"
    If PRItem.GetBySQL(SQLString) = True Then
        Do
            ItemDrop = Trim(ItemDrop) & "|#" & PRItem.ItemID & ";" & Trim(PRItem.Abbreviation)
            If PRItem.GetNext = False Then Exit Do
        Loop
    End If
    
    ' ------------------------------------------------------------------
    ' get the state unemployment rate from the company record
    ' **** multi state ****
    SUNRate = PRCompany.StateUnempPct

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

    SQLString = "SELECT * FROM QBUpdate ORDER BY Category, Type, RelatedID"
    rsInit SQLString, cn, rs
    SetGrid rs, fg
    
    With Me.fg
    
        .ColWidth(0) = 0      ' QBUpdateID
        .ColWidth(1) = 1500   ' Category - EE Wage/EE Tax/EE Ded/ER Tax
        .ColWidth(2) = 0      ' Post - Dr / Cr
        .ColWidth(3) = 1000    ' Per Job - Y / N
        .ColWidth(4) = 1000   ' Title
        .ColWidth(5) = 1000   ' Type
        .ColWidth(6) = 1000   ' RelatedID
        .ColWidth(7) = 2000   ' QBID
        .ColWidth(8) = 2000    ' Debit Amount
        .ColWidth(9) = 2000    ' Credit Amount

        .ColComboList(1) = CatDrop
        .ColComboList(3) = JobDrop
        .ColComboList(5) = TypeDrop
        .ColComboList(6) = ItemDrop
        .ColComboList(7) = QBDrop
    
    End With
    
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

Private Sub cmdLoadPRBatch_Click()

Dim QBCat, QBType As Byte
Dim QBRelated As Long

    rsPRBatch.MoveFirst
    Do
        If rsPRBatch!Select = False Then GoTo NxtPRBatch

        SQLString = "SELECT * FROM PRHist WHERE BatchID = " & rsPRBatch!BatchID
        If PRHist.GetBySQL(SQLString) = True Then
            Do
                If PRHist.GetNext = False Then Exit Do
            Loop
            
            ' earnings from PRDist
            QBCat = PREquate.GlobalTypeQB_EE_Wage
            SQLString = "SELECT * FROM PRDist WHERE BatchID = " & rsPRBatch!BatchID
            If PRDist.GetBySQL(SQLString) Then
                Do
                    If PRDist.ItemType = PREquate.ItemTypeRegPay Then
                        QBType = PREquate.qbItem_RegPay
                        QBRelated = 999999
                    ElseIf PRDist.ItemType = PREquate.ItemTypeOvtPay Then
                        QBType = PREquate.qbItem_OvtPay
                        QBRelated = 999999
                    Else
                        QBType = PREquate.ItemTypeOE
                        If PRDist.EmployerItemID = 0 Then
                            If PRItem.GetByID(PRDist.ItemID) Then
                                QBRelated = PRItem.EmployerItemID
                            End If
                        Else
                            QBRelated = PRDist.EmployerItemID
                        End If
                    End If
                    qbUpd QBCat, QBType, QBRelated, PRDist.Amount
                    
                    If PRDist.GetNext = False Then Exit Do
                Loop
            End If

            ' deductions from PRItemHist
            QBCat = PREquate.GlobalTypeQB_EE_Tax
            QBType = PREquate.qbItem_DED
            SQLString = "SELECT * FROM PRItemHist WHERE BatchID = " & rsPRBatch!BatchID & _
                        " AND ItemType = " & PREquate.ItemTypeDED
            If PRItemHist.GetBySQL(SQLString) Then
                Do
                    qbUpd QBCat, QBType, PRItemHist.EmployerItemID, PRItemHist.Amount
                    If PRItemHist.GetNext = False Then Exit Do
                Loop
            End If
        
        End If
NxtPRBatch:
        rsPRBatch.MoveNext
    Loop Until rsPRBatch.EOF

    rs.Requery

End Sub

Private Sub qbUpd(ByVal QBCat As Byte, _
                  ByVal QBType As Byte, _
                  ByVal QBRelated As Long, _
                  ByVal Amt As Currency)

Dim FFlag As Boolean

    FFlag = False
    If rs.RecordCount = 0 Then
    Else
        rs.MoveFirst
        Do
            If rs!Category = QBCat And rs!Type = QBType And rs!RelatedID = QBRelated Then
                FFlag = True
                Exit Do
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
        
    If FFlag = False Then
        rs.AddNew
        rs!Category = QBCat
                
        ' debit or credit???
        If QBType = PREquate.qbItem_RegPay Then rs!Post = "D"
        If QBType = PREquate.qbItem_OvtPay Then rs!Post = "D"
        If QBType = PREquate.qbItem_OE Then rs!Post = "D"
        If QBType = PREquate.qbItem_DED Then rs!Post = "C"
        
        rs!PerJob = 1           ' !!!
        rs!Title = "AAA"
        rs!Type = QBType
        rs!RelatedID = QBRelated
        rs!QBID = ""
        rs!DebitAmount = 0
        rs!CreditAmount = 0
    
    End If
    
    If rs!Post = "D" Then
        rs!DebitAmount = rs!DebitAmount + Amt
    Else
        rs!CreditAmount = rs!CreditAmount + Amt
    End If
        
    rs.Update

End Sub

