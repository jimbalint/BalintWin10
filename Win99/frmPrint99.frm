VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPrint99 
   Caption         =   "Print 1099 Form"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint99.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTotals 
      Caption         =   "&TOTALS"
      Height          =   615
      Left            =   7200
      TabIndex        =   13
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Top             =   9360
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHorz 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   9480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      Calculator      =   "frmPrint99.frx":030A
      Caption         =   "frmPrint99.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrint99.frx":0394
      Keys            =   "frmPrint99.frx":03B2
      Spin            =   "frmPrint99.frx":03FC
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   9360
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6855
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   12855
      _cx             =   22675
      _cy             =   12091
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
      Begin VB.CommandButton Command1 
         Caption         =   "&PRINT"
         Height          =   615
         Left            =   7320
         TabIndex        =   11
         Top             =   9120
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cmbForm 
      Height          =   360
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   11040
      TabIndex        =   0
      Top             =   9360
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbVertical 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   9480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      Calculator      =   "frmPrint99.frx":0424
      Caption         =   "frmPrint99.frx":0444
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrint99.frx":04AE
      Keys            =   "frmPrint99.frx":04CC
      Spin            =   "frmPrint99.frx":0516
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
   Begin VB.Label Label3 
      Caption         =   "01/15/2022"
      Height          =   255
      Left            =   11040
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "1099 Form:"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Year:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   975
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
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmPrint99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim FormID, RowCount, Rw As Long
Dim PRGlobalID As Long

Private Sub cmdTotals_Click()
    cmdSave_Click
    frmTotals.TaxYear = Me.cmbTaxYear
    frmTotals.FormType = GetFormType()
    frmTotals.Show vbModal
End Sub

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    
    Me.KeyPreview = True

    With Me
        
        .cmbForm.AddItem "1099-NEC"
        .cmbForm.AddItem "1099-MISC"
        .cmbForm.AddItem "1099-R"
        .cmbForm.AddItem "1099-INT"
        .cmbForm.AddItem "1099-DIV"
        .cmbForm.ListIndex = 0
    
        PopTaxYear .cmbTaxYear
    
    End With
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With Me.fg
        If Col = 1 Then     ' set select off? - delete all detail99 records for the payee
            If .TextMatrix(Row, 1) = False Then
                SQLString = " DELETE * FROM Detail99 WHERE PayeeID = " & .TextMatrix(Row, 0) & _
                            " AND TaxYear = " & Me.cmbTaxYear.text & _
                            " AND FormType = '" & GetFormType() & "'"
                cn.Execute SQLString
                For I = 5 To .Cols - 1
                    .TextMatrix(Row, I) = ""
                Next I
            End If
        End If
        .TextMatrix(Row, 1) = SetSelect(Row)
    End With
End Sub

Private Sub cmdPrint_Click()
    
    cmdSave_Click
    SaveNudge
    
    HorzNudge = Me.tdbHorz.Value
    VertNudge = Me.tdbVertical.Value
    
    With Me
        X = Mid(.cmbForm.text, 6)
        I = Me.cmbTaxYear.text
        PrintForm99 X, I, False
    End With

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 And Col <= 4 Then
        Cancel = True
    End If
End Sub

Private Sub cmdLoad_Click()

Dim ColCt As Integer

    ' free mode grid
    With Me.fg
        
        ' grid paramters
        .DataMode = flexDMFree
        .FixedCols = 5
        .FixedRows = 1
        .Rows = 1
        .BackColorAlternate = RGB(195, 195, 195)
        .ExplorerBar = flexExMove + flexExSort
        .Editable = flexEDKbdMouse

        ' get the form
        SQLString = " SELECT * FROM Form99 WHERE TaxYear = " & Me.cmbTaxYear.text & " " & _
                    " AND FormType = '" & GetFormType() & "'"
        If Form99.GetBySQL(SQLString) = False Then
            MsgBox "Form NF: " & Me.cmbTaxYear.text & " 1099-" & Me.cmbForm.text, vbExclamation
            GoBack
        End If
        
        FormID = Form99.FormID
        
        ' add the initial columns
        ' same for all 1099 forms
        
        .TextMatrix(0, 0) = "PayeeID"
        .ColData(0) = "PayeeID"
        .ColHidden(0) = True
        .ColWidth(0) = 1500
        
        .TextMatrix(0, 1) = "Select"
        .ColData(1) = "Select"
        .ColDataType(1) = flexDTBoolean
        .ColWidth(1) = 750
        
        .TextMatrix(0, 2) = "Payee #"
        .ColData(2) = "PayeeNumber"
        .ColDataType(2) = flexDTDouble
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Payee Name"
        .ColData(3) = "PayeeName"
        .ColDataType(3) = flexDTString
        .ColWidth(3) = 2000
        
        .TextMatrix(0, 4) = "Payee Fed ID"
        .ColData(4) = "FederalID"
        .ColDataType(4) = flexDTString
        .ColWidth(4) = 1500
        
        ColCt = 4
        
        ' get the fields of the form
        ' put to header line of grid
        ' coldata is boxname
        SQLString = " SELECT * FROM Field99 WHERE TaxYear = " & Me.cmbTaxYear.text & _
                    " AND FormType = '" & GetFormType() & "'" & _
                    " AND QuickEntry > 0 ORDER BY QuickEntry"
        If Field99.GetBySQL(SQLString) = False Then
            MsgBox "Fields NF: " & Me.cmbTaxYear.text & " 1099-" & Me.cmbForm.text, vbExclamation
            GoBack
        End If
        
        Do
            
            ColCt = ColCt + 1
            .Cols = ColCt + 1
            
            .TextMatrix(0, ColCt) = Field99.BTitle
            .ColData(ColCt) = Field99.BoxName
            
            If Field99.FieldFormat = Equate.fmtAmount Then
                .ColDataType(ColCt) = flexDTCurrency
                .ColFormat(ColCt) = "Currency"
                .ColWidth(ColCt) = 1300
            ElseIf Field99.FieldFormat = Equate.fmtString Then
                .ColDataType(ColCt) = flexDTString
                .ColWidth(ColCt) = 2000
            End If
            
            If Field99.GetNext = False Then Exit Do
        
        Loop

        
        ' load the payee data
        SQLString = " SELECT * FROM Payee99 ORDER BY PayeeNumber"
        If Payee99.GetBySQL(SQLString) = False Then
            MsgBox "No Payee info found!", vbExclamation
            GoBack
        End If
        
        Rw = 0

        Do
            With Me.fg
                
                ' inactive w/ no detail data - skip it
                If Payee99.Inactive = 1 Then
                    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & Payee99.PayeeID & _
                                " AND FormType = '" & GetFormType() & "'" & _
                                " AND TaxYear = " & Me.cmbTaxYear.text
                    If Detail99.GetBySQL(SQLString) = False Then
                        GoTo NextPayee
                    End If
                End If
                
                Rw = Rw + 1
                .Rows = Rw + 1
                
                ' load the info from Payee99
                .TextMatrix(Rw, 0) = Payee99.PayeeID
                .TextMatrix(Rw, 1) = False
                .TextMatrix(Rw, 2) = Payee99.PayeeNumber
                .TextMatrix(Rw, 3) = Payee99.PayeeName
                .TextMatrix(Rw, 4) = Payee99.FederalID
                
                ' load the detail data
                SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & Payee99.PayeeID & _
                            " AND FormType = '" & GetFormType() & "'" & _
                            " AND TaxYear = " & Me.cmbTaxYear.text
                If Detail99.GetBySQL(SQLString) = True Then

                    Do
                        For J = 5 To .Cols - 1
                            If .ColData(J) = Detail99.BoxName Then
                                If .ColFormat(J) = "Currency" Then
                                    .TextMatrix(Rw, J) = Format(Detail99.FieldValue, "Currency")
                                Else
                                    .TextMatrix(Rw, J) = Detail99.FieldValue
                                End If
                            End If
                        Next J
                        If Detail99.GetNext = False Then Exit Do
                    Loop
                End If
                
                .TextMatrix(Rw, 1) = SetSelect(Rw)
        
NextPayee:
                If Payee99.GetNext = False Then Exit Do
            
            End With
        
        Loop
        
        ' 2022-01-15 causing issue for new clients???
        ' .AutoSize 0, .Cols - 1
        
        .TabBehavior = flexTabCells
    
    End With
    
    ' load nudge
    PRGlobalID = 0
    With Me
        SQLString = " SELECT * FROM PRGlobal WHERE UserID = " & User.ID & _
                    " AND Description = '" & Me.cmbForm.text & "'"
        If PRGlobal.GetBySQL(SQLString) = False Then
            .tdbHorz.Value = 0
            .tdbVertical.Value = 0
        Else
            PRGlobalID = PRGlobal.GlobalID
            .tdbHorz.Value = PRGlobal.Var1
            .tdbVertical.Value = PRGlobal.Var2
        End If
    End With

End Sub

Private Sub SaveNudge()

    If PRGlobalID = 0 Then
        PRGlobal.Clear
        PRGlobal.UserID = User.ID
        PRGlobal.Description = Me.cmbForm.text
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    PRGlobal.Var1 = Me.tdbHorz.Value
    PRGlobal.Var2 = Me.tdbVertical.Value
    PRGlobal.Save (Equate.RecPut)

End Sub

Private Sub cmdSave_Click()

    With Me.fg
    
        For Rw = 1 To .Rows - 1
    
            ' reset the FieldValue for all for payee/form
            SQLString = " UPDATE Detail99 SET FieldValue = 'JimBo' WHERE " & _
                        " PayeeID = " & .TextMatrix(Rw, 0) & _
                        " AND FormType = '" & GetFormType() & "'" & _
                        " AND TaxYear = " & Me.cmbTaxYear.text
            cn.Execute SQLString
    
            For J = 5 To .Cols - 1
                
                X = Trim(.TextMatrix(Rw, J))
                If X <> "" Then
                    ' see if the detail record already exists
                    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & .TextMatrix(Rw, 0) & _
                                " AND FormType = '" & GetFormType() & "'" & _
                                " AND TaxYear = " & Me.cmbTaxYear.text & _
                                " AND BoxName = '" & .ColData(J) & "'"
                                
                    If Detail99.GetBySQL(SQLString) = True Then
                        SQLString = " UPDATE Detail99 SET FieldValue = '" & X & "' " & _
                                    " WHERE PayeeID = " & .TextMatrix(Rw, 0) & _
                                    " AND FormType = '" & GetFormType() & "'" & _
                                    " AND TaxYear = " & Me.cmbTaxYear.text & _
                                    " AND BoxName = '" & .ColData(J) & "'"
                        cn.Execute SQLString
                    Else
                        Detail99.Clear
                        Detail99.PayeeID = .TextMatrix(Rw, 0)
                        Detail99.FormType = GetFormType
                        Detail99.TaxYear = Me.cmbTaxYear.text
                        Detail99.BoxName = .ColData(J)
                        Detail99.FieldValue = X
                        Detail99.Save (Equate.RecAdd)
                    End If
                End If
                
            Next J
        
            ' delete old records not update
            SQLString = " DELETE * FROM Detail99 WHERE " & _
                        " PayeeID = " & .TextMatrix(Rw, 0) & _
                        " AND FormType = '" & GetFormType() & "' " & _
                        " AND TaxYear = " & Me.cmbTaxYear.text & _
                        " AND FieldValue = 'JimBo'"
            cn.Execute SQLString
        
        Next Rw
    
    End With

    SaveNudge

End Sub
Private Function SetSelect(ByVal fgRow As Long) As Boolean
Dim fgCol As Long

    SetSelect = False
    For fgCol = 5 To fg.Cols - 1
        If Trim(fg.TextMatrix(fgRow, fgCol)) <> "" Then
            SetSelect = True
            Exit Function
        End If
    Next fgCol
    
End Function

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Function GetFormType() As String

Dim ii As Integer

    ii = Me.cmbForm.ListIndex
    GetFormType = ""
    If ii = 0 Then GetFormType = "NEC"
    If ii = 1 Then GetFormType = "MISC"
    If ii = 2 Then GetFormType = "R"
    If ii = 3 Then GetFormType = "INT"
    If ii = 4 Then GetFormType = "DIV"

End Function
