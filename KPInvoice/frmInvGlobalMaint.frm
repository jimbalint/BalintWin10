VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInvGlobalMaint 
   Caption         =   "Invoicing Global Maintenance"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvGlobalMaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQBUpdate 
      Caption         =   "&QB ACCOUNTS UPDATE"
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   9360
      TabIndex        =   4
      Top             =   8880
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6135
      Left            =   570
      TabIndex        =   3
      Top             =   2280
      Width           =   10695
      _cx             =   18865
      _cy             =   10821
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
   Begin VB.ComboBox cmbGlobalType 
      Height          =   360
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Select Item:"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11295
   End
End
Attribute VB_Name = "frmInvGlobalMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String

Dim rs As New ADODB.Recordset
Dim rsQB As New ADODB.Recordset

Dim GType As Byte
Dim QBAcctDrop, QBTemplateDrop, QBItemDrop As String
Dim LoadFlag As Boolean

Dim rsPrinters As New ADODB.Recordset
Dim PrinterDrop As String
Dim JobDrop As String

Private Sub Form_Load()

    LoadFlag = True
    
    With Me.cmbGlobalType
        For I = 1 To 11
            X = ""
            If I = 1 Then J = InvEquate.GlobalTypeTruck:        X = "Truck"
            If I = 2 Then J = InvEquate.GlobalTypeTrailer:      X = "Trailer"
            If I = 3 Then J = InvEquate.GlobalTypeDriver:       X = "Driver"
            'If i = 4 Then j = InvEquate.GlobalTypeTerms:    X = "Terms"
            If I = 5 Then J = InvEquate.GlobalTypeComment:      X = "Comments"
            'If I = 7 Then J = InvEquate.GlobalTypeQBSetup:      X = "QB Setup"
            If I = 8 Then J = InvEquate.GlobalTypeInvPrinter:   X = "Invoice Printer"
            If I = 11 Then J = InvEquate.GlobalTypeVAdj:        X = "Printer Adjustment"
            If X <> "" Then
                .AddItem (X)
                .ItemData(.NewIndex) = I
            End If
        Next I
    End With

    LoadPrinters
    
    Me.lblCompanyName = PRCompany.Name

    Me.KeyPreview = True

    LoadFlag = False

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdQBUpdate_Click()
    frmQBAccts.Show vbModal
End Sub

Private Sub LoadPrinters()

    PrinterDrop = ""
    
    On Error Resume Next
    rsPrinters.Close
    On Error GoTo 0
    
    rsPrinters.CursorLocation = adUseClient
    rsPrinters.Fields.Append "PrinterName", adVarChar, 60, adFldIsNullable
    rsPrinters.Open , , adOpenDynamic, adLockOptimistic
    
    Set Prvw = New frmPreview
    For I = 0 To Prvw.vsp.NDevices - 1
        rsPrinters.AddNew
        rsPrinters!PrinterName = Prvw.vsp.Devices(I)
        rsPrinters.Update
    Next I
    
    If rsPrinters.RecordCount = 0 Then Exit Sub
    
    rsPrinters.Sort = "PrinterName"
    rsPrinters.MoveFirst
    Do
        PrinterDrop = PrinterDrop & "|" & rsPrinters!PrinterName
        rsPrinters.MoveNext
    Loop Until rsPrinters.EOF

End Sub

Private Sub cmbGlobalType_Click()
    
    If LoadFlag = True Then Exit Sub
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    fg.Clear
    
    With Me.cmbGlobalType
    
        If .ListIndex = -1 Then
            Exit Sub
        End If
    
        GType = .ItemData(.ListIndex)
        
        ' for QB setup - just two items in one InvGlobal record
        ' one line in the grid
        If GType = InvEquate.GlobalTypeQBSetup Then
        
            ' Me.cmdAdd.Enabled = False
            ' Me.cmdDelete.Enabled = False
            
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeQBSetup
            If InvGlobal.GetBySQL(SQLString) = False Then
                InvGlobal.Clear
                InvGlobal.CompanyID = PRCompany.CompanyID
                InvGlobal.TypeCode = InvEquate.GlobalTypeQBSetup
                InvGlobal.rsAdd
            End If
            
            On Error Resume Next
            rs.Close
            rsQB.Close
            On Error GoTo 0
            rsQB.CursorLocation = adUseClient
            rsQB.Fields.Append "AR_Account", adDouble
            rsQB.Fields.Append "Template", adDouble
            rsQB.Fields.Append "FreightItem", adDouble
            rsQB.Open , , adOpenDynamic, adLockOptimistic
            
            rsQB.AddNew

            ' >>>> field not found ????
            If QBAccount.GetByID(NumValue(InvGlobal.Var1)) = True Then
                rsQB!AR_Account = QBAccount.QBAccountID
            Else
                rsQB!AR_Account = 0
            End If
            
            If QBAccount.GetByID(NumValue(InvGlobal.Var2)) = True Then
                rsQB!Template = QBAccount.QBAccountID
            Else
                rsQB!Template = 0
            End If
            
            If InvStock.GetByID(NumValue(InvGlobal.Var3)) = True Then
                rsQB!FreightItem = InvStock.StockID
            Else
                rsQB!FreightItem = 0
            End If
            
            rsQB.Update
            
            SetGrid rsQB, fg
        
            QBAcctDrop = "|#0; "
            SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'AccountsReceivable' " & _
                        " ORDER BY Name"
            If QBAccount.GetBySQL(SQLString) = True Then
                Do
                    QBAcctDrop = QBAcctDrop & "|#" & QBAccount.QBAccountID & ";" & QBAccount.Name
                    If QBAccount.GetNext = False Then Exit Do
                Loop
            End If
            
            QBTemplateDrop = "|#0; "
            SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'TEMPLATE' " & _
                        " ORDER BY NAME"
            If QBAccount.GetBySQL(SQLString) = True Then
                Do
                    QBTemplateDrop = QBTemplateDrop & "|#" & QBAccount.QBAccountID & ";" & QBAccount.Name
                    If QBAccount.GetNext = False Then Exit Do
                Loop
            End If
        
            QBItemDrop = "!#0; "
            SQLString = "SELECT * FROM InvStock WHERE Description = 'Freight' AND JobID = 0"
            If InvStock.GetBySQL(SQLString) = True Then
                Do
                    QBItemDrop = QBItemDrop & "|#" & InvStock.StockID & ";" & InvStock.Description
                    If InvStock.GetNext = False Then Exit Do
                Loop
            End If
            
            fg.ColComboList(0) = QBAcctDrop
            fg.ColWidth(0) = 3000
            fg.ColComboList(1) = QBTemplateDrop
            fg.ColWidth(1) = 3000
            fg.ColComboList(2) = QBItemDrop
            fg.ColWidth(2) = 3000
        ElseIf GType = InvEquate.GlobalTypeInvPrinter Then
        
            On Error Resume Next
            rs.Close
            rsQB.Close
            On Error GoTo 0
            
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeInvPrinter
            If InvGlobal.GetBySQL(SQLString) = False Then
                InvGlobal.Clear
                InvGlobal.CompanyID = PRCompany.CompanyID
                InvGlobal.TypeCode = InvEquate.GlobalTypeInvPrinter
                InvGlobal.rsAdd
            End If
            
            SQLString = "SELECT GlobalID, CompanyID, TypeCode, Var1 FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeInvPrinter
            rsInit SQLString, cnDes, rs
            
'            rs.CursorLocation = adUseClient
'            rs.Fields.Append "PrinterName", adVarChar, 60, adFldIsNullable
'            rs.Open , , adOpenDynamic, adLockOptimistic
'
'            rs.AddNew
'            rs!PrinterName = Mid(InvGlobal.Var1, 1, 60)
'            rs.Update
            
            SetGrid rs, fg
            
            fg.ColWidth(0) = 0
            fg.ColWidth(1) = 0
            fg.ColWidth(2) = 0
            fg.ColComboList(3) = PrinterDrop
            fg.ColWidth(3) = 9000
            fg.TextMatrix(0, 3) = "Printer"
        
        ElseIf GType = InvEquate.GlobalTypeVAdj Then
        
            On Error Resume Next
            rs.Close
            rsQB.Close
            On Error GoTo 0
            
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeVAdj
            If InvGlobal.GetBySQL(SQLString) = False Then
                InvGlobal.Clear
                InvGlobal.CompanyID = PRCompany.CompanyID
                InvGlobal.TypeCode = InvEquate.GlobalTypeVAdj
                InvGlobal.Byte1 = 0
                InvGlobal.rsAdd
            End If
        
            SQLString = "SELECT GlobalID, CompanyID, TypeCode, Var1 FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeVAdj
            rsInit SQLString, cnDes, rs
        
'            rs.CursorLocation = adUseClient
'            rs.Fields.Append "PrinterAdjustment", adInteger
'            rs.Open , , adOpenDynamic, adLockOptimistic
'
'            rs.AddNew
'            rs!PrinterAdjustment = InvGlobal.Byte1
'            rs.Update
            
            SetGrid rs, fg
            
            fg.ColWidth(0) = 0
            fg.ColWidth(1) = 0
            fg.ColWidth(2) = 0
            fg.ColWidth(3) = 5000
            fg.TextMatrix(0, 3) = "Vertical Adjustment"
            
        Else        ' truck / trailer / driver
        
            Me.cmdAdd.Enabled = True
            Me.cmdDelete.Enabled = True
            
            On Error Resume Next
            rs.Close
            rsQB.Close
            On Error GoTo 0
            
            SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & GType & _
                        " ORDER BY DESCRIPTION"
            rsInit SQLString, cnDes, rs
            
            ' clear out blank entries
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do
                    If IsNull(rs!Description) Or IsNull(rs!Description) Then
                        rs.Delete
                    End If
                    rs.Update
                    rs.MoveNext
                Loop Until rs.EOF
            End If

            SetGrid rs, fg
        
            With fg
                    
                For I = 0 To .Cols - 1
                    .ColKey(I) = .TextMatrix(0, I)
                    If .TextMatrix(0, I) <> "Description" Then
                        .ColHidden(I) = True
                    Else
                        .ColWidth(I) = fg.Width - 200
                    End If
                Next I
                
                If GType = InvEquate.GlobalTypeTruck Then .TextMatrix(0, .ColIndex("Description")) = "No. - Truck - Lic."
                If GType = InvEquate.GlobalTypeTrailer Then .TextMatrix(0, .ColIndex("Description")) = "No. - Trailer - Lic."
                If GType = InvEquate.GlobalTypeDriver Then .TextMatrix(0, .ColIndex("Description")) = "Driver"
                'If GType = InvEquate.GlobalTypeTerms Then .TextMatrix(0, .ColIndex("Description")) = "Sales Terms"
                If GType = InvEquate.GlobalTypeComment Then .TextMatrix(0, .ColIndex("Description")) = "Invoice Comment"
                    
                ' sales terms are maintained in QB
                If GType = InvEquate.GlobalTypeTerms Then
                    MsgBox "Sales terms are maintained in QuickBooks", vbInformation
                    .Enabled = False
                    Me.cmdAdd.Enabled = False
                    Me.cmdDelete.Enabled = False
                Else
                    .Enabled = True
                    Me.cmdAdd.Enabled = True
                    Me.cmdDelete.Enabled = True
                End If
            
            End With
        
        End If
    
    End With

End Sub

Private Sub cmdAdd_Click()
    If Me.cmbGlobalType.ListIndex = -1 Then Exit Sub
    rs.AddNew
    rs!CompanyID = PRCompany.CompanyID
    rs!TypeCode = GType
    rs.Update
    fg.DataRefresh
End Sub

Private Sub cmdDelete_Click()
    If Me.cmbGlobalType.ListIndex = -1 Then Exit Sub
    If fg.Row <= 0 Then Exit Sub
    If MsgBox("OK to delete: " & Trim(rs!Description & "") & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    rs.Delete
    rs.Update
    fg.DataRefresh
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With Me.cmbGlobalType
        If .ItemData(.ListIndex) = InvEquate.GlobalTypeQBSetup Then
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeQBSetup
            boo = InvGlobal.GetBySQL(SQLString)
            InvGlobal.Var1 = rsQB!AR_Account
            InvGlobal.Var2 = rsQB!Template
            InvGlobal.Var3 = rsQB!FreightItem
            InvGlobal.rsPut
        ElseIf .ItemData(.ListIndex) = InvEquate.GlobalTypeInvPrinter Then
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeInvPrinter
            boo = InvGlobal.GetBySQL(SQLString)
            InvGlobal.Var1 = fg.Cell(flexcpTextDisplay, Row, Col)
            InvGlobal.rsPut
        ElseIf .ItemData(.ListIndex) = InvEquate.GlobalTypeVAdj Then
            SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                        " AND TypeCode = " & InvEquate.GlobalTypeVAdj
            boo = InvGlobal.GetBySQL(SQLString)
            InvGlobal.Byte1 = fg.Cell(flexcpTextDisplay, Row, Col)
            InvGlobal.rsPut
        End If
    End With
End Sub


