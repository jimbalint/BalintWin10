VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm1099 
   Caption         =   "Payroll 1099 Processing"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm1099.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5295
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   8055
      _cx             =   14208
      _cy             =   9340
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
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   1680
      TabIndex        =   12
      Top             =   9600
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHorzNudge 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   9360
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calculator      =   "frm1099.frx":030A
      Caption         =   "frm1099.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm1099.frx":03A0
      Keys            =   "frm1099.frx":03BE
      Spin            =   "frm1099.frx":0408
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
      MaxValueVT      =   6356997
      MinValueVT      =   5242885
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "CLEAR ALL"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdChkAll 
      Caption         =   "CHECK ALL"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&CALCULATE"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin TDBNumber6Ctl.TDBNumber tdbMinAmount 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calculator      =   "frm1099.frx":0430
      Caption         =   "frm1099.frx":0450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm1099.frx":04BA
      Keys            =   "frm1099.frx":04D8
      Spin            =   "frm1099.frx":0522
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
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbItem 
      Height          =   360
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   5055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   3960
      TabIndex        =   0
      Top             =   9600
      Width           =   1575
   End
   Begin TDBNumber6Ctl.TDBNumber tdbVertNudge 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   9840
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calculator      =   "frm1099.frx":054A
      Caption         =   "frm1099.frx":056A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm1099.frx":05DC
      Keys            =   "frm1099.frx":05FA
      Spin            =   "frm1099.frx":0644
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
      MaxValueVT      =   6356997
      MinValueVT      =   5242885
   End
   Begin VB.Label lblTotal 
      Caption         =   "Label3"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   3240
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Payroll Item:"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Year:"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   855
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frm1099"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim Flg As Boolean
Dim TaxYear As Long
Dim GlobalID As Long
Public rs As New ADODB.Recordset
Dim TlCount As Long
Dim TlAmount As Currency

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    Me.KeyPreview = True

    ' init tax year dropdown
    With Me.cmbTaxYear
        SQLString = "SELECT * FROM PRBatch ORDER BY YearMonth DESC"
        If PRBatch.GetBySQL(SQLString) = False Then
            MsgBox "No payroll data found!", vbExclamation
            GoBack
        End If
        Do
            TaxYear = Int(PRBatch.YearMonth / 100)
            If .ListCount > 0 Then
                Flg = False
                For j = 0 To .ListCount - 1
                    .ListIndex = j
                    If .Text = TaxYear Then
                        Flg = True
                        Exit For
                    End If
                Next j
                If Flg = False Then
                    .AddItem TaxYear
                End If
            Else
                .AddItem TaxYear
            End If
            If PRBatch.GetNext = False Then Exit Do
        Loop
        If .ListCount > 1 Then
            .ListIndex = 1
        Else
            .ListIndex = 0
        End If
    End With

    ' init the item dropdown
    With Me.cmbItem
        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 AND " & _
                   "ItemType = " & PREquate.ItemTypeOE & _
                   " ORDER BY ItemID"
        If PRItem.GetBySQL(SQLString) = False Then
            MsgBox "No items found!", vbExclamation
            GoBack
        End If
        
        Do
            .AddItem PRItem.Title
            .ItemData(.NewIndex) = PRItem.ItemID
            If PRItem.GetNext = False Then Exit Do
        Loop
        .ListIndex = 0
    End With
    
    tdbAmountSet Me.tdbMinAmount
                   
    ' get the defaults
    SQLString = "SELECT * FROM PRGlobal WHERE Description = '1099' AND " & _
                "UserID = " & PRCompany.CompanyID
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.Description = "1099"
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Var1 = "0"
        PRGlobal.Var2 = "0"
        PRGlobal.Var3 = "0"
        PRGlobal.Var4 = "0"
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    GlobalID = PRGlobal.GlobalID
    SetNudge Me.tdbHorzNudge
    SetNudge Me.tdbVertNudge
    Me.tdbMinAmount = PRGlobal.Var1
    Me.tdbHorzNudge = PRGlobal.Var2
    Me.tdbVertNudge = PRGlobal.Var3
                   
    If PRGlobal.Var4 <> "0" Then
        With Me.cmbItem
            For i = 0 To .ListCount - 1
                X = .ItemData(i)
                If X = PRGlobal.Var4 Then
                    .ListIndex = i
                    Exit For
                End If
            Next i
        End With
    End If
                   
    Me.lblTotal = ""

    ' !!!!!!!!!!!!!!!!
    ' cmdCalc_Click
    ' cmdPrint_Click
    ' !!!!!!!!!!!!!!!!

End Sub
Private Sub cmdCalc_Click()

    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Select", adBoolean
    rs.Fields.Append "EmployeeName", adChar, 30
    rs.Fields.Append "EmployeeNumber", adDouble
    rs.Fields.Append "Amount", adCurrency
    rs.Fields.Append "EmployeeID", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic

    TaxYear = Me.cmbTaxYear.Text
    i = TaxYear * 100 + 1
    j = TaxYear * 100 + 12
    SQLString = "SELECT * FROM PRDist WHERE YearMonth >= " & i & _
                " AND YearMonth <= " & j & _
                " AND EmployerItemID = " & Me.cmbItem.ItemData(Me.cmbItem.ListIndex) & _
                " ORDER BY EmployeeID"
    If PRDist.GetBySQL(SQLString) = False Then
        MsgBox "No amounts found for that item!", vbExclamation
        GoBack
    End If
    
    Me.MousePointer = vbHourglass
    Me.lblTotal = "Now scanning payroll data ... a"
    Me.Refresh
    
    Do
        
        Flg = True
        If rs.RecordCount > 0 Then
            rs.Find "EmployeeID = " & PRDist.EmployeeID, 0, adSearchForward, 1
            If rs.EOF Then Flg = False
        Else
            Flg = False
        End If
        
        If Flg = False Then
            If PREmployee.GetByID(PRDist.EmployeeID) = False Then
                MsgBox "Employee ID not found: " & PRDist.EmployeeID, vbExclamation
                GoBack
            End If
            rs.AddNew
            rs!EmployeeName = Mid(PREmployee.LFName, 1, 30)
            rs!EmployeeNumber = PREmployee.EmployeeNumber
            rs!Amount = 0
            rs!EmployeeID = PREmployee.EmployeeID
        End If
        rs!Amount = rs!Amount + PRDist.Amount
        rs.Update
        If PRDist.GetNext = False Then Exit Do
    Loop

    Me.lblTotal = "Now scanning payroll data ... b"
    Me.Refresh
    
    Me.MousePointer = vbArrow

    If rs.RecordCount = 0 Then
        MsgBox "No data found!", vbExclamation
        GoBack
    End If

    ' calc totals - use min amount
    TlCount = 0
    TlAmount = 0
    rs.MoveFirst
    Do
        If rs!Amount >= Me.tdbMinAmount Or Me.tdbMinAmount = 0 Then
            rs!Select = True
        End If
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
    
    SetGrid rs, Me.fg

    CalcTotals
    rs.MoveFirst
    fg.Row = 1

End Sub
Private Sub cmdChkAll_Click()
    MarkRecs True
End Sub

Private Sub cmdClearAll_Click()
    MarkRecs False
End Sub

Private Sub MarkRecs(ByVal sel As Boolean)
    If rs.RecordCount = 0 Then Exit Sub
    i = fg.Row
    rs.MoveFirst
    Do
        rs!Select = sel
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
    CalcTotals
    fg.Row = i
End Sub

Private Sub CalcTotals()

    If rs.RecordCount = 0 Then Exit Sub
    TlCount = 0
    TlAmount = 0
    i = fg.Row
    rs.MoveFirst
    Do
        If rs!Select = True Then
            TlCount = TlCount + 1
            TlAmount = TlAmount + rs!Amount
        End If
        rs.MoveNext
    Loop Until rs.EOF
    Me.lblTotal = "Totals: " & Format(TlCount, "#,##0") & " " & Format(TlAmount, "###,##0.00")
    Me.Refresh
    fg.Row = i

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdPrint_Click()

    ' save the screen settings
    If PRGlobal.GetByID(GlobalID) = True Then
        PRGlobal.Var1 = Me.tdbMinAmount.Value
        PRGlobal.Var2 = Me.tdbHorzNudge.Value
        PRGlobal.Var3 = Me.tdbVertNudge.Value
        PRGlobal.Var4 = Me.cmbItem.ItemData(Me.cmbItem.ListIndex)
        PRGlobal.Save (Equate.RecPut)
    End If
        
    HorzNudge = Me.tdbHorzNudge.Value
    VertNudge = Me.tdbVertNudge.Value
    
    PR1099 Me.cmbTaxYear.Text
    
    GoBack

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    CalcTotals
End Sub


