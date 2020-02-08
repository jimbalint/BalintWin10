VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmGlobalMaint 
   Caption         =   "Windows PR Global Maintenance"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
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
   ScaleHeight     =   8340
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbGlobalType 
      Height          =   360
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&E&XIT"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7095
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      _cx             =   20770
      _cy             =   12515
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
   Begin VB.Label Label1 
      Caption         =   "Select a category:"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmGlobalMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rw As Long
Dim SString As String
Dim SortCol As Byte
Dim SortType As Byte     ' 0=ascending 1=descending
Dim WkcDrop As String

Dim dbFileName As String
Dim dbFields(8) As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte

Dim GlobalType As Byte
Dim FirstFlag As Boolean
Dim SQLString As String
Dim DollarFmt, PercentFmt As String

Dim i, j, k As Long
Dim X, Y, Z As String

Dim CompanyDrop As String

Private Sub Form_Load()
    
    ' get rid of nulls
    rsInit "DELETE * FROM PRGlobal WHERE Isnull(TypeCode)", cnDes, rs
    
    ' drop down for company list
    CompanyDrop = ""
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    SQLString = "SELECT * FROM GLCompany ORDER BY Name"
    rsInit SQLString, cnDes, rs
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
            CompanyDrop = CompanyDrop & "|#" & rs!ID & ";" & rs!Name
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    ' set the constants for the file
    dbFileName = "PRGlobal"
    dbFields(0) = "GlobalID"
    dbFields(1) = "TypeCode"
    dbFields(2) = "Description"
    dbFields(3) = "Amount"
    dbFields(4) = "Percent"
    dbFields(5) = "Year"
    dbFields(6) = "Month"
    dbFields(7) = "Var1"
    dbFields(8) = "Var2"
    dbSortCol = 1
    dbSortDesc = False
    
    With Me.cmbGlobalType
        For i = 1 To 21
            Select Case i
                
                Case 1:     .AddItem "Company Option":      .ItemData(.NewIndex) = PREquate.GlobalTypeCompanyOption
                Case 2:     .AddItem "Education Level":     .ItemData(.NewIndex) = PREquate.GlobalTypeEducationLevel
                Case 3:     .AddItem "EIC Max Advance":     .ItemData(.NewIndex) = PREquate.GlobalTypeEICMaxAdv
                Case 4:     .AddItem "EIC Max Wage":        .ItemData(.NewIndex) = PREquate.GlobalTypeEICMaxWage
                Case 5:     .AddItem "Fed Unemp Max Wage":  .ItemData(.NewIndex) = PREquate.GlobalTypeFUNMax
                Case 6:     .AddItem "Fed Unemp Percnt":    .ItemData(.NewIndex) = PREquate.GlobalTypeFUNPct
                Case 7:     .AddItem "Med Tax Percent":     .ItemData(.NewIndex) = PREquate.GlobalTypeMEDPct
                Case 8:     .AddItem "Med Tax Addl Pct":    .ItemData(.NewIndex) = PREquate.GlobalTypeMEDAddPct
                Case 9:     .AddItem "Med Tax Addl Max$":   .ItemData(.NewIndex) = PREquate.GlobalTypeMEDAddAmt
                Case 10:    .AddItem "OH SD Tax Allowance": .ItemData(.NewIndex) = PREquate.GlobalTypeOHSDTaxAllow
                Case 11:    .AddItem "PR Check Prefix":     .ItemData(.NewIndex) = PREquate.GlobalTypePRCheckPrefix
                Case 12:    .AddItem "Race Code":           .ItemData(.NewIndex) = PREquate.GlobalTypeRaceCode
                Case 13:    .AddItem "SS Tax Max Wage":     .ItemData(.NewIndex) = PREquate.GlobalTypeSSMax
                Case 14:    .AddItem "SS Tax Percent":      .ItemData(.NewIndex) = PREquate.GlobalTypeSSPct
                Case 15:    .AddItem "Shift Code":          .ItemData(.NewIndex) = PREquate.GlobalTypeShiftCode
                Case 16:    .AddItem "Termination Code":    .ItemData(.NewIndex) = PREquate.GlobalTypeTerminationCode
                Case 17:    .AddItem "W2 Box 12":           .ItemData(.NewIndex) = PREquate.GlobalTypeW2Box12
                Case 18:    .AddItem "W2 Box 14":           .ItemData(.NewIndex) = PREquate.GlobalTypeW2Box14
                Case 19:    .AddItem "Work Comp":           .ItemData(.NewIndex) = PREquate.GlobalTypeWkcCat
                Case 20:    .AddItem "Other State ID":      .ItemData(.NewIndex) = PREquate.GlobalTypeOtherStateID
                Case 21:    .AddItem "OH SWT Multiplier":   .ItemData(.NewIndex) = PREquate.GlobalTypeOHMultiplier
            End Select
        
        Next i
    End With
    
    DollarFmt = "$##,###,##0.00"
    PercentFmt = "##0.00"
    
    Me.KeyPreview = True
    FirstFlag = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub cmbGlobalType_Click()
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
        
    GlobalType = Me.cmbGlobalType.ItemData(cmbGlobalType.ListIndex)
    
    GetSQLString
    rsInit GetSQLString, cnDes, rs
    SetGrid rs, fg

    fg.ColHidden(0) = True
    fg.ColHidden(1) = True
    
    fg.ColWidth(7) = 0
    fg.ColWidth(8) = 0
    
    ' adjust grid settings
    If GlobalType = PREquate.GlobalTypeEducationLevel Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeEICMaxAdv Then
        fg.TextMatrix(0, 3) = "EIC Max Advance"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeEICMaxWage Then
        fg.TextMatrix(0, 3) = "EIC Max Wage"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeFUNMax Then
        fg.TextMatrix(0, 3) = "FUN Max Wage"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeFUNPct Then
        fg.TextMatrix(0, 3) = "FUN Percent"
        fg.ColFormat(3) = PercentFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeMEDPct Then
        fg.TextMatrix(0, 3) = "MED Percent"
        fg.ColFormat(3) = PercentFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeMEDAddPct Then
        fg.TextMatrix(0, 3) = "MED Add Pct"
        fg.ColFormat(3) = PercentFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeMEDAddAmt Then
        fg.TextMatrix(0, 3) = "MED Add Theshhold Amt"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeRaceCode Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeSSMax Then
        fg.TextMatrix(0, 3) = "Soc Sec Max Wage"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeSSPct Then
        fg.TextMatrix(0, 3) = "Soc Sec Percent"
        fg.ColFormat(3) = PercentFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeShiftCode Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeTerminationCode Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeW2Box12 Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeW2Box14 Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeWkcCat Then
        fg.TextMatrix(0, 2) = "Work Comp Category"
        fg.TextMatrix(0, 4) = "Work Comp Pct"
        fg.ColFormat(4) = "#0.0000"
        fg.ColHidden(3) = True
        fg.ColHidden(5) = True
        fg.ColHidden(6) = True
        fg.ColWidth(2) = 5000
        fg.ColWidth(4) = 2000
    ElseIf GlobalType = PREquate.GlobalTypePRCheckPrefix Then
        SetGridDescOnly
    ElseIf GlobalType = PREquate.GlobalTypeOHSDTaxAllow Then
        fg.TextMatrix(0, 3) = "OH SD Tax Allowance"
        fg.ColFormat(3) = DollarFmt
        SetGridAmount
    ElseIf GlobalType = PREquate.GlobalTypeCompanyOption Then
        
        fg.ColWidth(3) = 0
        fg.ColWidth(4) = 0
        fg.ColWidth(5) = 0
        fg.ColWidth(6) = 0
        
        fg.ColHidden(2) = False
        fg.TextMatrix(0, 2) = "Company Option"
        fg.ColWidth(2) = 5000
        
        fg.ColHidden(7) = False
        fg.TextMatrix(0, 7) = "Setting"
        fg.ColWidth(7) = 1000
        
        fg.ColHidden(8) = False
        fg.TextMatrix(0, 8) = "Company"
        fg.ColWidth(8) = 3000
        fg.ColComboList(8) = CompanyDrop
    
    ElseIf GlobalType = PREquate.GlobalTypeOtherStateID Then
        
        fg.ColWidth(3) = 0
        fg.ColWidth(4) = 0
        fg.ColWidth(5) = 0
        fg.ColWidth(6) = 0
        
        fg.ColHidden(2) = False
        fg.TextMatrix(0, 2) = "Other State"
        fg.ColWidth(2) = 2000
        
        fg.ColHidden(7) = False
        fg.TextMatrix(0, 7) = "Other State ID"
        fg.ColWidth(7) = 2000
        
        fg.ColHidden(8) = False
        fg.TextMatrix(0, 8) = "Company"
        fg.ColWidth(8) = 5000
        fg.ColComboList(8) = CompanyDrop
    
    Else
        SetGridAmount
    End If

End Sub

Private Sub SetGridAmount()
    
    With fg
        fg.ColHidden(2) = True
        fg.ColWidth(2) = 2000
        fg.ColHidden(3) = False
        fg.ColWidth(3) = 3000
        fg.ColHidden(4) = True
        fg.ColHidden(5) = False
        fg.ColWidth(5) = 1000
        fg.ColHidden(6) = False
        fg.ColWidth(6) = 1000
    End With

End Sub


Private Sub SetGridDescOnly()

    With fg
        fg.ColHidden(2) = False
        fg.ColWidth(2) = 9500
        fg.ColHidden(3) = True
        fg.ColHidden(4) = True
        fg.ColHidden(5) = True
        fg.ColHidden(6) = True
    
        ' set header row
        If GlobalType = PREquate.GlobalTypeRaceCode Then
            fg.TextMatrix(0, 2) = "Enter Race Description"
        ElseIf GlobalType = PREquate.GlobalTypeEducationLevel Then
            fg.TextMatrix(0, 2) = "Enter Education Level"
        ElseIf GlobalType = PREquate.GlobalTypeShiftCode Then
            fg.TextMatrix(0, 2) = "Enter Shift Code"
        ElseIf GlobalType = PREquate.GlobalTypeTerminationCode Then
            fg.TextMatrix(0, 2) = "Enter Termination Code"
        ElseIf GlobalType = PREquate.GlobalTypeW2Box12 Then
            fg.TextMatrix(0, 2) = "W2 Box 12 Code"
        ElseIf GlobalType = PREquate.GlobalTypeW2Box14 Then
            fg.TextMatrix(0, 2) = "W2 Box 14 Code"
        End If
    
    End With

End Sub

Private Sub cmdAdd_Click()
    ' AddAdo rs, fg
    rs.AddNew
    rs!TypeCode = GlobalType
    rs!Description = ""
    rs!Amount = 0
    rs!Percent = 0
    rs!Year = 0
    rs!Month = 0
    rs!Var1 = ""
    rs!Var2 = ""
    rs.Update
    fg.DataRefresh
End Sub


Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

'Dim CurrID As Long
'
'    CurrID = fg.TextMatrix(fg.Row, 0)
'
'    rs.Close
'    rsInit GetSQLString, cn, rs
'    Set fg.DataSource = rs.DataSource
'
'    rw = fg.FindRow(CurrID, 0, 0)
'
'    fg.TopRow = rw
'    fg.Select rw, 0
'    fg.SetFocus
    
    
End Sub

Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
'    If Col = 1 Then     ' validates for number - must enter a value that does not already exist
'
'       If fg.EditText = "" Or fg.EditText = "0" Then
'          MsgBox "Enter a value!", vbExclamation + vbOKOnly
'          Cancel = True
'       End If
'
'       rw = fg.FindRow(fg.EditText, 0, 1)
'       If rw <> -1 Then
'          MsgBox "Number already exists!", vbExclamation + vbOKOnly
'          Cancel = True
'       End If
'
'    End If
'
'    If Col = 2 Then
'        fg.EditText = Trim(UCase(fg.EditText))
'    End If
    
End Sub

Private Sub cmdExit_Click()
    
    GoBack

End Sub

Private Sub cmdDelete_Click()
    
Dim DelConfirm As Integer
    
    If fg.Rows = 1 Then Exit Sub
    
    DelConfirm = MsgBox(Trim(fg.TextMatrix(fg.Row, 2)), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")

    If DelConfirm = vbNo Then
       fg.SetFocus
       Exit Sub
    End If

    rw = fg.Row
    rs.Delete
    fg.DataRefresh
    If rw = fg.Rows Then rw = fg.Rows - 1

    fg.Select rw, 0
    fg.ShowCell rw, 0

End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

'    ' clicking on a column header sorts based on that column
'    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then
'
'       ' toggle the sort order
'       If fg.MouseCol = dbSortCol Then
'          If dbSortDesc = False Then
'             dbSortDesc = True
'          Else
'             dbSortDesc = False
'          End If
'       Else
'          ' switch the column
'          fg.Cell(flexcpFontBold, 0, fg.MouseCol) = True
'          fg.Cell(flexcpFontBold, 0, dbSortCol) = False
'          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol)
'          dbSortCol = fg.MouseCol
'       End If
'
'       If dbSortDesc Then
'          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "-"
'       Else
'          fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "+"
'       End If
'
'       rs.Close
'
'       rsInit GetSQLString, cn, rs
'       Set fg.DataSource = rs.DataSource
'
'       fg.ShowCell 1, 1
'
'    End If
    
End Sub

Private Function GetSQLString() As String
    
Dim aa As Integer
    
' set the SQL string
'    x = "SELECT [Number],[Description] " & _
'        "FROM GLDescriptions ORDER BY [Number] DESC"

    GetSQLString = "SELECT "
    
    For aa = 0 To UBound(dbFields, 1)
        GetSQLString = GetSQLString & " [" & dbFields(aa) & "]"
        If aa <> UBound(dbFields, 1) Then GetSQLString = GetSQLString & ","
        GetSQLString = GetSQLString & " "
    Next aa
    
    ' GetSQLString = GetSQLString & "FROM " & dbFileName & " ORDER BY [" & dbFields(dbSortCol) & "]"
    
    GetSQLString = GetSQLString & " FROM PRGlobal WHERE TypeCode = " & GlobalType
    
'    If dbSortDesc Then
'       GetSQLString = GetSQLString & " DESC"
'    End If

End Function






