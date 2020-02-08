VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   14370
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   9015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   12495
      _cx             =   22040
      _cy             =   15901
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCol As New ADODB.Recordset
Dim RowNum, ColNum As Integer
Dim I, J, K As Long
Dim X, Y, Z As String

Private Sub Form_Load()

    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    rsCol.Open , , adOpenDynamic, adLockOptimistic
    
    ' columns for the matrix
    AddCol "Descr", "Descr", 5000
    
    I = 1800
    AddCol "Amt1", "Amt1", I
    AddCol "Amt2", "Amt2", I
    AddCol "Amt3", "Amt3", I
    
    I = 300
    AddCol "Edit1", "Edit1", I, adBoolean
    AddCol "Edit2", "Edit2", I, adBoolean
    AddCol "Edit3", "Edit3", I, adBoolean
    AddCol "Show1", "Show1", I, adBoolean
    AddCol "Show2", "Show2", I, adBoolean
    AddCol "Show3", "Show3", I, adBoolean

    RowNum = 0
    
    With Me.fg
    
        .FixedRows = 0
        .FixedCols = 0
        .Cols = 10
        .Rows = 99
        .Editable = flexEDKbdMouse
        
        I = 0
        rsCol.MoveFirst
        Do
            .ColWidth(I) = rsCol!Width
            .ColData(I) = rsCol!Abbrev
            If rsCol!DataType <> 0 Then
                .ColDataType(I) = rsCol!DataType
            End If
            If rsCol!Format <> 0 Then
                .ColFormat(I) = rsCol!Format
            End If
            I = I + 1
            rsCol.MoveNext
        Loop Until rsCol.EOF
    
        AddRow " 2 ) Wages, tips and other compensation", False, False, True, False, False, True
        AddRow " 3 ) Income tax withheld", False, False, True, False, False, True
        
        AddRow " 5a) Taxable social security wages", True, True, False, True, True, False
        AddRow " 5b) Taxable social security tips", True, True, False, True, True, False
        AddRow " 5c) Taxable Medicare wages & tips", True, True, False, True, True, False
        
        AddRow " 5d) Add Col 2 5a,Col 2 5b, Col 2 5c", False, False, False, False, False, True
        AddRow " 5e) Sec 3121(q) Notice and Demand-Tax due on unreported tips", False, False, True, False, False, True
        
        AddRow " 6e) Total taxes before adjustments", False, False, False, False, False, True
        AddRow " 7 ) Current qtr adj for fractions of cents", False, True, False, False, True, False
        AddRow " 8 ) Current qtr adj for sick pay", False, True, False, False, True, False
        AddRow " 9 ) Current qtr adj for tips and group-term life insurance", False, True, False, False, True, False
        
        AddRow "10 ) Total taxes after adjustments", False, False, False, False, False, True
        AddRow "11 ) Total deposits, incl prior qtr overpay", False, False, True, False, False, True
        
        AddRow "12a) COBRA premium asst payments", True, False, False, True, False, False
        AddRow "12b) Number of COBRA provided", True, False, False, True, False, False
        
        AddRow "13 ) Add lines 11 and 12a", False, False, False, False, False, True
        AddRow "14 ) Balance Due", False, False, False, False, False, True
        AddRow "15 ) Overpayment", False, False, False, False, True, False
        
        .Rows = RowNum
    
        .ColFormat(GetCol("Amt1")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt2")) = "##,###,##0.00-"
        .ColFormat(GetCol("Amt3")) = "##,###,##0.00-"
    
        ' color the grid
        For I = 1 To RowNum
            For J = 1 To 3
                If J = 1 Then K = GetCol("Show1")
                If J = 2 Then K = GetCol("Show2")
                If J = 3 Then K = GetCol("Show3")
                If .TextMatrix(I - 1, K) = "False" Then
                    .Select I - 1, K - 6
                    .CellBackColor = RGB(192, 192, 192)
                    .CellBackColor = RGB(100, 100, 100)
                End If
            Next J
        Next I
    
    End With

    Set941Val "5a", 2, 1234.56

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' can't edit control columns
    If Col = GetCol("Descr") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit1") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit2") Then Cancel = True: Exit Sub
    If Col = GetCol("Edit3") Then Cancel = True: Exit Sub
    If Col = GetCol("Show1") Then Cancel = True: Exit Sub
    If Col = GetCol("Show2") Then Cancel = True: Exit Sub
    If Col = GetCol("Show3") Then Cancel = True: Exit Sub
    
    ' flagged as not editable
    If fg.TextMatrix(Row, Col + 3) = "False" Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub AddRow(ByVal Title As String, _
                   ByVal Edit1 As Boolean, _
                   ByVal Edit2 As Boolean, _
                   ByVal Edit3 As Boolean, _
                   ByVal Show1 As Boolean, _
                   ByVal Show2 As Boolean, _
                   ByVal Show3 As Boolean)
                   
    With Me.fg
        
        .TextMatrix(RowNum, GetCol("Descr")) = Title
        
        .TextMatrix(RowNum, GetCol("Edit1")) = Edit1
        .TextMatrix(RowNum, GetCol("Edit2")) = Edit2
        .TextMatrix(RowNum, GetCol("Edit3")) = Edit3
        
        .TextMatrix(RowNum, GetCol("Show1")) = Show1
        .TextMatrix(RowNum, GetCol("Show2")) = Show2
        .TextMatrix(RowNum, GetCol("Show3")) = Show3
    
    End With
    
    RowNum = RowNum + 1
                   
End Sub

Private Sub AddCol(ByVal Title As String, _
                   ByVal Abbrev As String, _
                   ByVal Width As Long, _
                   Optional DType As Byte, _
                   Optional Fmt As String)

    rsCol.AddNew
    rsCol!Title = Mid(Title, 1, 30)
    rsCol!Abbrev = Mid(Abbrev, 1, 30)
    rsCol!Width = Width
    rsCol!Number = ColNum
    rsCol!DataType = DType
    rsCol!Format = Fmt
    rsCol.Update
    
    ColNum = ColNum + 1

End Sub

Private Function GetCol(ByVal ColData As String) As Long

    SQLString = "Abbrev = '" & ColData & "'"
    rsCol.Find SQLString, 0, adSearchForward, 1
    If rsCol.EOF Then
        GetCol = -1
    Else
        GetCol = rsCol!Number
    End If

End Function

Private Sub Set941Val(ByVal RowKey As String, ByVal ColNum As Byte, Amt As Currency)

    Dim fgRow, fgCol, fgI As Integer

    With Me.fg
        
        fgRow = .Row
        fgCol = .Col
        
        If ColNum < 1 Or ColNum > 3 Then
            MsgBox "Invalid ColNum: " + ColNum, vbExclamation
            End
        End If
        
        For fgI = 1 To .Rows
            If InStr(1, .TextMatrix(fgI - 1, GetCol("Descr")), RowKey, vbTextCompare) Then
                Exit For
            End If
        Next fgI
        
        If fgI = .Rows + 1 Then
            MsgBox "Row Key NF: " + RowKey, vbExclamation
            End
        End If
        
        .TextMatrix(fgI - 1, GetCol("Amt" & ColNum)) = Amt
    
    End With

End Sub
