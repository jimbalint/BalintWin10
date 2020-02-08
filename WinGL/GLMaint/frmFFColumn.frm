VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFFColumn 
   Caption         =   "GL Free Format Column Setup"
   ClientHeight    =   8850
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
   ScaleHeight     =   8850
   ScaleWidth      =   13455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Cancel          =   -1  'True
      Caption         =   "R&ESET"
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&RENAME"
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&COPY"
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&NEW"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbFFColumn 
      Height          =   360
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   4815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   11760
      TabIndex        =   2
      Top             =   8160
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6135
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   12615
      _cx             =   22251
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
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "frmFFColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim H, I, J, K As Long
Dim x, Y, Z As String
Public FFName As String
Dim GlobID As Long
Dim Ct1, Ct2, Ct3 As Long
Dim Fmt As String
Dim EditFlag As Boolean
Dim cmbPrev As String
Dim rsCol As New ADODB.Recordset
Dim ColNum As Integer
Dim LoadFlag As Boolean
Dim boo As Boolean

Dim TypeDrop As String
Dim FYDrop As String
Dim ColDrop As String
Dim PeriodDrop As String

Private Sub Form_Load()

    ' *** to do ***
    '  move column - renumber
    '  add - columns inclusive
    ' *************

    LoadFlag = True
    
    Me.lblCompanyName = GLCompany.Name

    ' drop from the company MDB file if necessary
    If TableExists("GLColumn", cn) = True Then
        cn.Execute "DROP TABLE GLColumn"
    End If
    
    ' create table if necessary IN PRGLOBAL
    If TableExists("GLFFColumn", cnDes) = False Then
        FFColumnCreate
    End If

    GridInit
    LoadCmb
    
    LoadFlag = False
        
    With Me
        .cmdSave.ToolTipText = "SAVE the current column definition set"
        .cmdNew.ToolTipText = "CREATE a NEW column definition set"
        .cmdDelete.ToolTipText = "DELETE the current column definition set"
        .cmdCopy.ToolTipText = "COPY the current column definition set"
        .cmdRename.ToolTipText = "RENAME the current column definition set"
        .cmdReset.ToolTipText = "RESET the current column definition set to last saved"
        .cmbFFColumn.ToolTipText = "Open an existing or create new column definition set"
        .KeyPreview = True
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    If EditFlag = True Then
        If MsgBox("Save changes?", vbQuestion + vbYesNo) = vbYes Then
            cmdSave_Click
        End If
    End If
    GoBack
End Sub
Private Sub cmdSave_Click()
    SaveColumns
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    EditFlag = True

End Sub
Private Sub fg_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
'    With fg
'        If OldCol = GetCol("Type") Then
'            If .TextMatrix(OldRow, GetCol("Type")) = Equate.ColYTD Then
'                .Cell(flexcpBackColor, OldRow, GetCol("StartNum"), OldRow, GetCol("EndNum")) = RGB(192, 192, 192)
'                .TextMatrix(OldRow, GetCol("StartNum")) = "N/A"
'            End If
'        End If
'    End With

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' set the list for the Start / End Value columns
    ' based on the Type
    With fg
    
        If Col = GetCol("StartNum") Or Col = GetCol("EndNum") Then
        
            Select Case .TextMatrix(Row, GetCol("Type"))
                Case 0:                     Cancel = True
                Case Equate.ColAdd:         .ComboList = ColDrop
                Case Equate.ColSubtract:    .ComboList = ColDrop
                Case Equate.ColDivide:      .ComboList = ColDrop
                Case Equate.ColMultiply:    .ComboList = ColDrop
                Case Equate.ColAvg:         .ComboList = ColDrop
                Case Equate.ColProj:        .ComboList = ColDrop
                Case Equate.ColCurrPd:      Cancel = True
                Case Equate.ColPriorPd:     Cancel = True
                Case Equate.ColYTD:         Cancel = True
                Case Equate.ColAllPd:       Cancel = True
                Case Equate.ColCustom:      .ComboList = PeriodDrop
            End Select
        
        Else
        
            .ComboList = ""
        
        End If
    
    End With

End Sub


Private Sub GridInit()

    TypeDrop = "|#0; "
    TypeDrop = TypeDrop & "|#" & Equate.ColAdd & ";ADD"
    TypeDrop = TypeDrop & "|#" & Equate.ColSubtract & ";SUBTRACT"
    TypeDrop = TypeDrop & "|#" & Equate.ColDivide & ";PERCENT"
    TypeDrop = TypeDrop & "|#" & Equate.ColMultiply & ";MULTIPLY"
    TypeDrop = TypeDrop & "|#" & Equate.ColAvg & ";AVERAGE"
    TypeDrop = TypeDrop & "|#" & Equate.ColProj & ";PROJECTED"
    TypeDrop = TypeDrop & "|#" & Equate.ColCurrPd & ";CURR PD"
    TypeDrop = TypeDrop & "|#" & Equate.ColPriorPd & ";PRIOR PD"
    TypeDrop = TypeDrop & "|#" & Equate.ColYTD & ";YTD"
    TypeDrop = TypeDrop & "|#" & Equate.ColAllPd & ";ALL PDS"
    TypeDrop = TypeDrop & "|#" & Equate.ColCustom & ";CUSTOM"
    
    FYDrop = "|#0;Curr Yr"
    For I = 1 To 10
        FYDrop = FYDrop & "|#" & I & ";Prior Yr " & I
    Next I

    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    rsCol.Open , , adOpenDynamic, adLockOptimistic


    With fg
            
        ColNum = 0
        AddCol "", "Col0", 0
        AddCol "Column #", "ColNum", 1000
        AddCol "Description", "Desc", 4000
        AddCol "Type", "Type", 1000
        AddCol "Fiscal Year", "FY", 1300
        AddCol "Start Value", "StartNum", 1300
        AddCol "End Value", "EndNum", 1300
        AddCol "Budget", "Budget", 1000, flexDTBoolean
        AddCol "Print Tab", "Tab", 0
        AddCol "Non Print", "NonPrt", 1000, flexDTBoolean
        AddCol "FFColumnID", "ID", 0
        
        .Rows = 21
        .Cols = rsCol.RecordCount
        
        .FixedRows = 1
        .FixedCols = 1
        
        .ExplorerBar = flexExMoveRows
        .AllowBigSelection = False
        .Editable = flexEDKbdMouse
            
        I = 0
        rsCol.MoveFirst
        Do
            .TextMatrix(0, I) = rsCol!Title
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
    
        .ColComboList(GetCol("Type")) = TypeDrop
        .ColComboList(GetCol("FY")) = FYDrop
    
    End With

    ColDrop = "|N / A"
    For I = 1 To 20
        ColDrop = ColDrop & "|Column " & I
    Next I
    
    PeriodDrop = ""
    For I = 1 To 13
        PeriodDrop = PeriodDrop & "|Period " & I
    Next I

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

Private Sub LoadCmb()
    
    ' init the combo of FFSched definitions
    With Me.cmbFFColumn
        .AddItem "<Add New>"
        .ItemData(.NewIndex) = 0
        
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFColumn & _
                    " AND UserID = " & GLCompany.ID & _
                    " ORDER BY Description"
        
        ' *** Global - for ALL companies ***
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFColumn & _
                    " ORDER BY Description"
        
        If PRGlobal.GetBySQL(SQLString) = True Then
            Do
                .AddItem PRGlobal.Description
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                If PRGlobal.GetNext = False Then Exit Do
            Loop
        End If
    End With

End Sub

Private Sub cmbFFColumn_Click()

    If LoadFlag = True Then Exit Sub
    
    ' save grid ???
    If EditFlag = True Then
        If MsgBox("Save changes to: " & cmbPrev, vbQuestion + vbYesNo) = vbYes Then
            cmdSave_Click
        End If
    End If
    
    With Me.cmbFFColumn
        
        ' add new
        If .ListIndex = 0 Then
            x = InputBox("Title for free format columns", _
                       "NEW Free Format Columns")
            If x <> "" Then
                SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFColumn & _
                            " AND UserID = " & GLCompany.ID & _
                            " AND Description = '" & x & "'"
                If PRGlobal.GetBySQL(SQLString) = True Then
                    MsgBox "This title already exists!", vbExclamation
                    Exit Sub
                End If
                
                PRGlobal.Clear
                PRGlobal.TypeCode = PREquate.GlobalTypeGLFFColumn
                PRGlobal.UserID = GLCompany.ID
                PRGlobal.Description = x
                PRGlobal.Save (Equate.RecAdd)
                
                .AddItem x
                .ItemData(.NewIndex) = PRGlobal.GlobalID
                .ListIndex = .ListCount - 1
                
            End If
        End If
        
        GlobID = .ItemData(.ListIndex)
    
    End With
    
    FFName = Me.cmbFFColumn
    LoadColumns
    cmbPrev = Me.cmbFFColumn

    Me.Show
    With fg
        If .Rows > 1 Then
            .Select 1, 2
        End If
    End With
    Me.Show
    fg.SetFocus
    


End Sub

Private Sub LoadColumns()

    For I = 1 To 20
        SQLString = "SELECT * FROM GLFFColumn WHERE GlobalID = " & GlobID & _
                    " AND ColNum = " & I
        
        With fg
        
            If GLFFColumn.GetBySQL(SQLString) = False Then
                GLFFColumn.Clear
                GLFFColumn.ColNum = I
                GLFFColumn.GlobalID = GlobID
                GLFFColumn.Save (Equate.RecAdd)
            End If
            
            .TextMatrix(I, GetCol("ColNum")) = I
            .TextMatrix(I, GetCol("Desc")) = GLFFColumn.Description
            .TextMatrix(I, GetCol("Type")) = GLFFColumn.ColType
            .TextMatrix(I, GetCol("FY")) = GLFFColumn.FiscalYear
            
            K = .TextMatrix(I, GetCol("Type"))
            x = ""
            If K = Equate.ColCustom Then
                x = "Period " & GLFFColumn.StartNum
            ElseIf K = Equate.ColAdd Or _
               K = Equate.ColSubtract Or _
               K = Equate.ColDivide Or _
               K = Equate.ColMultiply Or _
               K = Equate.ColAvg Or _
               K = Equate.ColProj Then
                    If GLFFColumn.StartNum = 99 Then
                        x = "N / A"
                    Else
                        x = "Column " & GLFFColumn.StartNum
                    End If
            End If
            .TextMatrix(I, GetCol("StartNum")) = x
            
            x = ""
            If K = Equate.ColCustom Then
                x = "Period " & GLFFColumn.EndNum
            ElseIf K = Equate.ColAdd Or _
               K = Equate.ColSubtract Or _
               K = Equate.ColDivide Or _
               K = Equate.ColMultiply Or _
               K = Equate.ColAvg Or _
               K = Equate.ColProj Then
                    If GLFFColumn.EndNum = 99 Then
                        x = "N / A"
                    Else
                        x = "Column " & GLFFColumn.EndNum
                    End If
            End If
            .TextMatrix(I, GetCol("EndNum")) = x
            
            .TextMatrix(I, GetCol("Budget")) = GLFFColumn.Budget
            .TextMatrix(I, GetCol("Tab")) = GLFFColumn.PrintTab
            .TextMatrix(I, GetCol("NonPrt")) = GLFFColumn.NonPrint
            .TextMatrix(I, GetCol("ID")) = GLFFColumn.FFColumnID
                        
        End With
    
    Next I
    
End Sub

Private Sub SaveColumns()

    For I = 1 To 20

        With fg
            
            If GLFFColumn.GetByID(CLng(.TextMatrix(I, GetCol("ID")))) = False Then
                MsgBox "GLFFColumn NF: " & .TextMatrix(I, GetCol("ID")), vbExclamation
                GoBack
            End If
            
            GLFFColumn.ColNum = I
            GLFFColumn.Description = Mid(.TextMatrix(I, GetCol("Desc")), 1, 30)
            GLFFColumn.ColType = .TextMatrix(I, GetCol("Type"))
            GLFFColumn.FiscalYear = .TextMatrix(I, GetCol("FY"))
            
            GLFFColumn.StartNum = tmxConvert(.TextMatrix(I, GetCol("StartNum")))
            GLFFColumn.EndNum = tmxConvert(.TextMatrix(I, GetCol("EndNum")))
            
            If .TextMatrix(I, GetCol("Budget")) = 0 Then
                GLFFColumn.Budget = 0
            Else
                GLFFColumn.Budget = 1
            End If
            
            GLFFColumn.PrintTab = .TextMatrix(I, GetCol("Tab"))
                        
            If .TextMatrix(I, GetCol("NonPrt")) = 0 Then
                GLFFColumn.NonPrint = 0
            Else
                GLFFColumn.NonPrint = 1
            End If
            
            GLFFColumn.Save (Equate.RecPut)
            
        End With
    
    Next I
    
    EditFlag = False

End Sub

Private Function tmxConvert(ByVal str As String) As Long

    tmxConvert = 0
    If IsNull(str) Then Exit Function
    If str = "" Then Exit Function
    If str = "N / A" Then
        tmxConvert = 99
        Exit Function
    End If
    If Len(str) < 8 Then Exit Function
    
    If IsNumeric(Mid(str, 8, 2)) Then
        tmxConvert = CLng(Mid(str, 8, 2))
    End If

End Function

Private Sub cmdNew_Click()
    
    x = InputBox("Name for new Free Format Columns:", "GL Free Format Columns")
    If x = "" Then Exit Sub
    
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFColumn
    PRGlobal.UserID = GLCompany.ID
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecAdd)
    
    GlobID = PRGlobal.GlobalID
    
    With Me.cmbFFColumn
        .AddItem x
        .ItemData(.NewIndex) = GlobID
        .ListIndex = .ListCount - 1
    End With

End Sub

Private Sub cmdDelete_Click()

    If Me.cmbFFColumn.ListIndex <= 0 Then Exit Sub
    
    If MsgBox("OK to delete the ENTIRE Free Format Columns " & vbCr & vbCr & _
              Me.cmbFFColumn & "?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
              
    EditFlag = False
    
    SQLString = "DELETE * FROM GLFFColumn WHERE GlobalID = " & GlobID
    cnDes.Execute SQLString
    
    SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & GlobID
    cnDes.Execute SQLString
    
    LoadFlag = True
    Me.cmbFFColumn.Clear
    LoadCmb
    
    LoadFlag = False

    With Me.cmbFFColumn
        If .ListCount > 1 Then
            .ListIndex = 1
        Else
            .ListIndex = 0
        End If
        GlobID = .ItemData(.ListIndex)
    End With

End Sub

Private Sub cmdCopy_Click()

    x = InputBox("Copy " & Me.cmbFFColumn & " to: ", "Free Format Columns")
    If x = "" Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFColumn
    PRGlobal.UserID = GLCompany.ID
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecAdd)
    
    GlobID = PRGlobal.GlobalID
    
    ' >>>>>>> copy column data <<<<<<<<<<
    For I = 1 To 20
        With fg
            GLFFColumn.Clear
            GLFFColumn.Description = .TextMatrix(I, GetCol("Desc")) & ""
            GLFFColumn.ColType = nNull(.TextMatrix(I, GetCol("Type")))
            GLFFColumn.FiscalYear = nNull(.TextMatrix(I, GetCol("FY")))
            
            Y = .TextMatrix(I, GetCol("StartNum")) & ""
            If InStr(1, Y, "Column", vbTextCompare) Then
                GLFFColumn.StartNum = CByte(Mid(Y, Len(Y) - 1, 2))
            Else
                GLFFColumn.StartNum = 0
            End If
            
            Y = .TextMatrix(I, GetCol("EndNum")) & ""
            If InStr(1, Y, "Column", vbTextCompare) Then
                GLFFColumn.EndNum = CByte(Mid(Y, Len(Y) - 1, 2))
            Else
                GLFFColumn.EndNum = 0
            End If
            
            GLFFColumn.Budget = .TextMatrix(I, GetCol("Budget"))
            GLFFColumn.PrintTab = nNull(.TextMatrix(I, GetCol("Tab")))
            GLFFColumn.NonPrint = .TextMatrix(I, GetCol("NonPrt"))
            GLFFColumn.ColNum = I
            GLFFColumn.GlobalID = GlobID
            GLFFColumn.Save (Equate.RecAdd)
        End With
    Next I

    With Me.cmbFFColumn
        LoadFlag = True
        .AddItem x
        .ItemData(.NewIndex) = GlobID
        LoadFlag = False
        .ListIndex = .ListCount - 1
    End With

    LoadColumns

    Me.MousePointer = vbArrow

End Sub

Private Sub cmdRename_Click()

    x = InputBox("Rename " & Me.cmbFFColumn & " to: ", "Free Format Columns Rename")
    If x = "" Then Exit Sub
    
    boo = PRGlobal.GetByID(GlobID)
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecPut)
    
    I = Me.cmbFFColumn.ListIndex
    
    LoadFlag = True
    Me.cmbFFColumn.Clear
    LoadCmb
    Me.cmbFFColumn.ListIndex = I
    LoadFlag = False

End Sub

Private Sub cmdReset_Click()

    If EditFlag = True Then
        If MsgBox("OK to discard all changes?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    EditFlag = False
    LoadColumns

End Sub


