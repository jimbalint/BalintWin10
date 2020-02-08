VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPWMaint 
   Caption         =   "Prevailing Wage by Department"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
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
   ScaleHeight     =   10305
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DEL"
      Height          =   735
      Left            =   9720
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   735
      Left            =   9720
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   9360
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5895
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   9015
      _cx             =   15901
      _cy             =   10398
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
   Begin VB.ComboBox cmbCounty 
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1920
      Width           =   4455
   End
   Begin VB.ComboBox cmbUnion 
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.ComboBox cmbWorkCraft 
      Height          =   360
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   6480
      TabIndex        =   0
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "County:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Union Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Craft:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmPWMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsWorkCat As New ADODB.Recordset
Dim rsCounty As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim GlobID As Long
Dim DropData As Long
Dim LoadFlag As Boolean

Dim i, j, k As Long
Dim X, Y, z As String

Private Sub Form_Load()

    ' stored individual entries in PRGlobal
    '   Var1 = Work Class
    '   Var2 = Union
    '   Var3 = County
    '   Var4 = Prev Wge Rate
    '   Var5 = OT Rate
    '   Var6 = Fringe Rate
    
    ' WorkClass and Union categories in PRGlobal
    ' County list in PRCounty

    LoadCombo Me.cmbWorkCraft, PREquate.GlobalTypePWCraft
    LoadCombo Me.cmbUnion, PREquate.GlobalTypePWUnion
    LoadCombo Me.cmbCounty, PREquate.GlobalTypePWCounty

    Me.fg.Visible = False
    Me.cmdSave.Enabled = False

    LoadFlag = True
    GlobID = 0
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeScreenDefault & _
                " AND UserID = " & User.ID & _
                " AND Description = 'PWMaint'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        GlobID = PRGlobal.GlobalID
        cmbPoint Me.cmbWorkCraft, PRGlobal.Var1
        cmbPoint Me.cmbUnion, PRGlobal.Var2
        cmbPoint Me.cmbCounty, PRGlobal.Var3
    End If
    LoadFlag = False
    
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdLoad_Click()

    ' all dropdowns must be selected
    If cmbVerify(Me.cmbWorkCraft) = False Then Exit Sub
    If cmbVerify(Me.cmbUnion) = False Then Exit Sub
    If cmbVerify(Me.cmbCounty) = False Then Exit Sub

    fg.Visible = True
    cmdSave.Enabled = True
    Me.cmbWorkCraft.Enabled = False
    Me.cmbUnion.Enabled = False
    Me.cmbCounty.Enabled = False

    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Classification", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "Total_PWR", adCurrency
    rs.Fields.Append "Overtime_Rate", adCurrency
    rs.Fields.Append "Fringe_Amount", adCurrency
    rs.Fields.Append "GlobalID", adDouble           ' keep this column last
    rs.Open , , adOpenDynamic, adLockOptimistic

    ' get the data from PRGlobal
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePWWage & _
                " AND Var1 = '" & Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex) & "'" & _
                " AND Var2 = '" & Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex) & "'" & _
                " AND Var3 = '" & Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex) & "'" & _
                " ORDER BY Description"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Do
            rs.AddNew
            rs!Classification = PRGlobal.Description
            rs!Total_PWR = PRGlobal.Var4
            rs!Overtime_Rate = PRGlobal.Var5
            rs!Fringe_Amount = PRGlobal.Var6
            rs!GlobalID = PRGlobal.GlobalID
            rs.Update
            
            ' clear the field
            PRGlobal.Byte10 = 0
            PRGlobal.Save (Equate.RecPut)
            
            If PRGlobal.GetNext = False Then Exit Do
        Loop
    End If
    
    SetGrid rs, fg
    
    With fg
        .ColWidth(fg.Cols - 1) = 0        ' don't show the last column
        .ColWidth(0) = 3500
        For i = 1 To 3
            .ColWidth(i) = 1470
        Next i
    End With

End Sub

Private Function cmbVerify(ByRef cmb As ComboBox) As Boolean

    cmbVerify = False
    
    With cmb
    
        If .ListIndex = -1 Then
            MsgBox "All selections must be chosen to load!", vbExclamation
            Exit Function
        End If
        
        If .ItemData(.ListIndex) = 0 Then
            MsgBox "All selections must be chosen to load!", vbExclamation
            Exit Function
        End If
    
    End With

    cmbVerify = True

End Function

Private Sub cmdSave_Click()

    fg.Visible = False
    cmdSave.Enabled = False
    Me.cmbWorkCraft.Enabled = True
    Me.cmbUnion.Enabled = True
    Me.cmbCounty.Enabled = True

    ' clear the Desc field for all
    ' will then know if any were deleted in the grid
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
        
            ' was added this session
            If IsNull(rs!GlobalID) Or rs!GlobalID = 0 Then
                ' add to prglobal - set Byte10 to 1
                PRGlobal.Clear
                PRGlobal.TypeCode = PREquate.GlobalTypePWWage
                PRGlobal.Description = rs!Classification
                PRGlobal.Var1 = Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex)
                PRGlobal.Var2 = Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex)
                PRGlobal.Var3 = Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex)
                PRGlobal.Var4 = rs!Total_PWR
                PRGlobal.Var5 = rs!Overtime_Rate
                PRGlobal.Var6 = rs!Fringe_Amount
                PRGlobal.Byte10 = 1
                PRGlobal.Save (Equate.RecAdd)
            Else
                ' find the prglobal record
                ' update and set byte10 to 1
                If PRGlobal.GetByID(rs!GlobalID) = False Then
                    MsgBox "Global record not found - " & rs!GlobalID, vbExclamation
                    GoBack
                End If
                PRGlobal.Var1 = Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex)
                PRGlobal.Var2 = Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex)
                PRGlobal.Var3 = Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex)
                PRGlobal.Var4 = rs!Total_PWR
                PRGlobal.Var5 = rs!Overtime_Rate
                PRGlobal.Var6 = rs!Fringe_Amount
                PRGlobal.Byte10 = 1
                PRGlobal.Save (Equate.RecPut)
            End If
            
            rs.MoveNext
        
        Loop Until rs.EOF
    
    End If
    
    ' delete all prglobal where byte10 = 0
    SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypePWWage & _
                " AND Var1 = '" & Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex) & "'" & _
                " AND Var2 = '" & Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex) & "'" & _
                " AND Var3 = '" & Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex) & "'" & _
                " AND Byte10 = 0"
    cnDes.Execute SQLString

End Sub
Private Sub cmdCancel_Click()
    
    fg.Visible = False
    cmdSave.Enabled = False
    Me.cmbWorkCraft.Enabled = True
    Me.cmbUnion.Enabled = True
    Me.cmbCounty.Enabled = True

End Sub

Private Sub cmdExit_Click()
    
    ' save the last three combo items used
    If Me.cmbWorkCraft.ListIndex = -1 Then
    ElseIf Me.cmbUnion.ListIndex = -1 Then
    ElseIf Me.cmbCounty.ListIndex = -1 Then
    Else
        If GlobID = 0 Then
            PRGlobal.Clear
            PRGlobal.TypeCode = PREquate.GlobalTypeScreenDefault
            PRGlobal.Description = "PWMaint"
            PRGlobal.UserID = User.ID
            PRGlobal.Var1 = Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex)
            PRGlobal.Var2 = Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex)
            PRGlobal.Var3 = Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex)
            PRGlobal.Save (Equate.RecAdd)
        Else
            If PRGlobal.GetByID(GlobID) Then
                PRGlobal.Var1 = Me.cmbWorkCraft.ItemData(Me.cmbWorkCraft.ListIndex)
                PRGlobal.Var2 = Me.cmbUnion.ItemData(Me.cmbUnion.ListIndex)
                PRGlobal.Var3 = Me.cmbCounty.ItemData(Me.cmbCounty.ListIndex)
                PRGlobal.Save (Equate.RecPut)
            End If
        End If
    
    End If
    
    GoBack

End Sub

Private Sub cmbWorkCraft_Click()
    If LoadFlag = True Then Exit Sub
    ClickCombo Me.cmbWorkCraft, PREquate.GlobalTypePWCraft
    If DropData <> 0 Then Exit Sub
    If frmPWAdd.OK = True Then
        LoadFlag = True
        LoadCombo Me.cmbWorkCraft, PREquate.GlobalTypePWCraft
        PointCombo Me.cmbWorkCraft, frmPWAdd.txtEntry
        LoadFlag = False
    End If
End Sub

Private Sub cmbUnion_Click()
    If LoadFlag = True Then Exit Sub
    ClickCombo Me.cmbUnion, PREquate.GlobalTypePWUnion
    If DropData <> 0 Then Exit Sub
    If frmPWAdd.OK = True Then
        LoadFlag = True
        LoadCombo Me.cmbUnion, PREquate.GlobalTypePWUnion
        PointCombo Me.cmbUnion, frmPWAdd.txtEntry
        LoadFlag = False
    End If
End Sub
Private Sub cmbCounty_Click()
    If LoadFlag = True Then Exit Sub
    ClickCombo Me.cmbCounty, PREquate.GlobalTypePWCounty
    If DropData <> 0 Then Exit Sub
    If frmPWAdd.OK = True Then
        LoadFlag = True
        LoadCombo Me.cmbCounty, PREquate.GlobalTypePWCounty
        PointCombo Me.cmbCounty, frmPWAdd.txtEntry
        LoadFlag = False
    End If
End Sub

Private Sub PointCombo(ByRef cmb As ComboBox, ByVal ComboString As String)

Dim CmbFlg As Boolean

    CmbFlg = False
    With cmb
        .ListIndex = 0      ' ???
        If .ListCount = 1 Then Exit Sub
        For i = 0 To .ListCount - 1
            .ListIndex = i
            If UCase(Trim(.Text)) = UCase(Trim(ComboString)) Then
                CmbFlg = True
                Exit For
            End If
        Next i
        If Not CmbFlg Then .ListIndex = -1
    End With

End Sub

Private Sub ClickCombo(ByRef cmb As ComboBox, ByVal GlobalType As Byte)
    
    If LoadFlag = True Then Exit Sub
    
    With cmb
        
        If .ListIndex = -1 Then Exit Sub
        DropData = .ItemData(.ListIndex)
    
        ' add new
        If DropData = 0 Then
            frmPWAdd.GlobalType = GlobalType
            frmPWAdd.Init
            frmPWAdd.Show vbModal
        End If
    
    End With

End Sub

Private Sub LoadCombo(ByRef cmb As ComboBox, ByVal GlobalType As Byte)
    
    With cmb
            
        .Clear
        
        If GlobalType <> PREquate.GlobalTypePWCounty Then
            .AddItem "<Add New>"
            .ItemData(.NewIndex) = 0
        End If
            
        If GlobalType = PREquate.GlobalTypePWCounty Then
            SQLString = "SELECT * FROM PRCounty ORDER BY CountyName"
            If PRCounty.GetBySQL(SQLString) Then
                Do
                    .AddItem PRCounty.CountyName
                    .ItemData(.NewIndex) = PRCounty.CountyID
                    If PRCounty.GetNext = False Then Exit Do
                Loop
            End If
        Else
            ' use PRGlobal for storage
            SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & GlobalType & _
                        " ORDER BY Description"
            If PRGlobal.GetBySQL(SQLString) Then
                Do
                    .AddItem PRGlobal.Description
                    .ItemData(.NewIndex) = PRGlobal.GlobalID
                    If PRGlobal.GetNext = False Then Exit Do
                Loop
            End If
        End If
    
    End With

End Sub

Private Sub cmdAdd_Click()
    rs.AddNew
End Sub

Private Sub cmdDelete_Click()
    rs.Delete
End Sub


