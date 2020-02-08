VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFFSchedule 
   Caption         =   "Free Format Schedules"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "R&ESET"
      Height          =   495
      Left            =   12840
      TabIndex        =   13
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAcctIns 
      Caption         =   "&INS ACCT"
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAcctDel 
      Caption         =   "DE&L ACCT"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&NEW"
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      ToolTipText     =   "NEW"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&RENAME"
      Height          =   495
      Left            =   11520
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&COPY"
      Height          =   495
      Left            =   10200
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdAcctAdd 
      Caption         =   "&ADD ACCTS"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   8520
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   14295
      _cx             =   25215
      _cy             =   10610
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin VB.ComboBox cmbFFSched 
      Height          =   345
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   0
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Left            =   1470
      TabIndex        =   6
      Top             =   840
      Width           =   11775
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "CompanyName"
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
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14415
   End
End
Attribute VB_Name = "frmFFSchedule"
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
Dim AccountDrop0 As String
Dim AccountDrop0TM As String
Dim AccountDropAll As String
Dim LoadFlag As Boolean
Dim boo As Boolean
Dim rw As Long

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.lblMsg1 = ""
    Fmt = "#,###,##0"

    If TableExists("GLFFSched", cn) = False Then
        GLFFSchedCreate
    End If
    
    LoadCmb
    
    ' list of accounts for ColCombo list for PctBase column
    I = 0
    GLAccount.GetAllAccounts
    AccountDrop0 = "|#0; "
    AccountDropAll = "|#0; "
    AccountDrop0TM = "|#0; "
    Me.Show
    Me.MousePointer = vbHourglass
    Do
        I = I + 1
        If I Mod 50 = 1 Then
            Me.lblMsg1 = "Now Loading Accounts ...." & I
            Me.Refresh
        End If
        If GLAccount.AcctType = "0" Then
            AccountDrop0 = AccountDrop0 & "|#" & GLAccount.Account & ";" & _
                          GLAccount.Account & vbTab & GLAccount.FullDesc
        End If
        If InStr(1, "0TM", GLAccount.AcctType, vbTextCompare) Then
            AccountDrop0TM = AccountDrop0TM & "|#" & GLAccount.Account & ";" & _
                          GLAccount.Account & vbTab & GLAccount.FullDesc
        End If
        AccountDropAll = AccountDropAll & "|#" & GLAccount.Account & ";" & _
                      GLAccount.Account & vbTab & GLAccount.FullDesc
        If GLAccount.GetNext = False Then Exit Do
    Loop

    ' recordset of columns
    rsCol.CursorLocation = adUseClient
    rsCol.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Abbrev", adVarChar, 30, adFldIsNullable
    rsCol.Fields.Append "Width", adDouble
    rsCol.Fields.Append "Number", adDouble
    rsCol.Fields.Append "DataType", adDouble
    rsCol.Fields.Append "Format", adVarChar, 30, adFldIsNullable
    rsCol.Open , , adOpenDynamic, adLockOptimistic

    ' setup the grid
    With Me.fg
        
        ColNum = 0
                
        ' add to rs in order of columns
        AddCol "", "Col0", 300
        AddCol "Account#", "Acct", 1300
        AddCol "Type", "Type", 470
        AddCol "Description", "Desc", 3880
        AddCol "Pct Base", "Pct", 1500
        AddCol "Print Tab", "Tab", 1000
        AddCol "Line Feeds", "LF", 1000
        AddCol "Sign Reverse", "Rev", 1300, flexDTBoolean
        AddCol "Alt Description", "AltD", 3050
        AddCol "FFSchedID", "ID", 0
        
        .Rows = 1
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
    
        .ColComboList(GetCol("Acct")) = AccountDropAll
        .ColComboList(GetCol("Pct")) = AccountDrop0TM
    
        .AllowSelection = False
        .AllowBigSelection = False
    
    End With

    With Me.cmbFFSched
        If .ListCount > 1 Then
            .ListIndex = 1
        Else
            .ListIndex = 0
        End If
        GlobID = .ItemData(.ListIndex)
    End With

    With Me
        .KeyPreview = True
        .cmdSave.ToolTipText = "SAVE the current account schedule"
        .cmdNew.ToolTipText = "Create NEW account schedule"
        .cmdDelete.ToolTipText = "DELETE the current account schedule"
        .cmdCopy.ToolTipText = "COPY the current acccount schedule"
        .cmdRename.ToolTipText = "RENAME the current account schedule"
        .cmdReset.ToolTipText = "RESET current schedule to last saved"
        .cmdAcctAdd.ToolTipText = "ADD a group of accounts"
        .cmdAcctDel.ToolTipText = "DELETE the current entry"
        .cmdAcctIns.ToolTipText = "INSERT an entry"
        .MousePointer = vbArrow
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    
    ' save grid ???
    If EditFlag = True Then
        If MsgBox("Save changes to: " & cmbPrev, vbQuestion + vbYesNo) = vbYes Then
            cmdSave_Click
        End If
    End If
    
    GoBack

End Sub

Private Sub cmbFFSched_Click()
    
    If LoadFlag = True Then Exit Sub
    
    ' save grid ???
    If EditFlag = True Then
        If MsgBox("Save changes to: " & cmbPrev, vbQuestion + vbYesNo) = vbYes Then
            cmdSave_Click
        End If
    End If
    
    With Me.cmbFFSched
        
        ' add new
        If .ListIndex = 0 Then
            x = InputBox("Title for free format schedule", _
                       "NEW Free Format Schedule")
            If x <> "" Then
                SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSched & _
                            " AND UserID = " & GLCompany.ID & _
                            " AND Description = '" & x & "'"
                If PRGlobal.GetBySQL(SQLString) = True Then
                    MsgBox "This title already exists!", vbExclamation
                    Exit Sub
                End If
                
                PRGlobal.Clear
                PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSched
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
    
    FFName = Me.cmbFFSched
    LoadAccounts
    cmbPrev = Me.cmbFFSched
    
    With fg
        If .Rows > 1 Then
            .Select 1, 2
        End If
    End With
    
    Me.Show
    fg.SetFocus

End Sub

Private Sub cmdAcctIns_Click()

    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    x = "" & vbTab & "0" & vbTab & "" & vbTab & "" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "" & vbTab & ""
    With fg
        If .Rows = 1 Or .Row = 0 Then
            .AddItem x
        Else
            .AddItem x, .Row
        End If
    End With

    EditFlag = True

'    If fg.Row = 0 And fg.Rows <> 1 Then
'        MsgBox "Click on the account grid at place to add accounts", vbInformation
'        Exit Sub
'    End If
'
'    EditFlag = True
'
'    InsertRows 1

End Sub

Private Sub InsertRows(ByVal InsRows As Long)

Dim CurrRow, OldMaxRows As Long
Dim OldRow, NewRow As Long

    If fg.Row = 0 Then Exit Sub

    CurrRow = fg.Row
    OldMaxRows = fg.Rows

    fg.Rows = fg.Rows + InsRows

    I = 0
    OldRow = OldMaxRows - 1
    NewRow = fg.Rows - 1

    Do

        For K = 0 To fg.Cols - 1
            fg.TextMatrix(NewRow, K) = fg.TextMatrix(OldRow, K)
            fg.TextMatrix(OldRow, K) = ""
        Next K

        OldRow = OldRow - 1
        NewRow = NewRow - 1

        If OldRow < CurrRow Then Exit Do

    Loop

End Sub

Private Sub cmdAcctDel_Click()

    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    If fg.Row = 0 Then Exit Sub
    
    If MsgBox("OK to delete: " & fg.TextMatrix(fg.Row, GetCol("Acct")) & " " & fg.TextMatrix(fg.Row, GetCol("Desc")), _
              vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    With fg
        For K = 0 To .Cols - 1
            If K <> GetCol("ID") Then
                .TextMatrix(.Row, K) = ""
            End If
        Next K
        .TextMatrix(.Row, GetCol("Desc")) = "*Deleted*"
    End With

    If fg.Row <> fg.Rows - 1 Then
        fg.Row = fg.Row + 1
    End If
    
    EditFlag = True

End Sub

Private Sub cmdAcctAdd_Click()
    
Dim SelNum, StartRow As Long
Dim aRow As Long
    
    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    rw = fg.Row
    
    With frmFFAddAcct
        
        .Show vbModal
        If .OK = False Then
            Unload frmFFAddAcct
            Exit Sub
        End If
    
        EditFlag = True
    
        With .fg
            
            For aRow = 1 To .Rows - 1
                
                If .TextMatrix(aRow, 0) = True Then
                    
                    boo = GLAccount.GetAccount(.TextMatrix(aRow, 1))
                    x = "" & vbTab & .TextMatrix(aRow, 1) & vbTab & GLAccount.TypeLevel & vbTab & GLAccount.FullDesc
                    If rw = 0 Then
                        Me.fg.AddItem x
                    Else
                        Me.fg.AddItem x, rw
                    End If
                    rw = rw + 1
                    
                    .TextMatrix(aRow, 0) = False   ' reset it
                    
                End If
            
            Next aRow
        
        End With
    
    End With

End Sub

Private Sub LoadAccounts()

    EditFlag = False
    fg.Rows = 1
    
    If Me.cmbFFSched.ListIndex <= 0 Then Exit Sub
    SQLString = "SELECT * FROM GLFFSched WHERE GlobalID = " & GlobID & _
                " ORDER BY SortOrder"
    If GLFFSched.GetBySQL(SQLString) = False Then Exit Sub

    Me.MousePointer = vbHourglass
    
    GLAccount.OpenRS
    I = 1
    Ct2 = GLFFSched.Records
    Do
        I = I + 1
        If I Mod 50 = 1 Then
            Me.lblMsg1 = "Now Loading Accounts: " & Format(I, Fmt) & " of " & Format(Ct2, Fmt)
            Me.Refresh
        End If
        With fg
            .Rows = I
            .TextMatrix(.Rows - 1, GetCol("Acct")) = GLFFSched.Account
            If GLAccount.GetAccount(GLFFSched.Account) = False Then
                .TextMatrix(.Rows - 1, GetCol("Desc")) = ""
            Else
                .TextMatrix(.Rows - 1, GetCol("Desc")) = GLAccount.FullDesc
            End If
            
            .TextMatrix(.Rows - 1, GetCol("Type")) = GLAccount.TypeLevel
            .TextMatrix(.Rows - 1, GetCol("Pct")) = GLFFSched.PercentBase
            .TextMatrix(.Rows - 1, GetCol("Tab")) = GLFFSched.PrintTab
            .TextMatrix(.Rows - 1, GetCol("LF")) = GLFFSched.LineFeeds
            .TextMatrix(.Rows - 1, GetCol("AltD")) = GLFFSched.AltDesc
            .TextMatrix(.Rows - 1, GetCol("ID")) = GLFFSched.FFSchedID
            .TextMatrix(.Rows - 1, GetCol("Rev")) = GLFFSched.SignReverse
        End With
        If GLFFSched.GetNext = False Then Exit Do
    
    Loop

    Me.lblMsg1 = ""
    Me.Refresh
    Me.MousePointer = vbArrow

    ' Me.fg.SetFocus

End Sub

Private Sub cmdSave_Click()

    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    ' handle the deleted accounts
    If fg.Rows > 1 Then
        I = 1
        Do
              
            If fg.TextMatrix(I, GetCol("Desc")) = "*Deleted*" Then
              
                K = GridValue(fg.TextMatrix(I, GetCol("ID")))
                If K <> 0 Then
                    SQLString = "DELETE * FROM GLFFSched WHERE FFSchedID = " & K
                    cn.Execute (SQLString)
                End If
              
                ' move the rows up
                If I = fg.Rows - 1 Then ' last row deleted - do nothing here
                Else
                    For J = I + 1 To fg.Rows - 1
                        For K = 0 To fg.Cols - 1
                            fg.TextMatrix(J - 1, K) = fg.TextMatrix(J, K)
                        Next K
                    Next J
                End If
         
                fg.Rows = fg.Rows - 1
                
            Else
                I = I + 1
            End If
      
            If I > fg.Rows - 1 Then Exit Do
      
        Loop
        
    End If
            
    For I = 1 To fg.Rows - 1
        
        If I Mod 50 = 1 Then
            Me.lblMsg1 = "Saving: " & Format(I, Fmt) & " of " & Format(fg.Rows - 1, Fmt)
            Me.Refresh
        End If
        
        With fg
            
            H = GridValue(.TextMatrix(I, GetCol("ID")))
            If H = 0 Then
                GLFFSched.Clear
                GLFFSched.GlobalID = GlobID
                GLFFSched.Save (Equate.RecAdd)
            Else
                If GLFFSched.GetByID(H) = True Then
                End If
            End If
                
            GLFFSched.Account = CLng(nNull(.TextMatrix(I, GetCol("Acct"))))
            GLFFSched.PercentBase = CLng(nNull(.TextMatrix(I, GetCol("Pct"))))
            GLFFSched.PrintTab = CByte(nNull(.TextMatrix(I, GetCol("Tab"))))
            GLFFSched.LineFeeds = CByte(nNull(.TextMatrix(I, GetCol("LF"))))
            GLFFSched.AltDesc = .TextMatrix(I, GetCol("AltD")) & ""
                    
            If GridValue(.TextMatrix(I, GetCol("Rev"))) = 0 Then
                GLFFSched.SignReverse = 0
            Else
                GLFFSched.SignReverse = 1
            End If
            
            GLFFSched.SortOrder = I
            
            GLFFSched.Save (Equate.RecPut)
    
        End With
    
    Next I

    EditFlag = False

    Me.lblMsg1 = ""
    Me.MousePointer = vbArrow
    Me.Refresh

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
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = GetCol("Desc") Then Cancel = True
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    EditFlag = True
    If Col = GetCol("Acct") Then
        GLAccount.GetAccount (fg.TextMatrix(Row, GetCol("Acct")))
        fg.TextMatrix(Row, GetCol("Desc")) = GLAccount.FullDesc
        fg.TextMatrix(Row, GetCol("Type")) = GLAccount.TypeLevel
    End If
End Sub

Private Sub LoadCmb()
    
    ' init the combo of FFSched definitions
    With Me.cmbFFSched
        .AddItem "<Add New>"
        .ItemData(.NewIndex) = 0
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeGLFFSched & _
                    " AND UserID = " & GLCompany.ID & _
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

Private Sub cmdDelete_Click()
    
    If Me.cmbFFSched.ListIndex <= 0 Then Exit Sub
    
    If MsgBox("OK to delete the ENTIRE Free Format Schedule " & vbCr & vbCr & _
              Me.cmbFFSched & "?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
              
    SQLString = "DELETE * FROM GLFFSched WHERE GlobalID = " & GlobID
    cn.Execute SQLString
    
    SQLString = "DELETE * FROM PRGlobal WHERE GlobalID = " & GlobID
    cnDes.Execute SQLString
    
    LoadFlag = True
    Me.cmbFFSched.Clear
    LoadCmb
    
    LoadFlag = False

    With Me.cmbFFSched
        If .ListCount > 1 Then
            .ListIndex = 1
        Else
            .ListIndex = 0
        End If
        GlobID = .ItemData(.ListIndex)
    End With
    
End Sub

Private Sub cmdCopy_Click()

    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    x = InputBox("Copy " & Me.cmbFFSched & " to: ", "Free Format Schedule")
    If x = "" Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSched
    PRGlobal.UserID = GLCompany.ID
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecAdd)
    
    GlobID = PRGlobal.GlobalID
    
    ' get the records from the fg
    With fg
        
        J = .Rows
        For I = 1 To J - 1
            
            If I Mod 50 = 1 Then
                Me.lblMsg1 = "Now copying " & Me.cmbFFSched & " to " & x & _
                             Format(I, Fmt) & " of " & Format(J - 1, Fmt)
                Me.Refresh
            End If
            
            boo = GLFFSched.GetByID(.TextMatrix(I, GetCol("ID")))
            GLFFSched.GlobalID = GlobID
            GLFFSched.Save (Equate.RecAdd)
            
        Next I
            
    End With

    With Me.cmbFFSched
        LoadFlag = True
        .AddItem x
        .ItemData(.NewIndex) = GlobID
        LoadFlag = False
        .ListIndex = .ListCount - 1
    End With

    Me.MousePointer = vbArrow

End Sub
Private Sub cmdReset_Click()
    
    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    If EditFlag = True Then
        If MsgBox("OK to discard all changes?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    EditFlag = False
    LoadAccounts
End Sub

Private Sub cmdRename_Click()

    If Me.cmbFFSched.ListIndex = -1 Then Exit Sub
    
    x = InputBox("Rename " & Me.cmbFFSched & " to: ", "Free Format Schedule Rename")
    If x = "" Then Exit Sub
    
    boo = PRGlobal.GetByID(GlobID)
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecPut)
    
    I = Me.cmbFFSched.ListIndex
    
    LoadFlag = True
    Me.cmbFFSched.Clear
    LoadCmb
    Me.cmbFFSched.ListIndex = I
    LoadFlag = False

End Sub

Private Sub cmdNew_Click()

    x = InputBox("Name for new Free Format Schedule:", "GL Free Format Schedule")
    If x = "" Then Exit Sub
    
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeGLFFSched
    PRGlobal.UserID = GLCompany.ID
    PRGlobal.Description = x
    PRGlobal.Save (Equate.RecAdd)
    
    GlobID = PRGlobal.GlobalID
    
    With Me.cmbFFSched
        .AddItem x
        .ItemData(.NewIndex) = GlobID
        .ListIndex = .ListCount - 1
    End With

End Sub

Private Function GridValue(ByVal Str As String) As Long

    GridValue = 0
    If IsNull(Str) Then Exit Function
    If Str = "" Then Exit Function
    If Str = "0" Then Exit Function
    GridValue = CLng(Str)

End Function
