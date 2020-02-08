VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPRCounty 
   Caption         =   "County File Maintenance"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
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
   ScaleHeight     =   9585
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   11040
      TabIndex        =   3
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   735
      Left            =   11040
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   735
      Left            =   11040
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   9255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _cx             =   17806
      _cy             =   16325
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
End
Attribute VB_Name = "frmPRCounty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim StateDrop As String
Dim OHIO As Long

Private Sub Form_Load()

    ' get rid of the blanks
    SQLString = "SELECT * FROM PRCounty"
    rsInit SQLString, cnDes, rs
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
            If IsNull(rs!CountyName) Then
                rs.Delete
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If

    SQLString = "SELECT * FROM PRState ORDER BY StateAbbrev"
    If PRState.GetBySQL(SQLString) = False Then
        MsgBox "No state data found!", vbExclamation
        GoBack
    End If
    StateDrop = ""
    Do
        If PRState.StateAbbrev = "OH" Then OHIO = PRState.StateID
        StateDrop = Trim(StateDrop) & "|#" & PRState.StateID & ";" & PRState.StateAbbrev
        If PRState.GetNext = False Then Exit Do
    Loop

    rsInit "SELECT * FROM PRCounty ORDER BY CountyName", cnDes, rs
    SetGrid rs, fg
    
    With Me.fg
        .ColWidth(0) = 0
        .ColWidth(1) = 4000
        .ColWidth(2) = 2000
        .ColWidth(3) = 1000
        .ColWidth(4) = 2000
        .ColComboList(3) = StateDrop
        .TextMatrix(0, 3) = "State"
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

Private Sub cmdAdd_Click()
    rs.AddNew
    rs!StateID = OHIO
    rs.Update
    rs.Requery
End Sub
Private Sub fg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 1 Or Col = 2 Then fg.EditText = UCase(fg.EditText)
End Sub

Private Sub cmdDelete_Click()
    
    If MsgBox("OK to delete: " & Trim(rs!CountyName) & "?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    ' don't allow if any cities assigned to this county
    SQLString = "SELECT * FROM PRCity ..."
End Sub




