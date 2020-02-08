VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   7920
      TabIndex        =   2
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   8880
      Width           =   2055
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7575
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      _cx             =   19500
      _cy             =   13361
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
Option Explicit

Dim rsQB As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim I, J, K As Long
Dim QBDrop As String


Private Sub Form_Load()

    SQLString = "SELECT * FROM QBAccount ORDER BY AccountType, AccountNumber"
    rsInit SQLString, cn, rsQB
    If rsQB.RecordCount = 0 Then
        MsgBox "no recs ..."
        End
    End If
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "QBAccount", adVarChar, 50, adFldIsNullable
    rs.Fields.Append "Amount", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    For I = 1 To 10
        rs.AddNew
        rs!Amount = 100 * I + I
        rs.Update
    Next I
    
    SetGrid rs, fg

    QBDrop = fg.BuildComboList(rsQB, "Name,AccountType", "QBID")
    
 MsgBox QBDrop
    
    fg.ColComboList(0) = QBDrop
    fg.ColWidth(0) = 3000
    
    
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    rs.MoveFirst
    Do
        MsgBox rs!QBAccount
        rs.MoveNext
    Loop Until rs.EOF
End Sub

