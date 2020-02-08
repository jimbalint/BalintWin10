VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmJournal 
   Caption         =   " General Ledger Journal List"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      _cx             =   11668
      _cy             =   10610
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FormatString    =   $"Journal.frx":0000
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
      TabBehavior     =   1
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
   Begin VB.CommandButton cmdSort 
      Caption         =   "&SORT"
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcn As ADODB.Connection
Private mrs As ADODB.Recordset

Private Sub cmdAdd_Click()
    AddAdo mrs, fg
End Sub

Private Sub cmdDelete_Click()
    mrs.Close
    SetAdo cn, mrs, "DELETE * FROM GLJournal WHERE JournalSource = " & fg.TextMatrix(fg.Row, 0)
    Set fg.DataSource = Nothing
    SetAdo cn, mrs, "SELECT [JournalSource],[JournalName] " & _
                     "FROM GLJournal WHERE JournalSource > 0 ORDER BY JournalSource"
    SetGrid mrs, fg
End Sub

Private Sub cmdSort_Click()
    
    Set fg.DataSource = Nothing
    SetAdo cn, mrs, "SELECT [JournalSource],[JournalName] " & _
                     "FROM GLJournal WHERE JournalSource > 0 ORDER BY JournalSource"
    SetGrid mrs, fg
    
    
'    Dim temp As Integer
'    On Error GoTo glErr
'    temp = mrs.Fields("JournalSource")
'    mrs.Requery
'    mrs.Find ("JournalSource = " & CStr(temp))
'    fg.Refresh
'    fg.SetFocus
'    Exit Sub
'glErr:
'    MsgBox Error(Err.Number)

End Sub

Private Sub Form_Load()
    ' get rid of nulls
    SetAdo cn, mrs, "DELETE * FROM GLJournal WHERE Isnull(JournalSource)"
    SetAdo cn, mrs, "SELECT [JournalSource],[JournalName] " & _
                     "FROM GLJournal WHERE JournalSource > 0 ORDER BY JournalSource"
    SetGrid mrs, fg
    frmJournal.Caption = " Journal Sources for " & glCompanyName
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub
