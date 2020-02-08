VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUsers 
   Caption         =   " System Users"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPassword 
      Caption         =   "&Password"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      _cx             =   8070
      _cy             =   6800
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      BackColorAlternate=   12632256
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Users.frx":0000
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
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   1
      VirtualData     =   -1  'True
      DataMember      =   "rsuser.datamember"
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
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton mdOK 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcn As ADODB.Connection
Private mrs As ADODB.Recordset

Private Sub cmdAdd_Click()
    AddAdo mrs, fg
'    mrs.Fields("Password") = InputBox("Enter PassWord")
End Sub

Private Sub cmdPassword_Click()
    frmPassword.Show vbModal, Me
    If Not "ERROR" = frmPassword.NewPassWord Then
        mrs.Fields("Password") = frmPassword.NewPassWord
        mrs.Update
        fg.Refresh
    End If
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    fg.AutoSize fg.Col
End Sub

Private Sub Form_Load()
    
    Set mcn = New ADODB.Connection
    mcn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\Balint\Data\GLSystem.mdb"
    mcn.Open
    
    SetAdo mcn, mrs, "select [Name],[Logon],[password] from Users"
    SetGrid mrs, fg
    fg.Cols = 3
    fg.TabBehavior = flexTabCells  ' tab to move between cells

End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub mdOK_Click()
    GoBack
End Sub


