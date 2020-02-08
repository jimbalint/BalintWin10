VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmItemSelect 
   Caption         =   "SELECT ITEMS TO INCLUDE/EXCLUDE"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
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
   ScaleHeight     =   7395
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6855
      _cx             =   12091
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsItem As New ADODB.Recordset

Private Sub Form_Load()
    
    ' set the record set with items
    rsItem.CursorLocation = adUseClient
    
    rsItem.Fields.Append "Select", adBoolean
    rsItem.Fields.Append "Title", adVarChar, 255, adFldIsNullable
    rsItem.Fields.Append "Type", adVarChar, 20, adFldIsNullable
    rsItem.Fields.Append "ItemID", adDouble
    
    rsItem.Open , , adOpenDynamic, adLockOptimistic
    
    ' populate with other earnings and deductions
    SQLString = "SELECT * FROM PRItem WHERE " & _
                "PRItem.ItemType = " & PREquate.ItemTypeOE & " OR " & _
                "PRItem.ItemType = " & PREquate.ItemTypeDED & " " & _
                "ORDER BY PRItem.ItemType, PRItem.ItemID"
    If Not PRItem.GetBySQL(SQLString) Then Unload Me
    
    Do
        rsItem.AddNew
        rsItem!Select = True
        rsItem!Title = PRItem.Title
        If PRItem.ItemType = PREquate.ItemTypeOE Then
            rsItem!Type = "OTH EARNG"
        Else
            rsItem!Type = "DEDUCTION"
        End If
        rsItem!ItemID = PRItem.ItemID
        rsItem.Update
        
        If Not PRItem.GetNext Then Exit Do
    Loop
    
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdOK_Click
    End Select
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub


