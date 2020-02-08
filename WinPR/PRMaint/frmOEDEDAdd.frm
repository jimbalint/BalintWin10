VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOEDEDAdd 
   Caption         =   "Add Other Earning / Deduction"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7185
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4575
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
      _cx             =   8281
      _cy             =   8070
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Msg2"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Msg1"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmOEDEDAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rs As New ADODB.Recordset
Dim rsEEItem As New ADODB.Recordset
Dim EmpID As Long

Private Sub Form_Load()
    
    Me.lblMsg1 = Trim(PRCompany.Name)
    Me.lblMsg2 = PREmployee.FLName
    EmpID = PREmployee.EmployeeID
    TaskID = 0
    
    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    Me.fg.AllowSelection = False
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
    End Select
    
End Sub


Private Sub cmdCancel_Click()
    TaskID = 0
    SelReturn
End Sub

Private Sub cmdSelect_Click()
    TaskID = rs!ItemID
    SelReturn
End Sub

Public Sub Init()

    rs.CursorLocation = adUseClient
   
    rs.Fields.Append "ItemID", adDouble
    rs.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "Type", adVarChar, 15, adFldIsNullable
    
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    ' load the employer items
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 AND Active = 1 " & _
                " AND (ItemType = " & PREquate.ItemTypeOE & _
                " OR ItemType = " & PREquate.ItemTypeDED & _
                " OR ItemType = " & PREquate.ItemTypeSDTax & ")" & _
                "ORDER BY ItemType, Title"
                
    If Not PRItem.GetBySQL(SQLString) Then
        MsgBox "No Active Items defined for the employer", vbExclamation
        Unload Me
    End If
    
    Do
    
        ' does the employee already have this one?
        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & PREmployee.EmployeeID & _
                    " AND EmployerItemID = " & PRItem.ItemID
        rsInit SQLString, cn, rsEEItem
        If rsEEItem.RecordCount = 0 Then
    
            rs.AddNew
            rs!ItemID = PRItem.ItemID
            rs!Title = Mid(PRItem.Title, 1, 30)
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                rs!Type = "Earning"
            Else
                rs!Type = "Deduction"
            End If
            rs.Update
        
        End If
        
        If Not PRItem.GetNext Then Exit Do
        
    Loop
    
    ' they already have 'em all !!!
    If rs.RecordCount = 0 Then
        Exit Sub
    End If
    
    rs.MoveFirst
    
    ' set the grid
    SetGrid rs, fg
    
    fg.ScrollBars = flexScrollBarVertical
    fg.SelectionMode = flexSelectionByRow
    fg.Editable = flexEDNone
    
    fg.ColHidden(0) = True
    fg.ColWidth(1) = 3000
    fg.ColWidth(2) = 1500

End Sub

Private Sub SelReturn()
    rs.Close
    Me.Hide
End Sub

Private Sub fg_DblClick()
    TaskID = rs!ItemID
    SelReturn
End Sub

