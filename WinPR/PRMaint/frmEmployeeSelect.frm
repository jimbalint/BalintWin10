VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEmployeeSelect 
   Caption         =   "Employee Select"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7605
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5175
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   8895
      _cx             =   15690
      _cy             =   9128
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
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmEmployeeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rs As New ADODB.Recordset

Private Sub Form_Load()
    
    ' setup temp record set
    rs.CursorLocation = adUseClient
    
    rs.Fields.Append "Select", adBoolean
    rs.Fields.Append "EmployeeNumber", adDouble
    rs.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    rs.Fields.Append "DeptNumber", adDouble
    rs.Fields.Append "DeptName", adVarChar, 80, adFldIsNullable
    rs.Fields.Append "EmployeeID", adDouble
    
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    ' fill the temp recordset
    SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
    If Not PREmployee.GetBySQL(SQLString) Then End ' ???
    Do
        rs.AddNew
        rs!Select = True
        rs!EmployeeNumber = PREmployee.EmployeeNumber
        rs!EmployeeName = Mid(PREmployee.LFName, 1, 80)
        
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            rs!DeptNumber = 0
            rs!DeptName = ""
        Else
            rs!DeptNumber = PRDepartment.DepartmentNumber
            rs!DeptName = PRDepartment.Name
        End If
        
        rs!EmployeeID = PREmployee.EmployeeID
        rs.Update
    
        If Not PREmployee.GetNext Then Exit Do
    Loop
    
    SetGrid rs, fg
    
End Sub

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub


