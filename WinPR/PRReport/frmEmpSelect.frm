VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEmpSelect 
   Caption         =   "Employee Selection"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
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
   ScaleHeight     =   8340
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChkAll 
      Caption         =   "&Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5078
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdUnchkAll 
      Caption         =   "Clear &All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2918
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4058
      TabIndex        =   1
      Top             =   7560
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5655
      Left            =   278
      TabIndex        =   0
      Top             =   1560
      Width           =   8895
      _cx             =   15690
      _cy             =   9975
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1898
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmEmpSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsEmp As New ADODB.Recordset
Public AllEmployees As Boolean
Public SelCount As Long

Private Sub Form_Load()
    
    ' setup temp record set
    rsEmp.CursorLocation = adUseClient
    rsEmp.Fields.Append "Select", adBoolean
    rsEmp.Fields.Append "EmployeeNumber", adDouble
    rsEmp.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "DeptNumber", adDouble
    rsEmp.Fields.Append "DeptName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "EmployeeID", adDouble
    
    rsEmp.Open , , adOpenDynamic, adLockOptimistic
    
    ' fill the temp recordset
    SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
    If Not PREmployee.GetBySQL(SQLString) Then End ' ???
    Do
        rsEmp.AddNew
        rsEmp!Select = True
        rsEmp!EmployeeNumber = PREmployee.EmployeeNumber
        rsEmp!EmployeeName = Mid(PREmployee.LFName, 1, 80)
        
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            rsEmp!DeptNumber = 0
            rsEmp!DeptName = ""
        Else
            rsEmp!DeptNumber = PRDepartment.DepartmentNumber
            rsEmp!DeptName = PRDepartment.Name
        End If
        
        rsEmp!EmployeeID = PREmployee.EmployeeID
        rsEmp.Update
    
        If Not PREmployee.GetNext Then Exit Do
    Loop
    
    SetGrid rsEmp, fg
    Me.lblCompanyName = PRCompany.Name
    AllEmployees = True
End Sub

Private Sub cmdChkAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Select = True
        rsEmp.Update
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
End Sub

Private Sub cmdUnchkAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Select = False
        rsEmp.Update
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
End Sub
Private Sub cmdOK_Click()
    
    AllEmployees = True
    SelCount = 0
    rsEmp.MoveFirst
    Do
        If rsEmp!Select = False Then
            AllEmployees = False
        Else
            SelCount = SelCount + 1
        End If
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    If SelCount = 0 Then
        MsgBox "You must select at least one employee", vbExclamation, "Employee Select"
        rsEmp.MoveFirst
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub


