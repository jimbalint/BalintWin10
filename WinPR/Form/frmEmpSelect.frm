VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEmpSelect 
   Caption         =   "Employee Selection"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
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
   ScaleHeight     =   8640
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Apply &Filters"
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
      Left            =   8060
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
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
      Left            =   8044
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.ComboBox cmbPayType 
      Height          =   360
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ComboBox cmbEmpStat 
      Height          =   360
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cmbDepts 
      Height          =   360
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
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
      Left            =   4800
      TabIndex        =   3
      Top             =   1200
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
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
      Left            =   1002
      TabIndex        =   1
      Top             =   7680
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5655
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   7050
      _cx             =   12435
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
   Begin VB.Label Label3 
      Caption         =   "Pay Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label emp 
      Caption         =   "Employee Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Departments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
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
      Height          =   375
      Left            =   743
      TabIndex        =   4
      Top             =   120
      Width           =   8895
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
Public SelString As String

Dim X As String
Dim SortCol As Byte
Dim SortType As Byte
Dim dbFileName As String
Dim dbSortDesc As Boolean
Dim dbSortCol As Byte
Dim dbFields(7) As String
Dim SQLStr As String


Private Sub Form_Load()
Dim n As Long
    
    ' setup temp record set
    rsEmp.CursorLocation = adUseClient
    rsEmp.Fields.Append "Select", adBoolean
    rsEmp.Fields.Append "EmpNo", adDouble
    rsEmp.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "DeptNo", adDouble
    rsEmp.Fields.Append "DeptName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "Active", adVarChar, 8, adFldIsNullable
    rsEmp.Fields.Append "Salaried", adVarChar, 8, adFldIsNullable
    rsEmp.Fields.Append "EmployeeID", adDouble

    rsEmp.Open , , adOpenDynamic, adLockOptimistic
'    flexScrollBarVertical (2)
    SetGrid rsEmp, fg

    dbFields(1) = "Emp # "
    dbFields(2) = "Emp Name "
    dbFields(3) = "Dpt # "
    dbFields(4) = "Dept Name "
    
    fg.ColWidth(1) = 1000
    fg.ColWidth(2) = 3000
    fg.ColWidth(3) = 800
    fg.ColWidth(4) = 3200
    
    Me.lblCompanyName = PRCompany.Name
    SelString = "All Employees Selected"
    AllEmployees = True

    Me.cmbDepts.AddItem "All"
'    Me.cmbDepts.ItemData(cmbDepts.NewIndex) = 0
    n = 0
    If PRDepartment.GetBySQL("SELECT * FROM PRDepartment ORDER BY DepartmentNumber") Then
        Do
            Me.cmbDepts.AddItem PRDepartment.DepartmentNumber & " " & PRDepartment.Name
            Me.cmbDepts.ItemData(cmbDepts.NewIndex) = PRDepartment.DepartmentID
            n = n + 1
            
            If Not PRDepartment.GetNext Then Exit Do
        Loop
    End If
    
    Me.cmbDepts.ListIndex = 0
    
    Me.cmbEmpStat.AddItem "All"
    Me.cmbEmpStat.AddItem "Active"
    Me.cmbEmpStat.AddItem "Inactive"
    Me.cmbEmpStat.ListIndex = 0
    
    Me.cmbPayType.AddItem "All"
    Me.cmbPayType.AddItem "Hourly"
    Me.cmbPayType.AddItem "Salary"
    Me.cmbPayType.ListIndex = 0

    PopRecordset

End Sub

Private Sub PopRecordset()
    
    ' fill the temp recordset

    SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
    If Not PREmployee.GetBySQL(SQLString) Then
        GoBack
    End If
    
    SelCount = 0
    AllEmployees = True
    
    Do
        
        rsEmp.AddNew
        rsEmp!Select = True
        rsEmp!EmpNo = PREmployee.EmployeeNumber
        rsEmp!EmployeeName = Mid(PREmployee.LFName, 1, 80)
        
        With Me.cmbDepts
            If .ListIndex <> 0 Then
                If .ItemData(.ListIndex) <> PREmployee.DepartmentID Then
                    rsEmp!Select = False
                End If
            End If
        End With
                
        With Me.cmbEmpStat
            If .ListIndex = 1 And PREmployee.Inactive = 1 Then
                rsEmp!Select = False
            ElseIf .ListIndex = 2 And PREmployee.Inactive = 0 Then
                rsEmp!Select = False
            End If
                
        End With
        
        With Me.cmbPayType
            If .ListIndex = 1 And PREmployee.Salaried = 1 Then
                rsEmp!Select = False
            ElseIf .ListIndex = 2 And PREmployee.Salaried = 0 Then
                rsEmp!Select = False
            End If
        End With
        
        If Not PRDepartment.GetByID(PREmployee.DepartmentID) Then
            rsEmp!DeptNo = 0
            rsEmp!Deptname = ""
        Else
            rsEmp!DeptNo = PRDepartment.DepartmentNumber
            rsEmp!Deptname = Mid(Trim(PRDepartment.Name), 1, 80)
        End If
        
        rsEmp!EmployeeID = PREmployee.EmployeeID
        If PREmployee.Inactive = 1 Then
            rsEmp!Active = "Inactive"
        Else
            rsEmp!Active = "Active"
        End If
        If PREmployee.Salaried = 0 Then
            rsEmp!Salaried = "Hourly"
        Else
            rsEmp!Salaried = "Salary"
        End If
        
        If rsEmp!Select = True Then
            SelCount = SelCount + 1
        Else
            AllEmployees = False
        End If
        
        rsEmp.Update

SkipIt:
        If Not PREmployee.GetNext Then Exit Do
    Loop
    
    If SelCount = 0 Then
        MsgBox "No employee matches this criteria", vbExclamation, "Employee Select"
    Else
        If frmEmpSelect.AllEmployees = False Then
            SelString = frmEmpSelect.SelCount & " Employees Selected"
        Else
            SelString = "ALL Employees Selected"
        End If
    End If
    
    rsEmp.MoveFirst


End Sub


Private Sub cmdFilter_Click()

    rsEmp.Close
    rsEmp.CursorLocation = adUseClient
    rsEmp.Fields.Append "Select", adBoolean
    rsEmp.Fields.Append "EmpNo", adDouble
    rsEmp.Fields.Append "EmployeeName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "DeptNo", adDouble
    rsEmp.Fields.Append "DeptName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "Active", adVarChar, 8, adFldIsNullable
    rsEmp.Fields.Append "Salaried", adVarChar, 8, adFldIsNullable
    rsEmp.Fields.Append "EmployeeID", adDouble

    rsEmp.Open , , adOpenDynamic, adLockOptimistic
    SetGrid rsEmp, fg
    
    PopRecordset
    SelString = rsEmp.RecordCount & " Employees Selected"

End Sub


Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'
'    fg.ColWidth(1) = 1000
'    fg.ColWidth(2) = 3000
'    fg.ColWidth(3) = 800
'    fg.ColWidth(4) = 3200
    
    ' clicking on a column header sorts based on that column
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then

    
        ' toggle the sort order
        If fg.MouseCol = dbSortCol Then
           If dbSortDesc = False Then
              dbSortDesc = True
           Else
              dbSortDesc = False
           End If
        Else
           ' switch the column
           fg.Cell(flexcpFontBold, 0, fg.MouseCol) = True
           fg.Cell(flexcpFontBold, 0, dbSortCol) = False
           fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol)
           dbSortCol = fg.MouseCol
        End If
    
        If dbSortDesc Then
           fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "-"
        Else
           fg.TextMatrix(0, dbSortCol) = dbFields(dbSortCol) & "+"
        End If
        
        If dbSortCol = 1 Then
            If dbSortDesc = True Then
                rsEmp.Sort = "EmpNo desc"
            Else
                rsEmp.Sort = "EmpNo"
            End If
        ElseIf dbSortCol = 2 Then
            If dbSortDesc = True Then
                rsEmp.Sort = "EmployeeName Desc"
            Else
                rsEmp.Sort = "Employeename"
            End If
        ElseIf dbSortCol = 3 Then
            If dbSortDesc = True Then
                rsEmp.Sort = "DeptNo Desc"
            Else
                rsEmp.Sort = "DeptNo"
            End If
        ElseIf dbSortCol = 4 Then
            If dbSortDesc = True Then
                rsEmp.Sort = "DeptName Desc"
            Else
                rsEmp.Sort = "DeptName"
            End If
        End If
        
        fg.ShowCell 1, 0
    
    End If
            
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
    If Me.cmbDepts.ListIndex <> 0 Then AllEmployees = False
    If Me.cmbEmpStat.ListIndex <> 0 Then AllEmployees = False
    If Me.cmbPayType.ListIndex <> 0 Then AllEmployees = False
    
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
        MsgBox "You must select at least One Employee", vbExclamation, "Employee Select"
        rsEmp.MoveFirst
        Exit Sub
    End If
    Me.SelString = SelCount & " Employees Selected"
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub


