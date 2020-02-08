VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPRTSPrint 
   Caption         =   "Time Sheet Print"
   ClientHeight    =   10125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   10125
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAllJobs 
      Caption         =   "A L L"
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CheckBox chkAllCustomers 
      Caption         =   "A L L"
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox chkAllEmployees 
      Caption         =   "A L L"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   3240
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fgWE 
      Height          =   1935
      Left            =   480
      TabIndex        =   18
      Top             =   1080
      Width           =   3855
      _cx             =   6800
      _cy             =   3413
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   8040
      TabIndex        =   12
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Default         =   -1  'True
      Height          =   615
      Left            =   6120
      TabIndex        =   11
      Top             =   9120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "   SORT / SUBTOTAL BY:    "
      Height          =   855
      Left            =   360
      TabIndex        =   17
      Top             =   8880
      Width           =   4575
      Begin VB.OptionButton optSortJob 
         Caption         =   "&JOB"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optSortEE 
         Caption         =   "&EMPLOYEE"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdJobClear 
      Caption         =   "CLEAR ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   8
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdJobAll 
      Caption         =   "SELECT ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCustClear 
      Caption         =   "CLEAR ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdCustAll 
      Caption         =   "SELECT ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmpClear 
      Caption         =   "CLEAR ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmpAll 
      Caption         =   "SELECT ALL"
      Height          =   735
      Left            =   7680
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   3600
      Width           =   6975
      _cx             =   12303
      _cy             =   1720
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
   Begin VSFlex8Ctl.VSFlexGrid fgCust 
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   6975
      _cx             =   12303
      _cy             =   1720
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
   Begin VSFlex8Ctl.VSFlexGrid fgJob 
      Height          =   975
      Left            =   480
      TabIndex        =   6
      Top             =   7200
      Width           =   6975
      _cx             =   12303
      _cy             =   1720
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
   Begin VB.Label Label3 
      Caption         =   "Week(s) Ended:"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Jobs:"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Customers:"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblEmp 
      Caption         =   "Employees:"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "lblCompanyName"
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
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmPRTSPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WEDate As Date
Dim rs As New ADODB.Recordset
Dim rsWE As New ADODB.Recordset
Dim RSEmp As New ADODB.Recordset
Dim rsCust As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim fgRW As Long
Dim EmpFilter As Boolean
Dim CustFilter As Boolean
Dim JobFilter As Boolean
Dim i, j, k As Long
Dim StartDate, EndDate As Date
Dim UseBillingRate As Boolean

Private Sub Form_Load()

    PrvwReturn = True
 
    If TableExists("PRTimeSheet", cn) = False Then PRTimeSheetCreate
    If TableExists("JCCustomer", cn) = False Then CustomerCreate
    If TableExists("JCJob", cn) = False Then JobCreate
 
    ' use billing rate ?
    UseBillingRate = False
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeScreenDefault & _
              " AND UserID = " & PRCompany.CompanyID & _
              " AND Description = 'TimeSheet'"
    If PRGlobal.GetBySQL(SQLString) = True And PRGlobal.Byte1 = 1 Then
        UseBillingRate = True
    End If
  
    ' ---------------------------------------------------------------
    ' Week(s) ended grid
    SQLString = "SELECT WEDate FROM PRTimeSheet ORDER BY WEDate DESC"
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then
        MsgBox "No timesheet date exists!", vbExclamation
        GoBack
    End If
    
    On Error Resume Next
    rsWE.Close
    On Error GoTo 0
    rsWE.CursorLocation = adUseClient
    rsWE.Fields.Append "Select", adBoolean
    rsWE.Fields.Append "WeekEnded", adDate
    rsWE.Open , , adOpenDynamic, adLockOptimistic
    
    rs.MoveFirst
    Do
        SQLString = "WeekEnded = " & rs!WEDate
        If rsWE.RecordCount > 0 Then
            rsWE.Find SQLString, 0, adSearchForward, 1
        End If
        If rsWE.EOF Or rsWE.RecordCount = 0 Then
            rsWE.AddNew
            If rs!WEDate = WEDate And WEDate <> 0 Then
                rsWE!Select = True
            Else
                rsWE!Select = False
            End If
            rsWE!WeekEnded = rs!WEDate
            rsWE.Update
        End If
        rs.MoveNext
    Loop Until rs.EOF
        
    rs.Close
    Set rs = Nothing
    
    rsWE.Sort = "WeekEnded DESC"
    SetGrid rsWE, Me.fgWE
    
    With fgWE
        .SelectionMode = flexSelectionByRow
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
    End With
    
    Me.chkAllCustomers = 1
    CustomerDisplay
    Me.chkAllEmployees = 1
    EmployeeDisplay
    Me.chkAllJobs = 1
    JobDisplay
    
    Me.chkAllCustomers.Enabled = False
    Me.chkAllEmployees.Enabled = False
    Me.chkAllJobs.Enabled = False
    
    ' ---------------------------------------------------------------
    
'    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'    ' fill in PRTimeSheet.CustomerID ALWAYS
'    SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate)
'    If PRTimeSheet.GetBySQL(SQLString) Then
'        Do
'            If PRTimeSheet.JobID <> 0 Then
'                If JCJob.GetByID(PRTimeSheet.JobID) Then
'                    PRTimeSheet.CustomerID = JCJob.ParentID
'                End If
'            Else
'                PRTimeSheet.CustomerID = 0
'            End If
'            PRTimeSheet.Save (Equate.RecPut)
'            If PRTimeSheet.GetNext = False Then Exit Do
'        Loop
'    End If
'    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'
'    With frmPRTimeSheet.cmbWEDate
'        WEDate = .ItemData(.ListIndex)
'    End With
'
'    rsEmp.CursorLocation = adUseClient
'    rsEmp.Fields.Append "Select", adBoolean
'    rsEmp.Fields.Append "Name", adVarChar, 40, adFldIsNullable
'    rsEmp.Fields.Append "Number", adDouble
'    rsEmp.Fields.Append "ID", adDouble
'    rsEmp.Open , , adOpenDynamic, adLockOptimistic
'
'    rsJob.CursorLocation = adUseClient
'    rsJob.Fields.Append "Select", adBoolean
'    rsJob.Fields.Append "Name", adVarChar, 40, adFldIsNullable
'    rsJob.Fields.Append "Number", adDouble
'    rsJob.Fields.Append "ID", adDouble
'    rsJob.Open , , adOpenDynamic, adLockOptimistic
'
'    SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(WEDate)
'    If PRTimeSheet.GetBySQL(SQLString) = False Then
'        MsgBox "No Time Sheet Data Found!", vbInformation
'        Unload Me
'    End If
'    Do
'
'        ' employees
'        rsEmp.Find "ID = " & PRTimeSheet.EmployeeID, 0, adSearchForward, 1
'        If rsEmp.EOF Then
'            If PREmployee.GetByID(PRTimeSheet.EmployeeID) = False Then
'                MsgBox "EE not found: " & PRTimeSheet.EmployeeID, vbExclamation
'                GoBack
'            End If
'            rsEmp.AddNew
'            rsEmp!Select = True
'            rsEmp!Name = Mid(PREmployee.LFName, 1, 40)
'            rsEmp!Number = PREmployee.EmployeeNumber
'            rsEmp!ID = PREmployee.EmployeeID
'            rsEmp.Update
'        End If
'
'        ' find the job record
'        If PRTimeSheet.JobID <> 0 Then
'
'            SQLString = "SELECT * FROM JCJob WHERE JobID = " & PRTimeSheet.JobID
'            If JCJob.GetBySQL(SQLString) = False Then
'                MsgBox "Job not found: " & PRTimeSheet.JobID, vbExclamation
'                GoBack
'            End If
'
'            ' customer
'            rsCust.Find "ID = " & JCJob.ParentID, 0, adSearchForward, 1
'            If rsCust.EOF Then
'                If JCCustomer.GetByID(JCJob.ParentID) = False Then
'                    MsgBox "Customer not found: " & JCJob.ParentID, vbExclamation
'                    GoBack
'                End If
'                rsCust.AddNew
'                rsCust!Select = True
'                rsCust!Name = Mid(JCCustomer.Name, 1, 40)
'                rsCust!Number = JCCustomer.CustomerID
'                rsCust!ID = JCCustomer.CustomerID
'                rsCust.Update
'            End If
'
'            ' job
'            rsJob.Find "ID = " & PRTimeSheet.JobID, 0, adSearchForward, 1
'            If rsJob.EOF Then
'                rsJob.AddNew
'                rsJob!Select = True
'                rsJob!Name = Mid(JCJob.FullName, 1, 40)
'                rsJob!Number = JCJob.JobID
'                rsJob!ID = JCJob.JobID
'                rsJob.Update
'            End If
'
'        Else
'
'            rsCust.Find "ID = 0", 0, adSearchForward, 1
'            If rsCust.EOF Then
'                rsCust.AddNew
'                rsCust!Select = True
'                rsCust!Name = Mid(PRCompany.Name, 1, 40)
'                rsCust!Number = 0
'                rsCust!ID = 0
'                rsCust.Update
'            End If
'
'            rsJob.Find "ID = 0", 0, adSearchForward, 1
'            If rsJob.EOF Then
'                rsJob.AddNew
'                rsJob!Select = True
'                rsJob!Name = Mid(PRCompany.Name, 1, 40)
'                rsJob!Number = 0
'                rsJob!ID = 0
'                rsJob.Update
'            End If
'
'        End If
'
'        If PRTimeSheet.GetNext = False Then Exit Do
'
'    Loop
'
'    ' sorts .....
'    rsEmp.Sort = "Name"
'    rsCust.Sort = "Name"
'    rsJob.Sort = "Name"
'
'    SetGrid rsEmp, Me.fgEmp
'    SetGrid rsCust, Me.fgCust
'    SetGrid rsJob, Me.fgJob
'
'    With fgEmp
'        .ColWidth(1) = 4000
'        .SelectionMode = flexSelectionByRow
'    End With
'
'    With fgCust
'        .ColWidth(1) = 4000
'        .SelectionMode = flexSelectionByRow
'    End With
'
'    With fgJob
'        .ColWidth(1) = 4000
'        .SelectionMode = flexSelectionByRow
'    End With
'
    
    Me.optSortEE = True
    Me.lblCompanyName = PRCompany.Name

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    If WEDate <> 0 Then
        Me.Hide
    Else
        GoBack
    End If
End Sub

Private Sub fgEmp_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
Private Sub fgCust_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
Private Sub fgJob_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub cmdEmpAll_Click()
    CheckSet fgEmp, RSEmp, True
End Sub

Private Sub cmdEmpClear_Click()
    CheckSet fgEmp, RSEmp, False
End Sub

Private Sub cmdCustAll_Click()
    CheckSet fgCust, rsCust, True
End Sub

Private Sub cmdCustClear_Click()
    CheckSet fgCust, rsCust, False
End Sub
Private Sub cmdJobAll_Click()
    CheckSet fgJob, rsJob, True
End Sub

Private Sub cmdJobClear_Click()
    CheckSet fgJob, rsJob, False
End Sub

Private Sub CheckSet(ByRef fg As VSFlexGrid, ByRef rs As ADODB.Recordset, ByVal Sel As Boolean)
    If rs.RecordCount = 0 Then Exit Sub
    fgRW = fg.Row
    rs.MoveFirst
    Do
        rs!Select = Sel
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
    fg.TopRow = fgRW
    fg.Select fgRW, 0
End Sub

Private Sub cmdPrint_Click()

    ' any customer filters applied
    CustFilter = False
'    rsCust.MoveFirst
'    Do
'        If rsCust!Select = False Then
'            CustFilter = True
'            Exit Do
'        End If
'        rsCust.MoveNext
'    Loop Until rsCust.EOF

    ' ***************************************************
    ' *** print all ***
    
    RSEmp.CursorLocation = adUseClient
    RSEmp.Fields.Append "Select", adBoolean
    RSEmp.Fields.Append "Name", adVarChar, 40, adFldIsNullable
    RSEmp.Fields.Append "Number", adDouble
    RSEmp.Fields.Append "ID", adDouble
    RSEmp.Open , , adOpenDynamic, adLockOptimistic

    rsJob.CursorLocation = adUseClient
    rsJob.Fields.Append "Select", adBoolean
    rsJob.Fields.Append "Name", adVarChar, 40, adFldIsNullable
    rsJob.Fields.Append "Number", adDouble
    rsJob.Fields.Append "ID", adDouble
    rsJob.Open , , adOpenDynamic, adLockOptimistic

    rsCust.CursorLocation = adUseClient
    rsCust.Fields.Append "Select", adBoolean
    rsCust.Fields.Append "Name", adVarChar, 40, adFldIsNullable
    rsCust.Fields.Append "Number", adDouble
    rsCust.Fields.Append "ID", adDouble
    rsCust.Open , , adOpenDynamic, adLockOptimistic

    rsWE.MoveFirst
    Do
        If rsWE!Select = False Then GoTo NxtWE

        SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & CLng(rsWE!WeekEnded)
        If PRTimeSheet.GetBySQL(SQLString) = False Then
            MsgBox "No Time Sheet Data Found!", vbInformation
            Unload Me
        End If
        Do
    
            If JCJob.GetByID(PRTimeSheet.JobID) Then
                PRTimeSheet.CustomerID = JCJob.ParentID
            End If
            PRTimeSheet.Save (Equate.RecPut)
    
            ' employees
            RSEmp.Find "ID = " & PRTimeSheet.EmployeeID, 0, adSearchForward, 1
            If RSEmp.EOF Then
                If PREmployee.GetByID(PRTimeSheet.EmployeeID) = False Then
                    MsgBox "EE not found: " & PRTimeSheet.EmployeeID, vbExclamation
                    GoBack
                End If
                RSEmp.AddNew
                RSEmp!Select = True
                RSEmp!Name = Mid(PREmployee.LFName, 1, 40)
                RSEmp!Number = PREmployee.EmployeeNumber
                RSEmp!ID = PREmployee.EmployeeID
                RSEmp.Update
            End If
    
            ' find the job record
            If PRTimeSheet.JobID <> 0 Then
    
                SQLString = "SELECT * FROM JCJob WHERE JobID = " & PRTimeSheet.JobID
                If JCJob.GetBySQL(SQLString) = False Then
                    MsgBox "Job not found: " & PRTimeSheet.JobID, vbExclamation
                    GoBack
                End If
    
                ' customer
                rsCust.Find "ID = " & JCJob.ParentID, 0, adSearchForward, 1
                If rsCust.EOF Then
                    If JCCustomer.GetByID(JCJob.ParentID) = False Then
                        MsgBox "Customer not found: " & JCJob.ParentID, vbExclamation
                        GoBack
                    End If
                    rsCust.AddNew
                    rsCust!Select = True
                    rsCust!Name = Mid(JCCustomer.Name, 1, 40)
                    rsCust!Number = JCCustomer.CustomerID
                    rsCust!ID = JCCustomer.CustomerID
                    rsCust.Update
                    PRTimeSheet.CustomerID = JCCustomer.CustomerID
                    PRTimeSheet.Save (Equate.RecPut)
                End If
    
                ' job
                rsJob.Find "ID = " & PRTimeSheet.JobID, 0, adSearchForward, 1
                If rsJob.EOF Then
                    rsJob.AddNew
                    rsJob!Select = True
                    rsJob!Name = Mid(JCJob.FullName, 1, 40)
                    rsJob!Number = JCJob.JobID
                    rsJob!ID = JCJob.JobID
                    rsJob.Update
                End If
    
            Else
    
                rsCust.Find "ID = 0", 0, adSearchForward, 1
                If rsCust.EOF Then
                    rsCust.AddNew
                    rsCust!Select = True
                    rsCust!Name = Mid(PRCompany.Name, 1, 40)
                    rsCust!Number = 0
                    rsCust!ID = 0
                    rsCust.Update
                End If
    
                rsJob.Find "ID = 0", 0, adSearchForward, 1
                If rsJob.EOF Then
                    rsJob.AddNew
                    rsJob!Select = True
                    rsJob!Name = Mid(PRCompany.Name, 1, 40)
                    rsJob!Number = 0
                    rsJob!ID = 0
                    rsJob.Update
                End If
    
            End If
    
            If PRTimeSheet.GetNext = False Then Exit Do
    
        Loop

NxtWE:
        rsWE.MoveNext
    Loop Until rsWE.EOF
    
    ' ***************************************************
    If Me.optSortEE = True Then
        TSPrint RSEmp, rsJob
    Else
        TSPrint rsJob, RSEmp
    End If

End Sub

Private Sub TSPrint(ByRef rs1 As ADODB.Recordset, ByRef rs2 As ADODB.Recordset)
    
Dim ID1, ID2 As Long
Dim JobName, EmpName, DeptName, EarnName As String
Dim SubTl(2, 8) As Currency
Dim FirstFlag As Boolean
Dim THrs As Currency
Dim LnCount As Long
    
    rs1.Sort = "Name"
    rs2.Sort = "Name"
    
    rsWE.Sort = "WeekEnded"
    ' get the date range for the header
    StartDate = 0
    EndDate = 0
    rsWE.MoveFirst
    Do
        If rsWE!Select = True Then
            If StartDate = 0 Then StartDate = rsWE!WeekEnded
            EndDate = rsWE!WeekEnded
        End If
        rsWE.MoveNext
    Loop Until rsWE.EOF
    
    PrtInit ("Land")
    SetFont 8, Equate.LandScape
    Columns = 150

    TSHeader

    rs1.MoveFirst
    Do
        
        If rs1!Select = False Then GoTo NextRS1
        
        For i = 1 To 8
            SubTl(1, i) = 0
        Next i
        
        FirstFlag = True
            
        rsWE.MoveFirst
        Do
            If rsWE!Select = False Then GoTo NextrsWE
                    
            rs2.MoveFirst
            Do
                If rs2!Select = False Then GoTo NextRS2
        
            
                If Me.optSortEE = True Then
                    
                    SQLString = "SELECT * FROM PRTimeSheet WHERE EmployeeID = " & rs1!ID & _
                                " AND JobID = " & rs2!ID & _
                                " AND WEDate = " & CLng(rsWE!WeekEnded) & _
                                " AND TotalHours <> 0 " & _
                                " ORDER BY TimeSheetID"
                Else
                    SQLString = "SELECT * FROM PRTimeSheet WHERE JobID = " & rs1!ID & _
                                " AND EmployeeID = " & rs2!ID & _
                                " AND WEDate = " & CLng(rsWE!WeekEnded) & _
                                " AND TotalHours <> 0 " & _
                                " ORDER BY TimeSheetID"
                End If
                
                If PRTimeSheet.GetBySQL(SQLString) = False Then GoTo NextRS2
                    
                Do
                
                    ' sub section header
                    If FirstFlag = True Then
                        PrintValue(1) = "   " & rs1!Name:   FormatString(1) = "a83"
                        PrintValue(2) = " ":                FormatString(2) = "~"
                        FormatPrint
                        Ln = Ln + 1
                        FirstFlag = False
                    End If
    
    '                ' skip for this customer?
    '                If PRTimeSheet.CustomerID = 0 Then
    '                Else
    '                    rsCust.Find "ID = " & PRTimeSheet.CustomerID, 0, adSearchForward, 1
    '                    If rsCust.EOF = False And rsCust!Select = False Then GoTo NextPRTimeSheet
    '                End If
                    If Me.optSortEE = True Then
                        EmpName = rs1!Number & " " & rs1!Name
                        JobName = rs2!Name
                    Else
                        JobName = rs1!Name
                        EmpName = rs2!Number & " " & rs2!Name
                    End If
                    
                    ' dept name
                    If PRDepartment.GetByID(PRTimeSheet.DepartmentID) Then
                        DeptName = PRDepartment.Name
                    Else
                        DeptName = PRTimeSheet.DepartmentID & " NF"
                    End If
                    
                    ' earnings type name
                    If PRTimeSheet.ItemID = 99991 Then
                        EarnName = "REG PAY"
                    ElseIf PRTimeSheet.ItemID = 99992 Then
                        EarnName = "OVT PAY"
                    Else
                        If PRItem.GetByID(PRTimeSheet.ItemID) Then
                            EarnName = PRItem.Title
                        Else
                            EarnName = PRTimeSheet.ItemID & " NF"
                        End If
                    End If
                    
                    ' print the line
                    If Me.optSortEE = True Then
                        PrintValue(1) = rs2!Name:           FormatString(1) = "a35"
                    Else
                        PrintValue(1) = rs2!Number & " " & rs2!Name: FormatString(1) = "a35"
                    End If
                    
                    PrintValue(2) = " ":                FormatString(2) = "a1"
                    PrintValue(3) = DeptName:           FormatString(3) = "a12"
                    PrintValue(4) = " ":                FormatString(4) = "a1"
                    PrintValue(5) = EarnName:           FormatString(5) = "a10"
                    PrintValue(6) = " ":                FormatString(6) = "a1"
                    
                    ' PrintValue(7) = PRTimeSheet.Note:   FormatString(7) = "a15"
                    PrintValue(7) = Format(PRTimeSheet.WEDate, "  mm/dd/yy")
                    FormatString(7) = "a13"
                    
                    k = 7
                    
                    If UseBillingRate = True Then
                        k = k + 1
                        PrintValue(k) = PRTimeSheet.BillingRate
                        FormatString(k) = "d8"
                    End If
                    
                    THrs = PRTimeSheet.SunHours + PRTimeSheet.MonHours + _
                           PRTimeSheet.TueHours + PRTimeSheet.WedHours + _
                           PRTimeSheet.ThuHours + PRTimeSheet.FriHours + _
                           PRTimeSheet.SatHours
                    
                    For i = 1 To 8
                        If i = 1 Then PrintValue(k + i) = PRTimeSheet.SunHours
                        If i = 2 Then PrintValue(k + i) = PRTimeSheet.MonHours
                        If i = 3 Then PrintValue(k + i) = PRTimeSheet.TueHours
                        If i = 4 Then PrintValue(k + i) = PRTimeSheet.WedHours
                        If i = 5 Then PrintValue(k + i) = PRTimeSheet.ThuHours
                        If i = 6 Then PrintValue(k + i) = PRTimeSheet.FriHours
                        If i = 7 Then PrintValue(k + i) = PRTimeSheet.SatHours
                        If i = 8 Then PrintValue(k + i) = THrs
                        SubTl(1, i) = SubTl(1, i) + CCur(PrintValue(k + i))
                        SubTl(2, i) = SubTl(2, i) + CCur(PrintValue(k + i))
                        If i = 8 Then
                            FormatString(k + i) = "d9"
                        Else
                            FormatString(k + i) = "d8"
                        End If
                    Next i
                    
                    PrintValue(k + i) = " ":        FormatString(k + i) = "~"
                    
                    ' bill rate not assigned
                    If UseBillingRate = True And PRTimeSheet.BillingRate = 0 Then
                        Prvw.vsp.Font.Bold = True
                        Prvw.vsp.Font.Italic = True
                    End If
                    
                    FormatPrint
                    
                    If UseBillingRate = True And PRTimeSheet.BillingRate = 0 Then
                        Prvw.vsp.Font.Bold = False
                        Prvw.vsp.Font.Italic = False
                    End If
                    
                    Ln = Ln + 1
                    If Ln >= MaxLines Then
                        FormFeed
                        TSHeader
                        FirstFlag = True
                    End If

                    LnCount = LnCount + 1

NextPRTimeSheet:
                    If PRTimeSheet.GetNext = False Then Exit Do
            
                Loop
            
NextRS2:
                rs2.MoveNext
            Loop Until rs2.EOF
NextrsWE:
            rsWE.MoveNext
        Loop Until rsWE.EOF

        If Me.optSortEE = True Then
            PrintValue(1) = "   Total For: " & rs1!Number & " " & rs1!Name
        Else
            PrintValue(1) = "   Total For: " & rs1!Name
        End If
        
        If UseBillingRate = True Then
            FormatString(1) = "a81"
        Else
            FormatString(1) = "a73"
        End If
        
        For i = 1 To 8
            PrintValue(1 + i) = SubTl(1, i)
            If i = 8 Then
                FormatString(1 + i) = "d9"
            Else
                FormatString(1 + i) = "d8"
            End If
        Next i
        PrintValue(1 + i) = " ":        FormatString(1 + i) = "~"
        FormatPrint
        Ln = Ln + 2

NextRS1:
        rs1.MoveNext
    Loop Until rs1.EOF

    Ln = Ln + 1

    PrintValue(1) = "   GRAND TOTALS:"
    
    If UseBillingRate = True Then
        FormatString(1) = "a81"
    Else
        FormatString(1) = "a73"
    End If
    
    For i = 1 To 8
        PrintValue(1 + i) = SubTl(2, i)
        If i = 8 Then
            FormatString(1 + i) = "d9"
        Else
            FormatString(1 + i) = "d8"
        End If
        
    Next i
    PrintValue(1 + i) = " ":        FormatString(1 + i) = "~"
    FormatPrint

    Prvw.vsp.EndDoc
    Prvw.Show vbModal

    RSEmp.Close
    Set RSEmp = Nothing
    rsJob.Close
    Set rsJob = Nothing
    rsCust.Close
    Set rsCust = Nothing
    Unload Me

End Sub

Private Sub TSHeader()
    
Dim DayString As String
    
    If StartDate = EndDate Then
        PageHeader "Time Sheet Report", "For the week ended: " & Format(StartDate, "mm/dd/yyyy"), "", ""
    Else
        PageHeader "Time Sheet Report", "For the weeks ended: " & Format(StartDate, "mm/dd/yyyy") & _
                   " To: " & Format(EndDate, "mm/dd/yyyy"), "", ""
    End If
                    
    Ln = Ln + 1
    
    ' print the line
    If Me.optSortEE = True Then
        PrintValue(1) = "Job Name":         FormatString(1) = "a35"
    Else
        PrintValue(1) = "Employee Name":    FormatString(1) = "a35"
    End If
    
    PrintValue(2) = " ":                FormatString(2) = "a1"
    PrintValue(3) = "Work Cat":         FormatString(3) = "a12"
    PrintValue(4) = " ":                FormatString(4) = "a1"
    PrintValue(5) = "Earng Type":       FormatString(5) = "a10"
    PrintValue(6) = " ":                FormatString(6) = "a1"
    PrintValue(7) = "  Wk End Dt":      FormatString(7) = "a12"

    k = 7
    
    If UseBillingRate = True Then
        k = k + 1
        PrintValue(k) = "Bill Rte":     FormatString(k) = "a8"
    End If

    ' show the date
    For i = 1 To 7
        k = k + 1
        DayString = Mid("SunMonTueWedThuFriSat", i * 3 - 2, 3)
        DayString = Trim(DayString) & " " & Day(WEDate - 7 + i)
        PrintValue(k) = DayString
        FormatString(k) = "r8"
    Next i
    PrintValue(k + 1) = "TOTAL":         FormatString(k + 1) = "r9"
    PrintValue(k + 2) = " ":             FormatString(k + 2) = "~"
    FormatPrint

    Ln = Ln + 2

End Sub

Private Sub fgWE_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub
Private Sub chkAllCustomers_Click()
    CustomerDisplay
End Sub

Private Sub CustomerDisplay()

    If Me.chkAllCustomers = 1 Then
        
        Me.fgCust.Visible = False
        Me.cmdCustAll.Visible = False
        Me.cmdCustClear.Visible = False
    
        On Error Resume Next
        rsCust.Close
        On Error GoTo 0
    
    Else
        
        Me.fgCust.Visible = True
        Me.cmdCustAll.Visible = True
        Me.cmdCustClear.Visible = True
    
'        fgRW = Me.fgWE.Row
'
'        rsCust.CursorLocation = adUseClient
'        rsCust.Fields.Append "Select", adBoolean
'        rsCust.Fields.Append "Name", adVarChar, 40, adFldIsNullable
'        rsCust.Fields.Append "Number", adDouble
'        rsCust.Fields.Append "ID", adDouble
'        rsCust.Open , , adOpenDynamic, adLockOptimistic
'
'        rsWE.MoveFirst
'        Do
'            SQLString = "SELECT * FROM PRTimeSheet WHERE WEDate = " & rsWE!WeekEnded
'            If PRTimeSheet.GetBySQL(SQLString) = True Then
'                Do
'                    SQLString = "ID = " & PRTimeSheet.CustomerID
'                    rsCust.Find SQLString, 0, adSearchForward, 1
'                    If rsCust.EOF = True Then
'                        if jccustomer.GetByID(prtimsheet
    
    End If






End Sub

Private Sub chkAllJobs_Click()
    JobDisplay
End Sub

Private Sub JobDisplay()

    If Me.chkAllJobs = 1 Then
        Me.fgJob.Visible = False
        Me.cmdJobAll.Visible = False
        Me.cmdJobClear.Visible = False
    Else
        Me.fgJob.Visible = True
        Me.cmdJobAll.Visible = True
        Me.cmdJobClear.Visible = True
    End If

End Sub
Private Sub chkAllEmployees_Click()
    EmployeeDisplay
End Sub

Private Sub EmployeeDisplay()

    If Me.chkAllEmployees = 1 Then
        Me.fgEmp.Visible = False
        Me.cmdEmpAll.Visible = False
        Me.cmdEmpClear.Visible = False
    Else
        Me.fgEmp.Visible = True
        Me.cmdEmpAll.Visible = True
        Me.cmdEmpClear.Visible = True
    End If

End Sub
