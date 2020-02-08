VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmItemDetail 
   Caption         =   "Item Detail Report"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8340
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShowRemain 
      Caption         =   "Show Remaining Value"
      Height          =   375
      Left            =   3683
      TabIndex        =   23
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "&DATE RANGE"
      Height          =   615
      Left            =   1523
      TabIndex        =   22
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtDisplay 
      Height          =   615
      Left            =   2723
      TabIndex        =   21
      Top             =   600
      Width           =   5775
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   5483
      TabIndex        =   16
      Top             =   2160
      Width           =   3495
      Begin VB.CommandButton cmdSelDed 
         Caption         =   "Select &All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   650
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearDed 
         Caption         =   "C&lear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   640
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "OE && Deduction Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   1950
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Please Select Up To FIVE (5)  Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   2940
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   735
      Left            =   1472
      TabIndex        =   12
      Top             =   1320
      Width           =   4215
      Begin VB.OptionButton optEmpNo 
         BackColor       =   &H80000016&
         Caption         =   "&Employee Number"
         Height          =   245
         Left            =   240
         TabIndex        =   15
         Top             =   400
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optChkDate 
         BackColor       =   &H80000016&
         Caption         =   "Check &Date"
         Height          =   245
         Left            =   2520
         TabIndex        =   13
         Top             =   400
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Order By"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Height          =   855
      Left            =   1043
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
      Begin VB.CommandButton cmdClearAll 
         BackColor       =   &H80000014&
         Caption         =   "&Clear All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelectAll 
         BackColor       =   &H80000016&
         Caption         =   "&Select All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Employee Listing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Height          =   735
      Left            =   6253
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
      Begin VB.Label lblEmpCount 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         Height          =   200
         Left            =   480
         TabIndex        =   11
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Employees Selected"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1740
      End
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
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   7440
      Width           =   2175
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
      Height          =   495
      Left            =   6983
      TabIndex        =   2
      Top             =   7440
      Width           =   2175
   End
   Begin VSFlex8Ctl.VSFlexGrid fgEmp 
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   4335
      _cx             =   7646
      _cy             =   5953
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
   Begin VSFlex8Ctl.VSFlexGrid fgItem 
      Height          =   3375
      Left            =   5160
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
      _cx             =   7646
      _cy             =   5953
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
      Caption         =   "COMPANY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   220
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmItemDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsEmp As New ADODB.Recordset
Public RSItem As New ADODB.Recordset
Public PEDate As Long
Public CheckDt As Long
Public EmpCount As Long
Public NoItems As Long

Private Sub Form_Load()
    Load_Grids
    Me.lblCompanyName.Caption = PRCompany.Name
    Me.KeyPreview = True

End Sub

Private Sub fgEmp_LostFocus()
    EmpCount = 0
    rsEmp.MoveFirst
    Do
        If rsEmp!Selected = True Then
            EmpCount = EmpCount + 1
        End If
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
    
End Sub

Private Sub fgItem_LostFocus()
    RSItem.MoveFirst
    NoItems = 0
    Do
        If RSItem!Selected = True Then
            NoItems = NoItems + 1
        End If
        RSItem.MoveNext
    Loop Until RSItem.EOF
    RSItem.MoveFirst
    If NoItems > 5 Then
        MsgBox "Please select ONLY FIVE (5) Items", vbCritical, "Item Detail Report"
        GoBack
    End If

End Sub

Private Sub cmdDateRange_Click()
    frmDateRange.lblProgram = "Item Detail"
    frmDateRange.Show vbModal
        
    If frmDateRange.optCheckDate = True Then
        OptDate = "CHECK DATE"
    ElseIf frmDateRange.optPEDate = True Then
        OptDate = "P/E DATE"
    End If
        
    If InitFlag = False Then Exit Sub   ' user exited
    
    If BatchNumbr > 0 Then
        If Not PRBatch.GetByID(BatchNumbr) Then
            MsgBox "PRBatch Not Found: " & BatchNumbr, vbCritical
            End
        End If
        PEDate = PRBatch.PEDate
        CheckDt = PRBatch.CheckDate
        OptDate = " "
        txtDisplay = "Batch: " & BatchNumbr & "  Period Ending: " & CDate(PEDate) & _
                     "  CheckDate: " & CDate(CheckDt)
        RangeType = PREquate.RangeTypeBatch

    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(Startdate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(Startdate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    PRBatchID = BatchNumbr
    Me.Refresh
End Sub

Private Sub cmdClearAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Selected = False
        rsEmp.Update
        rsEmp.MoveNext
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
End Sub

Private Sub cmdSelectAll_Click()
    rsEmp.MoveFirst
    Do
        rsEmp!Selected = True
        rsEmp.Update
        rsEmp.MoveNext
        
    Loop Until rsEmp.EOF
    rsEmp.MoveFirst
    lblEmpCount = "All Employees"
End Sub

Public Sub Load_Grids()
'  Loop Through Employee File for all Employees
    rsEmp.CursorLocation = adUseClient
    rsEmp.Fields.Append "Selected", adBoolean
    rsEmp.Fields.Append "EmpNo", adDouble
    rsEmp.Fields.Append "EmpName", adVarChar, 80, adFldIsNullable
    rsEmp.Fields.Append "EmpID", adDouble
    Me.lblCompanyName.Caption = PRCompany.Name
    
    rsEmp.Open , , adOpenDynamic, adLockOptimistic
    SQLString = "Select * from PREmployee ORDER BY LastName, FirstName"
    If Not PREmployee.GetBySQL(SQLString) Then
        MsgBox "No Employee Records were Found: ", vbCritical
        GoBack
    End If
    
    Do
        rsEmp.AddNew
        rsEmp!Selected = True
        rsEmp!EmpNo = PREmployee.EmployeeNumber
        rsEmp!EmpName = PREmployee.LFName
        rsEmp!EmpID = PREmployee.EmployeeID
        rsEmp.Update
        
        If Not PREmployee.GetNext Then Exit Do
    Loop
    SetGrid rsEmp, fgEmp
    
'  Loop Through PRItemHist for all Items    **********************************
    RSItem.CursorLocation = adUseClient
    RSItem.Fields.Append "Selected", adBoolean
    RSItem.Fields.Append "Type", adVarChar, 20
    RSItem.Fields.Append "Description", adVarChar, 80, adFldIsNullable
    RSItem.Fields.Append "ItemID", adDouble
    RSItem.Fields.Append "IsItHours", adBoolean
    RSItem.Fields.Append "MaxAmount", adCurrency
    
    RSItem.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " OR PRItem.itemtype = " & PREquate.ItemTypeSDTax & ")" & _
                " ORDER BY PRItem.ItemType, PRItem.ItemID"

    If Not PRItem.GetBySQL(SQLString) Then
'        Me.fgitem.Visible = False
    Else
        Do
            RSItem.AddNew
            RSItem!Selected = False
            RSItem.Fields("Type") = PRItem.ItemType
            RSItem.Fields("ItemID") = Trim(PRItem.ItemID)
            RSItem.Fields("Description") = PRItem.Abbreviation
             
            RSItem.Update
            If PRItem.ItemType = PREquate.ItemTypeOE Then
                RSItem.AddNew
                RSItem!Selected = False
                RSItem.Fields("Type") = 5
                RSItem.Fields("ItemID") = Trim(PRItem.ItemID)
                RSItem.Fields("Description") = PRItem.Abbreviation
                RSItem.Fields("IsItHours") = True
                RSItem.Update
            End If
            RSItem.Fields("MaxAmount") = PRItem.MaxAmount
            If Not PRItem.GetNext Then Exit Do
        Loop

    End If
    
    frmDateRange.lblClient = PRCompany.Name
    Me.KeyPreview = True
    SetGrid RSItem, fgItem
    fgItem.ColComboList(1) = "|#3;OTH EARN|#4;DEDUCT|#5;SD TAX|#6;OTH HRS"
    fgItem.ColWidth(1) = 1000
    fgItem.ColWidth(2) = 2800
    
End Sub

Private Sub cmdClearDed_Click()
    RSItem.MoveFirst
    Do
        RSItem!Selected = False
        RSItem.Update
        RSItem.MoveNext
    Loop Until RSItem.EOF
    RSItem.MoveFirst
End Sub

Private Sub cmdOK_Click()
    fgEmp_LostFocus
    fgItem_LostFocus
    If chkShowRemain = 1 Then
        If NoItems > 1 Then
            MsgBox "Please select ONLY ONE Item", vbExclamation, "Item Detail Report"
            GoBack
        End If
    End If
    If PRBatchID = 0 And Startdate = 0 And EndDate = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbExclamation, "Item Detail Report"
        GoBack
    ElseIf EmpCount = 0 Then
        MsgBox "Please select AT LEAST ONE Employee", vbExclamation, "Item Detail Report"
        GoBack
    ElseIf NoItems = 0 Then
        MsgBox "Please select AT LEAST ONE Item", vbExclamation, "Item Detail Report"
        GoBack
    Else
        InitFlag = True
        Me.Hide
        ItemDetail RangeType, PRBatchID, CLng(PEDate), CLng(CheckDt), CLng(Startdate), CLng(EndDate), OptDate
    End If
End Sub


Private Sub cmdSelDed_Click()
    RSItem.MoveFirst
    Do
        RSItem!Selected = True
        RSItem.Update
        RSItem.MoveNext
    Loop Until RSItem.EOF
    RSItem.MoveFirst
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub optChkDate_Click()
    
    ' can't show remaining if by date
    If optChkDate = True Then
        Me.chkShowRemain = 0
        Me.chkShowRemain.Enabled = False
    Else
        Me.chkShowRemain.Enabled = True
    End If

End Sub

Private Sub optEmpNo_Click()
    If optEmpNo = True Then
        Me.chkShowRemain.Enabled = True
    End If
End Sub
