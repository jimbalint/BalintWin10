VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDeposit 
   Caption         =   "Deposit Listing"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDateRange 
      Caption         =   "&Date Range"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   1020
      Width           =   975
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      HideSelection   =   0   'False
      Left            =   3533
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Options:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2153
      TabIndex        =   7
      Top             =   1920
      Width           =   6615
      Begin VB.CheckBox chkExclUnemp 
         Caption         =   "Exclude Unemployment?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   900
         Width           =   6015
      End
      Begin VB.CheckBox chkFedTaxDep 
         Caption         =   "Federal Tax Deposit Now?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   595
         Width           =   6015
      End
      Begin VB.CheckBox chkCombNetPay 
         Caption         =   "Combine Net Payroll with Tax Escrow?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   290
         Width           =   6135
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8693
      TabIndex        =   6
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   653
      TabIndex        =   5
      Top             =   7440
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   2895
      Left            =   2633
      TabIndex        =   4
      Top             =   3840
      Width           =   5655
      _cx             =   9975
      _cy             =   5106
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      ScrollBars      =   2
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
      Height          =   375
      Left            =   833
      TabIndex        =   10
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label1 
      Caption         =   "Deductions to Include With Escrow:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3293
      TabIndex        =   9
      Top             =   3480
      Width           =   4335
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public deds As New ADODB.Recordset

Private Sub Form_Load()

    ' BatchID assigned? - use it
    If PRBatchID <> 0 Then
        If Not PRBatch.GetByID(PRBatchID) Then
            MsgBox "Batch NF: " & PRBatchID, vbCritical
            End
        End If
        Me.cmdDaterange.Enabled = False
        Me.txtDisplay.Text = "Batch #: " & PRBatch.BatchID & _
                             " PE Date: " & Format(PRBatch.PEDate, "mm/dd/yy") & _
                             " Check Date: " & Format(PRBatch.CheckDate, "mm/dd/yy")
        RangeType = PREquate.RangeTypeBatch
    End If

    ' grid for deductions to add to escrow
    deds.CursorLocation = adUseClient
    deds.Fields.Append "UseDeduction", adBoolean
    deds.Fields.Append "Title", adVarChar, 20, adFldIsNullable
    deds.Fields.Append "ItemID", adDouble
    deds.Fields.Append "Amount", adCurrency
    deds.Fields.Append "DirDep", adBoolean
    deds.Fields.Append "Count", adDouble
   
    deds.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 " & _
                " AND (PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeSDTax & ")" & _
                " ORDER BY PRItem.ItemID"

    If Not PRItem.GetBySQL(SQLString) Then
        Me.fg.Visible = False
    Else
        Do
            deds.AddNew
            If PRItem.Escrow Then deds!UseDeduction = True
            deds.Fields("Title") = Mid(Trim(PRItem.Title), 1, 20)
            deds.Fields("ItemId") = PRItem.ItemID
            
            If PRItem.DirDepRpt = 1 Then
                deds.Fields("DirDep") = True
            Else
                deds.Fields("DirDep") = False
            End If
            
            deds!Amount = 0
            deds!Count = 0
            
            deds.Update
            
            If Not PRItem.GetNext Then
                Exit Do
            End If
        Loop
        
        SetGrid deds, fg
        fg.ColHidden(2) = True
        fg.ColWidth(1) = 5300
    End If
    
    ' screen defaults pre company
    Me.chkCombNetPay = 0
    Me.chkFedTaxDep = 0
    Me.chkExclUnemp = 0
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                " AND Var1 = 'DepositList'"
    If PRGlobal.GetBySQL(SQLString) Then
        If PRGlobal.Var2 = "1" Then Me.chkCombNetPay = 1
        If PRGlobal.Var3 = "1" Then Me.chkFedTaxDep = 1
        If PRGlobal.Var4 = "1" Then Me.chkExclUnemp = 1
    End If
    
    Me.lblCompanyName = PRCompany.Name
    frmDateRange.lblClient = PRCompany.Name
    Me.KeyPreview = True
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdDateRange_Click()
    frmDateRange.lblProgram = "Deposit Listing"
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
    Else
        If OptDate = "CHECK DATE" Then
            txtDisplay = "Check Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        Else
            txtDisplay = "P/E Date Range: " & Format(StartDate, "mm/dd/yyyy") & " - " & Format(EndDate, "mm/dd/yyyy")
        End If
    End If
    PRBatchID = BatchNumbr
    Me.Refresh

End Sub

Private Sub cmdOkay_Click()

    If PRBatchID = 0 And StartDate = 0 And EndDate = 0 And PEDate = 0 Then
        MsgBox "PLEASE SELECT A DATE RANGE", vbExclamation, "Direct Deposit Report"
        Exit Sub
    End If
        
    ' save the screen defaults per company
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                " AND Var1 = 'DepositList'"
    If PRGlobal.GetBySQL(SQLString) = False Then
        PRGlobal.Clear
        PRGlobal.Var1 = "DepositList"
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
    End If
        
    PRGlobal.Var2 = 0
    PRGlobal.Var3 = 0
    PRGlobal.Var4 = 0
    
    If Me.chkCombNetPay = 1 Then PRGlobal.Var2 = "1"
    If Me.chkFedTaxDep = 1 Then PRGlobal.Var3 = "1"
    If Me.chkExclUnemp = 1 Then PRGlobal.Var4 = "1"
    PRGlobal.Save (Equate.RecPut)
        
    InitFlag = True
    Me.Hide
    DepositListing RangeType, PRBatchID, CLng(Int(PEDate)), CLng(Int(CheckDt)), _
                   CLng(Int(StartDate)), CLng(Int(EndDate)), OptDate

End Sub



