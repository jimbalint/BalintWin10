VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmGLUpd 
   Caption         =   "Payroll to GL Update Maintenance"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11145
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
   ScaleHeight     =   9045
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdEmpCatDel 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmpCatAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   7800
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fgPRItem 
      Height          =   6135
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   5895
      _cx             =   10398
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
   Begin VSFlex8Ctl.VSFlexGrid fgEmpCat 
      Height          =   6135
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3855
      _cx             =   6800
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
   Begin VB.Label Label2 
      Caption         =   "Payroll Items"
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Employee Category"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmGLUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsEmpCat As New ADODB.Recordset
Dim rsPRItem As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim GLDrop As String
Dim InitFlag As Boolean

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    InitFlag = True
    
    LoadData
    GridSetup
    
    rsEmpCat.MoveFirst
    LoadPRItemGrid
    rsPRItem.MoveFirst
    
    InitFlag = False

    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub LoadData()

Dim x As String

    ' define EmpCat record set
    rsEmpCat.CursorLocation = adUseClient
    
    rsEmpCat.Fields.Append "sOrder", adVarChar, 20, adFldIsNullable
    rsEmpCat.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rsEmpCat.Fields.Append "EmpCatType", adInteger
    rsEmpCat.Fields.Append "RelatedID", adDouble
    
    rsEmpCat.Open , , adOpenDynamic, adLockOptimistic
    
    ' define PRItem record set
    rsPRItem.CursorLocation = adUseClient
    
    rsPRItem.Fields.Append "sOrder", adInteger
    rsPRItem.Fields.Append "Title", adVarChar, 20, adFldIsNullable
    rsPRItem.Fields.Append "GLAcctNum", adDouble
    rsPRItem.Fields.Append "GLItemType", adInteger
    rsPRItem.Fields.Append "ItemID", adDouble

    rsPRItem.Open , , adOpenDynamic, adLockOptimistic

    ' loop thru the existing file and load up the different emp categories defined
    SQLString = "SELECT * FROM PRGLUpd"
    If PRGLUpd.GetBySQL(SQLString) Then
    
        Do
            If PRGLUpd.GLType = PREquate.GLTypeEmployee Then
                x = "A" & Format(PRGLUpd.RelatedID, "000000")
            ElseIf PRGLUpd.GLType = PREquate.GLTypeDept Then
                x = "B" & Format(PRGLUpd.RelatedID, "000000")
            ElseIf PRGLUpd.GLType = PREquate.GLTypeCompany Then
                x = "C000000"
            Else
                MsgBox "Invalid PRGLUpd.GLType: " & PRGLUpd.GLType & " " & PRGLUpd.GLUpdID, vbCritical
                End
            End If
                    
            SQLString = "sOrder = '" & Trim(x) & "'"
            Debug.Print SQLString
            rsEmpCat.Find SQLString, 0, adSearchForward, 1
            If rsEmpCat.EOF Then
                rsEmpCat.AddNew
                rsEmpCat!sOrder = x
                If PRGLUpd.GLType = PREquate.GLTypeEmployee Then
                    If Not PREmployee.GetByID(PRGLUpd.RelatedID) Then
                        MsgBox "Employee ID Not Found: " & PRGLUpd.RelatedID, vbCritical
                        End
                    End If
                    rsEmpCat!Title = "EE#: " & PREmployee.EmployeeNumber & " " & PREmployee.LFName
                    rsEmpCat!EmpCatType = PREquate.GLTypeEmployee
                    rsEmpCat!RelatedID = PREmployee.EmployeeID
                ElseIf PRGLUpd.GLType = PREquate.GLTypeDept Then
                    If Not PRDepartment.GetByID(PRGLUpd.RelatedID) Then
                        MsgBox "Department ID Not Found: " & PRGLUpd.RelatedID, vbCritical
                        End
                    End If
                    rsEmpCat!Title = "DPT#: " & PRDepartment.DepartmentNumber & " " & PRDepartment.Name
                    rsEmpCat!EmpCatType = PREquate.GLTypeDept
                    rsEmpCat!RelatedID = PRDepartment.DepartmentID
                ElseIf PRGLUpd.GLType = PREquate.GLTypeCompany Then
                    rsEmpCat!Title = "COMPANY"
                    rsEmpCat!EmpCatType = PREquate.GLTypeCompany
                    rsEmpCat!RelatedID = 0
                End If
                
                rsEmpCat.Update
            End If
        
            If Not PRGLUpd.GetNext Then Exit Do
        
        Loop
    
    Else    ' none defined - set up the company record
    
        rsEmpCat.AddNew
        rsEmpCat!sOrder = "C000000"
        rsEmpCat!Title = "COMPANY"
        rsEmpCat!EmpCatType = PREquate.GLTypeCompany
        rsEmpCat!RelatedID = 0
        rsEmpCat.Update
    
    End If

    ' load the GL Accounts to a drop down string
    GLAccount.OpenRS
    If GLAccount.Account = 0 Then
        MsgBox "No GL Account records found!", vbCritical
        End
    End If
    
    Do
        If GLAccount.AcctType = "0" Then
            GLDrop = Trim(GLDrop) & "|#" & GLAccount.Account & ";" & GLAccount.Account & " " & Trim(GLAccount.FullDesc)
        End If
        If Not GLAccount.GetNext Then Exit Do
    Loop

    rsEmpCat.Sort = "sOrder"

End Sub

                           
Private Sub GridSetup()

    ' left side - fgEmpCat
    SetGrid rsEmpCat, fgEmpCat
    
    With fgEmpCat
        .ColHidden(0) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColWidth(1) = 3750
    End With
    
    ' right side - fgPRItem
    SetGrid rsPRItem, fgPRItem
    
    With fgPRItem
        .ColHidden(0) = True
        .ColHidden(3) = True
        .ColComboList(2) = GLDrop
        .ColWidth(1) = 2000
        .ColWidth(2) = 3580
    End With
    
End Sub

Private Sub LoadPRItemGrid()

    ' the rsEmpCat record has been added but not yet assigned
    If rsEmpCat!EmpCatType = 0 Then Exit Sub
    
    ' clear out the temp record set
    If rsPRItem.RecordCount > 0 Then
        rsPRItem.MoveFirst
        Do Until rsPRItem.RecordCount = 0
            rsPRItem.Delete
            rsPRItem.MoveNext
        Loop
    End If
        
    ' other earnings
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " AND PRItem.EmployeeID = 0" & _
                " ORDER BY PRItem.ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do
            x = "1" & Format(PRItem.ItemID, "000000")
            AddPRItem x, PRItem.Title, PREquate.GLItemTypeOE, PRItem.ItemID
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If
    
    ' deductions
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " AND PRItem.EmployeeID = 0" & _
                " ORDER BY PRItem.ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do
            x = "2" & Format(PRItem.ItemID, "000000")
            AddPRItem x, PRItem.Title, PREquate.GLItemTypeDed, PRItem.ItemID
            If Not PRItem.GetNext Then Exit Do
        Loop
    End If
    
    AddPRItem "3000000", "SS TAX", PREquate.GLItemTypeSSTax, 0
    AddPRItem "4000000", "MED TAX", PREquate.GLItemTypeMedTax, 0
    AddPRItem "5000000", "FWT TAX", PREquate.GLItemTypeFWTTax, 0
    AddPRItem "6000000", "SWT TAX", PREquate.GLItemTypeSWTTax, 0
    AddPRItem "7000000", "CWT TAX", PREquate.GLItemTypeCWTTax, 0
    AddPRItem "8000000", "FUN TAX", PREquate.GLItemTypeFUN, 0
    AddPRItem "9000000", "SUN TAX", PREquate.GLItemTypeSUN, 0
    AddPRItem "10000000", "GROSS PAY", PREquate.GLItemTypeGross, 0
    AddPRItem "11000000", "NET PAY", PREquate.GLItemTypeNet, 0
    AddPRItem "12000000", "SS EXP", PREquate.GLItemTypeSSExp, 0
    AddPRItem "13000000", "MED EXP", PREquate.GLItemTypeMEDExp, 0
    AddPRItem "14000000", "FUN EXP", PREquate.GLItemTypeFUNExp, 0
    AddPRItem "15000000", "SUN EXP", PREquate.GLItemTypeSUNExp, 0
    AddPRItem "16000000", "WKC EXP", PREquate.GLItemTypeWkcExp, 0
                
End Sub
Private Sub AddPRItem(ByVal sOrder As String, _
                           ByVal Title As String, _
                           ByVal GLItemType As Byte, _
                           ByVal ItemID As Long)

    rsPRItem.AddNew
    rsPRItem!sOrder = sOrder
    rsPRItem!Title = Title
    
    SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & rsEmpCat!EmpCatType & _
                " AND RelatedID = " & rsEmpCat!RelatedID & _
                " AND GLItemType = " & GLItemType
                
    If ItemID <> 0 Then
        SQLString = Trim(SQLString) & "  AND ItemID = " & ItemID
    End If
    
    If PRGLUpd.GetBySQL(SQLString) Then
        rsPRItem!GLAcctNum = PRGLUpd.GLAccountNum
    Else
        rsPRItem!GLAcctNum = 0
    End If
    rsPRItem!GLItemType = GLItemType
    rsPRItem!ItemID = ItemID
    rsPRItem.Update

End Sub

Private Sub cmdEmpCatDel_Click()
    If rsEmpCat!EmpCatType = PREquate.GLTypeCompany Then
        MsgBox "Company record can not be deleted!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("OK to delete " & rsEmpCat!Title, vbQuestion + vbOKCancel, "PR to GL Update Maint") = vbCancel Then
        Exit Sub
    End If
    
    rsEmpCat.Delete
    rsEmpCat.MoveFirst

End Sub
Private Sub fgEmpCat_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    
    ' save the data if not initial form load
    If InitFlag Then Exit Sub
    
    ' clear the data for the Emp Cat from PRGLUpd
    SQLString = "DELETE * FROM PRGLUpd WHERE GLType = " & fgEmpCat.TextMatrix(fgEmpCat.Row, 2) & _
                " AND RelatedID = " & fgEmpCat.TextMatrix(fgEmpCat.Row, 3)
    
    rsInit SQLString, cn, rs
    
    ' add the data in from the temp record sets
    rsPRItem.MoveFirst
    Do
        PRGLUpd.Clear
        PRGLUpd.GLType = fgEmpCat.TextMatrix(fgEmpCat.Row, 2)
        PRGLUpd.RelatedID = fgEmpCat.TextMatrix(fgEmpCat.Row, 3)
        PRGLUpd.GLItemType = rsPRItem!GLItemType
        PRGLUpd.ItemID = rsPRItem!ItemID
        PRGLUpd.GLAccountNum = rsPRItem!GLAcctNum
        PRGLUpd.Title = rsPRItem!Title
        PRGLUpd.Save (Equate.RecAdd)
        rsPRItem.MoveNext
        If rsPRItem.EOF Then Exit Do
    Loop

End Sub
Private Sub fgEmpCat_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    LoadPRItemGrid
    rsPRItem.MoveFirst
End Sub
Private Sub cmdExit_Click()
    ' >>> save final setup ...
    GoBack
End Sub

Private Sub cmdEmpCatAdd_Click()
    
    frmGLUpdAdd.Show vbModal
    
    ' user hit the exit button
    If frmGLUpdAdd.EmpCatType = 0 Then
        Unload frmGLUpdAdd
        Exit Sub
    End If

    ' add the selection to the rsEmpCat record set
    ' see if it already exists ...
    rsEmpCat.MoveFirst
    Do
        If rsEmpCat!EmpCatType = frmGLUpdAdd.EmpCatType And rsEmpCat!RelatedID = frmGLUpdAdd.RecId Then
            Exit Sub    ' will leave with pointed to that record on fgEmpCat
        End If
        rsEmpCat.MoveNext
        If rsEmpCat.EOF Then Exit Do
    Loop
    
    ' OK to add it
    rsEmpCat.AddNew
    If frmGLUpdAdd.EmpCatType = PREquate.GLTypeEmployee Then
        rsEmpCat!sOrder = "A" & Format(frmGLUpdAdd.RecId, "000000")
        rsEmpCat!Title = "EE#: " & frmGLUpdAdd.Number & " " & frmGLUpdAdd.Title
    Else
        rsEmpCat!sOrder = "B" & Format(frmGLUpdAdd.RecId, "000000")
        rsEmpCat!Title = "DPT#: " & frmGLUpdAdd.Number & " " & frmGLUpdAdd.Title
    End If
    rsEmpCat!EmpCatType = frmGLUpdAdd.EmpCatType
    rsEmpCat!RelatedID = frmGLUpdAdd.RecId
    rsEmpCat.Update
    
    rsEmpCat.Sort = "sOrder"
    
    ' find it again ???
    
End Sub


