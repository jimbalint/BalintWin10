VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmGLUpd 
   Caption         =   "Payroll to GL Update Maintenance"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11460
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
   ScaleHeight     =   9105
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fgGLType 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      _cx             =   7858
      _cy             =   12726
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
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7215
      Left            =   4920
      TabIndex        =   1
      Top             =   840
      Width           =   6375
      _cx             =   11245
      _cy             =   12726
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGLTypeDel 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdGLTypeAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   8400
      Width           =   1215
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
      TabIndex        =   5
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

Public trsGLT As New ADODB.Recordset
Dim trs As New ADODB.Recordset
Dim trsCopyFrom As New ADODB.Recordset
Dim TypeDrop, GLDrop As String
Dim InitFlag As Boolean
Dim Flg As Boolean
Dim GName As String
Dim EditFlag As Boolean
Dim ReportTitle As String
Dim LastID, LastType As Byte

Dim X, Y, z As String

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    InitFlag = True
    
    LoadData
    GridSetup
    
    
'  trsGLT.MoveFirst
'  Do
'      MsgBox trsGLT!GLType & vbCr & trsGLT!RelatedID & vbCr & trsGLT!GLName
'      trsGLT.MoveNext
'  Loop Until trsGLT.EOF
  trsGLT.MoveFirst
    
    InitFlag = False
    
    PopAcctGrid trsGLT!GLType, trsGLT!RelatedID, 0
    
    trsGLT.Sort = "GLType, RelatedID"
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub LoadData()

Dim X As String

    ' temp record set to track each GLType in the file
    trsGLT.CursorLocation = adUseClient
    
    trsGLT.Fields.Append "GLType", adInteger
    trsGLT.Fields.Append "RelatedID", adDouble
    trsGLT.Fields.Append "GLName", adVarChar, 30
    trsGLT.Fields.Append "OrigType", adInteger
    trsGLT.Open , , adOpenDynamic, adLockOptimistic

    SQLString = "SELECT * FROM PRGLUpd ORDER BY GLType, RelatedID"
    If Not PRGLUpd.GetBySQL(SQLString) Then
    
        trsGLT.AddNew
        trsGLT!GLType = PREquate.GLTypeCompany
        trsGLT!RelatedID = 0
        trsGLT!GLName = "COMPANY"
        trsGLT.Update
    
    Else
    
        ' load up the different types
        LastType = 0
        LastID = 0
        Do
            If LastType = 0 Or LastType <> PRGLUpd.GLType Or LastID <> PRGLUpd.RelatedID Then
                trsGLT.AddNew
                trsGLT!GLType = PRGLUpd.GLType
                trsGLT!RelatedID = PRGLUpd.RelatedID
                trsGLT!GLName = Mid(GetGLName(PRGLUpd.GLType, PRGLUpd.RelatedID), 1, 30)
                trsGLT.Update
            End If
            LastType = PRGLUpd.GLType
            LastID = PRGLUpd.RelatedID
            If Not PRGLUpd.GetNext Then Exit Do
        Loop
    
    End If
                
    ' create company record if DNE
    SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & PREquate.GLTypeCompany
    If PRGLUpd.GetBySQL(SQLString) = False Then
        trsGLT.AddNew
        trsGLT!GLType = PREquate.GLTypeCompany
        trsGLT!RelatedID = 0
        trsGLT!GLName = "COMPANY"
        trsGLT.Update
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

    trsGLT.Sort = "GLType, RelatedID"
    trsGLT.MoveFirst

End Sub
                          
Private Sub GridSetup()

    TypeDrop = "|#1;EMPLOYEE|#2;DEPARTMENT|#3;COMPANY|#4;*DELETED*"

    ' Left Side - GLType Grid
    trsGLT.Sort = "GLType, RelatedID"
    trsGLT.MoveFirst
    SetGrid trsGLT, fgGLType

    With fgGLType
    
        .ColComboList(0) = TypeDrop
        
        .ColHidden(1) = True
    
        .ColWidth(0) = 1500
        .ColWidth(2) = 2900
    
        .Editable = flexEDNone
        .AllowSelection = False
        .SelectionMode = flexSelectionByRow
    
    End With
    
End Sub
                          
Private Sub MakeGLType(ByVal GLType As Byte, ByVal RelatedID As Long, ByVal CopyFromID As Long)

'    ' define the copy from record set
'    ' record count will be zero if not used
'    On Error Resume Next
'    trsCopyFrom.Close
'    On Error GoTo 0
'    trsCopyFrom.CursorLocation = adUseClient
'    trsCopyFrom.Fields.Append "GLItemType", adInteger
'    trsCopyFrom.Fields.Append "ItemID", adDouble
'    trsCopyFrom.Fields.Append "GLAccountNum", adDouble
'    trsCopyFrom.Open , , adOpenDynamic, adLockOptimistic
'
'    ' make a copy from record set
'    If CopyFromID <> 0 Then
'
'        ' get the copy from info
'        trs.MoveFirst
'        Do
'            If trs!GLType = frmGLUpdAdd.CopyFromType And _
'               trs!RelatedID = frmGLUpdAdd.CopyFromID Then
'
'                trsCopyFrom.AddNew
'                trsCopyFrom!GLItemType = trs!GLItemType
'                trsCopyFrom!ItemID = trs!ItemID
'                trsCopyFrom!GLAccountNum = trs!GLAccountNum
'                trsCopyFrom.Update
'            End If
'            trs.MoveNext
'        Loop Until trs.EOF
'
'    End If
'
'    ' Regular / OVT
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeRegPay, 0, "REG PAY", "1000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeOvtPay, 0, "OVT PAY", "2000000"
'
'    ' add records to trs - read from PRGLUpd if the record exists
'    ' other earnings
'    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeOE & _
'                " AND PRItem.EmployeeID = 0" & _
'                " ORDER BY PRItem.ItemID"
'    If PRItem.GetBySQL(SQLString) Then
'        Do
'
'            X = "A" & Format(PRItem.ItemID, "000000")
'            AddPRItem GLType, RelatedID, PREquate.GLItemTypeOE, PRItem.ItemID, PRItem.Title, X
'
'            If Not PRItem.GetNext Then Exit Do
'
'        Loop
'    End If
'
'    ' deductions
'    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDED & _
'                " AND PRItem.EmployeeID = 0" & _
'                " ORDER BY PRItem.ItemID"
'    If PRItem.GetBySQL(SQLString) Then
'        Do
'
'            X = "B" & Format(PRItem.ItemID, "000000")
'            AddPRItem GLType, RelatedID, PREquate.GLItemTypeDed, PRItem.ItemID, PRItem.Title, X
'
'            If Not PRItem.GetNext Then Exit Do
'
'        Loop
'    End If
'
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSTax, 0, "SS TAX", "C000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSMatch, 0, "SS MATCH", "D000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMedTax, 0, "MED TAX", "E000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMedMatch, 0, "MED MATCH", "F000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFWTTax, 0, "FWT TAX", "G000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSWTTax, 0, "SWT TAX", "H000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeCWTTax, 0, "CWT TAX", "I000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFUN, 0, "FUN TAX", "J000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSUN, 0, "SUN TAX", "K000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeWkcTax, 0, "WKC TAX", "L000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeGross, 0, "GROSS PAY", "M000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeNet, 0, "NET PAY", "N000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSExp, 0, "SS EXPENSE", "O000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMEDExp, 0, "MED EXPENSE", "P000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFUNExp, 0, "FUN EXPENSE", "Q000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSUNExp, 0, "SUN EXPENSE", "R000000"
'    AddPRItem GLType, RelatedID, PREquate.GLItemTypeWkcExp, 0, "WKC EXPENSE", "S000000"
'
'    ' fill in account numbers and PRGLUpd RecIDs
'    trs.Filter = "GLType = " & GLType & " AND RelatedID = " & RelatedID
'    If trs.RecordCount = 0 Then ' ???
'        trs.Filter = adFilterNone
'        Exit Sub
'    End If
'    trs.MoveFirst
'    Do
'        SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & GLType & _
'                    " AND RelatedID = " & RelatedID & _
'                    " AND GLItemType = " & trs!GLItemType & _
'                    " AND ItemID = " & trs!ItemID
'        If PRGLUpd.GetBySQL(SQLString) Then
'            trs!GLAccountNum = PRGLUpd.GLAccountNum
'            trs!GLUpdID = PRGLUpd.GLUpdID
'            trs.Update
'        End If
'        trs.MoveNext
'    Loop Until trs.EOF
'
'    trs.Filter = adFilterNone

End Sub

Private Sub cmdExit_Click()
    On Error Resume Next
    GName = trsGLT!GLName
    On Error GoTo 0
    SaveFG
    GoBack
End Sub
Private Sub cmdGLTypeDel_Click()

    If trsGLT!GLType = PREquate.GLTypeCompany Then
        MsgBox "Company record delete not allowed!", vbInformation
        Exit Sub
    End If

    If MsgBox("OK to delete: " & fgGLType.TextMatrix(fgGLType.Row, 2) & "?", vbQuestion + vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    trs.MoveFirst
    Do
        If trs!GLUpdID <> 0 Then
            SQLString = "DELETE * FROM PRGLUpd WHERE GLUpdID = " & trs!GLUpdID
            cn.Execute SQLString
        End If
        trs.MoveNext
    Loop Until trs.EOF
    
    trsGLT.Delete
    trsGLT.MoveFirst
    
    PopAcctGrid trsGLT!GLType, trsGLT!RelatedID, 0

End Sub

Private Sub cmdGLTypeAdd_Click()
    
Dim FFlag As Boolean
    
    On Error Resume Next
    GName = trsGLT!GLName
    On Error GoTo 0
    SaveFG
    
    InitFlag = True
    
    frmGLUpdAdd.Show vbModal
    
    ' user hit the exit button
    If frmGLUpdAdd.GLType = 0 Then
        Unload frmGLUpdAdd
        Exit Sub
    End If

    ' see if the selection already exists
    trsGLT.Filter = "GLType = " & frmGLUpdAdd.GLType & " AND RelatedID = " & frmGLUpdAdd.RecID
    If trsGLT.RecordCount > 0 Then
        trsGLT.Filter = adFilterNone
        Unload frmGLUpdAdd
        Exit Sub
    Else
        trsGLT.Filter = adFilterNone
    End If
    
    ' add to the left side grid RS
    trsGLT.AddNew
    trsGLT!GLType = frmGLUpdAdd.GLType
    trsGLT!RelatedID = frmGLUpdAdd.RecID
    trsGLT!GLName = GetGLName(frmGLUpdAdd.GLType, frmGLUpdAdd.RecID)
    trsGLT.Update
    trsGLT.Sort = "GLType, RelatedID"
    
    PopAcctGrid trsGLT!GLType, trsGLT!RelatedID, frmGLUpdAdd.CopyFromID
    
    InitFlag = False
    EditFlag = True
    GName = trsGLT!GLName
    SaveFG
    InitFlag = True
    EditFlag = False
    
    ' point to the record just added
    trsGLT.MoveFirst
    Do
        If trsGLT!GLType = frmGLUpdAdd.GLType And trsGLT!RelatedID = frmGLUpdAdd.RecID Then
            Exit Do
        End If
        trsGLT.MoveNext
        If trsGLT.EOF Then
            MsgBox "EOF?"
            End
        End If
    Loop
    
    Unload frmGLUpdAdd
    
    InitFlag = False

End Sub

Private Function GetGLName(ByVal GLType As Byte, ByVal RelatedID As Long) As String

    If GLType = PREquate.GLTypeEmployee Then
        If Not PREmployee.GetByID(RelatedID) Then
            MsgBox "Employee ID NF: " & RelatedID, vbCritical
            End
        End If
        GetGLName = "EE#" & PREmployee.EmployeeNumber & " " & PREmployee.LFName
    ElseIf GLType = PREquate.GLTypeDept Then
        If Not PRDepartment.GetByID(RelatedID) Then
            MsgBox "Department ID NF: " & RelatedID, vbCritical
            End
        End If
        GetGLName = "DPT #" & PRDepartment.DepartmentNumber & " " & PRDepartment.Name
    ElseIf GLType = PREquate.GLTypeCompany Then
        GetGLName = "COMPANY"
    Else
        MsgBox "GLType Error: " & GLType, vbCritical
        End
    End If

End Function
Private Sub PopAcctGrid(ByVal GLType As Byte, ByVal RelatedID As Long, ByVal CopyFromID As Long)
    
    ' save from trs
    SaveFG
    
    ' define EmpCat record set
    On Error Resume Next
    trs.Close
    On Error GoTo 0
    
    ' 2016-03-19
    Set trs = New ADODB.Recordset
    trs.CursorLocation = adUseClient
    trs.Fields.Append "GLType", adInteger
    trs.Fields.Append "RelatedID", adDouble
    trs.Fields.Append "ItemTitle", adVarChar, 30, adFldIsNullable
    trs.Fields.Append "ItemID", adDouble
    trs.Fields.Append "GLItemType", adInteger
    trs.Fields.Append "GLAccountNum", adDouble
    trs.Fields.Append "GLUpdID", adDouble
    trs.Fields.Append "sOrder", adVarChar, 20, adFldIsNullable
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    ' define the copy from record set
    ' record count will be zero if not used
    On Error Resume Next
    trsCopyFrom.Close
    On Error GoTo 0
    
    Set trsCopyFrom = New ADODB.Recordset
    trsCopyFrom.CursorLocation = adUseClient
    trsCopyFrom.Fields.Append "GLItemType", adInteger
    trsCopyFrom.Fields.Append "ItemID", adDouble
    trsCopyFrom.Fields.Append "GLAccountNum", adDouble
    trsCopyFrom.Open , , adOpenDynamic, adLockOptimistic

    ' make a copy from record set
    If CopyFromID <> 0 Then

        SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & frmGLUpdAdd.GLType & _
                    " AND RelatedID = " & CopyFromID
        If PRGLUpd.GetBySQL(SQLString) = True Then

            Do

                trsCopyFrom.AddNew
                trsCopyFrom!GLItemType = PRGLUpd.GLItemType
                trsCopyFrom!ItemID = PRGLUpd.ItemID
                trsCopyFrom!GLAccountNum = PRGLUpd.GLAccountNum
                trsCopyFrom.Update
            
                If PRGLUpd.GetNext = False Then Exit Do
                
            Loop
        End If
    End If
        
    ' Regular / OVT
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeRegPay, 0, "REG PAY", "1000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeOvtPay, 0, "OVT PAY", "2000000"
        
    ' add records to trs - read from PRGLUpd if the record exists
    ' other earnings
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeOE & _
                " AND PRItem.EmployeeID = 0" & _
                " ORDER BY PRItem.ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do
            
            X = "A" & Format(PRItem.ItemID, "000000")
            AddPRItem GLType, RelatedID, PREquate.GLItemTypeOE, PRItem.ItemID, PRItem.Title, X
            
            If Not PRItem.GetNext Then Exit Do
            
        Loop
    End If
    
    ' deductions
    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " AND PRItem.EmployeeID = 0" & _
                " ORDER BY PRItem.ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do

            X = "B" & Format(PRItem.ItemID, "000000")
            AddPRItem GLType, RelatedID, PREquate.GLItemTypeDed, PRItem.ItemID, PRItem.Title, X

            If Not PRItem.GetNext Then Exit Do

        Loop
    End If
    
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSTax, 0, "SS TAX", "C000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSMatch, 0, "SS MATCH", "D000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMedTax, 0, "MED TAX", "E000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMedMatch, 0, "MED MATCH", "F000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFWTTax, 0, "FWT TAX", "G000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSWTTax, 0, "SWT TAX", "H000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeCWTTax, 0, "CWT TAX", "I000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFUN, 0, "FUN TAX", "J000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSUN, 0, "SUN TAX", "K000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeWkcTax, 0, "WKC TAX", "L000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeGross, 0, "GROSS PAY", "M000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeNet, 0, "NET PAY", "N000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSSExp, 0, "SS EXPENSE", "O000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeMEDExp, 0, "MED EXPENSE", "P000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeFUNExp, 0, "FUN EXPENSE", "Q000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeSUNExp, 0, "SUN EXPENSE", "R000000"
    AddPRItem GLType, RelatedID, PREquate.GLItemTypeWkcExp, 0, "WKC EXPENSE", "S000000"
    
    trs.MoveFirst
    Do
        SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & GLType & _
                    " AND RelatedID = " & RelatedID & _
                    " AND GLItemType = " & trs!GLItemType & _
                    " AND ItemID = " & trs!ItemID
        If PRGLUpd.GetBySQL(SQLString) Then
            trs!GLAccountNum = PRGLUpd.GLAccountNum
            trs!GLUpdID = PRGLUpd.GLUpdID
            trs.Update
        End If
        trs.MoveNext
    Loop Until trs.EOF

    trs.Sort = "GLType, RelatedID, sOrder"

    SetGrid trs, fg

    With fg
        
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(6) = True
        .ColHidden(7) = True
    
        .ColWidth(2) = 1800
        .ColWidth(5) = 3950
    
        .ColAlignment(5) = flexAlignLeftTop
    
        .ColComboList(5) = GLDrop
    
    End With

End Sub

Private Sub SaveFG()
    
    If InitFlag = False And EditFlag = True Then
        If MsgBox("OK to save changes for: " & GName, vbQuestion + vbYesNo) = vbYes Then
            trs.MoveFirst
            Do
                If trs!GLUpdID <> 0 Then
                    If PRGLUpd.GetByID(trs!GLUpdID) = False Then
                        MsgBox "GLUpd not found: " & trs!GLUpdID, vbExclamation
                        GoBack
                    End If
                Else
                    PRGLUpd.Clear
                    PRGLUpd.GLType = trs!GLType
                    PRGLUpd.RelatedID = trs!RelatedID
                    PRGLUpd.ItemID = trs!ItemID
                    PRGLUpd.GLItemType = trs!GLItemType
                    PRGLUpd.Save (Equate.RecAdd)
                End If
                
                PRGLUpd.GLAccountNum = trs!GLAccountNum
                PRGLUpd.Save (Equate.RecPut)
                
                trs.MoveNext
            Loop Until trs.EOF
        End If
    End If

End Sub

Private Sub AddPRItem(ByVal GLType As Byte, _
                      ByVal RelatedID As Long, _
                      ByVal GLItemType As Byte, _
                      ByVal ItemID As Long, _
                      ByVal Title As String, _
                      ByVal sOrder As String)
                       
    trs.AddNew
    trs!GLType = GLType
    trs!RelatedID = RelatedID
    trs!ItemTitle = Title
    trs!sOrder = sOrder
    trs!ItemID = ItemID
    trs!GLItemType = GLItemType
    trs!GLUpdID = 0
    
    trs!GLAccountNum = 0
    
    ' init to copy from?
    If trsCopyFrom.RecordCount <> 0 Then
        If GLItemType = PREquate.GLItemTypeOE Or GLItemType = PREquate.GLItemTypeDed Then
            trsCopyFrom.Find "ItemID = " & ItemID, 0, adSearchForward, 1
        Else
            trsCopyFrom.Find "GLItemType = " & GLItemType, 0, adSearchForward, 1
        End If
        If trsCopyFrom.EOF = False Then
            trs!GLAccountNum = trsCopyFrom!GLAccountNum
        End If
    End If
    
    trs.Update

End Sub

Private Sub fgGLType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    If InitFlag Then Exit Sub
    
    If fgGLType.TextMatrix(fgGLType.Row, 0) = "" Then Exit Sub
    If fgGLType.Row = 0 Then Exit Sub
    
    PopAcctGrid trsGLT!GLType, trsGLT!RelatedID, 0
    
    EditFlag = False

End Sub

Private Sub fgGLType_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    If InitFlag = True Then Exit Sub
    GName = trsGLT!GLName

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 Then
        Cancel = True
        Exit Sub
    Else
        EditFlag = True
    End If
End Sub

Private Sub cmdPrint_Click()

Dim fg_Row, fgGLT_Row As Long

    InitFlag = True

    ReportTitle = "PR to GL Update Accounts"
    
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Prvw.Caption = PRCompany.Name & " - " & Trim(ReportTitle)
    MaxLines = 65
    Ln = 0
    Pg = 0
    ' save the current rows
    fg_Row = fg.Row
    fgGLT_Row = fgGLType.Row
    
    GLAccount.OpenRS
    
    trsGLT.MoveFirst
    Do
        
        ' separate page for each
        If Ln <> 0 Then FormFeed
        PageHeader ReportTitle, trsGLT!GLName
        Ln = Ln + 1
        
        ' header
        Prvw.vsp.Font.Bold = True
        PrintValue(1) = " ":            FormatString(1) = "a3"
        PrintValue(2) = "Payroll Item": FormatString(2) = "a30"
        PrintValue(3) = " ":            FormatString(3) = "a3"
        PrintValue(4) = " GL Acct#":    FormatString(4) = "a9"
        PrintValue(5) = " ":            FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 2
        Prvw.vsp.Font.Bold = False
        
        SQLString = "SELECT * FROM PRGLUpd WHERE GLType = " & trsGLT!GLType & _
                    " AND RelatedID = " & trsGLT!RelatedID & _
                    " ORDER BY GLItemType, ItemID"
        If PRGLUpd.GetBySQL(SQLString) = True Then
            Do
                
                X = ""
                Select Case PRGLUpd.GLItemType
                    Case PREquate.GLItemTypeRegPay:         X = "Regular Pay"
                    Case PREquate.GLItemTypeOvtPay:         X = "Ovt Pay"
                    Case PREquate.GLItemTypeOE
                        If PRItem.GetByID(PRGLUpd.ItemID) = True Then
                            X = PRItem.Title
                        Else
                            X = "Oth Ern " & PRGLUpd.ItemID
                        End If
                    Case PREquate.GLItemTypeDed
                        If PRItem.GetByID(PRGLUpd.ItemID) = True Then
                            X = PRItem.Title
                        Else
                            X = "Deduction " & PRGLUpd.ItemID
                        End If
                    Case PREquate.GLItemTypeSSTax:          X = "SS Tax"
                    Case PREquate.GLItemTypeSSMatch:        X = "SS Match"
                    Case PREquate.GLItemTypeMedTax:         X = "Med Tax"
                    Case PREquate.GLItemTypeMedMatch:       X = "Med Match"
                    Case PREquate.GLItemTypeFWTTax:         X = "FWT Tax"
                    Case PREquate.GLItemTypeSWTTax:         X = "SWT Tax"
                    Case PREquate.GLItemTypeCWTTax:         X = "CWT Tax"
                    Case PREquate.GLItemTypeFUN:            X = "FUN Tax"
                    Case PREquate.GLItemTypeSUN:            X = "SUN Tax"
                    Case PREquate.GLItemTypeWkcTax:         X = "WKC Tax"
                    Case PREquate.GLItemTypeGross:          X = "Gross Pay"
                    Case PREquate.GLItemTypeNet:            X = "Net Pay"
                    Case PREquate.GLItemTypeSSExp:          X = "SS Expense"
                    Case PREquate.GLItemTypeMEDExp:         X = "MED Expense"
                    Case PREquate.GLItemTypeFUNExp:         X = "FUN Expense"
                    Case PREquate.GLItemTypeSUNExp:         X = "SUN Expense"
                    Case PREquate.GLItemTypeWkcExp:         X = "WKC Expense"
                End Select
                
                Y = ""
                If PRGLUpd.GLAccountNum <> 0 Then
                    If GLAccount.GetAccount(PRGLUpd.GLAccountNum) = True Then
                        Y = Mid(GLAccount.FullDesc, 1, 30)
                    End If
                End If
                
                PrintValue(1) = " ":            FormatString(1) = "a3"
                PrintValue(2) = X:              FormatString(2) = "a30"
                PrintValue(3) = " ":            FormatString(3) = "a3"
                                
                If PRGLUpd.GLAccountNum = 0 Then
                    PrintValue(4) = " ":        FormatString(4) = "~"
                Else
                    PrintValue(4) = PRGLUpd.GLAccountNum:       FormatString(4) = "a9"
                    PrintValue(5) = Y:                          FormatString(5) = "a30"
                    PrintValue(6) = " ":                        FormatString(6) = "~"
                End If
                FormatPrint
                Ln = Ln + 1
                
                If PRGLUpd.GetNext = False Then Exit Do
            
            Loop
        
        End If
        
        trsGLT.MoveNext
    
    Loop Until trsGLT.EOF
    
    PrvwReturn = True
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

    trs.MoveFirst
    trsGLT.MoveFirst
    
    InitFlag = False

End Sub




