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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   8400
      Width           =   1455
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
      TabIndex        =   5
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGLTypeDel 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   2880
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
   Begin VB.Label Label1 
      Caption         =   "ALL CHANGES WILL NOT BE SAVED !!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   9480
      TabIndex        =   7
      Top             =   8160
      Width           =   1455
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
      TabIndex        =   6
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


Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    InitFlag = True
    
    LoadData
    GridSetup
    
    trs.Filter = "GLType = " & trsGLT!GLType & " AND RelatedID = " & trsGLT!RelatedID
    
    InitFlag = False
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub LoadData()

Dim X As String

    ' define EmpCat record set
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
    
    ' temp record set to track each GLType in the file
    trsGLT.CursorLocation = adUseClient
    
    trsGLT.Fields.Append "GLType", adInteger
    trsGLT.Fields.Append "RelatedID", adDouble
    trsGLT.Fields.Append "GLName", adVarChar, 30
    trsGLT.Fields.Append "OrigType", adInteger
    trsGLT.Open , , adOpenDynamic, adLockOptimistic
    
    SQLString = "SELECT * FROM PRGLUpd"
    If Not PRGLUpd.GetBySQL(SQLString) Then
    
        ' get/create for the company always
        MakeGLType PREquate.GLTypeCompany, 0, 0
    
        trsGLT.AddNew
        trsGLT!GLType = PREquate.GLTypeCompany
        trsGLT!RelatedID = 0
        trsGLT!GLName = "COMPANY"
        trsGLT.Update
    
    Else
    
        ' load up the different types
        Do
            trsGLT.Filter = "GLType = " & PRGLUpd.GLType & " AND RelatedID = " & PRGLUpd.RelatedID
            If trsGLT.RecordCount = 0 Then
                trsGLT.Filter = adFilterNone
                trsGLT.AddNew
                trsGLT!GLType = PRGLUpd.GLType
                trsGLT!RelatedID = PRGLUpd.RelatedID
                trsGLT!GLName = Mid(GetGLName(PRGLUpd.GLType, PRGLUpd.RelatedID), 1, 30)
                trsGLT.Update
            Else
                trsGLT.Filter = adFilterNone
            End If
            
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
                
    ' fill in the different types
    trsGLT.MoveFirst
    Do
        MakeGLType trsGLT!GLType, trsGLT!RelatedID, 0
        trsGLT.MoveNext
    Loop Until trsGLT.EOF
    
    trs.Sort = "GLType, RelatedID, sOrder"
    
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
    
    ' Right Side - Item Grid
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
                          
Private Sub MakeGLType(ByVal GLType As Byte, ByVal RelatedID As Long, ByVal CopyFromID As Long)

    ' define the copy from record set
    ' record count will be zero if not used
    On Error Resume Next
    trsCopyFrom.Close
    On Error GoTo 0
    trsCopyFrom.CursorLocation = adUseClient
    trsCopyFrom.Fields.Append "GLItemType", adInteger
    trsCopyFrom.Fields.Append "ItemID", adDouble
    trsCopyFrom.Fields.Append "GLAccountNum", adDouble
    trsCopyFrom.Open , , adOpenDynamic, adLockOptimistic

    ' make a copy from record set
    If CopyFromID <> 0 Then

        ' get the copy from info
        trs.Filter = adFilterNone
        trs.MoveFirst
        Do
            If trs!GLType = frmGLUpdAdd.CopyFromType And _
               trs!RelatedID = frmGLUpdAdd.CopyFromID Then
                
                trsCopyFrom.AddNew
                trsCopyFrom!GLItemType = trs!GLItemType
                trsCopyFrom!ItemID = trs!ItemID
                trsCopyFrom!GLAccountNum = trs!GLAccountNum
                trsCopyFrom.Update
            End If
            trs.MoveNext
        Loop Until trs.EOF
    
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
    
    ' fill in account numbers and PRGLUpd RecIDs
    trs.Filter = "GLType = " & GLType & " AND RelatedID = " & RelatedID
    If trs.RecordCount = 0 Then ' ???
        trs.Filter = adFilterNone
        Exit Sub
    End If
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
                    
    trs.Filter = adFilterNone
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

Private Sub cmdExit_Click()
    ' >>> save final setup ...
    GoBack
End Sub
Private Sub cmdSave_Click()

    ' execute the deletes
    trsGLT.MoveFirst
    Do
        If trsGLT!GLType = 4 Then
            SQLString = "DELETE * FROM PRGLUpd WHERE GLType = " & trsGLT!OrigType & _
                        " AND RelatedID = " & trsGLT!RelatedID
            cn.Execute SQLString
        End If
        trsGLT.MoveNext
    Loop Until trsGLT.EOF
            
    ' update
    trs.Filter = adFilterNone
    trs.MoveFirst
    Do
        
        If trs!ItemID = -1 Then
        
            ' deleted
        
        Else
        
            If trs!GLUpdID <> 0 Then
                If Not PRGLUpd.GetByID(trs!GLUpdID) Then
                    MsgBox "PRGLUpdID NF: " & trs!GLUpdID, vbExclamation
                    GoBack
                End If
            Else
                PRGLUpd.Clear
            End If
            
            PRGLUpd.GLType = trs!GLType
            PRGLUpd.RelatedID = trs!RelatedID
            PRGLUpd.GLItemType = trs!GLItemType
            PRGLUpd.ItemID = trs!ItemID
            PRGLUpd.GLAccountNum = trs!GLAccountNum
            PRGLUpd.Title = trs!ItemTitle
            
            If trs!ItemID = -1 Then
                PRGLUpd.GLType = 99
            End If
            
            If trs!GLUpdID = 0 Then
                PRGLUpd.Save (Equate.RecAdd)
            Else
                PRGLUpd.Save (Equate.RecPut)
            End If
        
        End If
        
        trs.MoveNext
    
    Loop Until trs.EOF

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
    
    ' set a flag in trs
    If trs.RecordCount > 0 Then
        trs.Filter = adFilterNone
        trs.MoveFirst
        Do
            If trs!GLType = trsGLT!GLType And trs!RelatedID = trsGLT!RelatedID Then
                trs!ItemID = -1
                trs.Update
            End If
            trs.MoveNext
            If trs.EOF Then Exit Do
        Loop
    End If
    
    trsGLT!OrigType = trsGLT!GLType
    trsGLT!GLType = 4
    trsGLT.Update

    trs.Filter = "GLType = " & trsGLT!GLType & " AND RelatedID = " & trsGLT!RelatedID
    
    cmdSave_Click

End Sub

Private Sub cmdGLTypeAdd_Click()
    
Dim FFlag As Boolean
    
    frmGLUpdAdd.Show vbModal
    
    ' user hit the exit button
    If frmGLUpdAdd.GLType = 0 Then
        Unload frmGLUpdAdd
        Exit Sub
    End If

    ' see if the selection already exists
    trsGLT.Filter = "GLType = " & frmGLUpdAdd.GLType & " AND RelatedID = " & frmGLUpdAdd.RecId
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
    trsGLT!RelatedID = frmGLUpdAdd.RecId
    trsGLT!GLName = GetGLName(frmGLUpdAdd.GLType, frmGLUpdAdd.RecId)
    trsGLT.Update
    trsGLT.Sort = "GLType, RelatedID"
    
    MakeGLType frmGLUpdAdd.GLType, frmGLUpdAdd.RecId, frmGLUpdAdd.CopyFromID
    
    ' point to the record just added
    trsGLT.MoveFirst
    Do
        If trsGLT!GLType = frmGLUpdAdd.GLType And trsGLT!RelatedID = frmGLUpdAdd.RecId Then
            Exit Do
        End If
        trsGLT.MoveNext
        If trsGLT.EOF Then
            MsgBox "EOF?"
            End
        End If
    Loop
    
    ' filter the trs recordset for the GLType just added
    trs.Filter = "GLType = " & trsGLT!GLType & _
                 " AND RelatedID = " & trsGLT!RelatedID
    
    Unload frmGLUpdAdd
    
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
Private Sub fgGLType_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    
    If InitFlag Then Exit Sub
    
    If fgGLType.TextMatrix(fgGLType.Row, 0) = "" Then Exit Sub
    If fgGLType.Row = 0 Then Exit Sub
    
    trs.Filter = "GLType = " & fgGLType.TextMatrix(fgGLType.Row, 0) & _
                 " AND RelatedID = " & fgGLType.TextMatrix(fgGLType.Row, 1)
                 
End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 Then Cancel = True
End Sub



