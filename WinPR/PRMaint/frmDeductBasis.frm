VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDeductBasis 
   Caption         =   "Deduction Basis"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   8685
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6008
      TabIndex        =   2
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   2408
      TabIndex        =   1
      Top             =   7920
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4695
      Left            =   1575
      TabIndex        =   0
      Top             =   2880
      Width           =   6615
      _cx             =   11668
      _cy             =   8281
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Check the earning categories to be included in the basis for this deduction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label lblDeductName 
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
      Height          =   255
      Left            =   1035
      TabIndex        =   5
      Top             =   1680
      Width           =   7680
   End
   Begin VB.Label lblWho 
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
      Height          =   255
      Left            =   1035
      TabIndex        =   4
      Top             =   1200
      Width           =   7680
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
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmDeductBasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rsItemID As New ADODB.Recordset
Dim GlobID As Long

Public EmployeeID As Long
Public ItemID As Long   ' employer item ID for the deduction
Dim X, Y As String
Dim i, j As Long

Private Sub Form_Load()

    ' this records what earnings are to be excluded
    ' for the basis for deductions as a percent
    '
    ' PRGlobal Assignments:
    '       UserID              PR Company ID
    '       Description         Employer or Employee ItemID
    '       Var1                = 0 - Employer defn else EmployeeID
    '       Var2                Excluded earning ID's
    '
    
    Me.lblCompanyName = PRCompany.Name
    
    If EmployeeID = 0 Then
        Me.lblWho = "COMPANY DEFINITION"
    Else
        If PREmployee.GetByID(EmployeeID) = False Then
            MsgBox "Employee not found: " & EmployeeID, vbExclamation
            GoBack
        End If
        Me.lblWho = PREmployee.FLName
    End If
    
    ' *****************************************************************
    '    get the list of earnings
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Select", adBoolean
    rs.Fields.Append "Title", adVarChar, 30, adFldIsNullable
    rs.Fields.Append "ItemID", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    rs.AddNew
    rs!Select = True
    rs!Title = "REGULAR PAY"
    rs!ItemID = 99991
    rs.Update
    
    rs.AddNew
    rs!Select = True
    rs!Title = "OVERTIME PAY"
    rs!ItemID = 99992
    rs.Update
    
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0 AND ItemType = " & PREquate.ItemTypeOE & _
                " ORDER BY ItemID"
    If PRItem.GetBySQL(SQLString) Then
        Do
            rs.AddNew
            rs!Select = True
            rs!Title = PRItem.Title
            rs!ItemID = PRItem.ItemID
            rs.Update
            
            If PRItem.GetNext = False Then Exit Do
        Loop
    End If
    ' *****************************************************************

    ' title for the deduction from the employer item
    If PRItem.GetByID(ItemID) = False Then
        MsgBox "Employer Item not found: " & ItemID, vbExclamation
        GoBack
    End If
    Me.lblDeductName = PRItem.Title
    
    ' find the PRGlobal Record
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDeductBasis & _
                " AND UserID = " & PRCompany.CompanyID & _
                " AND Description = '" & ItemID & "'" & _
                " AND Var1 = '" & EmployeeID & "'"
 
    ' set the flags to off
    ' other earining items to EXCLUDE are the ones stored
    GlobID = 0
    If PRGlobal.GetBySQL(SQLString) = True Then
        GlobID = PRGlobal.GlobalID
        Set rsItemID = ParseString(PRGlobal.Var2, "/")
        If rsItemID.RecordCount > 0 Then
            rsItemID.MoveFirst
            Do
                SQLString = "ItemID = " & rsItemID!listvalue
                rs.Find SQLString, 0, adSearchForward, 1
                If rs.EOF = False Then
                    rs!Select = False
                    rs.Update
                End If
                rsItemID.MoveNext
            Loop Until rsItemID.EOF
        End If
    ElseIf EmployeeID <> 0 Then     ' ee defn DNE - init to employer
        SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeDeductBasis & _
                    " AND UserID = " & PRCompany.CompanyID & _
                    " AND Description = '" & ItemID & "'" & _
                    " AND Var1 = '0'"
        If PRGlobal.GetBySQL(SQLString) = True Then
            Set rsItemID = ParseString(PRGlobal.Var2, "/")
            If rsItemID.RecordCount > 0 Then
                rsItemID.MoveFirst
                Do
                    SQLString = "ItemID = " & rsItemID!listvalue
                    rs.Find SQLString, 0, adSearchForward, 1
                    If rs.EOF = False Then
                        rs!Select = False
                        rs.Update
                    End If
                    rsItemID.MoveNext
                Loop Until rsItemID.EOF
            End If
        End If
    End If

    SetGrid rs, fg

    With fg
        .ColWidth(1) = 4000
        .ColWidth(2) = 0
    End With

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdSave_Click()
    
    If GlobID = 0 Then
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeDeductBasis
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Description = ItemID
        PRGlobal.Var1 = EmployeeID
        PRGlobal.Save (Equate.RecAdd)
    Else
        If PRGlobal.GetByID(GlobID) = False Then
            MsgBox "PRGlobal not found: " & GlobID, vbExclamation
            GoBack
        End If
    End If

    PRGlobal.Var2 = ""
    
    rs.MoveFirst
    Do
        If rs!Select = False Then
            If PRGlobal.Var2 <> "" Then
                PRGlobal.Var2 = Trim(PRGlobal.Var2) & "/"
            End If
            PRGlobal.Var2 = Trim(PRGlobal.Var2) & rs!ItemID
        End If
        rs.MoveNext
    Loop Until rs.EOF
    
    PRGlobal.Save (Equate.RecPut)

    Unload Me

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


