VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmGLUpdAdd 
   Caption         =   "Add an Employee Category"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
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
   ScaleHeight     =   8385
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCopyFrom 
      Height          =   360
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   6135
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&SELECT"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5295
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   6975
      _cx             =   12303
      _cy             =   9340
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
   Begin VB.Frame fraEmpCatType 
      Caption         =   "  Employee Category Type  "
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   4335
      Begin VB.OptionButton optEmployee 
         Caption         =   "Employee"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optDepartment 
         Caption         =   "Department"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label Label1 
      Caption         =   "Copy Account Setup From:"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
End
Attribute VB_Name = "frmGLUpdAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GLType As Byte
Public RecID As Long
Public Title As String
Public Number As Long
Dim trs As New ADODB.Recordset
Public CopyFromID As Long
Public CopyFromType As Byte

Private Sub Form_Load()

    trs.CursorLocation = adUseClient
    
    trs.Fields.Append "RecID", adDouble
    trs.Fields.Append "Number", adDouble
    trs.Fields.Append "Name", adVarChar, 255, adFldIsNullable
    
    trs.Open , , adOpenDynamic, adLockOptimistic
    
    SetGrid trs, fg

    Me.optDepartment = True
    Me.optEmployee = False

    PopCmb Me.cmbCopyFrom, frmGLUpd.trsGLT
    LoadData

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub LoadData()

    ' clear out the temp record set
    If trs.RecordCount > 0 Then
        trs.MoveFirst
        Do Until trs.RecordCount = 0
            trs.Delete
            trs.MoveNext
        Loop
    End If

    If Me.optDepartment Then
        SQLString = "SELECT * FROM PRDepartment ORDER BY DepartmentNumber"
        ' no departments defined
        If Not PRDepartment.GetBySQL(SQLString) Then
            Me.optDepartment = False
            Me.optEmployee = True
        Else
            Do
                trs.AddNew
                trs!RecID = PRDepartment.DepartmentID
                trs!Number = PRDepartment.DepartmentNumber
                trs!Name = PRDepartment.Name
                trs.Update
                If Not PRDepartment.GetNext Then Exit Do
            Loop
        End If
    End If
    
    If Me.optEmployee Then
        SQLString = "SELECT * FROM PREmployee ORDER BY EmployeeNumber"
        If Not PREmployee.GetBySQL(SQLString) Then
            MsgBox "No Employees Found !!", vbCritical
            End
        End If
        Do
            trs.AddNew
            trs!RecID = PREmployee.EmployeeID
            trs!Number = PREmployee.EmployeeNumber
            trs!Name = PREmployee.LFName
            trs.Update
            If Not PREmployee.GetNext Then Exit Do
        Loop
    End If
    
    With fg
        .ColHidden(0) = True
        .ColWidth(1) = 2000
        .ColWidth(2) = 4000
        .SelectionMode = flexSelectionByRow
    End With
    trs.MoveFirst

End Sub

Private Sub cmdExit_Click()
    GLType = 0
    RecID = 0
    trs.Close
    Me.Hide
End Sub

Private Sub cmdSelect_Click()
    If Me.optDepartment Then
        GLType = PREquate.GLTypeDept
        CopyFromType = PREquate.GLTypeDept
    Else
        GLType = PREquate.GLTypeEmployee
        CopyFromType = PREquate.GLTypeEmployee
    End If
    RecID = fg.TextMatrix(fg.Row, 0)
    Number = fg.TextMatrix(fg.Row, 1)
    Title = fg.TextMatrix(fg.Row, 2)
    trs.Close
    
    With Me.cmbCopyFrom
        If .ListIndex = -1 Then
            CopyFromID = 0
        Else
            CopyFromID = .ItemData(.ListIndex)
        End If
    End With
    
    Me.Hide
End Sub

Private Sub optDepartment_Click()
    LoadData
End Sub

Private Sub optEmployee_Click()
    LoadData
End Sub

Private Sub PopCmb(ByRef cmb As ComboBox, ByRef rs As ADODB.Recordset)

'    SQLString = "SELECT * FROM PRGLUpd ORDER BY GLType"
'    if prglupd.GetBySQL(sqlstring)


    With cmb
        .Clear
        rs.MoveFirst
        Do
            .AddItem rs!GLName
            .ItemData(.NewIndex) = rs!RelatedID
            rs.MoveNext
        Loop Until rs.EOF
    End With

End Sub

