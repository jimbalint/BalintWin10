VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmJCList 
   Caption         =   "Customer / Job Maintenance"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
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
   ScaleHeight     =   10395
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQBRefresh 
      Caption         =   "REFRESH FROM QB"
      Height          =   975
      Left            =   10800
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   10800
      TabIndex        =   8
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton cmdJobDelete 
      Caption         =   "DE&LETE"
      Height          =   495
      Left            =   10800
      TabIndex        =   7
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdJobEdit 
      Caption         =   "ED&IT"
      Height          =   495
      Left            =   10800
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdJobAdd 
      Caption         =   "&A&DD"
      Height          =   495
      Left            =   10800
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustDelete 
      Caption         =   "&DELETE"
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustEdit 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   10800
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   10800
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid fgJob 
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   10215
      _cx             =   18018
      _cy             =   7435
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
   Begin VSFlex8Ctl.VSFlexGrid fgCustomer 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   10215
      _cx             =   18018
      _cy             =   7223
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
      Caption         =   "Label3"
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
      Left            =   1148
      TabIndex        =   11
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label Label2 
      Caption         =   "J O B S"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "C U S T O M E R S"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmJCList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCustomer As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim LoadFlag As Boolean
Dim fgFlag As Boolean
Dim SaveID As Long
Dim rw As Long
Dim CustSQL As String
Dim CityDrop As String

Private Sub Form_Load()

    ' city name drop down
    CityDrop = "|#0;NONE"
    SQLString = "SELECT * FROM PRCity ORDER BY CityName"
    If PRCity.GetBySQL(SQLString) Then
        Do
            CityDrop = Trim(CityDrop) & "|#" & PRCity.CityID & ";" & PRCity.CityName
            If PRCity.GetNext = False Then Exit Do
        Loop
    End If
    
    If TableExists("JCCustomer", cn) = False Then
        CustomerCreate
    End If
    
    If TableExists("JCJob", cn) = False Then
        JobCreate
    End If

    LoadFlag = True
    fgFlag = True

    CustPop

    LoadFlag = False
    
    JobPop
    
    fgFlag = False

    Me.lblCompanyName = PRCompany.Name

    Me.KeyPreview = True

    fgJob.ColWidth(3) = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub JobPop()

    If LoadFlag = True Then
        Exit Sub
    End If
    
    If rsCustomer.RecordCount = 0 Then
        ' start an empty recordset to avoid err when DB is empty
        SQLString = "SELECT JobID, Name, FullName, ParentID, CityID " & _
                    " FROM JCJob WHERE ParentID = " & -1
        rsInit SQLString, cn, rsJob
        Exit Sub
    End If
    
    If fgFlag = False Then
        rsJob.Close
    End If
    
    SQLString = "SELECT JobID, Name, FullName, ParentID, CityID " & _
                " FROM JCJob WHERE ParentID = " & rsCustomer!CustomerID
    rsInit SQLString, cn, rsJob
    SetGrid rsJob, fgJob
    
    fgJob.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fgJob.TabBehavior = flexTabCells                       ' tab moves between cells
    ' fgcustomer.HighLight = flexHighlightNever                   ' don't select ranges
    fgJob.SelectionMode = flexSelectionByRow
    fgJob.Editable = flexEDNone

    fgJob.ColWidth(0) = 1200
    fgJob.ColWidth(1) = 3800
    fgJob.ColWidth(2) = 3800
    fgJob.ColWidth(3) = 0
    fgJob.ColWidth(4) = 4000
    fgJob.ColComboList(4) = CityDrop
    
End Sub

Private Sub fgCustomer_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    JobPop
End Sub
Private Sub cmdCustEdit_Click()
    
    TaskID = rsCustomer!CustomerID
    frmJCEdit.Action = PREquate.ActionEdit
    frmJCEdit.EditJob = False
    frmJCEdit.Show vbModal
    If TaskID = 0 Then Exit Sub
    
    fgCustomerReset

End Sub
Private Sub cmdCustAdd_Click()

    TaskID = 0
    frmJCEdit.Action = PREquate.ActionAdd
    frmJCEdit.EditJob = False
    frmJCEdit.Show vbModal
    If TaskID = 0 Then Exit Sub
    
    fgCustomerReset

End Sub

Private Sub cmdCustDelete_Click()
    
    If rsCustomer.RecordCount = 0 Then Exit Sub
    
    If MsgBox("OK to delete customer: " & rsCustomer!CustomerID & vbCr & _
           rsCustomer!Name, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    X = rsCustomer!Name

    ' delete the associated jobs
    If rsJob.RecordCount > 0 Then
        rsJob.MoveFirst
        Do
            rsJob.Delete
            rsJob.MoveNext
        Loop Until rsJob.EOF
    End If
    
    ' delete the customer record
    LoadFlag = True
    rsCustomer.Delete
    LoadFlag = False
    
    TaskID = 0
    
    fgCustomerReset

End Sub
Private Sub cmdJobAdd_Click()

    frmJCEdit.ParentID = rsCustomer!CustomerID
    frmJCEdit.Action = PREquate.ActionAdd
    frmJCEdit.EditJob = True
    frmJCEdit.Show vbModal
    If TaskID = 0 Then Exit Sub
    
    rsJob.Requery

End Sub

Private Sub cmdJobEdit_Click()
    
    SaveID = rsJob!JobID
    
    frmJCEdit.Action = PREquate.ActionEdit
    frmJCEdit.EditJob = True
    TaskID = rsJob!JobID
    frmJCEdit.Show vbModal

    rsJob.Requery
    rsJob.Find "JobID = " & SaveID, 0, adSearchForward, 1

End Sub

Private Sub cmdJobDelete_Click()
    If rsJob.RecordCount = 0 Then Exit Sub
    If MsgBox("OK to delete job: " & rsJob!JobID & vbCr & _
           rsJob!Name, vbQuestion + vbYesNo) = vbNo Then Exit Sub
           
    rsJob.Delete
    JobPop
End Sub

Private Sub CustPop()
    SQLString = "SELECT CustomerID, Name, FullName FROM JCCustomer ORDER BY Name"
    CustSQL = SQLString
    rsInit SQLString, cn, rsCustomer
    
    SetGrid rsCustomer, fgCustomer
    
    fgCustomer.BackColorAlternate = RGB(192, 192, 192)          ' light gray
    fgCustomer.TabBehavior = flexTabCells                       ' tab moves between cells
    ' fgcustomer.HighLight = flexHighlightNever                   ' don't select ranges
    fgCustomer.SelectionMode = flexSelectionByRow
    fgCustomer.Editable = flexEDNone

    fgCustomer.ColWidth(0) = 1200
    fgCustomer.ColWidth(1) = 4000
    fgCustomer.ColWidth(2) = 4000
    
End Sub

Private Sub fgCustomerReset()
    
    LoadFlag = True
    
    rsCustomer.Close
    rsInit CustSQL, cn, rsCustomer
    Set fgCustomer.DataSource = rsCustomer.DataSource
    rsCustomer.Find "CustomerID = " & TaskID, 0, adSearchForward, 1
    rw = fgCustomer.FindRow(TaskID, 0, 0)
    If rw = -1 Then     ' not found - after delete
        If rsCustomer.RecordCount = 0 Then Exit Sub
        rsCustomer.MoveFirst
        rw = 1
    End If
    fgCustomer.TopRow = rw
    fgCustomer.Select rw, 0
    fgCustomer.SetFocus
    
    LoadFlag = False

    JobPop

End Sub

Private Sub fgCustomer_DblClick()
    cmdCustEdit_Click
End Sub

Private Sub fgJob_DblClick()
    cmdJobEdit_Click
End Sub

Private Sub cmdQBRefresh_Click()
    
'    If MsgBox("OK to overwrite ALL QB Customer and Job Info?", vbQuestion + vbYesNo, "QB Customer/Job Import") = vbNo Then
'        Exit Sub
'    End If
'
'    If TableExists("JCCustomer", cn) = False Then
'        CustomerCreate
'    End If
'
'    If TableExists("JCJob", cn) = False Then
'        JobCreate
'    End If
'
'    DoCustomerQueryRq "US", 5, 0
'
'    rsCustomer.Requery
'    fgCustomer.DataRefresh
'
'    rsJob.Requery
'    fgJob.DataRefresh
'
'    MsgBox "Import of QB Customer and Job Info Complete", vbInformation, "Balint Windows PR"

End Sub


