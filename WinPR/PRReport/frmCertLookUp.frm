VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCertLookUp 
   Caption         =   "Certified Payroll Report - Look Up"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCertLookUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClrAll 
      Caption         =   "CLEA&R ALL"
      Height          =   615
      Left            =   2520
      TabIndex        =   6
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "SELECT &ALL"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   7920
      TabIndex        =   4
      Top             =   8640
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6015
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   13335
      _cx             =   23521
      _cy             =   10610
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   615
      Left            =   10080
      TabIndex        =   0
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   7935
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmCertLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k As Long
Dim X, Y, Z As String
Dim boo As Boolean

Public rs As New ADODB.Recordset
Public OK As Boolean


Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
        
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
    End Select
End Sub

Public Sub Init()

    rsDelAll rs
    
    With rs
        
        .Filter = adFilterNone
        
        If UCase(frmCertReg.LUType) = "JOB" Then
            SQLString = "SELECT * FROM JCJob ORDER BY FullName"
            If JCJob.GetBySQL(SQLString) = False Then
                MsgBox "No Job Exist!", vbExclamation
                Me.Hide
            End If
            Me.lblTitle = "Select Union Job"
        Else
            SQLString = "SELECT * FROM PREmployee ORDER BY LastName, FirstName"
            If PREmployee.GetBySQL(SQLString) = False Then
                MsgBox "No Employees Exist!", vbExclamation
                Me.Hide
            End If
            Me.lblTitle = "Select Union Employee"
        End If
        
        Do
    
            ' don't add if already in the cert list
            If UCase(frmCertReg.LUType) = "JOB" Then
                SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeUnionJob & _
                            " AND Var1 = '" & JCJob.JobID & "'" & _
                            " AND UserID = " & PRCompany.CompanyID
            Else
                SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeUnionEmployee & _
                            " AND Var1 = '" & PREmployee.EmployeeID & "'" & _
                            " AND UserID = " & PRCompany.CompanyID
                            
            End If
            If PRGlobal.GetBySQL(SQLString) = False Then
                .AddNew
                If UCase(frmCertReg.LUType) = "JOB" Then
                    !ID = JCJob.JobID
                    !Number = JCJob.JobID
                    !Name = Mid(JCJob.FullName, 1, 40)
                Else
                    !ID = PREmployee.EmployeeID
                    !Number = PREmployee.EmployeeNumber
                    !Name = Mid(PREmployee.LFName, 1, 40)
                End If
                .Update
            End If
            
            If UCase(frmCertReg.LUType) = "JOB" Then
                If JCJob.GetNext = False Then Exit Do
            Else
                If PREmployee.GetNext = False Then Exit Do
            End If
        
        Loop
    
    End With

End Sub

Private Sub cmdCancel_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    OK = True
    Me.Hide
End Sub

Private Sub fg_DblClick()
    
    If rs!Select = True Then
        rs!Select = False
    Else
        rs!Select = True
    End If
    
End Sub

Private Sub cmdSelAll_Click()

    j = fg.Row
    rs.MoveFirst
    Do
        rs!Select = True
        rs.MoveNext
    Loop Until rs.EOF
    If j > 0 Then fg.Row = j

End Sub

Private Sub cmdClrAll_Click()

    j = fg.Row
    rs.MoveFirst
    Do
        rs!Select = False
        rs.MoveNext
    Loop Until rs.EOF
    If j > 0 Then fg.Row = j

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If fg.Col <> fg.ColIndex("Select") Then Cancel = True
End Sub


