VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTotals 
   Caption         =   "1099 Totals"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTotals.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   6975
      _cx             =   12303
      _cy             =   11033
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
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   8400
      Width           =   1575
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
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim rs As New ADODB.Recordset
Dim Count99 As Long
Dim LastBox As String
Dim Amt As Currency
Public TaxYear As Long
Public FormType As String

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.KeyPreview = True

    ' get the form count
    SQLString = " SELECT DISTINCT(PayeeID) FROM Detail99 " & _
                " WHERE FormType = '" & FormType & "' " & _
                " AND TaxYear = " & TaxYear
    rsInit SQLString, cn, rs
    rs.MoveLast
    Count99 = rs.RecordCount
    rs.Close
    
    If Count99 = 0 Then
        MsgBox "No detail found!", vbInformation
        Unload Me
    End If
    
    With Me.fg
        .Enabled = False
        .FixedCols = 0
        .FixedRows = 0
        .Cols = 2
        .Editable = flexEDNone
        .BackColorAlternate = RGB(180, 180, 180)
        .DataMode = flexDMFree
        .ColWidth(0) = 4500
        .ColWidth(1) = 1500
        .Rows = 2
        .ScrollBars = flexScrollBarNone
        .TextMatrix(0, 0) = "1099-" & Form99.FormType
        .TextMatrix(1, 0) = "Form Count"
        .TextMatrix(1, 1) = Count99
    End With
    
    SQLString = " SELECT * FROM Detail99 WHERE FormType = '" & FormType & "' " & _
                " AND TaxYear = " & TaxYear & _
                " ORDER BY BoxName"
    If Detail99.GetBySQL(SQLString) = False Then
        MsgBox "No detail found", vbInformation
        Unload Me
    End If

    Amt = 0
    LastBox = ""
    Do
        If LastBox <> "" And LastBox <> Detail99.BoxName Then DisplayTotal LastBox
        LastBox = Detail99.BoxName
        Amt = Amt + ParseAmt(Detail99.FieldValue)
        If Detail99.GetNext = False Then Exit Do
    Loop
    DisplayTotal LastBox

End Sub

Public Sub DisplayTotal(ByVal Box As String)

    SQLString = " SELECT * FROM Field99 WHERE FormType = '" & FormType & "'" & _
                " AND TaxYear = " & TaxYear & _
                " AND BoxName = '" & Box & "'"
    If Field99.GetBySQL(SQLString) = False Then
        X = "?"
    Else
        X = Field99.BoxName & " " & Trim(Field99.FieldTitle)
    End If

    With Me.fg
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = X
        .TextMatrix(.Rows - 1, 1) = Format(Amt, "Currency")
    End With
    Amt = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


