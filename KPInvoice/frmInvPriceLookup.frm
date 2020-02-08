VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInvPriceLookup 
   Caption         =   "Price Lookup"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvPriceLookup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10560
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "  Select Prices To Display  "
      Height          =   855
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   5415
      Begin VB.OptionButton optMasterPricing 
         Caption         =   "&Master Pricing"
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optCustomerPricing 
         Caption         =   "&Customer Pricing"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   7590
      TabIndex        =   3
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   4830
      TabIndex        =   2
      Top             =   9480
      Width           =   2055
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   7095
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   13815
      _cx             =   24368
      _cy             =   12515
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
   Begin VB.Label lblJobName 
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
      Height          =   855
      Left            =   6360
      TabIndex        =   7
      Top             =   960
      Width           =   7215
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
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "frmInvPriceLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Public rs As New ADODB.Recordset
Public JobID As Long

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim Flg As Boolean

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub
Private Sub cmdOK_Click()
    OK = True
    Me.Hide
End Sub

Private Sub cmdExit_Click()
    OK = False
    Me.Hide
End Sub

Public Sub Init()

    On Error Resume Next
    rs.Close
    On Error GoTo 0
    
    Me.lblJobName = JCJob.FullName
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "StockID1", adDouble
    rs.Fields.Append "Quantity1", adDouble
    rs.Fields.Append "Description1", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "Price1", adDouble
    rs.Fields.Append "X", adVarChar, 10
    rs.Fields.Append "StockID2", adDouble
    rs.Fields.Append "Quantity2", adDouble
    rs.Fields.Append "Description2", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "Price2", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic

    SetGrid rs, fg
    
    With fg
    
        For I = 0 To .Cols - 1
            .ColKey(I) = .TextMatrix(0, I)
        Next I
        
        .ColWidth(.ColIndex("StockID1")) = 0
        .ColWidth(.ColIndex("StockID2")) = 0
        
        .ColWidth(.ColIndex("Quantity1")) = 1200
        .ColWidth(.ColIndex("Quantity2")) = 1200
        
        .ColWidth(.ColIndex("Description1")) = 3500
        .ColWidth(.ColIndex("Description2")) = 3500
        
        .ColWidth(.ColIndex("Price1")) = 1200
        .ColWidth(.ColIndex("Price2")) = 1200
        
        .ColFormat(.ColIndex("Quantity1")) = "###,##0"
        .ColFormat(.ColIndex("Quantity2")) = "###,##0"
        
        .ColFormat(.ColIndex("Price1")) = "###,##0.0000"
        .ColFormat(.ColIndex("Price2")) = "###,##0.0000"
    
        .ColWidth(.ColIndex("X")) = 500
        .TextMatrix(0, .ColIndex("X")) = " "
    
    End With

    ' load the stock items - 2 columns
    SQLString = "SELECT * FROM InvStock WHERE JobID = " & JobID & _
                " AND Description <> 'Freight' " & _
                " AND StockSelect = True ORDER BY Description"
    If InvStock.GetBySQL(SQLString) = False Then
        boo = JCJob.GetByID(JobID)
        MsgBox "No stock items found for: " & JCJob.Name, vbExclamation
        Me.Hide
    End If

    I = 0
    Do
        
        I = I + 1
        If I Mod 2 = 1 Then
            rs.AddNew
            rs!StockID1 = InvStock.StockID
            rs!Quantity1 = 0
            rs!Description1 = Mid(InvStock.Description, 1, 40)
            rs!StockID2 = 0
            rs!X = "||||||||||"
            rs.Update
        Else
            rs!StockID2 = InvStock.StockID
            rs!Quantity2 = 0
            rs!Description2 = Mid(InvStock.Description, 1, 40)
            rs.Update
        End If
        
        If InvStock.GetNext = False Then Exit Do
    
    Loop

    Me.optCustomerPricing = True
    
    LoadPrices

End Sub

Private Sub LoadPrices()

    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    Do
        
        If rs!StockID1 <> 0 Then
            boo = InvStock.GetByID(rs!StockID1)
            If Me.optCustomerPricing = True Then
                rs!Price1 = InvStock.CustomerPrice
            Else
                rs!Price1 = InvStock.MasterPrice
            End If
        End If
        
        If rs!StockID2 <> 0 Then
            boo = InvStock.GetByID(rs!StockID2)
            If Me.optCustomerPricing = True Then
                rs!Price2 = InvStock.CustomerPrice
            Else
                rs!Price2 = InvStock.MasterPrice
            End If
        End If

        rs.MoveNext
    
    Loop Until rs.EOF
    
    rs.MoveFirst

End Sub

Private Sub optCustomerPricing_Click()
    LoadPrices
End Sub

Private Sub optMasterPricing_Click()
    LoadPrices
End Sub
