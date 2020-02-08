VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmNewBatch 
   Caption         =   "Add a new paryoll batch"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSortOrder 
      Height          =   360
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3240
      Width           =   3615
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4095
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   6255
      _cx             =   11033
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   4328
      TabIndex        =   6
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   8640
      Width           =   1095
   End
   Begin TDBNumber6Ctl.TDBNumber tdbIntStartCheck 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calculator      =   "frmNewBatch.frx":0000
      Caption         =   "frmNewBatch.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmNewBatch.frx":00A0
      Keys            =   "frmNewBatch.frx":00BE
      Spin            =   "frmNewBatch.frx":0108
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBDate6Ctl.TDBDate tdbDatePeriodEnding 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calendar        =   "frmNewBatch.frx":0130
      Caption         =   "frmNewBatch.frx":0230
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmNewBatch.frx":02AA
      Keys            =   "frmNewBatch.frx":02C8
      Spin            =   "frmNewBatch.frx":0326
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "11/01/2008"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39753
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate tdbDateCheck 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   661
      Calendar        =   "frmNewBatch.frx":034E
      Caption         =   "frmNewBatch.frx":044E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmNewBatch.frx":04B8
      Keys            =   "frmNewBatch.frx":04D6
      Spin            =   "frmNewBatch.frx":0534
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "11/01/2008"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   39753
      CenturyMode     =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Employee Sort Order:"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Select Deductions to use:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   -360
      TabIndex        =   9
      Top             =   2880
      Width           =   15
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Add/Edit a new payroll data entry batch"
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
      Left            =   1635
      TabIndex        =   8
      Top             =   720
      Width           =   4200
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
      Left            =   248
      TabIndex        =   7
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmNewBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rsItem As New ADODB.Recordset
Public BatchID As Long
Public SortOrder As Byte

Private Sub Form_Load()
    
    Response = False
    
    Me.lblCompanyName = Trim(PRCompany.Name)

    PRBatch.OpenRS

    If Me.BatchID <> 0 Then
        SQLString = "SELECT * FROM PRBatch WHERE BatchID = " & Me.BatchID
        If Not PRBatch.GetBySQL(SQLString) Then
            MsgBox "PRBatch Not Found: " & Me.BatchID, vbCritical
            End
        End If
        tdbDateSet Me.tdbDateCheck, PRBatch.CheckDate
        tdbDateSet Me.tdbDatePeriodEnding, PRBatch.PEDate
    
    Else
        
        tdbDateSet Me.tdbDateCheck, Int(Now()) + PRCompany.CheckDays
        tdbDateSet Me.tdbDatePeriodEnding, Int(Now())
    
    End If
    
    tdbIntegerSet Me.tdbIntStartCheck
    Me.tdbIntStartCheck.Format = "########0"
    Me.tdbIntStartCheck.DisplayFormat = ""
    
    Me.tdbIntStartCheck.Value = PRCompany.LastCheckNum + 1
    
    ' sort order
    With Me.cmbSortOrder
        .AddItem "By EE Number"
        .AddItem "By EE Name"
        .AddItem "By Dept By EE Number"
        .AddItem "By Dept By EE Name"
        .ListIndex = PRCompany.DfltSortOrder
    End With
    
    Me.KeyPreview = True
        
    InitItemList
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub

Private Sub cmdExit_Click()
    Response = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    SortOrder = Me.cmbSortOrder.ListIndex

    If Me.BatchID = 0 Then
        ' create the PRBatch record
        PRBatch.Clear
        PRBatch.CreateDate = Now()
        PRBatch.PEDate = Me.tdbDatePeriodEnding
        PRBatch.CheckDate = Me.tdbDateCheck
        PRBatch.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
        PRBatch.UserID = UserID
        PRBatch.Save (Equate.RecAdd)
    Else
        ' did the dates change? - if so update PRHist / PRDist / PRItemHist
        If PRBatch.PEDate <> Me.tdbDatePeriodEnding Or PRBatch.CheckDate <> Me.tdbDateCheck Then
        
            PRBatch.PEDate = Me.tdbDatePeriodEnding
            PRBatch.CheckDate = Me.tdbDateCheck
            PRBatch.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
            PRBatch.UserID = UserID
            PRBatch.Save (Equate.RecPut)
            
            SQLString = "SELECT * FROM PRHist WHERE BatchID = " & Me.BatchID
            If PRHist.GetBySQL(SQLString) Then
                Do
                    PRHist.PEDate = Me.tdbDatePeriodEnding
                    PRHist.CheckDate = Me.tdbDateCheck
                    PRHist.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
                    PRHist.Save (Equate.RecPut)
                    If Not PRHist.GetNext Then Exit Do
                Loop
            End If
            
            SQLString = "SELECT * FROM PRDist WHERE BatchID = " & Me.BatchID
            If PRDist.GetBySQL(SQLString) Then
                Do
                    PRDist.PEDate = Me.tdbDatePeriodEnding
                    PRDist.CheckDate = Me.tdbDateCheck
                    PRDist.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
                    PRDist.Save (Equate.RecPut)
                    If Not PRDist.GetNext Then Exit Do
                Loop
            End If
            
            SQLString = "SELECT * FROM PRItemHist WHERE BatchID = " & Me.BatchID
            If PRItemHist.GetBySQL(SQLString) Then
                Do
                    PRItemHist.PEDate = Me.tdbDatePeriodEnding
                    PRItemHist.CheckDate = Me.tdbDateCheck
                    PRItemHist.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
                    PRItemHist.Save (Equate.RecPut)
                    If Not PRItemHist.GetNext Then Exit Do
                Loop
            End If
            
        End If
    
    End If
    
    Response = True
    frmEntryDPT.StartCheckNumber = nNull(Me.tdbIntStartCheck)
    Me.Hide

End Sub

Private Sub tdbDatePeriodEnding_Change()
    If Me.BatchID <> 0 Then Exit Sub    ' don't override if editing existing batch
    Me.tdbDateCheck = Me.tdbDatePeriodEnding + PRCompany.CheckDays
End Sub
Private Sub InitItemList()
    
    ' set the record set with items
    rsItem.CursorLocation = adUseClient
    
    rsItem.Fields.Append "Select", adBoolean
    rsItem.Fields.Append "Title", adVarChar, 255, adFldIsNullable
    rsItem.Fields.Append "Type", adVarChar, 20, adFldIsNullable
    rsItem.Fields.Append "ItemID", adDouble
    
    rsItem.Open , , adOpenDynamic, adLockOptimistic
    
    ' populate with other earnings and deductions
    SQLString = "SELECT * FROM PRItem WHERE " & _
                "(PRItem.ItemType = " & PREquate.ItemTypeOE & " OR " & _
                "PRItem.ItemType = " & PREquate.ItemTypeDED & " OR " & _
                "PRItem.ItemType = " & PREquate.ItemTypeSDTax & ") " & _
                "AND PRItem.EmployeeID = 0 " & _
                "ORDER BY PRItem.ItemType, PRItem.ItemID"
    
    ' deductions only ???
    SQLString = "SELECT * FROM PRItem WHERE " & _
                "(PRItem.ItemType = " & PREquate.ItemTypeDED & _
                " OR PRItem.ItemType = " & PREquate.ItemTypeSDTax & ")" & _
                " AND PRItem.EmployeeID = 0 " & _
                "ORDER BY PRItem.ItemType, PRItem.ItemID"
    
    If PRItem.GetBySQL(SQLString) Then
    
        Do
            rsItem.AddNew
            rsItem!Select = True
            rsItem!Title = PRItem.ItemID & " " & Trim(PRItem.Title)
            If PRItem.ItemType = PREquate.ItemTypeDED Then
                rsItem!Type = "DEDUCTION"
            Else
                rsItem!Type = "SD TAX"
            End If
            rsItem!ItemID = PRItem.ItemID
            rsItem.Update
            
            If Not PRItem.GetNext Then Exit Do
        Loop
    
    End If
    
    SetGrid rsItem, fg
    
    fg.ColWidth(0) = 800
    fg.ColWidth(1) = 3500
    fg.ColWidth(2) = 1400
    fg.ColHidden(3) = True
    
End Sub

Private Sub cmdTimeSheet_Click()
    frmSelTimeSheets.Show vbModal
End Sub


