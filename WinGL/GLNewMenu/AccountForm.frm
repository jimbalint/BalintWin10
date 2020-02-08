VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmAccount 
   Caption         =   " Accounts"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAmounts 
      Caption         =   "A&MOUNTS"
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin TDBNumber6Ctl.TDBNumber txtAcctSearch 
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "AccountForm.frx":0000
      Caption         =   "AccountForm.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AccountForm.frx":008C
      Keys            =   "AccountForm.frx":00AA
      Spin            =   "AccountForm.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DEL"
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   3880
      Width           =   1335
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&EDIT"
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   2960
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   495
      Left            =   9600
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6135
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   5535
      _cx             =   9763
      _cy             =   10821
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FocusRect       =   3
      HighLight       =   2
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
      FixedCols       =   0
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
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdUnccheck 
      Caption         =   "&UnCheck All"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check All"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   9600
      TabIndex        =   6
      Top             =   6840
      Width           =   1335
   End
   Begin VB.ListBox lstType 
      Height          =   4110
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   840
      Width           =   3495
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
      Left            =   480
      TabIndex        =   13
      Top             =   120
      Width           =   9735
   End
   Begin VB.Label lblSearch 
      Caption         =   "Acct# &Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "FILE"
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mra As ADODB.Recordset
Private GotFocusColor As Long
Private LostFocusColor As Long
Dim FRow As Long

Dim SortCol As Integer
Dim SortDesc As Boolean


Private Sub cmdAdd_Click()

     FormType = Equate.FormAdd
     frmAcctEdit.Show vbModal
     
     ' update the grid
     fg.Rows = fg.Rows + 1
     fg.Row = fg.Rows - 1
     fg.TextMatrix(fg.Row, 0) = GLAccount.AcctType
     fg.TextMatrix(fg.Row, 1) = GLAccount.Account
     fg.TextMatrix(fg.Row, 2) = GLAccount.Description
     fg.AutoSize (2)
     fg.Col = 1
     fg.ShowCell fg.Row, 1
     fg.Sort = flexSortGenericAscending
     fg.SetFocus
     
'     ' ??? sort by acct after add
'     cmdSortAcct_Click

     ' select the acct just added
     fg.ShowCell fg.FindRow(GLAccount.Account, , 1, False), 1
     fg.Row = fg.FindRow(GLAccount.Account, , 1, False)
     fg.Col = 1

'    cmdEdit_Click
'    fg.Rows = fg.Rows + 1
'    fg.Row = fg.Rows - 1
'    fg.Col = 1
'    fg.TextMatrix(fg.Row, 0) = "0"
'    fg.TextMatrix(fg.Row, 1) = "4444"
'    fg.TextMatrix(fg.Row, 2) = ""
'    fg.ShowCell fg.Row, 1
'    fg.SetFocus
'
'    Dim mradd As ADODB.Recordset
'    SetAdo cn, mradd, "Select * from glAccount"
'
'    mradd!AcctType = "0"
'    mradd!Account = 0
'    mradd!Description = "New Account"
'    mradd!AllSchedules = False
'    mradd!AllStatements = False
'    mradd!BranchAcct = False
'    mradd!BSColumn = 0
'    mradd!ConsAcct = False
'    mradd!Date1 = 0
'    mradd!Date2 = 0
'    mradd!DollarSign = False
'    mradd!LineFeeds = 0
'    mradd!PrintTab = 0
'    mradd!SignRevSched = False
'    mradd!SignRevStmt = False
'    mradd!TotalLevel = 0
'    mradd!TotalOnLedger = False
'    mradd.AddNew
'    mradd.Update
'    mradd.Close
    
'    fg.Rows = fg.Rows + 1
'    fg.Row = fg.Rows - 1
'    fg.TextMatrix(fg.Row, 0) = mra!AcctType
'    fg.TextMatrix(fg.Row, 1) = mra!Account
'    fg.TextMatrix(fg.Row, 2) = mra!Description
'    fg.AutoSize (2)
'    fg.Col = 1
'    fg.ShowCell fg.Row, 1
'    fg.SetFocus
'    cmdEdit_Click
End Sub

Private Sub cmdAmounts_Click()
Dim x As String
    
    If fg.TextMatrix(fg.Row, 0) <> "0" Then Exit Sub
    
    If fg.Row = 0 Then Exit Sub
    
    x = "\Balint\BuckEdit.exe User=" & UserID & " " & _
        "Account=" & fg.TextMatrix(fg.Row, 1) & " " & _
        "Password=" & dbPwd
       
    ExecCmd (x)

    fg.SetFocus

End Sub

Private Sub cmdCheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstType.ListCount - 1
        lstType.Selected(ndx) = True
    Next ndx
End Sub

Private Sub cmdDelete_Click()
     
Dim MResp As Integer
Dim Acct As Long
     
     Response = GLAccount.GetAccount(CLng(fg.TextMatrix(fg.Row, 1)))
     Acct = GLAccount.Account
     
     If Response = False Then
        MsgBox "Account not found ?", vbCritical
        End
     End If
     
     MResp = MsgBox("Are you SURE you want to delete: " & vbCr & vbCr & _
            "Account # " & GLAccount.Account & " " & GLAccount.GetDesc & vbCr & _
            "Type: " & GLAccount.AcctType, vbExclamation + vbYesNo + vbDefaultButton2, "DELETE Account ?")
                 
     If MResp = vbNo Then
        fg.SetFocus
        Exit Sub
     End If
     
     If Not GLAccount.DeleteRecord(Acct) Then
        MsgBox "Delete failed for Acct # " & Acct, vbCritical + vbOKOnly
     End If
     
     LoadGrid
     
     fg.SetFocus
     
End Sub

Private Sub cmdEdit_Click()

     Response = GLAccount.GetAccount(CLng(fg.TextMatrix(fg.Row, 1)))
     If Response = False Then
        MsgBox "Account not found ?", vbCritical
        End
     End If
     
     GLAccount.AssignFields
     
     FormType = Equate.FormEdit
     frmAcctEdit.Show vbModal
     
     ' update the display
     fg.TextMatrix(fg.Row, 0) = GLAccount.AcctType
     fg.TextMatrix(fg.Row, 1) = GLAccount.Account
     fg.TextMatrix(fg.Row, 2) = GLAccount.GetDesc
     
'    If CNOpen(glFileName(0)) Then
'        frmAccount.strAccount = fg.TextMatrix(fg.Row, 1)
'        frmAccount.Show vbModal, Me
'        fg.TextMatrix(fg.Row, 0) = mra!AcctType
'        fg.TextMatrix(fg.Row, 1) = mra!Account
'        fg.TextMatrix(fg.Row, 2) = mra!Description
'        fg.AutoSize (2)
'    End If
    fg.SetFocus
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdLoad_Click()
    LoadGrid
    fg.SetFocus
End Sub

Private Sub cmdSortAcct_Click()
    fg.Select 1, 1
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdSortDesc_Click()
    fg.Select 1, 2
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdSortType_Click()
    fg.Select 1, 0
    fg.Sort = flexSortGenericAscending
    fg.SetFocus
End Sub

Private Sub cmdUnccheck_Click()
    Dim ndx As Integer
    For ndx = 0 To lstType.ListCount - 1
        lstType.Selected(ndx) = False
    Next ndx
End Sub

Private Sub fg_EnterCell()
    If fg.Row = 0 Then Exit Sub
    If IsNumeric(fg.TextMatrix(fg.Row, 1)) Then
'        If CNOpen(glFileName(0)) Then
        
        If CNOpen(DBName, dbPwd) Then
            SetAdo cn, mra, "Select * from glAccount where Account = " & fg.TextMatrix(fg.Row, 1)
        Else
        End If
    Else
    End If
End Sub

Private Sub fg_GotFocus()
    fg.BackColor = GotFocusColor
End Sub

Private Sub fg_LostFocus()
    fg.BackColor = LostFocusColor
End Sub

Private Sub Form_Load()
    
    frmAccount.Caption = " Accounts for " & glCompanyName
    Dim ndx As Integer
    
    lstType.AddItem " "   ' ?????????
    ndx = 1
    Do Until glTypeChar(ndx) = " "
       lstType.AddItem glTypeChar(ndx) & " " & glTypeName(ndx)
       ndx = ndx + 1
    Loop
    
    cmdCheck_Click
    
    Me.lblCompanyName = GLCompany.Name
    
    LoadGrid
    
    LostFocusColor = frmAccount.BackColor
    GotFocusColor = fg.BackColor
    
    lstType.BackColor = LostFocusColor

    ' select a whole row at a time
    fg.SelectionMode = flexSelectionByRow

End Sub

Private Sub LoadGrid()

Dim ct As Long

'    On Error GoTo glErr
    
'    Dim mrs As ADODB.Recordset
'    If CNOpen(DBName, Password) Then
'        SetAdo cn, mrs, "Select [Account],[AcctType],[Description] from glAccount"
'    End If
    
    
    On Error GoTo 0
    
    fg.Redraw = False
    fg.Rows = 1
    fg.Cols = 3
    fg.ColDataType(2) = flexDTString
    
    fg.ColAlignment(0) = flexAlignCenterCenter
    fg.ColAlignment(1) = flexAlignRightCenter
    fg.ColAlignment(2) = flexAlignLeftCenter
    
    fg.TextMatrix(0, 0) = "Type"
    fg.TextMatrix(0, 1) = "Acct# +"
    fg.TextMatrix(0, 2) = "Description"
    
    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
'    fg.TabBehavior = flexTabCells                       ' tab moves between cells
'    fg.HighLight = flexHighlightNever                   ' don't select ranges
    
    ' default sort by acct ascending
    SortCol = 1
    SortDesc = False
    fg.Cell(flexcpFontBold, 0, 1) = True
    
    frmProgress.lblMsg1 = GLCompany.Name
    frmProgress.lblMsg2 = "Gathering Account Information ..."
    frmProgress.Show
    ct = 0
    
    If GLAccount.GetAllAccounts Then
 
        Do
            
            ct = ct + 1
            If ct = 1 Or ct Mod 100 = 0 Then
               frmProgress.lblMsg2 = "On Account: " & GLAccount.Account
               frmProgress.lblMsg2.Refresh
            End If
            
            If lstType.Selected(glTypeByte((GLAccount.AcctType))) = True Then
                fg.Rows = fg.Rows + 1
                fg.TextMatrix(fg.Rows - 1, 0) = GLAccount.AcctType & ""
                fg.TextMatrix(fg.Rows - 1, 1) = GLAccount.Account
                fg.TextMatrix(fg.Rows - 1, 2) = GLAccount.GetDesc
            End If
            
            If Not GLAccount.GetNext Then Exit Do
        
        Loop
    
        fg.Row = 1
        fg.Col = 0
        fg.AutoSize (2)
    
        fg.Redraw = True
'        cmdSortAcct_Click   ' sort by acct by default
    
    End If
    
    fg.Redraw = True
    frmProgress.Hide
    GLDescription.OpenRS

glErr:
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub lblSearch_Click()
    Me.txtAcctSearch.SetFocus
End Sub

Private Sub lstType_GotFocus()
    lstType.BackColor = GotFocusColor
End Sub

Private Sub lstType_LostFocus()
    lstType.BackColor = LostFocusColor
End Sub

Private Sub txtAcctSearch_lostfocus()
    FRow = fg.FindRow(txtAcctSearch, , 1)
    
    If FRow = -1 Then
       MsgBox "Account not found " & txtAcctSearch, vbExclamation + vbOKOnly
       fg.Select 1, 1
       fg_GotFocus
       Exit Sub
    End If
    
    fg.ShowCell FRow, 1
    fg.Select FRow, 1

End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)

Dim Acct As String
Dim rw As Long

    ' clicking on a column header sorts based on that column
    If Button = 1 And Shift = 0 And fg.MouseRow = 0 Then

       If fg.MouseCol = SortCol Then
          
          ' toggle the sort order
          If SortDesc = False Then
             SortDesc = True
          Else
             SortDesc = False
          End If
       
       Else
          
          ' switch the column
          fg.Cell(flexcpFontBold, 0, fg.MouseCol) = True
          fg.Cell(flexcpFontBold, 0, SortCol) = False
          SortCol = fg.MouseCol
          SortDesc = False
       
       End If
       
       ' store the account currently on
       Acct = fg.TextMatrix(fg.Row, 1)
       
       Select Case SortCol
          Case 0        ' by type
               fg.Select 1, 0
               If SortDesc Then
                  fg.TextMatrix(0, 0) = "Type -"
                  fg.Sort = flexSortGenericDescending
               Else
                  fg.TextMatrix(0, 0) = "Type +"
                  fg.Sort = flexSortGenericAscending
               End If
               fg.TextMatrix(0, 1) = "Account"
               fg.TextMatrix(0, 2) = "Description"
          Case 1        ' by #
               fg.Select 1, 1
               If SortDesc Then
                  fg.TextMatrix(0, 1) = "Acct# -"
                  fg.Sort = flexSortGenericDescending
               Else
                  fg.TextMatrix(0, 1) = "Acct# +"
                  fg.Sort = flexSortGenericAscending
               End If
               fg.TextMatrix(0, 0) = "Type"
               fg.TextMatrix(0, 2) = "Description"
          Case 2        ' by description
               fg.Select 1, 2
               If SortDesc Then
                  fg.TextMatrix(0, 2) = "Description -"
                  fg.Sort = flexSortGenericDescending
               Else
                  fg.TextMatrix(0, 2) = "Description +"
                  fg.Sort = flexSortGenericAscending
               End If
               fg.TextMatrix(0, 0) = "Type"
               fg.TextMatrix(0, 1) = "Account"
       End Select
       
       ' find the row again
       fg.ShowCell 1, 1

'       rw = fg.FindRow(Acct, 1, 1)
'       fg.ShowCell rw, 1
'       fg.TopRow = rw
'       fg.SetFocus
       
    End If
    
End Sub

