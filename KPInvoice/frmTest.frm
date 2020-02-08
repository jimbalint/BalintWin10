VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   5760
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   661
      Calendar        =   "frmTest.frx":0000
      Caption         =   "frmTest.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmTest.frx":0164
      Keys            =   "frmTest.frx":0182
      Spin            =   "frmTest.frx":01E0
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
      Text            =   "08/13/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40403
      CenturyMode     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1900
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmTest.frx":0208
      Top             =   4800
      Width           =   3800
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3615
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   8775
      _cx             =   15478
      _cy             =   6376
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
   Begin VB.Label lblMsg1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   7935
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim P As String

Private Sub Command1_Click()

    PrtInit "Port"
    SetFont 12, Equate.Portrait
    
    Prvw.vsp.TextBox Me.Text1, 1000, 1000, 3800, 1900
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    End
    
End Sub

Private Sub Form_Load()

' QB Item variables
Dim ItemQuery As IItemQuery
Dim orItemRet As IORItemRet
Dim itemServiceAdd As IItemServiceAdd
Dim itemServiceRet As IItemServiceRet
    
' General QB variables
Dim requestMsgSet As IMsgSetRequest
Dim responseMsgSet As IMsgSetResponse
Dim ResponseList As IResponseList
Dim Response As IResponse
Dim ResponseType As Integer
Dim orItemRetList As IORItemRetList
    
Dim i, j, k As Long
Dim X, Y, z As String
Dim boo As Boolean

    
'    If DP_Init("Okidata ML 320-IBM") = False Then End
'
'    X = Chr(27) & Chr(80) & Chr(14) & "Hello ...."
'    DP_PrintLine X
'
'    DP_PrintLine vbFormFeed
'    DP_EndDoc
'
'    End
    
    InvHeader.OpenRS
    InvBody.OpenRS

    k = 1000
    For i = 1 To 5
        SQLString = "SELECT * FROM JCJob"
        boo = JCJob.GetBySQL(SQLString)
        Do
            k = k + 1
            InvHeader.Clear
            InvHeader.InvoiceNumber = k
            InvHeader.OrderDate = Now()
            InvHeader.SoldJobID = JCJob.JobID
            
            InvHeader.SoldAddr1 = JCJob.BillAddr1
            InvHeader.SoldAddr2 = JCJob.BillAddr2
            InvHeader.SoldAddr3 = JCJob.BillAddr3
            InvHeader.SoldAddr4 = JCJob.BillAddr4
            InvHeader.SoldCity = JCJob.BillCity
            InvHeader.SoldState = JCJob.BillState
            InvHeader.SoldZip = JCJob.BillZip
            
            InvHeader.ShipAddr1 = JCJob.ShipAddr1
            InvHeader.ShipAddr2 = JCJob.ShipAddr2
            InvHeader.ShipAddr3 = JCJob.ShipAddr3
            InvHeader.ShipAddr4 = JCJob.ShipAddr4
            InvHeader.ShipCity = JCJob.ShipCity
            InvHeader.ShipState = JCJob.ShipState
            InvHeader.ShipZip = JCJob.ShipZip
            
            InvHeader.rsAdd

            j = 0
            SQLString = "SELECT * FROM InvStock WHERE JobID = " & JCJob.JobID
            boo = InvStock.GetBySQL(SQLString)
            Do
                InvBody.Clear
                InvBody.HeaderID = InvHeader.HeaderID
                j = j + 1
                InvBody.LineNum = j
                InvBody.QtyOrdered = j * 10
                InvBody.QtyShipped = j * 10
                InvBody.Description = InvStock.Description
                InvBody.StockID = InvStock.StockID
                InvBody.Price = InvStock.CustomerPrice
                InvBody.Amount = InvBody.Price * InvBody.QtyShipped
                InvBody.rsAdd

                If InvStock.GetNext = False Then Exit Do
            Loop
            If JCJob.GetNext = False Then Exit Do
        Loop
    Next i
    End

End Sub

Private Sub Save()

'    P = "AAAAA" & vbCrLf & "BBBBBB" & vbCrLf & "CCCCC"
'    Me.Text1.Text = P
'
'    InvGlobalCreate True
'    StockCreate True
'    HeaderCreate True
'    BodyCreate True
'
'    End
'
'    Me.Show
'
'    If QBOpen(Me, Me.lblMsg1) = False Then End
'
'    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
'    requestMsgSet.Attributes.OnError = roeContinue
'
'    ' gather the QB SERVICE items that start with PR
'    ' record QB List ID in temp RS
'    Set ItemQuery = requestMsgSet.AppendItemQueryRq
'
'    ' get all of them - update active flag
'    ' ItemQuery.ORListQuery.ListFilter.ActiveStatus.SetValue asActiveOnly
'
'    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
'
'    If Not (responseMsgSet Is Nothing) Then
'
'        Me.lblMsg1 = "Now Parsing QB Item Query ..."
'        Me.Refresh
'
'        Set ResponseList = responseMsgSet.ResponseList
'        For i = 0 To ResponseList.Count - 1
'
'            Set Response = ResponseList.GetAt(i)
'            If Response.StatusCode <> 0 Then GoTo itemNxtI
'            If Response.Detail Is Nothing Then GoTo itemNxtI
'            ResponseType = Response.Type.GetValue
'            If ResponseType <> rtItemQueryRs Then GoTo itemNxtI
'
'            Set orItemRetList = Response.Detail
'            k = orItemRetList.Count - 1
'            For j = 0 To k
'
'                Me.lblMsg1 = "Item: " & j & " of: " & k
'                Me.Refresh
'
'                Set orItemRet = orItemRetList.GetAt(j)
'
'                ' service items
'                If (Not orItemRet.itemServiceRet Is Nothing) Then
'
'                    If (Not orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase Is Nothing) Then
'                        MsgBox "Service Item: " & orItemRet.itemServiceRet.Name.GetValue
'                    End If
'
'                End If
'
'                If (Not orItemRet.ItemInventoryRet Is Nothing) Then
'                    MsgBox "Inventory Item: " & orItemRet.ItemInventoryRet.Name.GetValue
'                End If
'
'                If (Not orItemRet.ItemNonInventoryRet Is Nothing) Then
'                    MsgBox "Non Inventory Item: " & orItemRet.ItemNonInventoryRet.Name.GetValue
'                End If
'
'            Next j
'
'itemNxtI:
'        Next i
'
'    End If
'
'
''                        ' QB Service item name is the format:
''                        ' PR_{EE Last/FirstName}/{Emp#}
''                        X = orItemRet.itemServiceRet.Name.GetValue
''                        If Mid(X, 1, 3) = "PR_" Then
''                            EmpNum = ParseEENum(X)
''                            If EmpNum <> 0 Then
''                                SQLString = "SELECT * FROM PREmployee WHERE EmployeeNumber = " & EmpNum
''                                If PREmployee.GetBySQL(SQLString) = True Then
''                                    rsEE.Find "EmployeeID = " & PREmployee.EmployeeID, 0, adSearchForward, 1
''                                    If rsEE.EOF Then
''                                        rsEE.AddNew
''                                        rsEE!EmployeeID = PREmployee.EmployeeID
''                                        rsEE!QBItemID = orItemRet.itemServiceRet.ListID.GetValue
''                                        rsEE.Update
''                                    Else
''                                        rsEE!QBItemID = orItemRet.itemServiceRet.ListID.GetValue
''                                    End If
''                                    rsEE.Update
''                                End If
''                            End If
''                        End If
''                    End If
''                End If
'
'
'
'
'    End

End Sub
