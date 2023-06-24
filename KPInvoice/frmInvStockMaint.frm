VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmInvStockMaint 
   Caption         =   "Stock File Maintenance"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvStockMaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnCheckAll 
      Caption         =   "UNCHECK ALL"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "CHECK ALL"
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   375
      Left            =   11640
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "&PRINT ALL"
      Height          =   615
      Left            =   10920
      TabIndex        =   10
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton cmdApplyMasterAll 
      Caption         =   "&APPLY PRICES FROM MASTER TO ALL CUSTOMERS"
      Height          =   615
      Left            =   3960
      TabIndex        =   9
      Top             =   9000
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   12360
      TabIndex        =   6
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmdApplyMaster 
      Caption         =   "&APPLY PRICES FROM MASTER TO THIS CUSTOMER"
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   9000
      Width           =   3015
   End
   Begin VB.CommandButton cmdQBUpdate 
      Caption         =   "UPDATE PRICES FROM &QB"
      Height          =   615
      Left            =   7440
      TabIndex        =   4
      Top             =   9000
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   615
      Left            =   9480
      TabIndex        =   3
      Top             =   9000
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6735
      Left            =   735
      TabIndex        =   2
      Top             =   2040
      Width           =   13095
      _cx             =   23098
      _cy             =   11880
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
   Begin VB.ComboBox cmbSelect 
      Height          =   360
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "lblMsg1"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   9840
      Width           =   13455
   End
   Begin VB.Label Label1 
      Caption         =   "Select Price List for:"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
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
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   12615
   End
End
Attribute VB_Name = "frmInvStockMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LoadFlag As Boolean
Dim rsQB As New ADODB.Recordset
Dim rsStock As New ADODB.Recordset
Dim rs As New ADODB.Recordset

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
    
' QB Sales Terms variables
Dim termsQuery As ITermsQuery
Dim orTermsRetList As IORTermsRetList
Dim orTermsRet As IORTermsRet

' QB Sales Tax Code Variables
Dim salesTaxCodeQuery As ISalesTaxCodeQuery
Dim salesTaxCodeRetList As ISalesTaxCodeRetList
Dim salesTaxCodeRet As ISalesTaxCodeRet

Dim i, j, k As Long
Dim x, y, z As String

Dim Cost, Price As Currency
Dim QBName, Description As String
Dim Active As Boolean
Dim ShowJobID As Long

Private Sub cmdLoad_Click()
    LoadGrid
End Sub
Private Sub Form_Load()

    lblMsg1 = ""
    LoadFlag = True
    Init

    ' start with the master prices displayed
    Me.cmbSelect.ListIndex = 0
    LoadFlag = False
    LoadGrid

    ' temp recordset for Cust/Job update from QB
    rsQB.CursorLocation = adUseClient
    rsQB.Fields.Append "JobID", adDouble
    rsQB.Open , , adOpenDynamic, adLockOptimistic

    Me.lblCompanyName = PRCompany.Name
    
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:   cmdExit_Click
        Case vbKeyF7:       ReCreate
    End Select
End Sub
Private Sub cmdCheckAll_Click()
    CheckSweep True
End Sub

Private Sub cmdUnCheckAll_Click()
    CheckSweep False
End Sub
Private Sub CheckSweep(ByVal boo As Boolean)
    If rs.RecordCount = 0 Then Exit Sub
    If ShowJobID = 0 Then Exit Sub
    i = fg.Row
    rs.MoveFirst
    Do
        rs!StockSelect = boo
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
    On Error Resume Next
    fg.Row = i
    On Error GoTo 0
End Sub


Private Sub ReCreate()
    
    If User.Logon <> "jim" Then Exit Sub
    x = "OK to delete and re-create ALL invoicing database tables ?"
    If MsgBox(x, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    ' close the current stock records set
    rs.Close
    
    InvHeader.rsClose
    SQLString = "DROP TABLE InvHeader"
    cn.Execute SQLString
    HeaderCreate
    
    InvBody.rsClose
    SQLString = "DROP TABLE InvBody"
    cn.Execute SQLString
    BodyCreate
    
    InvStock.rsClose
    SQLString = "DROP TABLE InvStock"
    cn.Execute SQLString
    StockCreate
    
    JCCustomer.rsClose
    SQLString = "DROP TABLE JCCustomer"
    cn.Execute SQLString
    CustomerCreate
    
    JCJob.rsClose
    SQLString = "DROP TABLE JCJob"
    cn.Execute SQLString
    JobCreate
    
    MsgBox "Invoicing database table re-create complete!", vbInformation
    
    GoBack
    
End Sub

Private Sub Init()

    ' init Combo w/ Customer List
    With Me.cmbSelect
        
        .Clear
        .AddItem "Master List"
        .ItemData(.NewIndex) = 0
        
'        SQLString = "SELECT * FROM JCJob WHERE Active = 1 ORDER BY FullName"
'        If JCJob.GetBySQL(SQLString) = True Then
'            Do
'                .AddItem JCJob.FullName
'                .ItemData(.NewIndex) = JCJob.JobID
'                If JCJob.GetNext = False Then Exit Do
'            Loop
'        End If
        
        
        ' price is per customer - not job !!!
        SQLString = "SELECT * FROM JCCustomer WHERE Active = 1 ORDER BY FullName"
        SQLString = "SELECT * FROM JCCustomer ORDER BY FullName"
        If JCCustomer.GetBySQL(SQLString) = True Then
            Do
                
                .AddItem JCCustomer.FullName
                
                ' use the CustomerID !!!
                ' prices are per customer !!!
                .ItemData(.NewIndex) = JCCustomer.CustomerID
                
'                ' get the JobID for the customer
'                SQLString = "SELECT * FROM JCJob WHERE ParentID = " & JCCustomer.CustomerID
'                If JCJob.GetBySQL(SQLString) Then
'                    .ItemData(.NewIndex) = JCJob.JobID
'                End If
                
                If JCCustomer.GetNext = False Then Exit Do
            
            Loop
        End If
    
        .ListIndex = 0
    
    End With

End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub
Private Sub cmdPrint_Click()
    
    PrintInit
    PrintPrices cmbSelect.ItemData(cmbSelect.ListIndex)
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

End Sub
Private Sub cmdPrintAll_Click()
    
Dim cmbCount As Long

    PrintInit
    
    With Me.cmbSelect
        For cmbCount = 0 To .ListCount - 1
            PrintPrices .ItemData(cmbCount)
            If cmbCount <> .ListCount - 1 Then FormFeed
        Next cmbCount
    End With
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
    
End Sub
Private Sub PrintInit()
    
    PrvwReturn = True
    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    y = "Stock File Price Listing"

End Sub
Private Sub PrintHeader()

    Prvw.FontBold = True
    Ln = 0
    PageHeader y, z
    Ln = Ln + 1
    PrintValue(1) = "Name":             FormatString(1) = "a30"
    PrintValue(2) = "Description":      FormatString(2) = "a30"
    PrintValue(3) = "Master Price":     FormatString(3) = "r15"
    PrintValue(4) = "Customer Price":   FormatString(4) = "r15"
    PrintValue(5) = " ":                FormatString(5) = "~"
    FormatPrint
    Ln = Ln + 2
    Prvw.FontBold = False

End Sub
Private Sub PrintPrices(ByVal lngJobID As Long)
    
    If lngJobID = 0 Then
        z = "Master List"
    Else
        boo = JCJob.GetByID(lngJobID)
        z = JCJob.FullName
    End If
    
    SQLString = "SELECT * FROM InvStock WHERE JobID = " & lngJobID & _
                " ORDER BY Description"
    If InvStock.GetBySQL(SQLString) = False Then Exit Sub
    PrintHeader
    Do
        PrintValue(1) = InvStock.QBName:        FormatString(1) = "a30"
        PrintValue(2) = InvStock.Description:   FormatString(2) = "a30"
        x = Format(InvStock.MasterPrice, "###,##0.0000")
        PrintValue(3) = x:                      FormatString(3) = "r15"
        x = Format(InvStock.CustomerPrice, "###,##0.0000")
        PrintValue(4) = x:                      FormatString(4) = "r15"
        PrintValue(5) = " ":                    FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        If Ln > MaxLines Then
            FormFeed
            PrintHeader
        End If
        If InvStock.GetNext = False Then Exit Do
    Loop

End Sub

Private Sub cmdApplyMaster_Click()
    If rs.RecordCount = 0 Then Exit Sub
    x = "OK to repleace all customer prices for: " & Me.cmbSelect & vbCr & "from the master list?"
    If MsgBox(x, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    rs.MoveFirst
    Do
        rs!CustomerPrice = rs!MasterPrice
        rs.Update
        rs.MoveNext
    Loop Until rs.EOF
    rs.MoveFirst
End Sub
Private Sub cmdApplyMasterAll_Click()
    If rs.RecordCount = 0 Then Exit Sub
    x = "OK to replace ALL customer prices with the master list entries?"
    If MsgBox(x, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    x = "*** WARNING ***" & vbCr & _
        "ALL customer prices will be replaced with the master list entries" & vbCr & _
        "OK to continue?"
    If MsgBox(x, vbYesNo + vbExclamation) = vbNo Then Exit Sub
    SQLString = "SELECT * FROM InvStock WHERE JobID <> 0"
    If InvStock.GetBySQL(SQLString) = False Then Exit Sub
    Do
        InvStock.CustomerPrice = InvStock.MasterPrice
        InvStock.rsPut
        If InvStock.GetNext = False Then Exit Do
    Loop
    LoadGrid
End Sub

Private Sub cmdQBUpdate_Click()

    If QBOpen(Me, Me.lblMsg1) = False Then
        GoBack
    End If
    
    ' update cusotmers / jobs / items from QB
    DoCustomerQueryRq "US", 5, 0

    ' update JCJob.ParentID for multi-level jobs
    If rsQB.RecordCount > 0 Then
        JobFill
    End If

    ItemUpdate
    
    TermsUpdate
    
    SalesTaxCodeUpdate
    
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection

    LoadFlag = True
    Init
    LoadFlag = False
    LoadGrid
    
End Sub

Private Sub SalesTaxCodeUpdate()

    Set requestMsgSet = Nothing
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue

    Set salesTaxCodeQuery = requestMsgSet.AppendSalesTaxCodeQueryRq
    salesTaxCodeQuery.metaData.SetValue mdNoMetaData
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    If Not (responseMsgSet Is Nothing) Then
    
        Me.lblMsg1 = "Now Parsing QB Sales Tax Code Query ..."
        Me.Refresh
        
        Set ResponseList = responseMsgSet.ResponseList
        For i = 0 To ResponseList.Count - 1
            Set Response = ResponseList.GetAt(i)
            If (Response.StatusCode = 0) Then
                If (Not Response.Detail Is Nothing) Then
                    ResponseType = Response.Type.GetValue
                    If (ResponseType = rtSalesTaxCodeQueryRs) Then
                        Set salesTaxCodeRetList = Response.Detail
                        For j = 0 To salesTaxCodeRetList.Count - 1
                            Set salesTaxCodeRet = salesTaxCodeRetList.GetAt(j)
                            
                            y = salesTaxCodeRet.ListID.GetValue
                            
                            If salesTaxCodeRet.Name.IsSet = True Then
                                QBName = salesTaxCodeRet.Name.GetValue
                            Else
                                QBName = ""
                            End If
                            
                            x = salesTaxCodeRet.IsTaxable.GetValue
                                                        
                            SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'SALESTAXCODE' AND " & _
                                        "QBID = '" & y & "'"
                            If QBAccount.GetBySQL(SQLString) = False Then
                                QBAccount.Clear
                                QBAccount.AccountType = "SALESTAXCODE"
                                QBAccount.QBID = y
                                QBAccount.Save (Equate.RecAdd)
                            End If
                            
                            QBAccount.Name = QBName
                            QBAccount.Description = x
                            QBAccount.Save (Equate.RecPut)
                        Next j
                    End If
                End If
            End If
        Next i
    End If

    Me.lblMsg1 = ""
    Me.Refresh

End Sub

Private Sub TermsUpdate()
    
    Me.MousePointer = vbHourglass
        
    InvGlobal.OpenRS
    
    Set requestMsgSet = Nothing
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue

    Set termsQuery = requestMsgSet.AppendTermsQueryRq
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    If Not (responseMsgSet Is Nothing) Then
    
        Me.lblMsg1 = "Now Parsing QB Sales Item Query ..."
        Me.Refresh
    
        Set ResponseList = responseMsgSet.ResponseList
        For i = 0 To ResponseList.Count - 1
            Set Response = ResponseList.GetAt(i)
            If (Response.StatusCode = 0) Then
                If (Not Response.Detail Is Nothing) Then
                    ResponseType = Response.Type.GetValue
                    If (ResponseType = rtTermsQueryRs) Then
                        Set orTermsRetList = Response.Detail
                        For j = 0 To orTermsRetList.Count - 1
                            Set orTermsRet = orTermsRetList.GetAt(j)
                            If orTermsRet.StandardTermsRet.Name.IsSet Then
                                SQLString = "SELECT * FROM InvGlobal WHERE TypeCode = " & InvEquate.GlobalTypeTerms & _
                                            " AND Var1 = '" & orTermsRet.StandardTermsRet.ListID.GetValue & "'" & _
                                            " AND CompanyID = " & PRCompany.CompanyID
                                If InvGlobal.GetBySQL(SQLString) = False Then
                                    InvGlobal.Clear
                                    InvGlobal.TypeCode = InvEquate.GlobalTypeTerms
                                    InvGlobal.Var1 = orTermsRet.StandardTermsRet.ListID.GetValue
                                    InvGlobal.Description = orTermsRet.StandardTermsRet.Name.GetValue
                                    InvGlobal.CompanyID = PRCompany.CompanyID
                                    InvGlobal.rsAdd
                                Else
                                    InvGlobal.Description = orTermsRet.StandardTermsRet.Name.GetValue
                                    InvGlobal.rsPut
                                End If
                            End If
                        Next j
                    End If
                End If
            End If
        Next i
                        
    End If
    
    Me.lblMsg1 = ""
    Me.Refresh
    Me.MousePointer = vbArrow

End Sub

Private Sub ItemUpdate()

    Dim rsQBID As New ADODB.Recordset

    Me.MousePointer = vbHourglass
        
    InvStock.OpenRS
    
    On Error Resume Next
    rsQBID.Close
    On Error GoTo 0
    rsQBID.CursorLocation = adUseClient
    rsQBID.Fields.Append "QBID", adVarChar, 50, adFldIsNullable
    rsQBID.Fields.Append "Flag", adInteger
    rsQBID.Open , , adOpenDynamic, adLockOptimistic
    
    Set requestMsgSet = Nothing
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    
    ' gather the QB SERVICE items that start with PR
    ' record QB List ID in temp RS
    Set ItemQuery = requestMsgSet.AppendItemQueryRq
    
    ' get all of them - update active flag
    ' ItemQuery.ORListQuery.ListFilter.ActiveStatus.SetValue asActiveOnly
    
    Set responseMsgSet = Nothing
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    If Not (responseMsgSet Is Nothing) Then
    
        Me.lblMsg1 = "Now Parsing QB Item Query ..."
        Me.Refresh
    
        Set ResponseList = responseMsgSet.ResponseList
        For i = 0 To ResponseList.Count - 1
        
            Set Response = ResponseList.GetAt(i)
            If Response.StatusCode <> 0 Then GoTo itemNxtI
            If Response.Detail Is Nothing Then GoTo itemNxtI
            ResponseType = Response.Type.GetValue
            If ResponseType <> rtItemQueryRs Then GoTo itemNxtI
            
            Set orItemRetList = Response.Detail
            k = orItemRetList.Count - 1
            For j = 0 To k
                
                Me.lblMsg1 = "Item: " & j & " of: " & k
                Me.Refresh
                
                Set orItemRet = orItemRetList.GetAt(j)
                            
                ' *** inventory items ***
                If (Not orItemRet.ItemInventoryRet Is Nothing) Then
                    x = orItemRet.ItemInventoryRet.ListID.GetValue
                    
                    QBName = Trim(orItemRet.ItemInventoryRet.Name.GetValue)
                    Description = QBName
                    
                    ' IsSet / IsEmpty not working for this field ?
                    y = ""
                    On Error Resume Next
                    y = orItemRet.ItemInventoryRet.SalesDesc.GetValue
                    On Error GoTo 0
                    If y <> "" Then Description = y
                    
'                    If orItemRet.ItemInventoryRet.SalesDesc.IsEmpty = False Then
'                    'If orItemRet.ItemInventoryRet.SalesDesc.IsSet = True Then
'                        Description = Trim(orItemRet.ItemInventoryRet.SalesDesc.GetValue)
'                    End If
'
'                    If orItemRet.ItemInventoryRet.PurchaseCost.IsSet Then
'                        Cost = orItemRet.ItemInventoryRet.PurchaseCost.GetValue
'                    Else
'                        Cost = 0
'                    End If
                    
                    If orItemRet.ItemInventoryRet.SalesPrice.IsSet Then
                        Price = orItemRet.ItemInventoryRet.SalesPrice.GetValue
                    Else
                        Price = 0
                    End If
                    If orItemRet.ItemInventoryRet.IsActive.IsSet Then
                        Active = orItemRet.ItemInventoryRet.IsActive.GetValue
                    Else
                        Active = False
                    End If
                    
                    MasterItemUpd x, 0, QBName, Description, Cost, Price, Active, True
                    
                    ' add to temp rs
                    rsQBID.AddNew
                    rsQBID!QBID = InvStock.QBID
                    rsQBID.Update
                
                End If
                
                ' *** non-inventory items ***
                If (Not orItemRet.ItemNonInventoryRet Is Nothing) Then
                    x = orItemRet.ItemNonInventoryRet.ListID.GetValue
                    QBName = Trim(orItemRet.ItemNonInventoryRet.Name.GetValue)
                    
                    Description = QBName
                    y = ""
                    On Error Resume Next
                    y = orItemRet.ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue
                    On Error GoTo 0
                    If y <> "" Then Description = y
                                        
'                    If orItemRet.ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc.IsSet Then
'                        Description = Trim(orItemRet.ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue)
'                    Else
'                        Description = ""
'                    End If
                    
                    Cost = 0        ' not used for non-inventory items
                    
                    If orItemRet.ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.IsSet Then
                        Price = orItemRet.ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue
                    Else
                        Price = 0
                    End If
                    If orItemRet.ItemNonInventoryRet.IsActive.IsSet Then
                        Active = orItemRet.ItemNonInventoryRet.IsActive.GetValue
                    Else
                        Active = False
                    End If
                    MasterItemUpd x, 0, QBName, Description, Cost, Price, Active, False
                    
                    ' add to temp rs
                    rsQBID.AddNew
                    rsQBID!QBID = InvStock.QBID
                    rsQBID.Update
                
                End If
                
                ' *** sales tax percentage ***
                If (Not orItemRet.ItemSalesTaxRet Is Nothing) Then
                    
                    x = orItemRet.ItemSalesTaxRet.ListID.GetValue
                    QBName = Trim(orItemRet.ItemSalesTaxRet.Name.GetValue)
                    
                    If orItemRet.ItemSalesTaxRet.ItemDesc.IsSet Then
                        Description = orItemRet.ItemSalesTaxRet.ItemDesc.GetValue
                    Else
                        Description = ""
                    End If
                                        
                    If orItemRet.ItemSalesTaxRet.TaxRate.IsSet Then
                        Price = orItemRet.ItemSalesTaxRet.TaxRate.GetValue
                    Else
                        Price = 0
                    End If
                                   
                    SQLString = "SELECT * FROM QBAccount WHERE AccountType = 'SALESTAX' AND " & _
                                "QBID = '" & x & "'"
                    If QBAccount.GetBySQL(SQLString) = False Then
                        QBAccount.Clear
                        QBAccount.AccountType = "SALESTAX"
                        QBAccount.QBID = x
                        QBAccount.Save (Equate.RecAdd)
                    End If
                    
                    QBAccount.Name = QBName
                    QBAccount.Description = Description
                    QBAccount.AccountNumber = CStr(Price * 10000)
                    QBAccount.Save (Equate.RecPut)
                
                End If
                
            Next j
                    
itemNxtI:
        Next i
    
    End If

    ' update name and description from master to all stock items
    Dim rsMaster As New ADODB.Recordset
    SQLString = " SELECT QBID, QBName, Description FROM InvStock WHERE JobID = 0 " + _
                " ORDER BY Description"
    rsInit SQLString, cn, rsMaster
    If rsMaster.RecordCount > 0 Then
        rsMaster.MoveFirst
        Do
            Me.lblMsg1 = "Now Updating Descriptions ... " + rsMaster!Description
            Me.Refresh
            SQLString = " UPDATE InvStock SET " + _
                        " QBName = '" + Trim(rsMaster!QBName) + "', " + _
                        " Description = '" + Trim(rsMaster!Description) + "' " + _
                        " WHERE JobID <> 0 AND QBID = '" + Trim(rsMaster!QBID) + "'"
            
            ' need fix for single quote in description/qbname
            On Error Resume Next
            cn.Execute SQLString
            On Error GoTo 0
            
            rsMaster.MoveNext
        Loop Until rsMaster.EOF
    End If

    rsMaster.Close

    ' ***********************
    ' items per job update done on demand
    ' during stock maint or invoicing
    ' ***********************

'    ' update to JCJob records
'    SQLString = "SELECT * FROM JCJob"
'    If JCJob.GetBySQL(SQLString) = False Then Exit Sub
'
'    SQLString = "SELECT * FROM InvStock WHERE JobID = 0 ORDER BY QBName"
'
'    rsInit SQLString, cn, rsStock
'
'    ' use custom code instead of rsInit
'    ' *** static cursor ***
'    Set rsStock = New ADODB.Recordset
'    rsStock.Source = SQLString
'    rsStock.ActiveConnection = cn
'    rsStock.CursorLocation = adUseServer
'    rsStock.CursorType = adOpenStatic
'    rsStock.LockType = adLockOptimistic
'    rsStock.Open
'
'    If rsStock.RecordCount = 0 Then Exit Sub
'    rsStock.MoveFirst
'    Do
'
'        Me.lblMsg1 = "Now updating stock file for: " & rsStock!Description
'        Me.Refresh
'
'        JCJob.GetFirst
'        Do
'            ItemUpd rsStock!QBID, _
'                    JCJob.JobID, _
'                    rsStock!QBName, _
'                    rsStock!Description, _
'                    rsStock!Cost, _
'                    rsStock!MasterPrice, _
'                    rsStock!Active, _
'                    rsStock!InventoryItem
'            If JCJob.GetNext = False Then Exit Do
'        Loop
'        rsStock.MoveNext
'
'    Loop Until rsStock.EOF
'
'    ' remove stock items not in QB
'    Dim rsStk As New ADODB.Recordset
'    rsStk.CursorLocation = adUseClient
'    rsStk.Fields.Append "StockID", adDouble
'    rsStk.Open , , adOpenDynamic, adLockOptimistic
'
'    SQLString = "SELECT * FROM InvStock"
'    If InvStock.GetBySQL(SQLString) = True Then
'        Do
'            SQLString = "QBID = '" & InvStock.QBID & "'"
'            rsQBID.Find SQLString, 0, adSearchForward, 1
'            If rsQBID.EOF = True Then
'                rsStk.AddNew
'                rsStk!StockID = InvStock.StockID
'                rsStk.Update
'            End If
'            If InvStock.GetNext = False Then Exit Do
'        Loop
'    End If
'
'    If rsStk.RecordCount > 0 Then
'        rsStk.MoveFirst
'        Do
'            SQLString = "DELETE * FROM InvStock WHERE StockID = " & rsStk!StockID
'            cn.Execute SQLString
'            rsStk.MoveNext
'        Loop Until rsStk.EOF
'    End If
'
'    rsQBID.Close
'    rsStk.Close

    Me.MousePointer = vbArrow
    Me.lblMsg1 = ""

End Sub
Private Sub cmbSelect_Click()
 
'    If LoadFlag = True Then Exit Sub
'    If Me.cmbSelect.ListIndex = 0 Then
'        Me.cmdApplyMaster.Enabled = False
'    Else
'        Me.cmdApplyMaster.Enabled = True
'    End If
'    LoadGrid

'    On Error Resume Next
'    rs.Close
'    On Error GoTo 0

End Sub

Private Sub LoadGrid(Optional DelaySeconds As Double)
  
Dim Time1, Time2 As Double
  
    If Me.cmbSelect.ListIndex = -1 Then Exit Sub
    If LoadFlag = True Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    ' = -1 - clear the screen if switching CMB and load button not hit
    If ShowJobID <> -1 Then
        With Me.cmbSelect
            ShowJobID = .ItemData(.ListIndex)
        End With
    End If
    
    ' update the stock items from the master stock items
    ' for this job id
    ItemUpd ShowJobID

    SQLString = "SELECT " & _
                "StockSelect, " & _
                "QBName, " & _
                "Description, " & _
                "MasterPrice, " & _
                "CustomerPrice, " & _
                "LastDate " & _
                "FROM InvStock WHERE JobID = " & ShowJobID & " " & _
                "AND Description <> 'Freight' " & _
                "ORDER BY QBName"
    rsInit SQLString, cn, rs
    
    ' pause ....
    Time1 = GetSecs(Now())
    Do
        Time2 = GetSecs(Now())
        If Time2 < Time1 Then Exit Do   ' change in days ...
        If Time2 - Time1 > DelaySeconds Then Exit Do
    Loop
    
    SetGrid rs, fg
    
    With fg
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
        Next i
        .TextMatrix(0, 0) = "Display"
        .TextMatrix(0, 1) = "QB Name"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Master Price"
        .TextMatrix(0, 4) = "Customer Price"
        .TextMatrix(0, 5) = "Last Sale Date"
        
        .ColFormat(.ColIndex("LastDate")) = "mm/dd/yyyy"
        .ColFormat(.ColIndex("MasterPrice")) = "###,###,##0.0000"
        .ColFormat(.ColIndex("CustomerPrice")) = "###,###,##0.0000"
        
        If ShowJobID = 0 Then
            .ColHidden(.ColIndex("StockSelect")) = True
            .ColHidden(.ColIndex("CustomerPrice")) = True
        Else
            .ColHidden(.ColIndex("StockSelect")) = False
            .ColHidden(.ColIndex("CustomerPrice")) = False
        End If
    
        .ColWidth(.ColIndex("StockSelect")) = 900
        .ColWidth(.ColIndex("QBName")) = 2400
        .ColWidth(.ColIndex("Description")) = 3800
        .ColWidth(.ColIndex("MasterPrice")) = 1700
        .ColWidth(.ColIndex("CustomerPrice")) = 1700
        .ColWidth(.ColIndex("LastDate")) = 1400
    
        .AutoSearch = flexSearchFromTop
    
    End With

    Me.MousePointer = vbArrow

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With fg
        If .ColKey(.Col) = "QBName" Then Cancel = True
        If .ColKey(.Col) = "Description" Then Cancel = True
        If .ColKey(.Col) = "MasterPrice" Then Cancel = True
        If .ColKey(.Col) = "LastDate" Then Cancel = True
    End With
End Sub


Public Sub DoCustomerQueryRq(country As String, MajorVersion As Integer, MinorVersion As Integer)
  
'  On Error GoTo Errs
  
    'We want to know if we've begun a session so we can end it if an
    'error sends us to the exception handler.
    
    ' Create the message set request object for the specific version messages.
    Set requestMsgSet = SessMgr.CreateMsgSetRequest(country, MajorVersion, MinorVersion)
    requestMsgSet.Attributes.OnError = roeContinue
  
    Me.lblMsg1 = "Building Customer Query ... "
    Me.Refresh
    
    BuildCustomerQueryRq requestMsgSet, country
  
    ' Perform the request and obtain a response from QuickBooks.
    Me.lblMsg1 = "Performing QB Data Request ... "
    Me.Refresh
    
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
  
    ' Close the session and connection with QuickBooks.
  
    Me.lblMsg1 = "Parsing QB Data ... "
    Me.Refresh
    
    ParseCustomerQueryRs responseMsgSet, country
  
    Exit Sub
  
Errs:
    MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error"
  
    ' SampleCodeForm.ErrorMsg.Text = Err.Description
    ' Me.Label1 = Err.Description
  
End Sub
  
Public Sub BuildCustomerQueryRq(requestMsgSet As IMsgSetRequest, country As String)
  
  If (requestMsgSet Is Nothing) Then
    MsgBox "Nuthin"
    Exit Sub
  End If
  
  'Add the request to the message set request object.
  Dim customerQuery As ICustomerQuery
  Set customerQuery = requestMsgSet.AppendCustomerQueryRq
  
  ' !!! *** problem if parent inactive / job active ***
  customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue asAll
  
End Sub
 
Public Sub ParseCustomerQueryRs(responseMsgSet As IMsgSetResponse, country As String)
  
Dim Ct, Recs As Long
Dim CCount As Long
      
  If (responseMsgSet Is Nothing) Then
    MsgBox "responseMsgSet = Nuthin"
    Exit Sub
  End If
  
  Set ResponseList = responseMsgSet.ResponseList
  
  If (ResponseList Is Nothing) Then
    MsgBox "responseList = Nuthin"
    Exit Sub
  End If
  
  ' Go through all of the responses in the list.
  
  Recs = ResponseList.Count
  
  For i = 0 To ResponseList.Count - 1
    
    Set Response = ResponseList.GetAt(i)
  
    ' Check the status returned for the response.
    If (Response.StatusCode = 0) Then
  
      ' Check to make sure the response is of the type we are expecting.
      If (Not Response.Detail Is Nothing) Then
        ResponseType = Response.Type.GetValue
        ' Check for CustomerQueryRs.
        If (ResponseType = rtCustomerQueryRs) Then
          Dim customerRetList As ICustomerRetList
          Set customerRetList = Response.Detail
          For j = 0 To customerRetList.Count - 1
          CCount = customerRetList.Count
            If j Mod 10 = 1 Then
                Me.lblMsg1 = "Getting Job Info: " & Format(j, "#,###,##0") & " of: " & Format(CCount, "#,###,##0")
                Me.Refresh
            End If
            ParseCustomerRet customerRetList.GetAt(j), country
          Next j
        End If
      End If
    End If
  Next i
End Sub
  
Private Sub ParseCustomerRet(customerRet As ICustomerRet, country As String)
  
Dim listID1 As String
Dim timeCreated2 As Date
Dim timeModified3 As Date
Dim editSequence4 As String
Dim name5 As String
Dim fullName6 As String
Dim isActive7 As Boolean
Dim fullName8 As String
Dim listID8 As String
Dim sublevel9 As Long
Dim companyName10 As String
Dim salutation11 As String
Dim firstName12 As String
Dim middleName13 As String
Dim lastName14 As String
Dim suffix15 As String
Dim addr116 As String
Dim addr217 As String
Dim addr318 As String
Dim addr419 As String
Dim city20 As String
Dim state21 As String
Dim county22 As String
Dim province23 As String
Dim postalCode24 As String
Dim country25 As String
Dim addr126 As String
Dim addr227 As String
Dim addr328 As String
Dim addr429 As String
Dim city30 As String
Dim state31 As String
Dim county32 As String
Dim province33 As String
Dim postalCode34 As String
Dim country35 As String
Dim printAs36 As String
Dim phone37 As String
Dim mobile38 As String
Dim pager39 As String
Dim altPhone40 As String
Dim fax41 As String
Dim email42 As String
Dim email43 As String
Dim contact44 As String
Dim altContact45 As String
Dim fullName46 As String
Dim listID46 As String
Dim fullName47 As String
Dim listID47 As String
Dim fullName48 As String
Dim listID48 As String
Dim balance49 As Double
Dim totalBalance50 As Double
Dim fullName51 As String
Dim listID51 As String
Dim fullName52 As String
Dim listID52 As String
Dim fullName53 As String
Dim listID53 As String
Dim fullName54 As String
Dim listID54 As String
Dim resaleNumber55 As String
Dim accountNumber56 As String
Dim creditLimit57 As Double
Dim fullName58 As String
Dim listID58 As String
Dim creditCardNumber59 As String
Dim expirationMonth60 As Long
Dim expirationYear61 As Long
Dim nameOnCard62 As String
Dim creditCardAddress63 As String
Dim creditCardPostalCode64 As String
Dim JobStatus65 As ENJobStatus
Dim jobStartDate66 As Date
Dim jobProjectedEndDate67 As Date
Dim jobEndDate68 As Date
Dim jobDesc69 As String
Dim fullName70 As String
Dim listID70 As String
Dim notes71 As String
Dim isStatementWithParent72 As Boolean
Dim deliveryMethod73 As ENDeliveryMethod
Dim fullName74 As String
Dim listID74 As String
  
  If (customerRet Is Nothing) Then
    Exit Sub
  End If
  
  'Go through all of the elements of ICustomerRet.
  
  ' Get the value of the ICustomerRet.ListID element.
  listID1 = customerRet.ListID.GetValue
  
  ' Get the value of the ICustomerRet.TimeCreated element.
  timeCreated2 = customerRet.TimeCreated.GetValue
  
  ' Get the value of the ICustomerRet.TimeModified element.
  timeModified3 = customerRet.TimeModified.GetValue
  
  ' Get the value of the ICustomerRet.EditSequence element.
  editSequence4 = customerRet.EditSequence.GetValue
  
  ' Get the value of the ICustomerRet.Name element.
  name5 = customerRet.Name.GetValue
  
  ' Get the value of the ICustomerRet.FullName element.
  fullName6 = customerRet.FullName.GetValue
  
  ' Get the value of the ICustomerRet.IsActive element.
  isActive7 = False
  If (Not customerRet.IsActive Is Nothing) Then
    isActive7 = customerRet.IsActive.GetValue
  End If
  
  ' Get the value of the ICustomerRet.ParentRef element.
  If (Not customerRet.ParentRef Is Nothing) Then
    ' Get the FullName value.
    fullName8 = customerRet.ParentRef.FullName.GetValue
  
    ' Get the ListID value.
    listID8 = customerRet.ParentRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.Sublevel element.
  sublevel9 = customerRet.Sublevel.GetValue
  
  ' Get the value of the ICustomerRet.CompanyName element.
  If (Not customerRet.CompanyName Is Nothing) Then
    companyName10 = customerRet.CompanyName.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Salutation element.
  If (Not customerRet.Salutation Is Nothing) Then
    salutation11 = customerRet.Salutation.GetValue
  End If
  
  ' Get the value of the ICustomerRet.FirstName element.
  If (Not customerRet.FirstName Is Nothing) Then
    firstName12 = customerRet.FirstName.GetValue
  End If
  
  ' Get the value of the ICustomerRet.MiddleName element.
  If (Not customerRet.MiddleName Is Nothing) Then
    middleName13 = customerRet.MiddleName.GetValue
  End If
  
  ' Get the value of the ICustomerRet.LastName element.
  If (Not customerRet.LastName Is Nothing) Then
    lastName14 = customerRet.LastName.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Suffix element.
  If (Not customerRet.Suffix Is Nothing) Then
    suffix15 = customerRet.Suffix.GetValue
  End If
  
  ' Get the value of the ICustomerRet.BillAddress element.
  If (Not customerRet.BillAddress Is Nothing) Then
    ' Get the value of the IAddress.Addr1 element.
    If (Not customerRet.BillAddress.Addr1 Is Nothing) Then
      addr116 = customerRet.BillAddress.Addr1.GetValue
    End If
  
    ' Get the value of the IAddress.Addr2 element.
    If (Not customerRet.BillAddress.Addr2 Is Nothing) Then
      addr217 = customerRet.BillAddress.Addr2.GetValue
    End If
  
    ' Get the value of the IAddress.Addr3 element.
    If (Not customerRet.BillAddress.Addr3 Is Nothing) Then
      addr318 = customerRet.BillAddress.Addr3.GetValue
    End If
  
    ' Get the value of the IAddress.Addr4 element.
    If (Not customerRet.BillAddress.Addr4 Is Nothing) Then
      addr419 = customerRet.BillAddress.Addr4.GetValue
    End If
  
    ' Get the value of the IAddress.City element.
    If (Not customerRet.BillAddress.City Is Nothing) Then
      city20 = customerRet.BillAddress.City.GetValue
    End If
  
    If (country = "US") Then
      ' Get the value of the IAddress.State element.
      If (Not customerRet.BillAddress.State Is Nothing) Then
        state21 = customerRet.BillAddress.State.GetValue
      End If
  
    End If
    If (country = "UK") Then
      ' Get the value of the IAddress.County element.
      If (Not customerRet.BillAddress.County Is Nothing) Then
        county22 = customerRet.BillAddress.County.GetValue
      End If
  
    End If
    If (country = "CA") Then
      ' Get the value of the IAddress.Province element.
      If (Not customerRet.BillAddress.Province Is Nothing) Then
        province23 = customerRet.BillAddress.Province.GetValue
      End If
  
    End If
    ' Get the value of the IAddress.PostalCode element.
    If (Not customerRet.BillAddress.PostalCode Is Nothing) Then
      postalCode24 = customerRet.BillAddress.PostalCode.GetValue
    End If
  
    ' Get the value of the IAddress.Country element.
    If (Not customerRet.BillAddress.country Is Nothing) Then
      country25 = customerRet.BillAddress.country.GetValue
    End If
  
  End If
  
  ' Get the value of the ICustomerRet.ShipAddress element.
  If (Not customerRet.ShipAddress Is Nothing) Then
    ' Get the value of the IAddress.Addr1 element.
    If (Not customerRet.ShipAddress.Addr1 Is Nothing) Then
      addr126 = customerRet.ShipAddress.Addr1.GetValue
    End If
  
    ' Get the value of the IAddress.Addr2 element.
    If (Not customerRet.ShipAddress.Addr2 Is Nothing) Then
      addr227 = customerRet.ShipAddress.Addr2.GetValue
    End If
  
    ' Get the value of the IAddress.Addr3 element.
    If (Not customerRet.ShipAddress.Addr3 Is Nothing) Then
      addr328 = customerRet.ShipAddress.Addr3.GetValue
    End If
  
    ' Get the value of the IAddress.Addr4 element.
    If (Not customerRet.ShipAddress.Addr4 Is Nothing) Then
      addr429 = customerRet.ShipAddress.Addr4.GetValue
    End If
  
    ' Get the value of the IAddress.City element.
    If (Not customerRet.ShipAddress.City Is Nothing) Then
      city30 = customerRet.ShipAddress.City.GetValue
    End If
  
    If (country = "US") Then
      ' Get the value of the IAddress.State element.
      If (Not customerRet.ShipAddress.State Is Nothing) Then
        state31 = customerRet.ShipAddress.State.GetValue
      End If
  
    End If
    If (country = "UK") Then
      ' Get the value of the IAddress.County element.
      If (Not customerRet.ShipAddress.County Is Nothing) Then
        county32 = customerRet.ShipAddress.County.GetValue
      End If
  
    End If
    If (country = "CA") Then
      ' Get the value of the IAddress.Province element.
      If (Not customerRet.ShipAddress.Province Is Nothing) Then
        province33 = customerRet.ShipAddress.Province.GetValue
      End If
  
    End If
    ' Get the value of the IAddress.PostalCode element.
    If (Not customerRet.ShipAddress.PostalCode Is Nothing) Then
      postalCode34 = customerRet.ShipAddress.PostalCode.GetValue
    End If
  
    ' Get the value of the IAddress.Country element.
    If (Not customerRet.ShipAddress.country Is Nothing) Then
      country35 = customerRet.ShipAddress.country.GetValue
    End If
  
  End If
  
  ' Get the value of the ICustomerRet.PrintAs element.
  If (Not customerRet.PrintAs Is Nothing) Then
    printAs36 = customerRet.PrintAs.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Phone element.
  If (Not customerRet.Phone Is Nothing) Then
    phone37 = customerRet.Phone.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Mobile element.
  If (Not customerRet.Mobile Is Nothing) Then
    mobile38 = customerRet.Mobile.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Pager element.
  If (Not customerRet.Pager Is Nothing) Then
    pager39 = customerRet.Pager.GetValue
  End If
  
  ' Get the value of the ICustomerRet.AltPhone element.
  If (Not customerRet.AltPhone Is Nothing) Then
    altPhone40 = customerRet.AltPhone.GetValue
  End If
  
  ' Get the value of the ICustomerRet.Fax element.
  If (Not customerRet.Fax Is Nothing) Then
    fax41 = customerRet.Fax.GetValue
  End If
  
  If (country = "US") Then
    ' Get the value of the ICustomerRet.Email element.
    If (Not customerRet.Email Is Nothing) Then
      email42 = customerRet.Email.GetValue
    End If
  
  End If
  If Not (country = "US") Then
    ' Get the value of the ICustomerRet.Email element.
    If (Not customerRet.Email Is Nothing) Then
      email43 = customerRet.Email.GetValue
    End If
  
  End If
  ' Get the value of the ICustomerRet.Contact element.
  If (Not customerRet.Contact Is Nothing) Then
    contact44 = customerRet.Contact.GetValue
  End If
  
  ' Get the value of the ICustomerRet.AltContact element.
  If (Not customerRet.AltContact Is Nothing) Then
    altContact45 = customerRet.AltContact.GetValue
  End If
  
  ' Get the value of the ICustomerRet.CustomerTypeRef element.
  If (Not customerRet.CustomerTypeRef Is Nothing) Then
    ' Get the FullName value.
    fullName46 = customerRet.CustomerTypeRef.FullName.GetValue
  
    ' Get the ListID value.
    listID46 = customerRet.CustomerTypeRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.TermsRef element.
  If (Not customerRet.TermsRef Is Nothing) Then
    ' Get the FullName value.
    fullName47 = customerRet.TermsRef.FullName.GetValue
  
    ' Get the ListID value.
    listID47 = customerRet.TermsRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.SalesRepRef element.
  If (Not customerRet.SalesRepRef Is Nothing) Then
    ' Get the FullName value.
    fullName48 = customerRet.SalesRepRef.FullName.GetValue
  
    ' Get the ListID value.
    listID48 = customerRet.SalesRepRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.Balance element.
  If (Not customerRet.Balance Is Nothing) Then
    balance49 = customerRet.Balance.GetValue
  End If
  
  ' Get the value of the ICustomerRet.TotalBalance element.
  If (Not customerRet.TotalBalance Is Nothing) Then
    totalBalance50 = customerRet.TotalBalance.GetValue
  End If
  
  If (country = "CA") Then
    ' Get the value of the ICustomerRet.TaxCodeRef element.
    If (Not customerRet.TaxCodeRef Is Nothing) Then
      ' Get the FullName value.
      fullName51 = customerRet.TaxCodeRef.FullName.GetValue
  
      ' Get the ListID value.
      listID51 = customerRet.TaxCodeRef.ListID.GetValue
  
    End If
  
  End If
  If (country = "UK") Then
    ' Get the value of the ICustomerRet.TaxCodeRef element.
    If (Not customerRet.TaxCodeRef Is Nothing) Then
      ' Get the FullName value.
      fullName52 = customerRet.TaxCodeRef.FullName.GetValue
  
      ' Get the ListID value.
      listID52 = customerRet.TaxCodeRef.ListID.GetValue
  
    End If
  
  End If
  If (country = "US") Then
    ' Get the value of the ICustomerRet.SalesTaxCodeRef element.
    If (Not customerRet.SalesTaxCodeRef Is Nothing) Then
      ' Get the FullName value.
      fullName53 = customerRet.SalesTaxCodeRef.FullName.GetValue
  
      ' Get the ListID value.
      listID53 = customerRet.SalesTaxCodeRef.ListID.GetValue
  
    End If
  
  End If
  If (country = "US") Then
    ' Get the value of the ICustomerRet.ItemSalesTaxRef element.
    If (Not customerRet.ItemSalesTaxRef Is Nothing) Then
      ' Get the FullName value.
      fullName54 = customerRet.ItemSalesTaxRef.FullName.GetValue
  
      ' Get the ListID value.
      listID54 = customerRet.ItemSalesTaxRef.ListID.GetValue
  
    End If
  
  End If
  ' Get the value of the ICustomerRet.ResaleNumber element.
  If (Not customerRet.ResaleNumber Is Nothing) Then
    resaleNumber55 = customerRet.ResaleNumber.GetValue
  End If
  
  ' Get the value of the ICustomerRet.AccountNumber element.
  If (Not customerRet.AccountNumber Is Nothing) Then
    accountNumber56 = customerRet.AccountNumber.GetValue
  End If
  
  ' Get the value of the ICustomerRet.CreditLimit element.
  If (Not customerRet.CreditLimit Is Nothing) Then
    creditLimit57 = customerRet.CreditLimit.GetValue
  End If
  
  ' Get the value of the ICustomerRet.PreferredPaymentMethodRef element.
  If (Not customerRet.PreferredPaymentMethodRef Is Nothing) Then
    ' Get the FullName value.
    fullName58 = customerRet.PreferredPaymentMethodRef.FullName.GetValue
  
    ' Get the ListID value.
    listID58 = customerRet.PreferredPaymentMethodRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.CreditCardInfo element.
  If (Not customerRet.CreditCardInfo Is Nothing) Then
    ' Get the value of the ICreditCardInfo.CreditCardNumber element.
    If (Not customerRet.CreditCardInfo.CreditCardNumber Is Nothing) Then
      creditCardNumber59 = customerRet.CreditCardInfo.CreditCardNumber.GetValue
    End If
  
    ' Get the value of the ICreditCardInfo.ExpirationMonth element.
    If (Not customerRet.CreditCardInfo.ExpirationMonth Is Nothing) Then
      expirationMonth60 = customerRet.CreditCardInfo.ExpirationMonth.GetValue
    End If
  
    ' Get the value of the ICreditCardInfo.ExpirationYear element.
    If (Not customerRet.CreditCardInfo.ExpirationYear Is Nothing) Then
      expirationYear61 = customerRet.CreditCardInfo.ExpirationYear.GetValue
    End If
  
    ' Get the value of the ICreditCardInfo.NameOnCard element.
    If (Not customerRet.CreditCardInfo.NameOnCard Is Nothing) Then
      nameOnCard62 = customerRet.CreditCardInfo.NameOnCard.GetValue
    End If
  
    ' Get the value of the ICreditCardInfo.CreditCardAddress element.
    If (Not customerRet.CreditCardInfo.CreditCardAddress Is Nothing) Then
      creditCardAddress63 = customerRet.CreditCardInfo.CreditCardAddress.GetValue
    End If
  
    ' Get the value of the ICreditCardInfo.CreditCardPostalCode element.
    If (Not customerRet.CreditCardInfo.CreditCardPostalCode Is Nothing) Then
      creditCardPostalCode64 = customerRet.CreditCardInfo.CreditCardPostalCode.GetValue
    End If
  
  End If
  
  ' Get the value of the ICustomerRet.JobStatus element.
  If (Not customerRet.JobStatus Is Nothing) Then
    JobStatus65 = customerRet.JobStatus.GetValue
  End If
  
  ' Get the value of the ICustomerRet.JobStartDate element.
  If (Not customerRet.JobStartDate Is Nothing) Then
    jobStartDate66 = customerRet.JobStartDate.GetValue
  End If
  
  ' Get the value of the ICustomerRet.JobProjectedEndDate element.
  If (Not customerRet.JobProjectedEndDate Is Nothing) Then
    jobProjectedEndDate67 = customerRet.JobProjectedEndDate.GetValue
  End If
  
  ' Get the value of the ICustomerRet.JobEndDate element.
  If (Not customerRet.JobEndDate Is Nothing) Then
    jobEndDate68 = customerRet.JobEndDate.GetValue
  End If
  
  ' Get the value of the ICustomerRet.JobDesc element.
  If (Not customerRet.JobDesc Is Nothing) Then
    jobDesc69 = customerRet.JobDesc.GetValue
  End If
  
  ' Get the value of the ICustomerRet.JobTypeRef element.
  If (Not customerRet.JobTypeRef Is Nothing) Then
    ' Get the FullName value.
    fullName70 = customerRet.JobTypeRef.FullName.GetValue
  
    ' Get the ListID value.
    listID70 = customerRet.JobTypeRef.ListID.GetValue
  
  End If
  
  ' Get the value of the ICustomerRet.Notes element.
  If (Not customerRet.Notes Is Nothing) Then
    notes71 = customerRet.Notes.GetValue
  End If
  
  ' Get the value of the ICustomerRet.IsStatementWithParent element.
  If (Not customerRet.IsStatementWithParent Is Nothing) Then
    isStatementWithParent72 = customerRet.IsStatementWithParent.GetValue
  End If
  
  ' Get the value of the ICustomerRet.DeliveryMethod element.
  If (Not customerRet.DeliveryMethod Is Nothing) Then
    deliveryMethod73 = customerRet.DeliveryMethod.GetValue
  End If
  
  If (country = "US") Then
    ' Get the value of the ICustomerRet.PriceLevelRef element.
    If (Not customerRet.PriceLevelRef Is Nothing) Then
      ' Get the FullName value.
      fullName74 = customerRet.PriceLevelRef.FullName.GetValue
  
      ' Get the ListID value.
      listID74 = customerRet.PriceLevelRef.ListID.GetValue
  
    End If
  
  End If
  ' Get the value of the ICustomerRet.DataExtRetList element.
  If (Not customerRet.DataExtRetList Is Nothing) Then
    For j = 0 To customerRet.DataExtRetList.Count - 1
      Dim dataExtRet75 As IDataExtRet
      Set dataExtRet75 = customerRet.DataExtRetList.GetAt(j)
      ' Get the value of the IDataExtRet.OwnerID element.
      If (Not dataExtRet75.OwnerID Is Nothing) Then
        Dim ownerID76 As String
        ownerID76 = dataExtRet75.OwnerID.GetValue
      End If
  
      ' Get the value of the IDataExtRet.DataExtName element.
      Dim dataExtName77 As String
      dataExtName77 = dataExtRet75.DataExtName.GetValue
  
      ' Get the value of the IDataExtRet.DataExtType element.
      Dim dataExtType78 As ENDataExtType
      dataExtType78 = dataExtRet75.DataExtType.GetValue
  
      ' Get the value of the IDataExtRet.DataExtValue element.
      Dim dataExtValue79 As String
      dataExtValue79 = dataExtRet75.DataExtValue.GetValue
  
    Next j
  
  End If
  
  If Not (country = "US") Then
    ' Get the value of the ICustomerRet.CurrencyRef element.
    If (Not customerRet.CurrencyRef Is Nothing) Then
      ' Get the FullName value.
      Dim fullName80 As String
      fullName80 = customerRet.CurrencyRef.FullName.GetValue
  
      ' Get the ListID value.
      Dim listID80 As String
      listID80 = customerRet.CurrencyRef.ListID.GetValue
  
    End If
  
  End If
  If (country = "UK") Then
    ' Get the value of the ICustomerRet.BusinessNumber element.
    If (Not customerRet.BusinessNumber Is Nothing) Then
      Dim businessNumber81 As String
      businessNumber81 = customerRet.BusinessNumber.GetValue
    End If
  
  End If
  If Not (country = "US") Then
    ' Get the value of the ICustomerRet.IsUsingCustomerTaxCode element.
    If (Not customerRet.IsUsingCustomerTaxCode Is Nothing) Then
      Dim isUsingCustomerTaxCode82 As Boolean
      isUsingCustomerTaxCode82 = customerRet.IsUsingCustomerTaxCode.GetValue
    End If
  
  End If

    ' update the Job tables
  
  
    ' filters applied???
'    If frmJCGetQBData.chkAllData = 0 Then
'
'        ' filter by status?
'        frmJCGetQBData.rs.Find "JobStatus = " & JobStatus65, 0, adSearchForward, 1
'        If frmJCGetQBData.rs!Select = False Then Exit Sub
'
'        ' filter by date
'        If Int(timeModified3) < frmJCGetQBData.tdbStartDate.Value Then Exit Sub
'        If Int(timeModified3) > frmJCGetQBData.tdbEndDate.Value Then Exit Sub
'
'
'    End If

    ' parent ID is not assigned - is a customer record
    If IsNull(listID8) Or listID8 = "" Then
        
        ' does the customer record already exist?
        If JCCustomer.GetByQBID(listID1) = False Then
            JCCustomer.Clear
            JCCustomer.QBID = listID1
            JCCustomer.Save (Equate.RecAdd)
        End If
 
        JCCustomer.Name = name5
        JCCustomer.FullName = fullName6
        JCCustomer.CompanyName = companyName10
        JCCustomer.FirstName = firstName12
        JCCustomer.LastName = lastName14
        JCCustomer.MidInit = middleName13
        
        JCCustomer.BillAddr1 = addr116
        JCCustomer.BillAddr2 = addr217
        JCCustomer.BillAddr3 = addr318
        JCCustomer.BillAddr4 = addr419
        JCCustomer.BillCity = city20
        JCCustomer.BillState = state21
        JCCustomer.BillZip = postalCode24
        
        JCCustomer.ShipAddr1 = addr126
        JCCustomer.ShipAddr2 = addr227
        JCCustomer.ShipAddr3 = addr328
        JCCustomer.ShipAddr4 = addr429
        JCCustomer.ShipCity = city30
        JCCustomer.ShipState = state31
        JCCustomer.ShipZip = postalCode34
        
        JCCustomer.QBTaxCode = listID53
        JCCustomer.QBTaxItem = listID54
        
        JCCustomer.Save (Equate.RecPut)
    
        ' add a ORIG job record for the customer
        SQLString = "SELECT * FROM JCJob WHERE " & _
                    "QBParentID = '" & Trim(JCCustomer.QBID) & "' " & _
                    "AND QBID = 'ORIG'"
        If JCJob.GetBySQL(SQLString) = False Then
            JCJob.Clear
            JCJob.QBParentID = JCCustomer.QBID
            JCJob.QBID = "ORIG"
            JCJob.Save (Equate.RecAdd)
        End If
        
        JCJob.Name = JCCustomer.Name
        JCJob.FullName = JCCustomer.FullName
        JCJob.CompanyName = JCCustomer.CompanyName
        JCJob.FirstName = JCCustomer.FirstName
        JCJob.LastName = JCCustomer.LastName
        JCJob.MidInit = JCCustomer.MidInit
                                                                    
        JCJob.BillAddr1 = JCCustomer.BillAddr1
        JCJob.BillAddr2 = JCCustomer.BillAddr2
        JCJob.BillAddr3 = JCCustomer.BillAddr3
        JCJob.BillAddr4 = JCCustomer.BillAddr4
        JCJob.BillCity = JCCustomer.BillCity
        JCJob.BillState = JCCustomer.BillState
        JCJob.BillZip = JCCustomer.BillZip
                                                                    
        JCJob.ShipAddr1 = JCCustomer.ShipAddr1
        JCJob.ShipAddr2 = JCCustomer.ShipAddr2
        JCJob.ShipAddr3 = JCCustomer.ShipAddr3
        JCJob.ShipAddr4 = JCCustomer.ShipAddr4
        JCJob.ShipCity = JCCustomer.ShipCity
        JCJob.ShipState = JCCustomer.ShipState
        JCJob.ShipZip = JCCustomer.ShipZip
        JCJob.Terms = listID47
        
        JCJob.JobStatus = CByte(JobStatus65)
        
        JCJob.ParentID = JCCustomer.CustomerID
        JCJob.StartDate = timeModified3
        JCJob.QBTaxCode = listID53
        
        If isActive7 = True Then
            JCJob.Active = 1
        Else
            JCJob.Active = 0
        End If
        
        JCJob.Save (Equate.RecPut)
    
    Else        ' parent ID filled in - is a job of existing customer
    
        If JCJob.GetByQBID(listID1) = False Then
            JCJob.Clear
            JCJob.QBID = listID1
            JCJob.QBParentID = listID8
            JCJob.Save (Equate.RecAdd)
        End If
        
        JCJob.Name = name5
        JCJob.FullName = fullName6
        JCJob.CompanyName = companyName10
        JCJob.FirstName = firstName12
        JCJob.LastName = lastName14
        JCJob.MidInit = middleName13
        
        JCJob.BillAddr1 = addr116
        JCJob.BillAddr2 = addr217
        JCJob.BillAddr3 = addr318
        JCJob.BillAddr4 = addr419
        JCJob.BillCity = city20
        JCJob.BillState = state21
        JCJob.BillZip = postalCode24
        
        JCJob.ShipAddr1 = addr126
        JCJob.ShipAddr2 = addr227
        JCJob.ShipAddr3 = addr328
        JCJob.ShipAddr4 = addr429
        JCJob.ShipCity = city30
        JCJob.ShipState = state31
        JCJob.ShipZip = postalCode34
        
        JCJob.JobStatus = CByte(JobStatus65)
        JCJob.StartDate = timeModified3
        
        JCJob.Terms = listID47
        
        If JCCustomer.GetByQBID(JCJob.QBParentID) Then
            JCJob.ParentID = JCCustomer.CustomerID
        Else
            ' not found - is a multi level job beneath the customer
            rsQB.AddNew
            rsQB!JobID = JCJob.JobID
            rsQB.Update
        End If
        
        JCJob.QBTaxCode = listID53
        
        If isActive7 = True Then
            JCJob.Active = 1
        Else
            JCJob.Active = 0
        End If
        
        JCJob.Save (Equate.RecPut)
    
    End If

    ' *** try ShipTo first ***
    ' auto update City if not assigned
    'AssignCity JCJob.ShipCity, JCJob.ShipState
    'AssignCity JCJob.BillCity, JCJob.BillState
    
End Sub
 
Private Sub JobFill()

Dim QBID As String
Dim PID As Long
Dim boo As Boolean

    ' fill in JCJob.ParentID for jobs more than one level deep
    rsQB.MoveFirst
    Do
        
        If JCJob.GetByID(rsQB!JobID) = False Then
            MsgBox "JobID not found: " & rsQB!JobID, vbExclamation
            GoBack
        End If
                
        PID = 0
                
        ' loop up the ladder
        Do
            ' if not found - at first level below customer
            '  use this JCJob.ParentID
            If JCJob.GetByQBID(JCJob.QBParentID) = False Then
                PID = JCJob.ParentID
                Exit Do
            End If
        Loop
            
        If PID = 0 Then
            MsgBox "ParentQBID not found: " & rsQB!JobID, vbExclamation
            GoBack
        End If
            
        ' reget the original job record
        boo = JCJob.GetByID(rsQB!JobID)
        JCJob.ParentID = PID
        JCJob.Save (Equate.RecPut)
        
        rsQB.MoveNext
    
    Loop Until rsQB.EOF

End Sub


