VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
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
   ScaleHeight     =   7275
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   6480
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9975
      _cx             =   17595
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
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qbXMLRP As New QBXMLRP2Lib.RequestProcessor2

Dim SessMgr As New QBFC5Lib.QBSessionManager

Dim requestMsgSet As IMsgSetRequest
Dim responseMsgSet As IMsgSetResponse
Dim responseList As IResponseList
Dim response As IResponse
Dim responseType As Integer
Dim orItemRetList As IORItemRetList

Dim invoiceAdd As IInvoiceAdd
Dim orInvoiceLineAdd1 As IORInvoiceLineAdd
Dim orInvoiceLineAddORElement2 As String
Dim orRateORElement3 As String
Dim orRatePriceLevelORElement4 As String
Dim dataExt5 As IDataExt
Dim dataExt6 As IDataExt
Dim orDiscountLineAddORElement7 As String
Dim orSalesTaxLineAddORElement8 As String
Dim invoiceRet As IInvoiceRet

Dim ClassQuery As IClassQuery
Dim classRetList As IClassRetList
Dim classRet As IClassRet
Dim rsClass As New ADODB.Recordset

Dim templateQuery As ITemplateQuery
Dim templateRetList As ITemplateRetList
Dim templateRet As ITemplateRet
Dim rsTpl As New ADODB.Recordset

Dim ItemQuery As IItemQuery
Dim orItemRet As IORItemRet
Dim rsItem As New ADODB.Recordset

Dim itemServiceAdd As IItemServiceAdd
Dim itemServiceRet As IItemServiceRet

Dim Ticket, QBFileName As String
Dim OC As Variant

Dim QBID, QBTaxCode, QBTaxItem As String

Dim I, j, K, M As Integer
Dim X, Y, Z As String

Private Sub Form_Load()

    Me.Show

    Me.lblMsg1 = "Open Connection"
    Me.Refresh
    SessMgr.OpenConnection2 "", "Balint Accounting", ctLocalQBD
    
    Me.lblMsg1 = "Begin Session"
    Me.Refresh
    SessMgr.BeginSession "", omDontCare
    
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue

    Me.KeyPreview = True

    GetItems
    
    SessMgr.EndSession
    SessMgr.BeginSession "", omDontCare
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    
    QB_AddItem
    
'    GetItems
'    rsItem.MoveFirst
''    SetGrid rsItem, fg
'
'    GetClass
'    rsClass.MoveFirst
''    SetGrid rsClass, fg
'
'    GetTemplates
'    rsTpl.MoveFirst
''    SetGrid rsTpl, fg
'
'    AddInvoice

    SessMgr.EndSession
    SessMgr.CloseConnection

    Me.lblMsg1 = "Ready ..."
    Me.Refresh

End Sub

Private Sub QB_AddItem()

    Set itemServiceAdd = requestMsgSet.AppendItemServiceAddRq
    
    itemServiceAdd.Name.SetValue "Test A1"
    itemServiceAdd.IsActive.SetValue True
    ' itemServiceAdd.ParentRef.ListID.SetValue ""
    ' itemServiceAdd.ParentRef.FullName.SetValue ""
    itemServiceAdd.SalesTaxCodeRef.FullName.SetValue "Tax"
    itemServiceAdd.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.SetValue "220000-934380913"
    
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    If responseMsgSet Is Nothing Then
        MsgBox "No response ..."
        Exit Sub
    End If
    
    Set responseList = responseMsgSet.responseList
    If responseList Is Nothing Then
        MsgBox "No Reponse List"
        Exit Sub
    End If
    
    Set response = responseList.GetAt(0)
    If response.StatusCode <> 0 Then
        MsgBox "Status Code: " & response.StatusCode, vbExclamation
        Exit Sub
    End If
        
    If (response.Detail Is Nothing) Then
        MsgBox "Response Detail is nothing", vbExclamation
        Exit Sub
    End If
    
    responseType = response.Type.GetValue
    If (responseType <> rtItemServiceAddRs) Then
        MsgBox "Invalid response type", vbExclamation
        Exit Sub
    End If
    
    Set itemServiceRet = response.Detail
    
    If itemServiceRet Is Nothing Then
        MsgBox "itemServiceRet is nothing", vbExclamation
        Exit Sub
    End If
    
    MsgBox itemServiceRet.ListID.GetValue
        
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
    
End Sub


Private Sub AddInvoice()

    Me.lblMsg1 = "Add Invoice ...."
    Me.Refresh

    ' cust/job QB ListID
    QBID = "6E0001-1197764528"
    
    ' Adams - taxable
    QBID = "6E0001-1197764528"
    QBTaxCode = "10000-999021789"
    QBTaxItem = "310000-1197753788"
    
    ' Chapman - non taxable
    QBID = "520000-1071508459"
    QBTaxCode = "20000-999021789"
    QBTaxItem = "310000-1197753788"
    
    Set invoiceAdd = requestMsgSet.AppendInvoiceAddRq
    
    invoiceAdd.CustomerRef.ListID.SetValue QBID
    
    ' invoiceAdd.CustomerRef.FullName.SetValue "ab"
    ' invoiceAdd.ClassRef.FullName.SetValue "ab"
    ' invoiceAdd.ClassRef.ListID.SetValue "ab"
    
    invoiceAdd.ARAccountRef.ListID.SetValue "40000-934380912"
    ' invoiceAdd.ARAccountRef.FullName.SetValue "ab"
    
    invoiceAdd.TemplateRef.ListID.SetValue rsTpl!ListID
    ' invoiceAdd.TemplateRef.FullName.SetValue "ab"
    
    invoiceAdd.TxnDate.SetValue #3/13/2010#
    
    ' QB auto numbers it .... SWEET !!!
    ' invoiceAdd.RefNumber.SetValue "1001"
    
    ' invoiceAdd.BillAddress.Addr1.SetValue "val"
    ' invoiceAdd.BillAddress.Addr2.SetValue "val"
    ' invoiceAdd.BillAddress.Addr3.SetValue "val"
    ' invoiceAdd.BillAddress.Addr4.SetValue "val"
    ' invoiceAdd.BillAddress.City.SetValue "val"
    ' invoiceAdd.BillAddress.State.SetValue "val"
    ' invoiceAdd.BillAddress.PostalCode.SetValue "val"
    ' invoiceAdd.BillAddress.Country.SetValue "val"
    
    ' invoiceAdd.ShipAddress.Addr1.SetValue "val"
    ' invoiceAdd.ShipAddress.Addr2.SetValue "val"
    ' invoiceAdd.ShipAddress.Addr3.SetValue "val"
    ' invoiceAdd.ShipAddress.Addr4.SetValue "val"
    ' invoiceAdd.ShipAddress.City.SetValue "val"
    ' invoiceAdd.ShipAddress.State.SetValue "val"
    ' invoiceAdd.ShipAddress.PostalCode.SetValue "val"
    ' invoiceAdd.ShipAddress.Country.SetValue "val"
    
    invoiceAdd.IsPending.SetValue False
    
    ' invoiceAdd.PONumber.SetValue "val"
    ' invoiceAdd.TermsRef.FullName.SetValue "ab"
    ' invoiceAdd.TermsRef.ListID.SetValue "ab"
    
    ' >>>>>>>>>>>>>> due date comes from customer terms
    ' invoiceAdd.DueDate.SetValue #4/1/2010#
    
    ' invoiceAdd.SalesRepRef.FullName.SetValue "ab"
    ' invoiceAdd.SalesRepRef.ListID.SetValue "ab"
    ' invoiceAdd.FOB.SetValue "val"
    ' invoiceAdd.ShipDate.SetValue #12/31/2003#
    ' invoiceAdd.ShipMethodRef.FullName.SetValue "ab"
    ' invoiceAdd.ShipMethodRef.ListID.SetValue "ab"
    
    ' ??? by customer???
    ' invoiceAdd.ItemSalesTaxRef.FullName.SetValue "10000-999021789"
    invoiceAdd.ItemSalesTaxRef.ListID.SetValue QBTaxItem
    
    ' invoiceAdd.CustomerSalesTaxCodeRef.FullName.SetValue "ab"
    invoiceAdd.CustomerSalesTaxCodeRef.ListID.SetValue QBTaxCode
    
    ' invoiceAdd.Memo.SetValue "val"
    ' invoiceAdd.CustomerMsgRef.FullName.SetValue "ab"
    ' invoiceAdd.CustomerMsgRef.ListID.SetValue "ab"
    
    invoiceAdd.IsToBePrinted.SetValue True
    
    
    For j = 1 To 4
                    
        ' Append an element to the list and save the element in orInvoiceLineAdd1 so we can set its values.
        Set orInvoiceLineAdd1 = invoiceAdd.ORInvoiceLineAddList.Append
  
        orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.ListID.SetValue rsItem!ListID
        ' orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue "ab"
        
        orInvoiceLineAdd1.InvoiceLineAdd.Desc.SetValue "Desc: " & j
        orInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue j + 10
      
        orRatePriceLevelORElement4 = "Rate"
        If (orRatePriceLevelORElement4 = "Rate") Then
            ' Set the value of the IORRatePriceLevel.Rate element.
            orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue j * 10
        ElseIf (orRatePriceLevelORElement4 = "RatePercent") Then
            ' Set the value of the IORRatePriceLevel.RatePercent element.
            orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.RatePercent.SetValue 2#
        ElseIf (orRatePriceLevelORElement4 = "PriceLevelRef") Then
            ' Set the FullName value.
            orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.PriceLevelRef.FullName.SetValue "ab"
  
            ' Set the ListID value.
            orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.PriceLevelRef.ListID.SetValue "ab"
        End If
  
        orInvoiceLineAdd1.InvoiceLineAdd.ClassRef.ListID.SetValue rsClass!ListID
        ' orInvoiceLineAdd1.InvoiceLineAdd.ClassRef.FullName.SetValue "ab"
        
        orInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue (j + 10) * (j * 10)
        
        orInvoiceLineAdd1.InvoiceLineAdd.ServiceDate.SetValue #3/1/2010#
  
        ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        ' *********** SALES TAX ********************
        orInvoiceLineAdd1.InvoiceLineAdd.SalesTaxCodeRef.ListID.SetValue QBTaxCode
        ' orInvoiceLineAdd1.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue "ab"
        
        orInvoiceLineAdd1.InvoiceLineAdd.IsTaxable.SetValue True
        
        ' %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
        ' orInvoiceLineAdd1.InvoiceLineAdd.OverrideItemAccountRef.FullName.SetValue "ab"
        ' orInvoiceLineAdd1.InvoiceLineAdd.OverrideItemAccountRef.ListID.SetValue "ab"
  
'        'Add multiple elements to the list. In this case we will add 5 elements.
'        For K = 0 To 4
'            ' Append an element to the list and save the element in dataExt5 so we can set its values.
'            Set dataExt5 = orInvoiceLineAdd1.InvoiceLineAdd.DataExtList.Append
'
'            ' Set the value of the IDataExt.OwnerID element.
'            dataExt5.OwnerID.SetValue "{22E8C9DC-320B-450d-962A-87CF7246D080}"
'
'            ' Set the value of the IDataExt.DataExtName element.
'            dataExt5.DataExtName.SetValue "val"
'
'            ' Set the value of the IDataExt.DataExtValue element.
'            dataExt5.DataExtValue.SetValue "val"
'
'        Next K
         
'        ' Set the value of the IInvoiceLineAdd.defMacro element.
'        orInvoiceLineAdd1.InvoiceLineAdd.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
  
    Next j
  
'    ' Only can set one of the OR elements.
'    ' We will portray this restriction by using an If/Then/Else.
'    orDiscountLineAddORElement7 = "Amount"
'    If (orDiscountLineAddORElement7 = "Amount") Then
'        ' Set the value of the IORDiscountLineAdd.Amount element.
'        invoiceAdd.DiscountLineAdd.ORDiscountLineAdd.Amount.SetValue 2#
'
'    ElseIf (orDiscountLineAddORElement7 = "RatePercent") Then
'        ' Set the value of the IORDiscountLineAdd.RatePercent element.
'        invoiceAdd.DiscountLineAdd.ORDiscountLineAdd.RatePercent.SetValue 2#
'
'    End If
'
'    ' Set the value of the IDiscountLineAdd.IsTaxable element.
'    invoiceAdd.DiscountLineAdd.IsTaxable.SetValue True
'
'    ' Set the FullName value.
'    invoiceAdd.DiscountLineAdd.AccountRef.FullName.SetValue "ab"
'
'    ' Set the ListID value.
'    invoiceAdd.DiscountLineAdd.AccountRef.ListID.SetValue "ab"
  
'    ' Only can set one of the OR elements.
'    ' We will portray this restriction by using an If/Then/Else.
'    orSalesTaxLineAddORElement8 = "Amount"
'    If (orSalesTaxLineAddORElement8 = "Amount") Then
'        ' Set the value of the IORSalesTaxLineAdd.Amount element.
'        invoiceAdd.SalesTaxLineAdd.ORSalesTaxLineAdd.Amount.SetValue 2#
'
'    ElseIf (orSalesTaxLineAddORElement8 = "RatePercent") Then
'        ' Set the value of the IORSalesTaxLineAdd.RatePercent element.
'        invoiceAdd.SalesTaxLineAdd.ORSalesTaxLineAdd.RatePercent.SetValue 2#
'
'    End If
          
'    ' Set the FullName value.
'    invoiceAdd.SalesTaxLineAdd.AccountRef.FullName.SetValue "ab"
'
'    ' Set the ListID value.
'    invoiceAdd.SalesTaxLineAdd.AccountRef.ListID.SetValue "ab"
'
'    ' Set the value of the IShippingLineAdd.Amount element.
'    invoiceAdd.ShippingLineAdd.Amount.SetValue 2#
'
'    ' Set the FullName value.
'    invoiceAdd.ShippingLineAdd.AccountRef.FullName.SetValue "ab"
'
'    ' Set the ListID value.
'    invoiceAdd.ShippingLineAdd.AccountRef.ListID.SetValue "ab"
'
'    ' Set the value of the IInvoiceAdd.IncludeRetElementList element.
'    invoiceAdd.IncludeRetElementList.Add "val"
'
'    ' Set the value of the IInvoiceAdd.defMacro element.
'    invoiceAdd.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
        
            
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    Set responseList = responseMsgSet.responseList
  
    If (responseList Is Nothing) Then
        MsgBox "ResponseList is nothing .... Ha Ha .... ", vbExclamation
        End
    End If
    
    For I = 0 To responseList.Count - 1
        
        Set response = responseList.GetAt(I)
  
        ' Check the status returned for the response.
        If (response.StatusCode <> 0) Then GoTo InvParseNxtI
        If (response.Detail Is Nothing) Then GoTo InvParseNxtI
        responseType = response.Type.GetValue
        If responseType <> rtInvoiceAddRs Then GoTo InvParseNxtI
        Set invoiceRet = response.Detail
        If invoiceRet Is Nothing Then
            MsgBox "invoiceRet is nothing ... Boo Hoo ...", vbExclamation
            End
        End If
        
        MsgBox "Txn ID: " & invoiceRet.TxnID & vbCr & invoiceRet.TxnNumber
  
InvParseNxtI:
    Next I

End Sub

Private Sub GetClass()
    
    On Error Resume Next
    rsClass.Close
    On Error GoTo 0
    rsClass.CursorLocation = adUseClient
    rsClass.Fields.Append "ListID", adVarChar, 255, adFldIsNullable
    rsClass.Fields.Append "Name", adVarChar, 31, adFldIsNullable
    rsClass.Fields.Append "FullName", adVarChar, 159, adFldIsNullable
    rsClass.Fields.Append "Active", adBoolean
    rsClass.Fields.Append "SubLevel", adDouble
    rsClass.Fields.Append "ParentRefFullName", adVarChar, 159, adFldIsNullable
    rsClass.Fields.Append "ParentRefListID", adVarChar, 255, adFldIsNullable
    rsClass.Fields.Append "ParentRefType", adVarChar, 255, adFldIsNullable
    rsClass.Open , , adOpenDynamic, adLockOptimistic
    
    Me.lblMsg1 = "Start Class Query"
    Me.Refresh
    Set ClassQuery = requestMsgSet.AppendClassQueryRq
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    
    If responseMsgSet Is Nothing Then
        MsgBox "No Class items found"
        Exit Sub
    End If
    
    Me.lblMsg1 = "Parse Class Query"
    Me.Refresh
    Set responseList = responseMsgSet.responseList
        
    If responseList Is Nothing Then
        MsgBox "No Class items found"
        Exit Sub
    End If
        
    For I = 0 To responseList.Count - 1
        
        Set response = responseList.GetAt(I)
  
        ' Check the status returned for the response.
        If response.StatusCode <> 0 Then GoTo clsNxtI
  
        ' Check to make sure the response is of the type we are expecting.
        If response.Detail Is Nothing Then GoTo clsNxtI
        
        ' Check for ClassQueryrsclass.
        responseType = response.Type.GetValue
        If responseType <> rtClassQueryRs Then GoTo clsNxtI
          
        Set classRetList = response.Detail
        K = classRetList.Count - 1
        For j = 0 To K
            
            Me.lblMsg1 = "Parse " & j & " of: " & K
            Me.Refresh
            
            Set classRet = classRetList.GetAt(j)
            If (Not classRet Is Nothing) Then
                rsClass.AddNew
                rsClass!ListID = classRet.ListID.GetValue
                rsClass!Name = classRet.Name.GetValue
                rsClass!FullName = classRet.FullName.GetValue
                rsClass!Active = classRet.IsActive.GetValue
                If Not classRet.ParentRef Is Nothing Then
                    rsClass!ParentRefFullName = classRet.ParentRef.FullName.GetValue
                    rsClass!ParentRefListID = classRet.ParentRef.ListID.GetValue
                    rsClass!ParentRefType = classRet.ParentRef.Type.GetAsString
                End If
                rsClass!Sublevel = classRet.Sublevel.GetValue
                rsClass.Update
            End If
        
        Next j
  
clsNxtI:
    Next I

    ' SetGrid rsClass, fg

End Sub

Private Sub GetItems()
    
    rsItem.CursorLocation = adUseClient
    rsItem.Fields.Append "ListID", adVarChar, 255, adFldIsNullable
    rsItem.Fields.Append "Name", adVarChar, 31, adFldIsNullable
    rsItem.Fields.Append "FullName", adVarChar, 159, adFldIsNullable
    rsItem.Fields.Append "Description", adVarChar, 255, adFldIsNullable
    rsItem.Fields.Append "Price", adCurrency
    rsItem.Open , , adOpenDynamic, adLockOptimistic
    
    Me.lblMsg1 = "Start Item Query"
    Me.Refresh
    
    Set ItemQuery = requestMsgSet.AppendItemQueryRq
    
    ' filters
    ItemQuery.ORListQuery.ListFilter.ActiveStatus.SetValue asActiveOnly
    
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)

    If responseMsgSet Is Nothing Then
        MsgBox "No Items found"
        Exit Sub
    End If
    
    Me.lblMsg1 = "Parse Item Query"
    Me.Refresh

    Set responseList = responseMsgSet.responseList
    For I = 0 To responseList.Count - 1
    
        Set response = responseList.GetAt(I)
        If response.StatusCode <> 0 Then GoTo itemNxtI
        If response.Detail Is Nothing Then GoTo itemNxtI
        responseType = response.Type.GetValue
        If responseType <> rtItemQueryRs Then GoTo itemNxtI
        
        Set orItemRetList = response.Detail
        K = orItemRetList.Count - 1
        For j = 0 To K
            
            Me.lblMsg1 = "Item: " & j & " of: " & K
            Me.Refresh
            
            Set orItemRet = orItemRetList.GetAt(j)
                        
            ' service items
            If (Not orItemRet.itemServiceRet Is Nothing) Then
                If (Not orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase Is Nothing) Then
                    rsItem.AddNew
                    rsItem!ListID = orItemRet.itemServiceRet.ListID.GetValue
                    rsItem!Name = orItemRet.itemServiceRet.Name.GetValue
                    rsItem!FullName = orItemRet.itemServiceRet.FullName.GetValue
                    rsItem!Price = orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue
                    
                    ' 4095 max characters from QB
'                    If orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc.IsEmpty = False Then
'                        rsItem!Description = Mid(orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue, 1, 255)
'                    End If
                        
                    rsItem.Update
                End If
            End If
        
        Next j
                
itemNxtI:
    Next I

    ' SetGrid rsItem, fg

End Sub

Private Sub GetTemplates()

    rsTpl.CursorLocation = adUseClient
    rsTpl.Fields.Append "ListID", adVarChar, 255, adFldIsNullable
    rsTpl.Fields.Append "Name", adVarChar, 31, adFldIsNullable
    rsTpl.Fields.Append "Type", adVarChar, 255, adFldIsNullable
    rsTpl.Open , , adOpenDynamic, adLockOptimistic

    Set templateQuery = requestMsgSet.AppendTemplateQueryRq
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)

    If responseMsgSet Is Nothing Then
        MsgBox "No Templates found"
        Exit Sub
    End If
    
    Me.lblMsg1 = "Template Query"
    Me.Refresh
    
    Set responseList = responseMsgSet.responseList
        
    If responseList Is Nothing Then
        MsgBox "No Templates found"
        Exit Sub
    End If
        
    For I = 0 To responseList.Count - 1
        
        Set response = responseList.GetAt(I)
  
        ' Check the status returned for the response.
        If (response.StatusCode <> 0) Then GoTo tplNxtI
  
        ' Check to make sure the response is of the type we are expecting.
        If (response.Detail Is Nothing) Then GoTo tplNxtI
        responseType = response.Type.GetValue
        If (responseType <> rtTemplateQueryRs) Then GoTo tplNxtI
        
        Set templateRetList = response.Detail
        For j = 0 To templateRetList.Count - 1
            Set templateRet = templateRetList.GetAt(j)
            If templateRet Is Nothing Then GoTo tplNxtJ
            If templateRet.IsActive Is Nothing Then GoTo tplNxtJ
            If templateRet.IsActive.GetValue = False Then GoTo tplNxtJ
            
            ' active - OK to add
            rsTpl.AddNew
            rsTpl!ListID = templateRet.ListID.GetValue
            rsTpl!Name = templateRet.Name.GetValue
            rsTpl!Type = templateRet.Type.GetAsString
        
tplNxtJ:
        Next j
          
tplNxtI:
    Next I

    ' SetGrid rsTpl, fg

End Sub


Private Sub cmdExit_Click()
    End
End Sub




