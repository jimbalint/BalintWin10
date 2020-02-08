VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   6795
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMsg1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Dim requestMsgSet As IMsgSetRequest
Dim responseMsgSet As IMsgSetResponse
Dim billAdd As IBillAdd
Dim expenseLineAdd1 As IExpenseLineAdd

Dim ResponseList As IResponseList

Dim QBCount As Long
Dim D2 As Date
Dim j As Integer
    

Private Sub Form_Load()

    Me.Show
    
    If QBOpen(Me, Me.lblMsg1) = False Then End

    D2 = DateSerial(2010, 5, 31)

    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    BuildBillAddRq requestMsgSet, "US"
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)

    ParseBillAddRs responseMsgSet, "US"


    SessMgr.EndSession
    SessMgr.CloseConnection
    
    MsgBox "Ok ..."
    End

End Sub

Public Sub BuildBillAddRq(requestMsgSet As IMsgSetRequest, country As String)
 
    If (requestMsgSet Is Nothing) Then
        Exit Sub
    End If
 
    'Add the request to the message set request object.
    Set billAdd = requestMsgSet.AppendBillAddRq
 
    'Set the elements of IBillAdd.
 
    ' Set the FullName value.
    ' billAdd.VendorRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    billAdd.VendorRef.ListID.SetValue "80001-1273563975"
 
    ' Set the FullName value.
    ' billAdd.APAccountRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    ' billAdd.APAccountRef.ListID.SetValue "ab"
 
    ' Set the value of the IBillAdd.TxnDate element.
    billAdd.TxnDate.SetValue D2
 
    ' Set the value of the IBillAdd.DueDate element.
    billAdd.DueDate.SetValue D2
 
    ' Set the value of the IBillAdd.RefNumber element.
    ' 20 char max
    billAdd.RefNumber.SetValue "PR20100501"
 
    ' Set the FullName value.
    ' billAdd.TermsRef.FullName.SetValue "ab"
 
    ' Set the ListID value.
    ' billAdd.TermsRef.ListID.SetValue "ab"
 
    ' Set the value of the IBillAdd.Memo element.
    billAdd.Memo.SetValue "PR Memo"
 
'    If (country = "US") Then
'        ' Set the value of the IBillAdd.LinkToTxnIDList element.
'        billAdd.LinkToTxnIDList.Add "val"
'    End If
    
    'Add multiple elements to the list. In this case we will add 5 elements.
    For j = 0 To 1
        
        ' Append an element to the list and save the element in expenseLineAdd1 so we can set its values.
        Set expenseLineAdd1 = billAdd.ExpenseLineAddList.Append
 
        ' Set the FullName value.
        ' expenseLineAdd1.AccountRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        expenseLineAdd1.AccountRef.ListID.SetValue "3A0000-1270916891"
 
        ' Set the value of the IExpenseLineAdd.Amount element.
        expenseLineAdd1.Amount.SetValue 5001 + j
 
        ' Set the value of the IExpenseLineAdd.Memo element.
        expenseLineAdd1.Memo.SetValue "Expense Memo" & j
 
        ' Set the FullName value.
        ' expenseLineAdd1.CustomerRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        If j = 0 Then
            expenseLineAdd1.CustomerRef.ListID.SetValue "10000-1270922690"      ' cust A
        Else
            expenseLineAdd1.CustomerRef.ListID.SetValue "20000-1270923734"      ' cust B
        End If
        
        ' Set the FullName value.
        ' expenseLineAdd1.ClassRef.FullName.SetValue "ab"
 
        ' Set the ListID value.
        ' expenseLineAdd1.ClassRef.ListID.SetValue "ab"
 
        ' Set the value of the IExpenseLineAdd.BillableStatus element.
        ' expenseLineAdd1.BillableStatus.SetValue bsBillable
 
        If Not (country = "US") Then
            ' Set the FullName value.
            expenseLineAdd1.TaxCodeRef.FullName.SetValue "ab"
 
            ' Set the ListID value.
            expenseLineAdd1.TaxCodeRef.ListID.SetValue "ab"
 
        End If
        
        ' Set the value of the IExpenseLineAdd.defMacro element.
        QBCount = QBCount + 1
        expenseLineAdd1.defMacro.SetValue "AA:" & QBCount
 
    Next j
 
'    'Add multiple elements to the list. In this case we will add 5 elements.
'    Dim orItemLineAdd2 As IORItemLineAdd
'    Dim k As Integer
'    For k = 0 To 4
'
'        ' Append an element to the list and save the element in orItemLineAdd2 so we can set its values.
'        Set orItemLineAdd2 = billAdd.ORItemLineAddList.Append
'
'        ' Only can set one of the OR elements.
'        ' We will portray this restriction by using an If/Then/Else.
'        Dim orItemLineAddORElement3 As String
'        orItemLineAddORElement3 = "ItemLineAdd"
'        If (orItemLineAddORElement3 = "ItemLineAdd") Then
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.ItemRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.ItemRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemLineAdd.Desc element.
'            orItemLineAdd2.ItemLineAdd.Desc.SetValue "val"
'
'            ' Set the value of the IItemLineAdd.Quantity element.
'            orItemLineAdd2.ItemLineAdd.Quantity.SetValue 2#
'
'            ' Set the value of the IItemLineAdd.Cost element.
'            orItemLineAdd2.ItemLineAdd.Cost.SetValue 2#
'
'            ' Set the value of the IItemLineAdd.Amount element.
'            orItemLineAdd2.ItemLineAdd.Amount.SetValue 2#
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.CustomerRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.CustomerRef.ListID.SetValue "ab"
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.ClassRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.ClassRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemLineAdd.BillableStatus element.
'            orItemLineAdd2.ItemLineAdd.BillableStatus.SetValue bsBillable
'
'            ' Set the FullName value.
'            orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.ListID.SetValue "ab"
'
'            If Not (country = "US") Then
'                ' Set the FullName value.
'                orItemLineAdd2.ItemLineAdd.TaxCodeRef.FullName.SetValue "ab"
'
'                ' Set the ListID value.
'                orItemLineAdd2.ItemLineAdd.TaxCodeRef.ListID.SetValue "ab"
'
'            End If
'            If (country = "US") Then
'                ' Set the value of the ILinkToTxn.TxnID element.
'                orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnID.SetValue "val"
'
'                ' Set the value of the ILinkToTxn.TxnLineID element.
'                orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnLineID.SetValue "val"
'
'            End If
'        ElseIf (orItemLineAddORElement3 = "ItemGroupLineAdd") Then
'            ' Set the FullName value.
'            orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.FullName.SetValue "ab"
'
'            ' Set the ListID value.
'            orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.ListID.SetValue "ab"
'
'            ' Set the value of the IItemGroupLineAdd.Desc element.
'            orItemLineAdd2.ItemGroupLineAdd.Desc.SetValue "val"
'
'            ' Set the value of the IItemGroupLineAdd.Quantity element.
'            orItemLineAdd2.ItemGroupLineAdd.Quantity.SetValue 2#
'
'        End If
'
'    Next k
'
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.Tax1Total element.
'        billAdd.Tax1Total.SetValue 2#
'
'    End If
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.Tax2Total element.
'        billAdd.Tax2Total.SetValue 2#
'
'    End If
'    If Not (country = "US") Then
'        ' Set the value of the IBillAdd.ExchangeRate element.
'        billAdd.ExchangeRate.SetValue 2.5
'
'    End If
'    If (country = "UK") Then
'        ' Set the value of the IBillAdd.AmountIncludesVAT element.
'        billAdd.AmountIncludesVAT.SetValue True
'
'    End If
'    If (country = "US") Then
'        ' Set the value of the IBillAdd.IncludeRetElementList element.
'        billAdd.IncludeRetElementList.Add "val"
'
'    End If
    
    ' Set the value of the IBillAdd.defMacro element.
    QBCount = QBCount + 1
    billAdd.defMacro.SetValue "BB:" & QBCount
 
End Sub

Public Sub ParseBillAddRs(responseMsgSet As IMsgSetResponse, country As String)
 
    If (responseMsgSet Is Nothing) Then
        Exit Sub
    End If
 
    Set ResponseList = responseMsgSet.ResponseList
    If (ResponseList Is Nothing) Then
        Exit Sub
    End If
 
    ' Go through all of the responses in the list.
    Dim i As Integer
    For i = 0 To ResponseList.Count - 1
        Dim Response As IResponse
        Set Response = ResponseList.GetAt(i)
 
 MsgBox Response.StatusCode & vbCr & Response.StatusMessage & vbCr & Response.StatusSeverity
 
        ' Check the status returned for the response.
        If (Response.StatusCode = 0) Then
 
            ' Check to make sure the response is of the type we are expecting.
            If (Not Response.Detail Is Nothing) Then
                Dim ResponseType As Integer
                ResponseType = Response.Type.GetValue
                Dim j As Integer
                ' Check for BillAddRs.
                If (ResponseType = rtBillAddRs) Then
                    Dim billRet As IBillRet
                    Set billRet = Response.Detail
                    ParseBillRet billRet, country
                End If
            End If
        End If
    Next i
End Sub

Private Sub ParseBillRet(billRet As IBillRet, country As String)
 
    If (billRet Is Nothing) Then
        Exit Sub
    End If
 
    'Go through all of the elements of IBillRet.
 
    ' Get the value of the IBillRet.TxnID element.
    Dim txnID1 As String
    txnID1 = billRet.TxnID.GetValue
 
    ' Get the value of the IBillRet.TimeCreated element.
    Dim timeCreated2 As Date
    timeCreated2 = billRet.TimeCreated.GetValue
 
    ' Get the value of the IBillRet.TimeModified element.
    Dim timeModified3 As Date
    timeModified3 = billRet.TimeModified.GetValue
 
    ' Get the value of the IBillRet.EditSequence element.
    Dim editSequence4 As String
    editSequence4 = billRet.EditSequence.GetValue
 
    ' Get the value of the IBillRet.TxnNumber element.
    If (Not billRet.TxnNumber Is Nothing) Then
        Dim txnNumber5 As Long
        txnNumber5 = billRet.TxnNumber.GetValue
    End If
 
    ' Get the value of the IBillRet.VendorRef element.
    ' Get the FullName value.
    Dim fullName6 As String
    fullName6 = billRet.VendorRef.FullName.GetValue
 
    ' Get the ListID value.
    Dim listID6 As String
    listID6 = billRet.VendorRef.ListID.GetValue
 
    ' Get the value of the IBillRet.APAccountRef element.
    If (Not billRet.APAccountRef Is Nothing) Then
        ' Get the FullName value.
        Dim fullName7 As String
        fullName7 = billRet.APAccountRef.FullName.GetValue
 
        ' Get the ListID value.
        Dim listID7 As String
        listID7 = billRet.APAccountRef.ListID.GetValue
 
    End If
 
    ' Get the value of the IBillRet.TxnDate element.
    Dim txnDate8 As Date
    txnDate8 = billRet.TxnDate.GetValue
 
    ' Get the value of the IBillRet.DueDate element.
    If (Not billRet.DueDate Is Nothing) Then
        Dim dueDate9 As Date
        dueDate9 = billRet.DueDate.GetValue
    End If
 
    ' Get the value of the IBillRet.AmountDue element.
    Dim amountDue10 As Double
    amountDue10 = billRet.AmountDue.GetValue
 
    ' Get the value of the IBillRet.RefNumber element.
    If (Not billRet.RefNumber Is Nothing) Then
        Dim refNumber11 As String
        refNumber11 = billRet.RefNumber.GetValue
    End If
 
    ' Get the value of the IBillRet.TermsRef element.
    If (Not billRet.TermsRef Is Nothing) Then
        ' Get the FullName value.
        Dim fullName12 As String
        fullName12 = billRet.TermsRef.FullName.GetValue
 
        ' Get the ListID value.
        Dim listID12 As String
        listID12 = billRet.TermsRef.ListID.GetValue
 
    End If
 
    ' Get the value of the IBillRet.Memo element.
    If (Not billRet.Memo Is Nothing) Then
        Dim memo13 As String
        memo13 = billRet.Memo.GetValue
    End If
 
    ' Get the value of the IBillRet.IsPaid element.
    If (Not billRet.IsPaid Is Nothing) Then
        Dim isPaid14 As Boolean
        isPaid14 = billRet.IsPaid.GetValue
    End If
 
    ' Get the value of the IBillRet.LinkedTxnList element.
    If (Not billRet.LinkedTxnList Is Nothing) Then
        Dim j As Integer
        For j = 0 To billRet.LinkedTxnList.Count - 1
            Dim linkedTxn15 As ILinkedTxn
            Set linkedTxn15 = billRet.LinkedTxnList.GetAt(j)
            ' Get the value of the ILinkedTxn.TxnID element.
            Dim txnID16 As String
            txnID16 = linkedTxn15.TxnID.GetValue
 
            ' Get the value of the ILinkedTxn.TxnType element.
            Dim txnType17 As ENTxnType
            txnType17 = linkedTxn15.TxnType.GetValue
 
            ' Get the value of the ILinkedTxn.TxnDate element.
            Dim txnDate18 As Date
            txnDate18 = linkedTxn15.TxnDate.GetValue
 
            ' Get the value of the ILinkedTxn.RefNumber element.
            If (Not linkedTxn15.RefNumber Is Nothing) Then
                Dim refNumber19 As String
                refNumber19 = linkedTxn15.RefNumber.GetValue
            End If
 
            ' Get the value of the ILinkedTxn.LinkType element.
            If (Not linkedTxn15.LinkType Is Nothing) Then
                Dim linkType20 As ENLinkType
                linkType20 = linkedTxn15.LinkType.GetValue
            End If
 
            ' Get the value of the ILinkedTxn.Amount element.
            Dim amount21 As Double
            amount21 = linkedTxn15.Amount.GetValue
 
            ' Get the value of the ILinkedTxn.TxnLineDetailList element.
            If (Not linkedTxn15.TxnLineDetailList Is Nothing) Then
                Dim k As Integer
                For k = 0 To linkedTxn15.TxnLineDetailList.Count - 1
                    Dim txnLineDetail22 As ITxnLineDetail
                    Set txnLineDetail22 = linkedTxn15.TxnLineDetailList.GetAt(k)
                    ' Get the value of the ITxnLineDetail.TxnLineID element.
                    Dim txnLineID23 As String
                    txnLineID23 = txnLineDetail22.TxnLineID.GetValue
 
                    ' Get the value of the ITxnLineDetail.Amount element.
                    Dim amount24 As Double
                    amount24 = txnLineDetail22.Amount.GetValue
 
                Next k
 
            End If
 
        Next j
 
    End If
 
    ' Get the value of the IBillRet.ExpenseLineRetList element.
    If (Not billRet.ExpenseLineRetList Is Nothing) Then
        Dim m As Integer
        For m = 0 To billRet.ExpenseLineRetList.Count - 1
            Dim expenseLineRet25 As IExpenseLineRet
            Set expenseLineRet25 = billRet.ExpenseLineRetList.GetAt(m)
            ' Get the value of the IExpenseLineRet.TxnLineID element.
            Dim txnLineID26 As String
            txnLineID26 = expenseLineRet25.TxnLineID.GetValue
 
            ' Get the value of the IExpenseLineRet.AccountRef element.
            If (Not expenseLineRet25.AccountRef Is Nothing) Then
                ' Get the FullName value.
                Dim fullName27 As String
                fullName27 = expenseLineRet25.AccountRef.FullName.GetValue
 
                ' Get the ListID value.
                Dim listID27 As String
                listID27 = expenseLineRet25.AccountRef.ListID.GetValue
 
            End If
 
            ' Get the value of the IExpenseLineRet.Amount element.
            If (Not expenseLineRet25.Amount Is Nothing) Then
                Dim amount28 As Double
                amount28 = expenseLineRet25.Amount.GetValue
            End If
 
            ' Get the value of the IExpenseLineRet.Memo element.
            If (Not expenseLineRet25.Memo Is Nothing) Then
                Dim memo29 As String
                memo29 = expenseLineRet25.Memo.GetValue
            End If
 
            ' Get the value of the IExpenseLineRet.CustomerRef element.
            If (Not expenseLineRet25.CustomerRef Is Nothing) Then
                ' Get the FullName value.
                Dim fullName30 As String
                fullName30 = expenseLineRet25.CustomerRef.FullName.GetValue
 
                ' Get the ListID value.
                Dim listID30 As String
                listID30 = expenseLineRet25.CustomerRef.ListID.GetValue
 
            End If
 
            ' Get the value of the IExpenseLineRet.ClassRef element.
            If (Not expenseLineRet25.ClassRef Is Nothing) Then
                ' Get the FullName value.
                Dim fullName31 As String
                fullName31 = expenseLineRet25.ClassRef.FullName.GetValue
 
                ' Get the ListID value.
                Dim listID31 As String
                listID31 = expenseLineRet25.ClassRef.ListID.GetValue
 
            End If
 
            ' Get the value of the IExpenseLineRet.BillableStatus element.
            If (Not expenseLineRet25.BillableStatus Is Nothing) Then
                Dim billableStatus32 As ENBillableStatus
                billableStatus32 = expenseLineRet25.BillableStatus.GetValue
            End If
 
            If Not (country = "US") Then
                ' Get the value of the IExpenseLineRet.TaxCodeRef element.
                If (Not expenseLineRet25.TaxCodeRef Is Nothing) Then
                    ' Get the FullName value.
                    Dim fullName33 As String
                    fullName33 = expenseLineRet25.TaxCodeRef.FullName.GetValue
 
                    ' Get the ListID value.
                    Dim listID33 As String
                    listID33 = expenseLineRet25.TaxCodeRef.ListID.GetValue
 
                End If
 
            End If
            If (country = "UK") Then
                ' Get the value of the IExpenseLineRet.Tax1Amount element.
                If (Not expenseLineRet25.Tax1Amount Is Nothing) Then
                    Dim tax1Amount34 As Double
                    tax1Amount34 = expenseLineRet25.Tax1Amount.GetValue
                End If
 
            End If
        Next m
 
    End If
 
    ' Get the value of the IBillRet.ORItemLineRetList element.
    If (Not billRet.ORItemLineRetList Is Nothing) Then
        Dim n As Integer
        For n = 0 To billRet.ORItemLineRetList.Count - 1
            Dim orItemLineRet35 As IORItemLineRet
            Set orItemLineRet35 = billRet.ORItemLineRetList.GetAt(n)
            ' Get the value of the IORItemLineRet.ItemLineRet element.
            If (Not orItemLineRet35.ItemLineRet Is Nothing) Then
                ' Get the value of the IItemLineRet.TxnLineID element.
                Dim txnLineID36 As String
                txnLineID36 = orItemLineRet35.ItemLineRet.TxnLineID.GetValue
 
                ' Get the value of the IItemLineRet.ItemRef element.
                If (Not orItemLineRet35.ItemLineRet.ItemRef Is Nothing) Then
                    ' Get the FullName value.
                    Dim fullName37 As String
                    fullName37 = orItemLineRet35.ItemLineRet.ItemRef.FullName.GetValue
 
                    ' Get the ListID value.
                    Dim listID37 As String
                    listID37 = orItemLineRet35.ItemLineRet.ItemRef.ListID.GetValue
 
                End If
 
                ' Get the value of the IItemLineRet.Desc element.
                If (Not orItemLineRet35.ItemLineRet.Desc Is Nothing) Then
                    Dim desc38 As String
                    desc38 = orItemLineRet35.ItemLineRet.Desc.GetValue
                End If
 
                ' Get the value of the IItemLineRet.Quantity element.
                If (Not orItemLineRet35.ItemLineRet.Quantity Is Nothing) Then
                    Dim quantity39 As Double
                    quantity39 = orItemLineRet35.ItemLineRet.Quantity.GetValue
                End If
 
                ' Get the value of the IItemLineRet.Cost element.
                If (Not orItemLineRet35.ItemLineRet.Cost Is Nothing) Then
                    Dim cost40 As Double
                    cost40 = orItemLineRet35.ItemLineRet.Cost.GetValue
                End If
 
                ' Get the value of the IItemLineRet.Amount element.
                If (Not orItemLineRet35.ItemLineRet.Amount Is Nothing) Then
                    Dim amount41 As Double
                    amount41 = orItemLineRet35.ItemLineRet.Amount.GetValue
                End If
 
                ' Get the value of the IItemLineRet.CustomerRef element.
                If (Not orItemLineRet35.ItemLineRet.CustomerRef Is Nothing) Then
                    ' Get the FullName value.
                    Dim fullName42 As String
                    fullName42 = orItemLineRet35.ItemLineRet.CustomerRef.FullName.GetValue
 
                    ' Get the ListID value.
                    Dim listID42 As String
                    listID42 = orItemLineRet35.ItemLineRet.CustomerRef.ListID.GetValue
 
                End If
 
                ' Get the value of the IItemLineRet.ClassRef element.
                If (Not orItemLineRet35.ItemLineRet.ClassRef Is Nothing) Then
                    ' Get the FullName value.
                    Dim fullName43 As String
                    fullName43 = orItemLineRet35.ItemLineRet.ClassRef.FullName.GetValue
 
                    ' Get the ListID value.
                    Dim listID43 As String
                    listID43 = orItemLineRet35.ItemLineRet.ClassRef.ListID.GetValue
 
                End If
 
                ' Get the value of the IItemLineRet.BillableStatus element.
                If (Not orItemLineRet35.ItemLineRet.BillableStatus Is Nothing) Then
                    Dim billableStatus44 As ENBillableStatus
                    billableStatus44 = orItemLineRet35.ItemLineRet.BillableStatus.GetValue
                End If
 
                If Not (country = "US") Then
                    ' Get the value of the IItemLineRet.TaxCodeRef element.
                    If (Not orItemLineRet35.ItemLineRet.TaxCodeRef Is Nothing) Then
                        ' Get the FullName value.
                        Dim fullName45 As String
                        fullName45 = orItemLineRet35.ItemLineRet.TaxCodeRef.FullName.GetValue
 
                        ' Get the ListID value.
                        Dim listID45 As String
                        listID45 = orItemLineRet35.ItemLineRet.TaxCodeRef.ListID.GetValue
 
                    End If
 
                End If
                If (country = "UK") Then
                    ' Get the value of the IItemLineRet.Tax1Amount element.
                    If (Not orItemLineRet35.ItemLineRet.Tax1Amount Is Nothing) Then
                        Dim tax1Amount46 As Double
                        tax1Amount46 = orItemLineRet35.ItemLineRet.Tax1Amount.GetValue
                    End If
 
                End If
            End If
 
            ' Get the value of the IORItemLineRet.ItemGroupLineRet element.
            If (Not orItemLineRet35.ItemGroupLineRet Is Nothing) Then
                ' Get the value of the IItemGroupLineRet.TxnLineID element.
                Dim txnLineID47 As String
                txnLineID47 = orItemLineRet35.ItemGroupLineRet.TxnLineID.GetValue
 
                ' Get the value of the IItemGroupLineRet.ItemGroupRef element.
                ' Get the FullName value.
                Dim fullName48 As String
                fullName48 = orItemLineRet35.ItemGroupLineRet.ItemGroupRef.FullName.GetValue
 
                ' Get the ListID value.
                Dim listID48 As String
                listID48 = orItemLineRet35.ItemGroupLineRet.ItemGroupRef.ListID.GetValue
 
                ' Get the value of the IItemGroupLineRet.Desc element.
                If (Not orItemLineRet35.ItemGroupLineRet.Desc Is Nothing) Then
                    Dim desc49 As String
                    desc49 = orItemLineRet35.ItemGroupLineRet.Desc.GetValue
                End If
 
                ' Get the value of the IItemGroupLineRet.Quantity element.
                If (Not orItemLineRet35.ItemGroupLineRet.Quantity Is Nothing) Then
                    Dim quantity50 As Double
                    quantity50 = orItemLineRet35.ItemGroupLineRet.Quantity.GetValue
                End If
 
                ' Get the value of the IItemGroupLineRet.TotalAmount element.
                Dim totalAmount51 As Double
                totalAmount51 = orItemLineRet35.ItemGroupLineRet.TotalAmount.GetValue
 
                ' Get the value of the IItemGroupLineRet.ItemLineRetList element.
                If (Not orItemLineRet35.ItemGroupLineRet.ItemLineRetList Is Nothing) Then
                    Dim p As Integer
                    For p = 0 To orItemLineRet35.ItemGroupLineRet.ItemLineRetList.Count - 1
                        Dim itemLineRet52 As IItemLineRet
                        Set itemLineRet52 = orItemLineRet35.ItemGroupLineRet.ItemLineRetList.GetAt(p)
                        ' Get the value of the IItemLineRet.TxnLineID element.
                        Dim txnLineID53 As String
                        txnLineID53 = itemLineRet52.TxnLineID.GetValue
 
                        ' Get the value of the IItemLineRet.ItemRef element.
                        If (Not itemLineRet52.ItemRef Is Nothing) Then
                            ' Get the FullName value.
                            Dim fullName54 As String
                            fullName54 = itemLineRet52.ItemRef.FullName.GetValue
 
                            ' Get the ListID value.
                            Dim listID54 As String
                            listID54 = itemLineRet52.ItemRef.ListID.GetValue
 
                        End If
 
                        ' Get the value of the IItemLineRet.Desc element.
                        If (Not itemLineRet52.Desc Is Nothing) Then
                            Dim desc55 As String
                            desc55 = itemLineRet52.Desc.GetValue
                        End If
 
                        ' Get the value of the IItemLineRet.Quantity element.
                        If (Not itemLineRet52.Quantity Is Nothing) Then
                            Dim quantity56 As Double
                            quantity56 = itemLineRet52.Quantity.GetValue
                        End If
 
                        ' Get the value of the IItemLineRet.Cost element.
                        If (Not itemLineRet52.Cost Is Nothing) Then
                            Dim cost57 As Double
                            cost57 = itemLineRet52.Cost.GetValue
                        End If
 
                        ' Get the value of the IItemLineRet.Amount element.
                        If (Not itemLineRet52.Amount Is Nothing) Then
                            Dim amount58 As Double
                            amount58 = itemLineRet52.Amount.GetValue
                        End If
 
                        ' Get the value of the IItemLineRet.CustomerRef element.
                        If (Not itemLineRet52.CustomerRef Is Nothing) Then
                            ' Get the FullName value.
                            Dim fullName59 As String
                            fullName59 = itemLineRet52.CustomerRef.FullName.GetValue
 
                            ' Get the ListID value.
                            Dim listID59 As String
                            listID59 = itemLineRet52.CustomerRef.ListID.GetValue
 
                        End If
 
                        ' Get the value of the IItemLineRet.ClassRef element.
                        If (Not itemLineRet52.ClassRef Is Nothing) Then
                            ' Get the FullName value.
                            Dim fullName60 As String
                            fullName60 = itemLineRet52.ClassRef.FullName.GetValue
 
                            ' Get the ListID value.
                            Dim listID60 As String
                            listID60 = itemLineRet52.ClassRef.ListID.GetValue
 
                        End If
 
                        ' Get the value of the IItemLineRet.BillableStatus element.
                        If (Not itemLineRet52.BillableStatus Is Nothing) Then
                            Dim billableStatus61 As ENBillableStatus
                            billableStatus61 = itemLineRet52.BillableStatus.GetValue
                        End If
 
                        If Not (country = "US") Then
                            ' Get the value of the IItemLineRet.TaxCodeRef element.
                            If (Not itemLineRet52.TaxCodeRef Is Nothing) Then
                                ' Get the FullName value.
                                Dim fullName62 As String
                                fullName62 = itemLineRet52.TaxCodeRef.FullName.GetValue
 
                                ' Get the ListID value.
                                Dim listID62 As String
                                listID62 = itemLineRet52.TaxCodeRef.ListID.GetValue
 
                            End If
 
                        End If
                        If (country = "UK") Then
                            ' Get the value of the IItemLineRet.Tax1Amount element.
                            If (Not itemLineRet52.Tax1Amount Is Nothing) Then
                                Dim tax1Amount63 As Double
                                tax1Amount63 = itemLineRet52.Tax1Amount.GetValue
                            End If
 
                        End If
                    Next p
 
                End If
 
            End If
 
        Next n
 
    End If
 
    If Not (country = "US") Then
        ' Get the value of the IBillRet.Tax1Total element.
        If (Not billRet.Tax1Total Is Nothing) Then
            Dim tax1Total64 As Double
            tax1Total64 = billRet.Tax1Total.GetValue
        End If
 
    End If
    If Not (country = "US") Then
        ' Get the value of the IBillRet.Tax2Total element.
        If (Not billRet.Tax2Total Is Nothing) Then
            Dim tax2Total65 As Double
            tax2Total65 = billRet.Tax2Total.GetValue
        End If
 
    End If
    If Not (country = "US") Then
        ' Get the value of the IBillRet.ExchangeRate element.
        If (Not billRet.ExchangeRate Is Nothing) Then
            Dim exchangeRate66 As Single
            exchangeRate66 = billRet.ExchangeRate.GetValue
        End If
 
    End If
    ' Get the value of the IBillRet.OpenAmount element.
    If (Not billRet.OpenAmount Is Nothing) Then
        Dim openAmount67 As Double
        openAmount67 = billRet.OpenAmount.GetValue
    End If
 
    ' Get the value of the IBillRet.DataExtRetList element.
    If (Not billRet.DataExtRetList Is Nothing) Then
        Dim q As Integer
        For q = 0 To billRet.DataExtRetList.Count - 1
            Dim dataExtRet68 As IDataExtRet
            Set dataExtRet68 = billRet.DataExtRetList.GetAt(q)
            ' Get the value of the IDataExtRet.OwnerID element.
            If (Not dataExtRet68.OwnerID Is Nothing) Then
                Dim ownerID69 As String
                ownerID69 = dataExtRet68.OwnerID.GetValue
            End If
 
            ' Get the value of the IDataExtRet.DataExtName element.
            Dim dataExtName70 As String
            dataExtName70 = dataExtRet68.DataExtName.GetValue
 
            ' Get the value of the IDataExtRet.DataExtType element.
            Dim dataExtType71 As ENDataExtType
            dataExtType71 = dataExtRet68.DataExtType.GetValue
 
            ' Get the value of the IDataExtRet.DataExtValue element.
            Dim dataExtValue72 As String
            dataExtValue72 = dataExtRet68.DataExtValue.GetValue
 
        Next q
 
    End If
 
    If (country = "UK") Then
        ' Get the value of the IBillRet.AmountIncludesVAT element.
        If (Not billRet.AmountIncludesVAT Is Nothing) Then
            Dim amountIncludesVAT73 As Boolean
            amountIncludesVAT73 = billRet.AmountIncludesVAT.GetValue
        End If
 
    End If
End Sub

