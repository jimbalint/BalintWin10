VERSION 5.00
Begin VB.Form frmQBCheckUpdate 
   Caption         =   "Update Payroll NET PAY Check info to QB"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
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
   ScaleHeight     =   8205
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoName 
      Caption         =   "Don't include Employee Name in QB check memo field"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   5880
      Width           =   5295
   End
   Begin VB.ComboBox cmbPayee 
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4080
      Width           =   5775
   End
   Begin VB.ComboBox cmbExpense 
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5280
      Width           =   5775
   End
   Begin VB.ComboBox cmbChecking 
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4680
      Width           =   5775
   End
   Begin VB.CommandButton cmdGetQB 
      Caption         =   "REFRESH QB CHART OF ACCOUNTS"
      Height          =   1095
      Left            =   2558
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   5078
      TabIndex        =   6
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2558
      TabIndex        =   5
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Payee:"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   8895
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   8895
   End
   Begin VB.Label Label3 
      Caption         =   "Expense Account:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblPRInfo 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   8895
   End
   Begin VB.Label Label2 
      Caption         =   "Checking Account:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Works best when the QuickBooks File is open to refresh or update check information!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1455
      Left            =   4965
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmQBCheckUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GlobalID As Long
Dim QBIDChk, QBIDExp, QBIDPay As String
Dim rs As New ADODB.Recordset
Dim rsPay As New ADODB.Recordset
Dim QBOpened As Boolean

Private Sub Form_Load()

Dim DateFmt As String

    ' set to TRUE if chart of accts is refreshed
    ' if so - don't need to open connection again
    QBOpened = False

    If TableExists("QBAccount", cn) = False Then
        QBAccountCreate
    End If

    Me.lblCompanyName = PRCompany.Name
    
    If PRBatch.GetByID(frmBatchList.BatchID) = False Then
        MsgBox "PR Batch info not found: " & frmBatchList.BatchID, vbExclamation
        GoBack
    End If
    
    DateFmt = "mm/dd/yy"
    Me.lblPRInfo = "PE Date: " & Format(PRBatch.PEDate, DateFmt) & " Chk Date: " & _
                   Format(PRBatch.CheckDate, DateFmt) & vbCr & _
                   "Check Count: " & PRBatch.RecCount

    Me.lblMsg1 = ""
    Me.lblMsg2 = ""

    ' accounts previously chosen for this company
    SQLString = "SELECT * FROM PRGlobal WHERE UserID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & PREquate.GlobalTypeQBPRChk
    If PRGlobal.GetBySQL(SQLString) = True Then
        GlobalID = PRGlobal.GlobalID
        QBIDChk = PRGlobal.Var1 & ""
        QBIDExp = PRGlobal.Var2 & ""
        QBIDPay = PRGlobal.Var3 & ""
        Me.chkNoName = PRGlobal.Byte1
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeQBPRChk
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
        GlobalID = PRGlobal.GlobalID
        QBIDChk = ""
        QBIDExp = ""
        QBIDPay = ""
    End If

    LoadQBAccts

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

Dim ChkQBID, ExpQBID, PayQBID As String

    If Me.cmbExpense.ListIndex = -1 Then
        MsgBox "Expense Account has not been chosen!", vbExclamation
        Exit Sub
    End If
    
    If Me.cmbChecking.ListIndex = -1 Then
        MsgBox "Checking Account has not been chosen!", vbExclamation
        Exit Sub
    End If

    ' store selections to PRGlobal
    If PRGlobal.GetByID(GlobalID) = False Then
        MsgBox "Global Error: " & GlobalID, vbExclamation
        GoBack
    End If

    rs.Find "LIndex = " & Me.cmbChecking.ListIndex, 0, adSearchForward, 1
    If rs.EOF Then
        MsgBox "Select error!", vbExclamation
        GoBack
    End If
    PRGlobal.Var1 = rs!QBID
    ChkQBID = rs!QBID

    rs.Find "LIndex = " & Me.cmbExpense.ListIndex, 0, adSearchForward, 1
    If rs.EOF Then
        MsgBox "Select error!", vbExclamation
        GoBack
    End If
    PRGlobal.Var2 = rs!QBID
    ExpQBID = rs!QBID

    rsPay.Find "LIndex = " & Me.cmbPayee.ListIndex, 0, adSearchForward, 1
    If rsPay.EOF Then
        MsgBox "Select error!", vbExclamation
        GoBack
    End If
    PRGlobal.Var3 = rsPay!QBID
    PayQBID = rsPay!QBID

    PRGlobal.Byte1 = Me.chkNoName
    
    PRGlobal.Save (Equate.RecPut)

    Me.MousePointer = vbHourglass
    
    DoCheckAddRq "US", 5, 0, frmBatchList.BatchID, _
                 ChkQBID, Me.cmbChecking, _
                 ExpQBID, Me.cmbExpense, _
                 PayQBID, Me.cmbPayee

    MsgBox "Check Update Complete!", vbInformation

    Me.MousePointer = vbArrow
    
    Unload Me

End Sub

Private Sub cmdGetQB_Click()
    frmQBAccts.Show vbModal
    LoadQBAccts
End Sub

Private Sub LoadQBAccts()

Dim PayCount, QBCount As Long
Dim ChkListIndex, ExpListIndex, PayListIndex As Long

    Me.cmbChecking.Clear
    Me.cmbExpense.Clear
    Me.cmbPayee.Clear

    ' temp record set to store QB combo info
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "LIndex", adDouble
    rs.Fields.Append "QBID", adVarChar, 50, adFldIsNullable
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    On Error Resume Next
    rsPay.Close
    On Error GoTo 0
    rsPay.CursorLocation = adUseClient
    rsPay.Fields.Append "LIndex", adDouble
    rsPay.Fields.Append "QBID", adVarChar, 50, adFldIsNullable
    rsPay.Open , , adOpenDynamic, adLockOptimistic
    
    QBCount = 0
    PayCount = 0
    ChkListIndex = -1
    ExpListIndex = -1
    PayListIndex = -1

    SQLString = "SELECT * FROM QBAccount ORDER BY Name "
    If QBAccount.GetBySQL(SQLString) = False Then Exit Sub
    
    Do
        
        If QBAccount.AccountType <> "VENDOR" And QBAccount.AccountType <> "TEMPLATE" Then
            
            With Me.cmbChecking
                .AddItem QBAccount.Name
            End With
            With Me.cmbExpense
                .AddItem QBAccount.Name
            End With
        
            rs.AddNew
            rs!LIndex = QBCount
            rs!QBID = QBAccount.QBID
            rs.Update
            
            ' store listindex for this company
            If QBIDChk = QBAccount.QBID Then ChkListIndex = QBCount
            If QBIDExp = QBAccount.QBID Then ExpListIndex = QBCount
            
            QBCount = QBCount + 1
        
        ElseIf QBAccount.AccountType = "VENDOR" Then
            
            Me.cmbPayee.AddItem QBAccount.Name
        
            rsPay.AddNew
            rsPay!LIndex = PayCount
            rsPay!QBID = QBAccount.QBID
            rsPay.Update
            
            If QBIDPay = QBAccount.QBID Then PayListIndex = PayCount
            
            PayCount = PayCount + 1
        
        End If
        
        If QBAccount.GetNext = False Then Exit Do
    
    Loop

    Me.cmbChecking.ListIndex = ChkListIndex
    Me.cmbExpense.ListIndex = ExpListIndex
    Me.cmbPayee.ListIndex = PayListIndex

End Sub

Private Sub DoCheckAddRq(ByVal country As String, _
                        ByVal MajorVersion As Integer, _
                        ByVal MinorVersion As Integer, _
                        ByVal BatchID As Long, _
                        ByVal QBIDChk As String, _
                        ByVal QBChkName As String, _
                        ByVal QBIDExp As String, _
                        ByVal QBExpName As String, _
                        ByVal QBIDPay As String, _
                        ByVal QBPayName As String)
  
  ' On Error GoTo Errs
  
'  On Error GoTo Errs
  
    If PRBatch.GetByID(BatchID) = False Then
        MsgBox "PR Batch NF: " & BatchID, vbExclamation
        GoBack
    End If
  
    SQLString = "SELECT * FROM PRHist WHERE BatchID = " & BatchID & _
                " ORDER BY CheckNumber"
    If PRHist.GetBySQL(SQLString) = False Then
        MsgBox "No Payroll data to export!", vbExclamation
        GoBack
    End If
  
    ' =====================================================================
    
    ' start session and open connection
    If QBOpen(Me, Me.lblMsg2) = False Then GoBack
    
    ' ================================================================
  
    ' Create the message set request object for the specific version messages.
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = SessMgr.CreateMsgSetRequest(country, MajorVersion, MinorVersion)
    requestMsgSet.Attributes.OnError = roeContinue
  
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    Do
        If PREmployee.GetByID(PRHist.EmployeeID) = False Then
            MsgBox "Employee NF: " & PRHist.EmployeeID, vbExclamation
            GoBack
        End If
        Me.lblMsg1 = "Building Check Add Request: " & PREmployee.LFName
        Me.Refresh
        BuildCheckAddRq requestMsgSet, country, _
                        QBIDChk, QBChkName, _
                        QBIDExp, QBExpName, _
                        QBIDPay, QBPayName
        If PRHist.GetNext = False Then Exit Do
    Loop
    ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    ' Perform the request and obtain a response from QuickBooks.
    Dim responseMsgSet As IMsgSetResponse
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
    Unload Me
  
    ' ParseCheckAddRs responseMsgSet, country
  
    Exit Sub
  
Errs:
    MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error"
  
    ' SampleCodeForm.ErrorMsg.Text = Err.Description
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
End Sub

Private Sub BuildCheckAddRq(ByVal requestMsgSet As IMsgSetRequest, _
                           ByVal country As String, _
                           ByVal QBIDChk As String, _
                           ByVal QBChkName As String, _
                           ByVal QBIDExp As String, _
                           ByVal QBExpName As String, _
                           ByVal QBIDPay As String, _
                           ByVal QBPayName As String)
  
  If (requestMsgSet Is Nothing) Then
    Exit Sub
  End If
  
  'Add the request to the message set request object.
  Dim checkAdd As ICheckAdd
  Set checkAdd = requestMsgSet.AppendCheckAddRq
  
  'Set the elements of ICheckAdd.
  
  ' Set the FullName value.
  checkAdd.AccountRef.FullName.SetValue frmQBCheckUpdate.cmbChecking
  
  ' Set the ListID value.
  checkAdd.AccountRef.ListID.SetValue QBIDChk
  
  ' Set the FullName value.
  checkAdd.PayeeEntityRef.FullName.SetValue QBPayName
  
  ' Set the ListID value.
  checkAdd.PayeeEntityRef.ListID.SetValue QBIDPay
  
  ' Set the value of the ICheckAdd.RefNumber element.
  checkAdd.RefNumber.SetValue PRHist.CheckNumber
  
  ' Set the value of the ICheckAdd.TxnDate element.
  checkAdd.TxnDate.SetValue PRHist.CheckDate
  
  ' Set the value of the ICheckAdd.Memo element.
  If Me.chkNoName = 0 Then
      checkAdd.Memo.SetValue PREmployee.LFName
  Else
      checkAdd.Memo.SetValue "Emp#: " & PREmployee.EmployeeNumber
  End If
  
  ' Set the value of the IAddress.Addr1 element.
  checkAdd.Address.Addr1.SetValue ""
  
  ' Set the value of the IAddress.Addr2 element.
  checkAdd.Address.Addr2.SetValue ""
  
  ' Set the value of the IAddress.Addr3 element.
  checkAdd.Address.Addr3.SetValue ""
  
  ' Set the value of the IAddress.Addr4 element.
  checkAdd.Address.Addr4.SetValue ""
  
  ' Set the value of the IAddress.City element.
  checkAdd.Address.City.SetValue ""
  
  If (country = "US") Then
    ' Set the value of the IAddress.State element.
    checkAdd.Address.State.SetValue ""
  
  End If
  If (country = "UK") Then
    ' Set the value of the IAddress.County element.
    checkAdd.Address.County.SetValue ""
  
  End If
  If (country = "CA") Then
    ' Set the value of the IAddress.Province element.
    checkAdd.Address.Province.SetValue ""
  
  End If
  ' Set the value of the IAddress.PostalCode element.
  checkAdd.Address.PostalCode.SetValue ""
  
  ' Set the value of the IAddress.Country element.
  checkAdd.Address.country.SetValue "l"
  
  ' Set the value of the ICheckAdd.IsToBePrinted element.
  checkAdd.IsToBePrinted.SetValue False
  
  'Add multiple elements to the list. In this case we will add 5 elements.
  Dim expenseLineAdd1 As IExpenseLineAdd
    
  ' Append an element to the list and save the element in expenseLineAdd1 so we can set its values.
  Set expenseLineAdd1 = checkAdd.ExpenseLineAddList.Append
  
  ' Set the FullName value.
  expenseLineAdd1.AccountRef.FullName.SetValue QBExpName
  
  ' Set the ListID value.
  expenseLineAdd1.AccountRef.ListID.SetValue QBIDExp
  
  ' Set the value of the IExpenseLineAdd.Amount element.
  expenseLineAdd1.Amount.SetValue PRHist.Net
  
  ' Set the value of the IExpenseLineAdd.Memo element.
  expenseLineAdd1.Memo.SetValue PREmployee.LFName
  
  ' ****************************************
  ' * Cust Ref
  ' Set the FullName value.
  ' expenseLineAdd1.CustomerRef.FullName.SetValue "ab"
  
  ' Set the ListID value.
  ' expenseLineAdd1.CustomerRef.ListID.SetValue "ab"
  
  ' ****************************************
  
  ' ****************************************
  ' * Class Ref
  ' Set the FullName value.
  ' expenseLineAdd1.ClassRef.FullName.SetValue "ab"
  
  ' Set the ListID value.
  ' expenseLineAdd1.ClassRef.ListID.SetValue "ab"
  
  ' Set the value of the IExpenseLineAdd.BillableStatus element.
  ' expenseLineAdd1.BillableStatus.SetValue bsBillable
  
'    If Not (country = "US") Then
'      ' Set the FullName value.
'      expenseLineAdd1.TaxCodeRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      expenseLineAdd1.TaxCodeRef.ListID.SetValue "ab"
'
'    End If
'    ' Set the value of the IExpenseLineAdd.defMacro element.
'    expenseLineAdd1.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
  
'  'Add multiple elements to the list. In this case we will add 5 elements.
'  Dim orItemLineAdd2 As IORItemLineAdd
'  Dim k As Integer
'  For k = 0 To 4
'    ' Append an element to the list and save the element in orItemLineAdd2 so we can set its values.
'    Set orItemLineAdd2 = checkAdd.ORItemLineAddList.Append
'
'    ' Only can set one of the OR elements.
'    ' We will portray this restriction by using an If/Then/Else.
'    Dim orItemLineAddORElement3 As String
'    orItemLineAddORElement3 = "ItemLineAdd"
'    If (orItemLineAddORElement3 = "ItemLineAdd") Then
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.ItemRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.ItemRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemLineAdd.Desc element.
'      orItemLineAdd2.ItemLineAdd.Desc.SetValue "val"
'
'      ' Set the value of the IItemLineAdd.Quantity element.
'      orItemLineAdd2.ItemLineAdd.Quantity.SetValue 2#
'
'      ' Set the value of the IItemLineAdd.Cost element.
'      orItemLineAdd2.ItemLineAdd.Cost.SetValue 2#
'
'      ' Set the value of the IItemLineAdd.Amount element.
'      orItemLineAdd2.ItemLineAdd.Amount.SetValue 2#
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.CustomerRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.CustomerRef.ListID.SetValue "ab"
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.ClassRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.ClassRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemLineAdd.BillableStatus element.
'      orItemLineAdd2.ItemLineAdd.BillableStatus.SetValue bsBillable
'
'      ' Set the FullName value.
'      orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemLineAdd.OverrideItemAccountRef.ListID.SetValue "ab"
'
'      If Not (country = "US") Then
'        ' Set the FullName value.
'        orItemLineAdd2.ItemLineAdd.TaxCodeRef.FullName.SetValue "ab"
'
'        ' Set the ListID value.
'        orItemLineAdd2.ItemLineAdd.TaxCodeRef.ListID.SetValue "ab"
'
'      End If
'      If (country = "US") Then
'        ' Set the value of the ILinkToTxn.TxnID element.
'        orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnID.SetValue "val"
'
'        ' Set the value of the ILinkToTxn.TxnLineID element.
'        orItemLineAdd2.ItemLineAdd.LinkToTxn.TxnLineID.SetValue "val"
'
'      End If
'    ElseIf (orItemLineAddORElement3 = "ItemGroupLineAdd") Then
'      ' Set the FullName value.
'      orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.FullName.SetValue "ab"
'
'      ' Set the ListID value.
'      orItemLineAdd2.ItemGroupLineAdd.ItemGroupRef.ListID.SetValue "ab"
'
'      ' Set the value of the IItemGroupLineAdd.Desc element.
'      orItemLineAdd2.ItemGroupLineAdd.Desc.SetValue "val"
'
'      ' Set the value of the IItemGroupLineAdd.Quantity element.
'      orItemLineAdd2.ItemGroupLineAdd.Quantity.SetValue 2#
'
'    End If
'
'  Next k
'
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.Tax1Total element.
'    checkAdd.Tax1Total.SetValue 2#
'
'  End If
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.Tax2Total element.
'    checkAdd.Tax2Total.SetValue 2#
'
'  End If
'  If Not (country = "US") Then
'    ' Set the value of the ICheckAdd.ExchangeRate element.
'    checkAdd.ExchangeRate.SetValue 2.5
'
'  End If
'  If (country = "UK") Then
'    ' Set the value of the ICheckAdd.AmountIncludesVAT element.
'    checkAdd.AmountIncludesVAT.SetValue True
'
'  End If
'  If (country = "US") Then
'    ' Set the value of the ICheckAdd.IncludeRetElementList element.
'    checkAdd.IncludeRetElementList.Add "val"
'
'  End If
'  ' Set the value of the ICheckAdd.defMacro element.
'  checkAdd.defMacro.SetValue "TxnID:" & Format(Now, "yyyymmddhhmmss")
  
End Sub




