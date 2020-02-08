VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmQBInvUpdate 
   Caption         =   "Update Invoicing to QuickBooks"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate tdbStartDate 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Calendar        =   "frmQBInvUpdate.frx":0000
      Caption         =   "frmQBInvUpdate.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBInvUpdate.frx":016A
      Keys            =   "frmQBInvUpdate.frx":0188
      Spin            =   "frmQBInvUpdate.frx":01E6
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
      Text            =   "04/12/2010"
      ValidateMode    =   0
      ValueVT         =   7536647
      Value           =   40280
      CenturyMode     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Invoices to update:  "
      Height          =   975
      Left            =   4800
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optRecent 
         Caption         =   "Recent"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin TDBDate6Ctl.TDBDate tdbInvoiceDate 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Calendar        =   "frmQBInvUpdate.frx":020E
      Caption         =   "frmQBInvUpdate.frx":030E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBInvUpdate.frx":037C
      Keys            =   "frmQBInvUpdate.frx":039A
      Spin            =   "frmQBInvUpdate.frx":03F8
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
      Text            =   "03/19/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40256
      CenturyMode     =   0
   End
   Begin VB.ComboBox cmbQBTemplate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   8640
      Width           =   6855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4680
      TabIndex        =   6
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2040
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton cmdQBRefresh 
      Caption         =   "Refresh QB Chart of Accounts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   9240
      Width           =   3375
   End
   Begin VB.ComboBox cmbQBEmp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7680
      Width           =   6855
   End
   Begin VB.ComboBox cmbQBAR 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6720
      Width           =   6855
   End
   Begin TDBDate6Ctl.TDBDate tdbEndDate 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   2400
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      Calendar        =   "frmQBInvUpdate.frx":0420
      Caption         =   "frmQBInvUpdate.frx":0520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBInvUpdate.frx":0586
      Keys            =   "frmQBInvUpdate.frx":05A4
      Spin            =   "frmQBInvUpdate.frx":0602
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
      Text            =   "04/12/2010"
      ValidateMode    =   0
      ValueVT         =   7536647
      Value           =   40280
      CenturyMode     =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   8520
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label5 
      Caption         =   "QuickBooks Invoice Template:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   8160
      Width           =   3495
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   8055
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   8055
   End
   Begin VB.Label lblPRInfo 
      Alignment       =   2  'Center
      Caption         =   "PR Info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   840
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Works Best when the QuickBooks data file is open!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   3600
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "QuickBooks Employee Billing Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   7200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "QuickBooks A/R Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmQBInvUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DateFmt As String
Dim QBIDAR, QBIDEmp, QBIDTPL As String
Dim rs As New ADODB.Recordset
Dim ARListIndex, EmpListIndex As Long
Dim QBOpened As Boolean
Dim GlobalID As Long

Dim rsTPL As New ADODB.Recordset
Dim TPLListIndex As Long

Dim i, j, k As Long
Dim X, Y, Z As String

Dim EmpID, EmpNum As Long
Dim PayCount, QBCount As Long

Dim rsPR As New ADODB.Recordset
Dim rsEE As New ADODB.Recordset
Dim boo As Boolean
Dim MsgResponse As Variant
Dim Flg As Boolean

' General QB variables
Dim requestMsgSet As IMsgSetRequest
Dim responseMsgSet As IMsgSetResponse
Dim ResponseList As IResponseList
Dim Response As IResponse
Dim ResponseType As Integer
Dim orItemRetList As IORItemRetList

' QB Item variables
Dim ItemQuery As IItemQuery
Dim orItemRet As IORItemRet
Dim itemServiceAdd As IItemServiceAdd
Dim itemServiceRet As IItemServiceRet

' QB Invoice Variables
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

Dim InvSeqNum, LastJobID As Long


Private Sub Form_Load()

    ' set to TRUE if chart of accts is refreshed
    ' if so - don't need to open connection again
    QBOpened = False

    Me.optRecent = True

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
                " AND TypeCode = " & PREquate.GlobalTypeQBInv
    If PRGlobal.GetBySQL(SQLString) = True Then
        GlobalID = PRGlobal.GlobalID
        QBIDAR = PRGlobal.Var1 & ""
        QBIDEmp = PRGlobal.Var2 & ""
        QBIDTPL = PRGlobal.Var3 & ""
    Else
        PRGlobal.Clear
        PRGlobal.TypeCode = PREquate.GlobalTypeQBInv
        PRGlobal.UserID = PRCompany.CompanyID
        PRGlobal.Save (Equate.RecAdd)
        GlobalID = PRGlobal.GlobalID
        QBIDAR = ""
        QBIDEmp = ""
        QBIDTPL = ""
    End If

    LoadQBAccts
    LoadQBTemplates

    Me.tdbInvoiceDate = PRBatch.PEDate
    Me.tdbEndDate = PRBatch.PEDate
    Me.tdbStartDate = PRBatch.PEDate - 6

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
Private Sub cmdQBRefresh_Click()
    frmQBAccts.Show vbModal
    LoadQBAccts
    LoadQBTemplates
End Sub

Private Sub LoadQBTemplates()

    Me.cmbQBTemplate.Clear
    
    On Error Resume Next
    rsTPL.Close
    On Error GoTo 0
    rsTPL.CursorLocation = adUseClient
    rsTPL.Fields.Append "LIndex", adDouble
    rsTPL.Fields.Append "QBID", adVarChar, 50, adFldIsNullable
    rsTPL.Open , , adOpenDynamic, adLockOptimistic
    
    TPLListIndex = -1
    QBCount = 0
    
    With Me.cmbQBTemplate
    
        SQLString = "SELECT * FROM QBAccount " & _
                    "WHERE AccountType = 'TEMPLATE' " & _
                    "ORDER BY Name "
        If QBAccount.GetBySQL(SQLString) = False Then Exit Sub
        
        Do
            
            .AddItem QBAccount.Name
            
            rsTPL.AddNew
            rsTPL!LIndex = QBCount
            rsTPL!QBID = QBAccount.QBID
            rsTPL.Update
            
            ' store listindex for this company
            If QBIDTPL = QBAccount.QBID Then TPLListIndex = QBCount
                    
            QBCount = QBCount + 1
            
            If QBAccount.GetNext = False Then Exit Do
        
        Loop
    
        .ListIndex = TPLListIndex

    End With

End Sub


Private Sub LoadQBAccts()

Dim ChkListIndex, ExpListIndex, PayListIndex As Long

    Me.cmbQBAR.Clear
    Me.cmbQBEmp.Clear

    ' temp record set to store QB combo info
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    rs.CursorLocation = adUseClient
    rs.Fields.Append "LIndex", adDouble
    rs.Fields.Append "QBID", adVarChar, 50, adFldIsNullable
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    QBCount = 0
    
    ARListIndex = -1
    EmpListIndex = -1
    
    SQLString = "SELECT * FROM QBAccount " & _
                "WHERE AccountType <> 'VENDOR' " & _
                "AND AccountType <> 'TEMPLATE' " & _
                "ORDER BY Name "
    If QBAccount.GetBySQL(SQLString) = False Then Exit Sub
    
    Do
        
        With Me.cmbQBAR
            .AddItem QBAccount.Name
        End With
        With Me.cmbQBEmp
            .AddItem QBAccount.Name
        End With
    
        rs.AddNew
        rs!LIndex = QBCount
        rs!QBID = QBAccount.QBID
        rs.Update
        
        ' store listindex for this company
        If QBIDAR = QBAccount.QBID Then ARListIndex = QBCount
        If QBIDEmp = QBAccount.QBID Then EmpListIndex = QBCount
                
        QBCount = QBCount + 1
        
        If QBAccount.GetNext = False Then Exit Do
    
    Loop

    Me.cmbQBAR.ListIndex = ARListIndex
    Me.cmbQBEmp.ListIndex = EmpListIndex

End Sub

Private Function ParseEENum(ByVal InString) As Long

Dim EString As String
Dim e As Long
Dim eFlag As Boolean

    ' InString is in the format:
    ' PR_{EE Last/First Name}/{EmpNum}
    ' return the employee number

    ParseEENum = 0
    If IsNull(InString) Then Exit Function
    If InString = "" Then Exit Function
    If Len(InString) <= 8 Then Exit Function
    
    e = InStr(1, InString, "/", vbTextCompare)
    If e <= 0 Then Exit Function
    
    ParseEENum = CLng(Mid(InString, e + 1, Len(InString) - 1 + 1))

End Function

Private Function AddQBItem(ByVal EmployeeID As Long) As String
    
Dim QBName, QBFullName As String
Dim QBln As Long
    
    ' make new session each time
    NewQBSession
    
    AddQBItem = ""
    
    If PREmployee.GetByID(EmployeeID) = False Then Exit Function
    
    ' max length of 31 for QB Item Name
    QBName = "PR_" & PREmployee.LFName & "/" & PREmployee.EmployeeNumber
    If Len(QBName) > 31 Then
        X = Mid(PREmployee.LFName, 1, 29 - Len("PR_/" & PREmployee.EmployeeNumber))
        QBName = "PR_" & X & "/" & PREmployee.EmployeeNumber
    End If
    
    QBFullName = PREmployee.LFName
    
    Set itemServiceAdd = requestMsgSet.AppendItemServiceAddRq
    
    itemServiceAdd.Name.SetValue QBName
    itemServiceAdd.IsActive.SetValue True
    
    ' itemServiceAdd.ORSalesPurchase.SalesAndPurchase.SalesDesc.SetValue QBFullName
    itemServiceAdd.ORSalesPurchase.SalesOrPurchase.Desc.SetValue QBFullName
    
    ' itemServiceAdd.ParentRef.ListID.SetValue ""
    ' itemServiceAdd.ParentRef.FullName.SetValue ""
    itemServiceAdd.SalesTaxCodeRef.FullName.SetValue "Tax"
    itemServiceAdd.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.SetValue QBIDEmp
    
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    If responseMsgSet Is Nothing Then
        MsgBox "No response ..."
        Exit Function
    End If
    
    Set ResponseList = responseMsgSet.ResponseList
    If ResponseList Is Nothing Then
        MsgBox "No Reponse List"
        Exit Function
    End If
    
    Set Response = ResponseList.GetAt(0)
    If Response.StatusCode <> 0 Then
        MsgBox "Status Code: " & Response.StatusCode, vbExclamation
        Exit Function
    End If
        
    If (Response.Detail Is Nothing) Then
        MsgBox "Response Detail is nothing", vbExclamation
        Exit Function
    End If
    
    ' *******************************************************
    ResponseType = Response.Type.GetValue
    
    If (ResponseType <> rtItemServiceAddRs) Then
        MsgBox "Invalid response type", vbExclamation
        Exit Function
    End If
    
    Set itemServiceRet = Response.Detail
    
    If itemServiceRet Is Nothing Then
        MsgBox "itemServiceRet is nothing", vbExclamation
        Exit Function
    End If
    
    AddQBItem = itemServiceRet.ListID.GetValue

End Function


Private Sub cmdOK_Click()
    
Dim THrs As Currency
    
    If Me.optAll = True Then
        If MsgBox("OK to create QB Invoices" & vbCr & vbCr & _
                    "Even if the payroll data has already been updated?", vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    
    If PRGlobal.GetByID(GlobalID) Then
        
        SQLString = "LIndex = " & Me.cmbQBAR.ListIndex
        rs.Find SQLString, 0, adSearchForward, 1
        If rs.EOF = False Then
            PRGlobal.Var1 = rs!QBID
            QBIDAR = rs!QBID
        End If
        
        SQLString = "LIndex = " & Me.cmbQBEmp.ListIndex
        rs.Find SQLString, 0, adSearchForward, 1
        If rs.EOF = False Then
            PRGlobal.Var2 = rs!QBID
            QBIDEmp = rs!QBID
        End If

        SQLString = "LIndex = " & Me.cmbQBTemplate.ListIndex
        rsTPL.Find SQLString, 0, adSearchForward, 1
        If rsTPL.EOF = False Then
            PRGlobal.Var3 = rsTPL!QBID
            QBIDTPL = rsTPL!QBID
        End If
        
        PRGlobal.Save (Equate.RecPut)
    
    End If

    ' make sure QB item exists for each Employee in batch
    ' Item Name - PR_{LFName}/EE# (max 31 chars)
    ' get temp RS of employees to bill in this batch
    ' update QBID field from existing items
    ' add item codes as needed
    '
    ' sort Billing data by:
    ' Customer / Employee / Pay Type / Billing Rate
    '    summ hours and bill amount - store first/last date
    
            
    ' open QB connection if necessary
    If QBOpened = False Then
        If QBOpen(Me, Me.lblMsg1) = False Then GoBack
    End If
        
    Me.lblMsg1 = "Now initializing data storage ..."
    Me.Refresh
        
    ' record set for each invoice line
    On Error Resume Next
    rsPR.Close
    On Error GoTo 0
    rsPR.CursorLocation = adUseClient
    rsPR.Fields.Append "JobID", adDouble
    rsPR.Fields.Append "EmployeeID", adDouble
    rsPR.Fields.Append "PRItemID", adDouble
    rsPR.Fields.Append "BillingRate", adCurrency
    rsPR.Fields.Append "Hours", adCurrency
    rsPR.Fields.Append "Amount", adCurrency
    rsPR.Fields.Append "InvSeqNum", adDouble
    rsPR.Fields.Append "TSRecID", adDouble
    rsPR.Fields.Append "StartDate", adDate
    rsPR.Fields.Append "EndDate", adDate
    rsPR.Open , , adOpenDynamic, adLockOptimistic
    
    ' recordset for each employee
    ' get the QBItem ID
    On Error Resume Next
    rsEE.Close
    On Error GoTo 0
    rsEE.CursorLocation = adUseClient
    rsEE.Fields.Append "EmployeeID", adDouble
    rsEE.Fields.Append "QBItemID", adVarChar, 50, adFldIsNullable
    rsEE.Open , , adOpenDynamic, adLockOptimistic
    
    Me.lblMsg1 = "Now gathering payroll data ..."
    Me.Refresh
    
    ' use PRDist
    If Me.optAll = False Then
        SQLString = "SELECT * FROM PRDist WHERE BatchID = " & PRBatch.BatchID & _
                    " AND BillingRate <> 0" & _
                    " AND QBInvoiceID = '' " & _
                    " AND JobID <> 999999"
    Else
        SQLString = "SELECT * FROM PRDist WHERE BatchID = " & PRBatch.BatchID & _
                    " AND BillingRate <> 0 " & _
                    " AND JobID <> 999999"
    End If
    
    If PRDist.GetBySQL(SQLString) = False Then
        MsgBox "No Billable entries exist for this batch!", vbExclamation
        GoBack
        Exit Sub
    End If
    
    Do
    
        ' add employee record?
        SQLString = "EmployeeID = " & PRDist.EmployeeID
        rsEE.Find SQLString, 0, adSearchForward, 1
        If rsEE.EOF Then
            rsEE.AddNew
            rsEE!EmployeeID = PRDist.EmployeeID
            rsEE!QBItemID = ""
            rsEE.Update
        End If

        ' add billing record
        ' 02/22/2011 No group - add a record for each PRDist
        Flg = False
        
        With rsPR
            
'            If .RecordCount > 0 Then
'                .MoveFirst
'                Do
'
''                    02/07/2009 - don't split by ItemID
''                    If PRDist.JobID = !JobID And _
''                       PRDist.EmployeeID = !EmployeeID And _
''                       ItmID() = !PRItemID And _
''                       PRDist.BillingRate = !BillingRate Then
'
'                    If PRDist.JobID = !JobID And _
'                       PRDist.EmployeeID = !EmployeeID And _
'                       PRDist.BillingRate = !BillingRate Then
'
'                        !Hours = !Hours + PRDist.Hours
'                        .Update
'                        Flg = True
'                        Exit Do
'                    End If
'                    .MoveNext
'                Loop Until .EOF
'            End If
                
            If Flg = False Then
                .AddNew
                !JobID = PRDist.JobID
                !EmployeeID = PRDist.EmployeeID
                !PRItemID = ItmID
                
                ' 2013-03-16 - add round
                !BillingRate = Round(PRDist.BillingRate, 2)
                
                !Hours = PRDist.Hours
                !StartDate = Me.tdbStartDate
                !EndDate = Me.tdbEndDate
                !TSRecID = PRDist.DistID
                .Update
            End If
        
        End With
        If PRDist.GetNext = False Then Exit Do
    
    Loop
        
    Me.lblMsg1 = "Now updating QB Item Codes ..."
    Me.Refresh
    
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue
    
    ' gather the QB SERVICE items that start with PR
    ' record QB List ID in temp RS
    Set ItemQuery = requestMsgSet.AppendItemQueryRq
    
    ' filters
    ItemQuery.ORListQuery.ListFilter.ActiveStatus.SetValue asActiveOnly
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
                            
                ' service items
                If (Not orItemRet.itemServiceRet Is Nothing) Then
                    If (Not orItemRet.itemServiceRet.ORSalesPurchase.SalesOrPurchase Is Nothing) Then
                        
                        ' QB Service item name is the format:
                        ' PR_{EE Last/FirstName}/{Emp#}
                        X = orItemRet.itemServiceRet.Name.GetValue
                        If Mid(X, 1, 3) = "PR_" Then
                            EmpNum = ParseEENum(X)
                            If EmpNum <> 0 Then
                                SQLString = "SELECT * FROM PREmployee WHERE EmployeeNumber = " & EmpNum
                                If PREmployee.GetBySQL(SQLString) = True Then
                                    rsEE.Find "EmployeeID = " & PREmployee.EmployeeID, 0, adSearchForward, 1
                                    If rsEE.EOF Then
                                        rsEE.AddNew
                                        rsEE!EmployeeID = PREmployee.EmployeeID
                                        rsEE!QBItemID = orItemRet.itemServiceRet.ListID.GetValue
                                        rsEE.Update
                                    Else
                                        rsEE!QBItemID = orItemRet.itemServiceRet.ListID.GetValue
                                    End If
                                    rsEE.Update
                                End If
                            End If
                        End If
                    End If
                End If
            
            Next j
                    
itemNxtI:
        Next i
    
    End If
    
    ' add QB Items where necessary for each employee
    rsEE.MoveFirst
    Do
        If rsEE!QBItemID = "" Then
            rsEE!QBItemID = AddQBItem(rsEE!EmployeeID)
            If rsEE!QBItemID = "" Then
                MsgBox "QB Item Add Error: " & rsEE!EmployeeID, vbExclamation
                GoBack
            End If
            rsEE.Update
        Else
            ' MsgBox rsEE!QBItemID & vbCr & rsEE!EmployeeID
        End If
        rsEE.MoveNext
    Loop Until rsEE.EOF
    
    NewQBSession
    
    Me.lblMsg1 = "Now Creating Invoices ...."
    Me.Refresh
    
    If rsPR.RecordCount = 0 Then
        MsgBox "There is no invoice data for this batch!", vbInformation
        GoBack
    End If
    
    ' sort the RS and create the frickin' invoices !!!
    LastJobID = 0
    InvSeqNum = 0
    rsPR.Sort = "JobID, EmployeeID, PRItemID"
    rsPR.MoveFirst
    Do
        
        ' break in job or first job
        If LastJobID = 0 Or rsPR!JobID <> LastJobID Then
        
            ' finish the last invoice
'            If LastJobID <> 0 Then
'                InvoiceFinish
'            End If
        
            If JCJob.GetByID(rsPR!JobID) = False Then
                MsgBox "JobID Not Found: " & rsPR!JobID, vbExclamation
                GoBack
            End If
        
            ' start the new invoice
            Set invoiceAdd = requestMsgSet.AppendInvoiceAddRq
            If JCJob.QBID = "ORIG" Then
                invoiceAdd.CustomerRef.ListID.SetValue JCJob.QBParentID
            Else
                invoiceAdd.CustomerRef.ListID.SetValue JCJob.QBID
            End If
            
            invoiceAdd.ARAccountRef.ListID.SetValue QBIDAR
            invoiceAdd.TemplateRef.ListID.SetValue QBIDTPL
            invoiceAdd.TxnDate.SetValue Me.tdbInvoiceDate.Value
            invoiceAdd.IsPending.SetValue False
            invoiceAdd.IsToBePrinted.SetValue True
            
            ' get that tax code / item from the customer record
            If JCCustomer.GetByID(JCJob.ParentID) = False Then
                MsgBox "Customer record not found for Job #: " & JCJob.JobID, vbExclamation
                GoBack
            End If
                
            If JCCustomer.QBTaxItem = "" Then
                MsgBox "QB Tax Item Not Set: " & JCCustomer.Name, vbExclamation
                GoBack
            End If
                
            If JCCustomer.QBTaxCode = "" Then
                MsgBox "QB Tax Code Not Set: " & JCCustomer.Name, vbExclamation
                GoBack
            End If
                
            invoiceAdd.ItemSalesTaxRef.ListID.SetValue JCCustomer.QBTaxItem
            invoiceAdd.CustomerSalesTaxCodeRef.ListID.SetValue JCCustomer.QBTaxCode
        
            InvSeqNum = InvSeqNum + 1
        
        End If
        LastJobID = rsPR!JobID
    
        ' add the invoice line item
        rsEE.Find "EmployeeID = " & rsPR!EmployeeID, 0, adSearchForward, 1
        If rsEE.EOF Then
            MsgBox "Employee Error: " & rsPR!EmployddID, vbExclamation
            GoBack
        End If
        Set orInvoiceLineAdd1 = invoiceAdd.ORInvoiceLineAddList.Append
        orInvoiceLineAdd1.InvoiceLineAdd.ItemRef.ListID.SetValue rsEE!QBItemID
        
        ' invoice line description
        If PREmployee.GetByID(rsPR!EmployeeID) = False Then
            MsgBox "EmployeeID Not Found: " & rsPR!EmployeeID, vbExclamation
            GoBack
        End If
        
        If Len(PREmployee.LFName) >= 25 Then
            Y = Mid(PREmployee.LFName, 1, 25) & " "
        Else
            Y = PREmployee.LFName & Space(26 - Len(PREmployee.LFName))
        End If
        
        X = Y & _
            Format(rsPR!StartDate, "mm/dd/yy") & " To: " & _
            Format(rsPR!EndDate, "mm/dd/yy") & " "
        
        ' time description
        If rsPR!PRItemID = 99991 Then
        ElseIf rsPR!PRItemID = 99992 Then
            X = X & " Over Time"
        Else
            If PRItem.GetByID(rsPR!PRItemID) = False Then
                MsgBox "PR Item Not Found: " & rsPR!EmployeeID & vbCr & _
                       rsPR!JobID & vbCr & _
                       Format(rsPR!StartDate, "mm/dd/yy") & vbCr & _
                       rsPR!PRItemID, vbExclamation
                GoBack
            End If
            X = X & PRItem.Abbreviation
        End If
        orInvoiceLineAdd1.InvoiceLineAdd.Desc.SetValue X
        
        orInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue rsPR!Hours
        orInvoiceLineAdd1.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue rsPR!BillingRate
        orInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue SuperRound(rsPR!Hours, rsPR!BillingRate)
        orInvoiceLineAdd1.InvoiceLineAdd.ServiceDate.SetValue CLng(rsPR!EndDate)
        
        orInvoiceLineAdd1.InvoiceLineAdd.SalesTaxCodeRef.ListID.SetValue JCCustomer.QBTaxCode
        orInvoiceLineAdd1.InvoiceLineAdd.IsTaxable.SetValue True
        
        rsPR!InvSeqNum = InvSeqNum
        rsPR.Update
        
        rsPR.MoveNext
    
    Loop Until rsPR.EOF
    
    ' process the invoice request batch
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
    Set ResponseList = responseMsgSet.ResponseList
  
    If (ResponseList Is Nothing) Then
        MsgBox "Invoice Add - response is nothing!", vbExclamation
        GoBack
    End If
  
    For i = 0 To ResponseList.Count - 1
        
        Set Response = ResponseList.GetAt(i)
        ' Check the status returned for the response.
        If Response.StatusCode >= 1000 Then
            MsgBox Response.StatusCode & vbCr & _
                   Response.StatusMessage, vbExclamation
            GoTo InvParseNxtI
        End If
        If (Response.Detail Is Nothing) Then GoTo InvParseNxtI
        ResponseType = Response.Type.GetValue
        If ResponseType <> rtInvoiceAddRs Then GoTo InvParseNxtI
        Set invoiceRet = Response.Detail
        If invoiceRet Is Nothing Then
            MsgBox "invoiceRet is nothing ", vbExclamation
            End
        End If
        
        ' update the TxnID back
        rsPR.MoveFirst
        Do
            If rsPR!InvSeqNum = i + 1 Then
                If PRDist.GetByID(rsPR!TSRecID) Then
                    PRDist.QBInvoiceID = invoiceRet.TxnID.GetValue
                    PRDist.Save (Equate.RecPut)
                End If
            End If
            rsPR.MoveNext
        Loop Until rsPR.EOF
        
InvParseNxtI:
    Next i
    
    MsgBox ResponseList.Count & " Invoices have been updated", vbInformation
    
    SessMgr.EndSession
    SessMgr.CloseConnection
    
    Unload Me

End Sub

Private Sub NewQBSession()
        
    SessMgr.EndSession
    SessMgr.BeginSession "", omDontCare
    Set requestMsgSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    requestMsgSet.Attributes.OnError = roeContinue

End Sub

Private Function ItmID() As Long
    If PRDist.DistType = PREquate.DistTypeReg Then
        ItmID = 99991
    ElseIf PRDist.DistType = PREquate.DistTypeOT Then
        ItmID = 99992
    Else
        ItmID = PRDist.EmployerItemID
    End If
End Function
