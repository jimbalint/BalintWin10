VERSION 5.00
Begin VB.Form frmQBAccts 
   Caption         =   "Get QB Chart of Accounts"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   5865
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   6735
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   6735
   End
   Begin VB.Label lblMsgA 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   6495
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
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmQBAccts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RepQ As QBFC13Lib.ICustomDetailReportQuery
Dim RepQ2 As QBFC13Lib.IGeneralDetailReportQuery

Dim RequestSet As QBFC13Lib.IMsgSetRequest
Dim ResponseSet As QBFC13Lib.IMsgSetResponse
Dim qResponse As QBFC13Lib.IResponse
Dim RepRet As QBFC13Lib.IReportRet
Dim orReportData As QBFC13Lib.IORReportData

' for chart of account query
Dim AccQ As QBFC13Lib.IAccountQuery
Dim RetList As IAccountRetList
Dim ItemRet As QBFC13Lib.IAccountRet

Dim rsQB As New ADODB.Recordset

Dim nRequest As Long
Dim index As Long
Dim Ct As Long

Dim i, j, k, l, m As Long
Dim X, Y, Z As String
Dim IconType As Integer
Dim QBCount As Long

' **** get vendor variables ***
Dim Response As IResponse
Dim VendorQuery As IVendorQuery
Dim ResponseType As Integer
Dim ResponseList As IResponseList
Dim VendorRetList As IVendorRetList
Dim VendorRet As IVendorRet
' **** get vendor variables ***

' **** get template variables ***
Dim templateQuery As ITemplateQuery
Dim templateRetList As ITemplateRetList
Dim templateRet As ITemplateRet
' **** get template variables ***

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    Me.lblMsgA = "OK to load QB chart of accounts?" & vbCr & vbCr & _
                 "Works Best when the QB client file is opened"
                 
    Me.lblMsg1 = ""
    Me.lblMsg2 = ""
                 
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
    
Dim fString As String
Dim qbFlag As Boolean
Dim QBID As String
    
    Me.MousePointer = vbHourglass
    
    Me.lblMsg2 = "Now clearing existing records ...."
    Me.Refresh
    
    If TableExists("QBAccount", cn) = False Then
        QBAccountCreate
    End If
    
    ' ========================================================
    
    If QBOpen(Me, Me.lblMsg2) = False Then
        Unload Me
        GoBack
    End If
    
    ' ============================================================
    
    Me.lblMsg2 = "Now clearing existing data ..."
    Me.Refresh
    
'    *** retain the records - just update the fields ***
'    SQLString = "DELETE * FROM QBAccount"
'    cn.Execute SQLString
    
    SQLString = "SELECT * FROM QBAccount"
    rsInit SQLString, cn, rsQB
    
    Me.lblMsg2 = "Now Getting QB Chart of Accounts"
    Me.Refresh
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    
    Set AccQ = RequestSet.AppendAccountQueryRq
    
'    Set AccQ = RequestSet.AppendAccountQueryRq.ORAccountListQuery.FullNameList
    
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
    Set qResponse = ResponseSet.ResponseList.GetAt(nRequest)

    ' check for errors
    If qResponse.StatusCode <> 0 Then
       
       If qResponse.StatusCode <= 499 Then
          IconType = vbInformation
       ElseIf qResponse.StatusCode <= 999 Then
          IconType = vbExclamation
       Else
          IconType = vbCritical
       End If
       
       MsgBox qResponse.StatusMessage & vbCrLf & _
              "Status Code: " & qResponse.StatusCode, IconType
              
       If qResponse.StatusCode >= 1000 Then  ' exit completely
          SessMgr.EndSession
          SessMgr.CloseConnection
          GoBack
       End If
    
    End If

    Set RetList = qResponse.Detail
    
    If RetList Is Nothing Then Exit Sub   ' no accounts ???
    
    QBCount = 0
    j = RetList.Count
    
    k = 0
        
    For i = 0 To j - 1
                
        Set ItemRet = RetList.GetAt(i)
        If (Not ItemRet Is Nothing) Then
            If (Not ItemRet.Name Is Nothing) Then
                k = k + 1
                
                ' rsqb.Find bombs out .... WTF ????
                qbFlag = False
                QBID = Trim(ItemRet.ListID.GetValue & "")
                If rsQB.RecordCount > 0 Then
                    rsQB.MoveFirst
                    Do
                        QBID = Trim(ItemRet.ListID.GetValue & "")
                        If Trim(rsQB!QBID & "") = QBID Then
                            qbFlag = True
                            Exit Do
                        End If
                        rsQB.MoveNext
                    Loop Until rsQB.EOF
                End If
                If qbFlag = False Then
                    rsQB.AddNew
                    rsQB!QBID = QBID
                    rsQB.Update
                End If
                      
                rsQB!Name = ItemRet.Name.GetValue
              
                If Not (ItemRet.Desc Is Nothing) Then
                    rsQB!Description = ItemRet.Desc.GetValue
                Else
                    rsQB!Description = ""
                End If
                rsQB!AccountType = ItemRet.AccountType.GetAsString
                
                ' 2017-05-24
                ' update the AccountNumber field also !!!
                ' needed for sales tax pct update
            
                rsQB.Update
              
                QBCount = QBCount + 1
                Me.lblMsg1 = Format(QBCount, "#,###,##0") & " " & rsQB!Name
                Me.Refresh
                                  
'                ' assign the account number
'                xdbAccts(k, 4) = CLng(GetNumber(xdbAccts(k, 1)))    ' from the QB account description
'
'                ' use the QB acct number if it is there
'                If Not (ItemRet.AccountNumber Is Nothing) Then
'                    X = ItemRet.AccountNumber.GetValue
'                    If IsNumeric(X) Then
'                        xdbAccts(k, 4) = CLng(ItemRet.AccountNumber.GetValue)
'                    End If
'                End If
            
            End If
        End If
    Next i

    Set ItemRet = Nothing
    Set RetList = Nothing
    Set qResponse = Nothing
    Set ResponseSet = Nothing
    Set RequestSet = Nothing

    ' ***************************************************************************************
    ' *** get the vendors also ***
    Me.lblMsg2 = "Now Getting QB Vendors"
    Me.Refresh
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    RequestSet.Attributes.OnError = roeContinue
    Set VendorQuery = RequestSet.AppendVendorQueryRq
'    VendorQuery.metaData.SetValue mdNoMetaData
'    VendorQuery.iterator.SetValue itStart
'    VendorQuery.iteratorID.SetEmpty
'    VendorQuery.ORVendorListQuery.FullNameList.Add ""
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
    Set ResponseList = ResponseSet.ResponseList
    If ResponseList Is Nothing Then
        ' MsgBox "No Vendor Info Found", vbExclamation
        GoTo GetTemplates
    End If
    
    For i = 0 To ResponseList.Count - 1
        Set Response = ResponseList.GetAt(i)
        
'    ' check for errors
'    If qResponse.StatusCode <> 0 Then
'
'       If qResponse.StatusCode <= 499 Then
'          IconType = vbInformation
'       ElseIf qResponse.StatusCode <= 999 Then
'          IconType = vbExclamation
'       Else
'          IconType = vbCritical
'       End If
'
'       MsgBox qResponse.StatusMessage & vbCrLf & _
'              "Status Code: " & qResponse.StatusCode, IconType
'
'       If qResponse.StatusCode >= 1000 Then  ' exit completely
'          SessMgr.EndSession
'          SessMgr.CloseConnection
'          End
'       End If
'
'    End If
        
        If Response.StatusCode = 0 Then
            If (Not Response.Detail Is Nothing) Then
                ResponseType = Response.Type.GetValue
                If ResponseType = rtVendorQueryRs Then
                    Set VendorRetList = Response.Detail
                    For j = 0 To VendorRetList.Count - 1
                        Set VendorRet = VendorRetList.GetAt(j)
                        
                        If (Not VendorRet.ListID Is Nothing) Then
                        
                            ' rsqb.Find bombs out .... WTF ????
                            qbFlag = False
                            QBID = Trim(VendorRet.ListID.GetValue & "")
                            If rsQB.RecordCount > 0 Then
                                rsQB.MoveFirst
                                Do
                                    QBID = Trim(VendorRet.ListID.GetValue & "")
                                    If Trim(rsQB!QBID & "") = QBID Then
                                        qbFlag = True
                                        Exit Do
                                    End If
                                    rsQB.MoveNext
                                Loop Until rsQB.EOF
                            End If
                            If qbFlag = False Then
                                rsQB.AddNew
                                rsQB!QBID = QBID
                                rsQB.Update
                            End If
                        
                            rsQB!Name = VendorRet.Name.GetValue
                            rsQB!Description = ""
                            rsQB!AccountType = "VENDOR"
                            rsQB.Update
                        
                        End If
                    
                    Next j
                End If
            End If
        End If
    Next i
    
    Set ItemRet = Nothing
    Set RetList = Nothing
    Set qResponse = Nothing
    Set ResponseSet = Nothing
    Set RequestSet = Nothing
    
    ' ***********************************************************************************
    ' get Invoice Templates
GetTemplates:
    
    Me.lblMsg2 = "Now getting invoice templates..."
    Me.Refresh
    
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    RequestSet.Attributes.OnError = roeContinue
    Set templateQuery = RequestSet.AppendTemplateQueryRq
    Set ResponseSet = SessMgr.DoRequests(RequestSet)

    If ResponseSet Is Nothing Then
        ' MsgBox "No Templates found"
        GoTo CloseForm
    End If
    
    Set ResponseList = ResponseSet.ResponseList
        
    If ResponseList Is Nothing Then
        ' MsgBox "No Templates found"
        GoTo CloseForm
    End If
        
    For i = 0 To ResponseList.Count - 1
        
        Set Response = ResponseList.GetAt(i)
  
        ' Check the status returned for the response.
        If (Response.StatusCode <> 0) Then GoTo tplNxtI
  
        ' Check to make sure the response is of the type we are expecting.
        If (Response.Detail Is Nothing) Then GoTo tplNxtI
        ResponseType = Response.Type.GetValue
        If (ResponseType <> rtTemplateQueryRs) Then GoTo tplNxtI
        
        Set templateRetList = Response.Detail
        For j = 0 To templateRetList.Count - 1
            Set templateRet = templateRetList.GetAt(j)
            If templateRet Is Nothing Then GoTo tplNxtJ
            If templateRet.IsActive Is Nothing Then GoTo tplNxtJ
            If templateRet.IsActive.GetValue = False Then GoTo tplNxtJ
            
            ' rsqb.Find bombs out .... WTF ????
            qbFlag = False
            QBID = Trim(templateRet.ListID.GetValue & "")
            If rsQB.RecordCount > 0 Then
                rsQB.MoveFirst
                Do
                    QBID = Trim(templateRet.ListID.GetValue & "")
                    If Trim(rsQB!QBID & "") = QBID Then
                        qbFlag = True
                        Exit Do
                    End If
                    rsQB.MoveNext
                Loop Until rsQB.EOF
            End If
            If qbFlag = False Then
                rsQB.AddNew
                rsQB!QBID = QBID
                rsQB.Update
            End If

            rsQB!Name = templateRet.Name.GetValue
            rsQB!Description = templateRet.Type.GetAsString
            rsQB!AccountType = "TEMPLATE"
            rsQB.Update

tplNxtJ:
        Next j
          
tplNxtI:
    Next i
        
    
CloseForm:
    Me.lblMsg1 = ""
    Me.lblMsg2 = "Job Complete ..."
    Me.Refresh
    MsgBox Format(QBCount, "#,###,##0") & " QB Accounts have been loaded", vbOKOnly + vbInformation, "Load QB Chart of Accounts"

    SessMgr.EndSession
    SessMgr.CloseConnection

    Me.MousePointer = vbArrow

    Unload Me

End Sub


