VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmJCGetQBData 
   Caption         =   "Import QB Customer and Job Data"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8835
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   2655
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
      _cx             =   6800
      _cy             =   4683
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin TDBDate6Ctl.TDBDate tdbStartDate 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   661
      Calendar        =   "frmJCGetQBData.frx":0000
      Caption         =   "frmJCGetQBData.frx":0100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCGetQBData.frx":0178
      Keys            =   "frmJCGetQBData.frx":0196
      Spin            =   "frmJCGetQBData.frx":01F4
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
      Text            =   "02/03/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40212
      CenturyMode     =   0
   End
   Begin VB.CheckBox chkAllData 
      Caption         =   "ALL CUSTOMER / JOB DATA"
      Height          =   375
      Left            =   2783
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   7320
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   7320
      Width           =   2055
   End
   Begin TDBDate6Ctl.TDBDate tdbEndDate 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   2040
      Width           =   3615
      _Version        =   65536
      _ExtentX        =   6376
      _ExtentY        =   661
      Calendar        =   "frmJCGetQBData.frx":021C
      Caption         =   "frmJCGetQBData.frx":031C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmJCGetQBData.frx":0390
      Keys            =   "frmJCGetQBData.frx":03AE
      Spin            =   "frmJCGetQBData.frx":040C
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
      Text            =   "02/03/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40212
      CenturyMode     =   0
   End
   Begin VB.Label lblQBMsg 
      Alignment       =   2  'Center
      Caption         =   "QB Message...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   8160
      Width           =   9015
   End
   Begin VB.Label Label3 
      Caption         =   "JOB STATUS:"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Works best when the company QuickBooks file is already opened"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   1110
      TabIndex        =   4
      Top             =   6120
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OK to import QuickBooks Customer and Job Data?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   930
      TabIndex        =   3
      Top             =   5520
      Width           =   8175
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "frmJCGetQBData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StatusDrop As String
Public rs As New ADODB.Recordset
Dim I, J, K As Long
Dim X, Y, Z As String

' recordset to fill in multi-level QB job levels
Dim rsQB As New ADODB.Recordset

Private Sub Form_Load()

    Me.WindowState = vbNormal
    Me.lblCompanyName = PRCompany.Name
    Me.lblQBMsg = ""
    
    ' keep track of JCJob records that are more than
    ' one level deep - used to fill in JCJob.ParentID
    ' JCJob.ParentID is the ID of the JCCustomer record
    rsQB.CursorLocation = adUseClient
    rsQB.Fields.Append "JobID", adDouble
    rsQB.Open , , adOpenDynamic, adLockOptimistic
    
    rs.CursorLocation = adUseClient
    rs.Fields.Append "Select", adBoolean
    rs.Fields.Append "JobStatus", adDouble
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    StatusDrop = ""
    For I = 0 To 5
        X = ""
        If I = PREquate.qbJobStatus_Awarded Then X = "Awarded"
        If I = PREquate.qbJobStatus_Closed Then X = "Closed"
        If I = PREquate.qbJobStatus_InProgress Then X = "In Progress"
        If I = PREquate.qbJobStatus_None Then X = "None"
        If I = PREquate.qbJobStatus_NotAwarded Then X = "Not Awarded"
        If I = PREquate.qbJobStatus_Pending Then X = "Pending"
        StatusDrop = Trim(StatusDrop) & "|#" & I & ";" & X
            
        rs.AddNew
        rs!Select = True
        rs!JobStatus = I
        rs.Update
    
    Next I
    
    rs.Sort = "JobStatus"
    SetGrid rs, fg
    
    With fg
        .SelectionMode = flexSelectionByRow
        .ColComboList(1) = StatusDrop
        .ColWidth(1) = 2000
    End With
    
    Me.chkAllData = 1
    Me.tdbStartDate.Enabled = False
    Me.tdbEndDate.Enabled = False
    Me.fg.Enabled = False
    Me.tdbStartDate = DateSerial(1980, 1, 1)
    Me.tdbEndDate = Int(Now())
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF7: DelAll
    End Select
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    ' all job are marked to inactive just
    ' just before the update actually starts
    ' to minimize the chance jobs are left inactive
    ' if the update bombs out
    
    QBProcess
    Me.Hide
    
    ' GoBack
    
End Sub

Private Sub QBProcess()

    Me.MousePointer = vbHourglass
    
    If TableExists("JCCustomer", cn) = False Then
        CustomerCreate
    End If
    
    If TableExists("JCJob", cn) = False Then
        JobCreate
    End If
    
    DoCustomerQueryRq "US", 5, 0
    
    ' update JCJob.ParentID for multi-level jobs
    If rsQB.RecordCount > 0 Then
        JobFill
    End If
    
    MsgBox "Import of QB Customer and Job Info Complete", vbInformation, "Balint Windows PR"
    
    Me.MousePointer = vbArrow
    
End Sub

Private Sub DelAll()

    If MsgBox("OK to delete ALL Job and Customer Data?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    SQLString = "DELETE * FROM JCJob"
    cn.Execute SQLString
    
    SQLString = "DELETE * FROM JCCustomer"
    cn.Execute SQLString

    MsgBox "All Cust/Job records deleted", vbInformation

End Sub
Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub chkAllData_Click()
    
    If Me.chkAllData = 1 Then
        Me.tdbStartDate.Enabled = False
        Me.tdbEndDate.Enabled = False
        Me.fg.Enabled = False
    Else
        Me.tdbStartDate = Int(Now() - 30)
        Me.tdbEndDate = Int(Now())
        Me.tdbStartDate.Enabled = True
        Me.tdbEndDate.Enabled = True
        Me.fg.Enabled = True
    End If

End Sub

Public Sub DoCustomerQueryRq(country As String, MajorVersion As Integer, MinorVersion As Integer)
  
'  On Error GoTo Errs
  
    'We want to know if we've begun a session so we can end it if an
    'error sends us to the exception handler.
    
    If QBOpen(Me, Me.lblQBMsg) = False Then
        GoBack
    End If
    
    ' Create the message set request object for the specific version messages.
    Dim requestMsgSet As IMsgSetRequest
    Set requestMsgSet = SessMgr.CreateMsgSetRequest(country, MajorVersion, MinorVersion)
    requestMsgSet.Attributes.OnError = roeContinue
  
    Me.lblQBMsg = "Building Customer Query ... "
    Me.Refresh
    
    BuildCustomerQueryRq requestMsgSet, country
  
    ' Perform the request and obtain a response from QuickBooks.
    Dim responseMsgSet As IMsgSetResponse
    
    Me.lblQBMsg = "Performing QB Data Request ... "
    Me.Refresh
    
    Set responseMsgSet = SessMgr.DoRequests(requestMsgSet)
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
    Me.lblQBMsg = "Parsing QB Data ... "
    Me.Refresh
    
    ParseCustomerQueryRs responseMsgSet, country
  
    Unload frmProgress
  
    SessMgr.EndSession
    SessMgr.CloseConnection
  
    Exit Sub
  
Errs:
    MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error"
  
    ' SampleCodeForm.ErrorMsg.Text = Err.Description
    ' Me.Label1 = Err.Description
  
    ' Close the session and connection with QuickBooks.
    SessMgr.EndSession
    SessMgr.CloseConnection
  
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
  
 ' ***********************
 Exit Sub
  
  
  'Set the elements of ICustomerQuery.
  
  If (country = "US") Then
    ' Set the value of the ICustomerQuery.metaData element.
    customerQuery.metaData.SetValue mdNoMetaData
  
  End If
  If (country = "US") Then
    ' Set the value of the ICustomerQuery.iterator element.
    customerQuery.iterator.SetValue itStart
  
  End If
  
  If (country = "US") Then
    ' Set the value of the ICustomerQuery.iteratorID element.
    customerQuery.iteratorID.SetValue "val"
  
  End If
  ' Only can set one of the OR elements.
  ' We will portray this restriction by using an If/Then/Else.
  Dim orCustomerListQueryORElement1 As String
  orCustomerListQueryORElement1 = "ListID"
  If (orCustomerListQueryORElement1 = "ListID") Then
    ' Set the value of the IORCustomerListQuery.ListIDList element.
    customerQuery.ORCustomerListQuery.ListIDList.Add "val"
  
  ElseIf (orCustomerListQueryORElement1 = "FullName") Then
    ' Set the value of the IORCustomerListQuery.FullNameList element.
    customerQuery.ORCustomerListQuery.FullNameList.Add "val"
  
  ElseIf (orCustomerListQueryORElement1 = "CustomerListFilter") Then
    ' Set the value of the ICustomerListFilter.MaxReturned element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue 10
  
    ' Set the value of the ICustomerListFilter.ActiveStatus element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.ActiveStatus.SetValue asActiveOnly
  
    ' Set the value of the ICustomerListFilter.FromModifiedDate element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.FromModifiedDate.SetValue #12/31/2003 9:35:00 AM#, False
  
    ' Set the value of the ICustomerListFilter.ToModifiedDate element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.ToModifiedDate.SetValue #12/31/2003 9:35:00 AM#, False
  
    ' Only can set one of the OR elements.
    ' We will portray this restriction by using an If/Then/Else.
    Dim orNameFilterORElement2 As String
    orNameFilterORElement2 = "NameFilter"
    If (orNameFilterORElement2 = "NameFilter") Then
      ' Set the value of the INameFilter.MatchCriterion element.
      customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue mcStartsWith
  
      ' Set the value of the INameFilter.Name element.
      customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameFilter.Name.SetValue "val"
  
    ElseIf (orNameFilterORElement2 = "NameRangeFilter") Then
      ' Set the value of the INameRangeFilter.FromName element.
      customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue "val"
  
      ' Set the value of the INameRangeFilter.ToName element.
      customerQuery.ORCustomerListQuery.CustomerListFilter.ORNameFilter.NameRangeFilter.ToName.SetValue "val"
  
    End If
  
    ' Set the value of the ITotalBalanceFilter.Operator element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.TotalBalanceFilter.Operator.SetValue oLessThan
  
    ' Set the value of the ITotalBalanceFilter.Amount element.
    customerQuery.ORCustomerListQuery.CustomerListFilter.TotalBalanceFilter.Amount.SetValue 2#
  
  End If
  
  If (country = "US") Then
    ' Set the value of the ICustomerQuery.IncludeRetElementList element.
    customerQuery.IncludeRetElementList.Add "val"
  
  End If
  
  ' Set the value of the ICustomerQuery.OwnerIDList element.
  customerQuery.OwnerIDList.Add "{22E8C9DC-320B-450d-962A-87CF7246D080}"
  
End Sub
 

Public Sub ParseCustomerQueryRs(responseMsgSet As IMsgSetResponse, country As String)
  
Dim Ct, Recs As Long
Dim CCount As Long
      
  If (responseMsgSet Is Nothing) Then
    MsgBox "responseMsgSet = Nuthin"
    Exit Sub
  End If
  
  Dim ResponseList As IResponseList
  Set ResponseList = responseMsgSet.ResponseList
  
  If (ResponseList Is Nothing) Then
    MsgBox "responseList = Nuthin"
    Exit Sub
  End If
  
    ' set all job active to zero
    ' if deleted from QB it will stay inactive
    ' if still in QB - the active flag will be update
    SQLString = "SELECT * FROM JCJob WHERE Active = 1"
    If JCJob.GetBySQL(SQLString) = True Then
        Do
            JCJob.Active = 0
            JCJob.Save (Equate.RecPut)
            If JCJob.GetNext = False Then Exit Do
        Loop
    End If
  
  ' Go through all of the responses in the list.
  Dim I As Integer
  
  Recs = ResponseList.Count
  
  For I = 0 To ResponseList.Count - 1
    
    Dim Response As IResponse
    Set Response = ResponseList.GetAt(I)
  
    ' Check the status returned for the response.
    If (Response.StatusCode = 0) Then
  
      ' Check to make sure the response is of the type we are expecting.
      If (Not Response.Detail Is Nothing) Then
        Dim ResponseType As Integer
        ResponseType = Response.Type.GetValue
        Dim J As Integer
        ' Check for CustomerQueryRs.
        If (ResponseType = rtCustomerQueryRs) Then
          Dim customerRetList As ICustomerRetList
          Set customerRetList = Response.Detail
          For J = 0 To customerRetList.Count - 1
          CCount = customerRetList.Count
            If J Mod 10 = 1 Then
                Me.lblQBMsg = "Getting Job Info: " & Format(J, "#,###,##0") & " of: " & Format(CCount, "#,###,##0")
                Me.Refresh
            End If
            ParseCustomerRet customerRetList.GetAt(J), country
          Next J
        End If
      End If
    End If
  Next I
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
    Dim J As Integer
    For J = 0 To customerRet.DataExtRetList.Count - 1
      Dim dataExtRet75 As IDataExtRet
      Set dataExtRet75 = customerRet.DataExtRetList.GetAt(J)
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
  
    Next J
  
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
    If frmJCGetQBData.chkAllData = 0 Then
        
        ' filter by status?
        frmJCGetQBData.rs.Find "JobStatus = " & JobStatus65, 0, adSearchForward, 1
        If frmJCGetQBData.rs!Select = False Then Exit Sub
        
        ' filter by date
        If Int(timeModified3) < frmJCGetQBData.tdbStartDate.Value Then Exit Sub
        If Int(timeModified3) > frmJCGetQBData.tdbEndDate.Value Then Exit Sub


    End If

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
        
        JCJob.JobStatus = CByte(JobStatus65)
        
        JCJob.ParentID = JCCustomer.CustomerID
        JCJob.StartDate = timeModified3
        JCJob.QBTaxCode = listID53
        
        If isActive7 = True Then
            JCJob.Active = 1
        Else
            JCJob.Active = 0
        
            ' remove all stock items - if table exists
            If TableExists("InvStock", cn) Then
                SQLString = "DELETE * FROM InvStock WHERE JobID = " & JCJob.JobID
                cn.Execute SQLString
            End If
        
        End If
        
        JCJob.Terms = listID47
                
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
            
            ' remove all stock items - if table exists
            If TableExists("InvStock", cn) Then
                SQLString = "DELETE * FROM InvStock WHERE JobID = " & JCJob.JobID
                cn.Execute SQLString
            End If
        
        End If
        
        JCJob.Terms = listID47
                
        JCJob.Save (Equate.RecPut)
    
    End If

    ' *** try ShipTo first ***
    ' auto update City if not assigned
    AssignCity JCJob.ShipCity, JCJob.ShipState
    AssignCity JCJob.BillCity, JCJob.BillState
    
End Sub
 
Private Sub AssignCity(ByVal Cty As String, ByVal Ste As String)
    
    If JCJob.CityID <> 0 Then Exit Sub
    If IsNull(Cty) Then Exit Sub
    If Cty = "" Then Exit Sub
    If IsNull(Ste) Then Exit Sub
    If Ste = "" Then Exit Sub
    
    SQLString = "SELECT * FROM PRCity"
    If PRCity.GetBySQL(SQLString) = True Then
        Do
            If UCase(Trim(Cty)) = UCase(Trim(PRCity.CityName)) Then
                If PRState.GetByID(PRCity.StateID) = True Then
                    If UCase(Trim(Ste)) = UCase(Trim(PRState.StateAbbrev)) Then
                        JCJob.CityID = PRCity.CityID
                        JCJob.Save (Equate.RecPut)
                        Exit Do
                    End If
                End If
            End If
            If PRCity.GetNext = False Then Exit Do
        Loop
    End If

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

