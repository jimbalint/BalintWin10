VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   4815
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   8535
      _cx             =   15055
      _cy             =   8493
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblMsg1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' for chart of account query
Dim SessMgr As New QBSessionManager

' Dim RepQ As qbfc13lib.ICustomDetailReportQuery
Dim RepQ As QBFC13Lib.ICustomDetailReportQuery

' Dim RepQ2 As qbfc13lib.IGeneralDetailReportQuery
Dim RepQ2 As QBFC13Lib.IGeneralDetailReportQuery

Dim RequestSet As QBFC13Lib.IMsgSetRequest
Dim ResponseSet As QBFC13Lib.IMsgSetResponse
Dim qResponse As QBFC13Lib.IResponse
Dim RepRet As QBFC13Lib.IReportRet
Dim orReportData As QBFC13Lib.IORReportData
Dim ResponseList As IResponseList
Dim Response As IResponse

' for chart of account query
Dim AccQ As QBFC13Lib.IAccountQuery
Dim RetList As IAccountRetList
Dim ItemRet As QBFC13Lib.IAccountRet

Dim rsQB As New ADODB.Recordset

Dim nRequest As Long
Dim index As Long
Dim Ct As Long

Dim i, j, k, l, m As Long
Dim x, y, z As String
Dim IconType As Integer

Dim JnlAddReq As IJournalEntryAdd
Dim orJournalLine1 As IORJournalLine

Private Sub Form_Load()

    ' get the QB Account
    SQLString = "SELECT * FROM QBAccount"
    rsInit SQLString, cn, rsQB
    If rsQB.RecordCount = 0 Then
        MsgBox "No QBAccount records found ..."
        End
    End If
    rsQB.MoveFirst
    
    frmProgress.lblMsg2 = "Create MsgSetRequest .... "
    frmProgress.Show
    Set RequestSet = SessMgr.CreateMsgSetRequest("US", 5, 0)
    RequestSet.Attributes.OnError = roeContinue
    
    Set JnlAddReq = RequestSet.AppendJournalEntryAddRq
    
    JnlAddReq.TxnDate.SetValue Int(Now())
    JnlAddReq.RefNumber.SetValue "111"
    JnlAddReq.Memo.SetValue "Memo..."
    ' JnlAddReq.IsAdjustment.SetValue True
        
    
    ' get the QB Job
    If JCJob.GetByID(976) = False Then
        MsgBox "Job NF"
        End
    End If
    
    ' get the DR account - Misc Exp
    If QBAccount.GetByID(1272) = False Then
        MsgBox "Acct NF"
        End
    End If
    
    ' debit info - Job A $75
    Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
    orJournalLine1.JournalDebitLine.TxnLineID.SetValue 1
    orJournalLine1.JournalDebitLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
    orJournalLine1.JournalDebitLine.AccountRef.ListID.SetValue Trim(QBAccount.QBID)
    orJournalLine1.JournalDebitLine.Amount.SetValue 75#
    orJournalLine1.JournalDebitLine.Memo.SetValue "DebitMemo"
    orJournalLine1.JournalDebitLine.EntityRef.ListID.SetValue Trim(JCJob.QBParentID)
    
    ' get the QB Job
    If JCJob.GetByID(977) = False Then
        MsgBox "Job NF"
        End
    End If
    
    ' debit info - Job B $25
    Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
    orJournalLine1.JournalDebitLine.TxnLineID.SetValue 2
    orJournalLine1.JournalDebitLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
    orJournalLine1.JournalDebitLine.AccountRef.ListID.SetValue Trim(QBAccount.QBID)
    orJournalLine1.JournalDebitLine.Amount.SetValue 25#
    orJournalLine1.JournalDebitLine.Memo.SetValue "DebitMemo"
    orJournalLine1.JournalDebitLine.EntityRef.ListID.SetValue Trim(JCJob.QBParentID)
    
    ' get the CR account
    If QBAccount.GetByID(1273) = False Then
        MsgBox "Acct NF"
        End
    End If
    
    ' credit info
    Set orJournalLine1 = JnlAddReq.ORJournalLineList.Append
    orJournalLine1.JournalCreditLine.TxnLineID.SetValue 3
    orJournalLine1.JournalCreditLine.AccountRef.FullName.SetValue Trim(QBAccount.Name)
    orJournalLine1.JournalCreditLine.AccountRef.ListID.SetValue Trim(QBAccount.QBID)
    orJournalLine1.JournalCreditLine.Amount.SetValue 100#
    orJournalLine1.JournalCreditLine.Memo.SetValue "CreditMemo"
    ' orJournalLine1.JournalCreditLine.EntityRef.ListID.SetValue Trim(JCJob.QBID)
    
    frmProgress.Caption = "Opening QB Session"
    frmProgress.lblMsg1 = ""
    frmProgress.lblMsg2 = "Now opening QuickBooks Session .... "
    frmProgress.Refresh
    SessMgr.OpenConnection2 "", "Balint Accounting", ctLocalQBD
    
    frmProgress.Caption = "Begin QB Session"
    frmProgress.lblMsg2 = "Now Beginning QuickBooks Session .... "
    frmProgress.Refresh
    SessMgr.BeginSession "", omDontCare
    
    frmProgress.Caption = "Begin QB Session"
    frmProgress.lblMsg2 = "Do Requests ...."
    frmProgress.Refresh
    Set ResponseSet = SessMgr.DoRequests(RequestSet)
    
    If ResponseSet Is Nothing Then
        MsgBox "ResponseSet = nothing "
        End
    End If
    
    Set ResponseList = ResponseSet.ResponseList
    If (ResponseList Is Nothing) Then
        MsgBox "ResponseList = nothing "
        End
    End If
    
    For i = 0 To ResponseList.Count - 1
        
        Set Response = ResponseList.GetAt(i)
 
        ' Check the status returned for the response.
        If (Response.StatusCode = 0) Then
 
            ' Check to make sure the response is of the type we are expecting.
            If (Not Response.Detail Is Nothing) Then
                Dim ResponseType As Integer
                ResponseType = Response.Type.GetValue
                Dim j As Integer
                ' Check for JournalEntryAddRs.
                If (ResponseType = rtJournalEntryAddRs) Then
'                    Dim journalEntryRet As IJournalEntryRet
'                    Set journalEntryRet = response.Detail
'                    ParseJournalEntryRet journalEntryRet, country
                End If
            End If
        Else
            MsgBox Response.StatusCode & vbCr & Response.StatusMessage & vbCr & Response.StatusSeverity
        End If
    Next i
    
    
    frmProgress.Caption = "Begin QB Session"
    frmProgress.lblMsg2 = "Close Connection..."
    frmProgress.Refresh
    SessMgr.EndSession
    SessMgr.CloseConnection
    
    End

End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub QBGetAccounts()
    
    If TableExists("QBAccount", cn) = True Then
        cn.Execute "DROP TABLE QBAccount"
    End If
    
    QBAccountCreate
    
    SQLString = "SELECT * FROM QBAccount"
    rsInit SQLString, cn, rsQB
    
    frmProgress.Caption = "Opening QB Session"
    frmProgress.lblMsg1 = ""
    frmProgress.lblMsg2 = "Now opening QuickBooks Session .... "
    frmProgress.Show
    
    SessMgr.OpenConnection2 "", "Balint Accounting", ctLocalQBD
    
    frmProgress.Caption = "Begin QB Session"
    frmProgress.lblMsg2 = "Now Beginning QuickBooks Session .... "
    frmProgress.Show
    
    SessMgr.BeginSession "", omDontCare
    
    frmProgress.Caption = "Get QB Chart of Accounts"
    frmProgress.lblMsg2 = "Now Getting QB Chart of Accounts"
    frmProgress.Show
    
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
          End
       End If
    
    End If

    Set RetList = qResponse.Detail
    
    If RetList Is Nothing Then Exit Sub   ' no accounts ???
    
    j = RetList.Count
    
    k = 0
        
    For i = 0 To j - 1
                
        Set ItemRet = RetList.GetAt(i)
        If (Not ItemRet Is Nothing) Then
            If (Not ItemRet.Name Is Nothing) Then
                k = k + 1
                
                rsQB.AddNew
                rsQB!QBID = ItemRet.ListID.GetValue
                rsQB!Name = ItemRet.Name.GetValue
                
                If Not (ItemRet.Desc Is Nothing) Then
                    rsQB!Description = ItemRet.Desc.GetValue
                Else
                    rsQB!Description = ""
                End If
                rsQB!AccountType = ItemRet.AccountType.GetAsString
              
                rsQB.Update
              
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
    
End Sub


