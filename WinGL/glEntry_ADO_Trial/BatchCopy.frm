VERSION 5.00
Begin VB.Form BatchCopy 
   Caption         =   "COPY BATCH"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BatchCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleMode       =   0  'User
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbJournal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   " Pick Journal Source"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.ComboBox cmbPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "BatchCopy.frx":030A
      Left            =   6000
      List            =   "BatchCopy.frx":030C
      TabIndex        =   2
      Text            =   "cmbPeriod"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.ComboBox cmbFiscalYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Text            =   "cmbFiscalYear"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5790
      TabIndex        =   4
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label txtCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label7 
      Caption         =   "JOURNAL SOURCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label txtCredits 
      Caption         =   "CREDITS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label lblUpdated 
      Caption         =   "Update User and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblCreated 
      Caption         =   "Created User and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label txtDebits 
      Caption         =   "DEBITS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label txtRecord 
      Caption         =   "RECORDS IN BATCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "FISCAL PERIOD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "FISCAL YEAR:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblBatchNumber 
      Caption         =   "BATCH NUMBER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "BatchCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BatchNumberC As Long
Public BatchFrom As Long

Public userOK As Boolean
Dim jou As New XArrayDB

Dim SessMgr As New QBSessionManager

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

Dim xdbAccts As New XArrayDB

Dim nRequest As Long
Dim index As Long
Dim ct As Long

Dim I, j, k, l As Long
Dim x, Y, z As String

Dim GLHAccount As Long
Dim GLHFiscalYear As Integer
Dim GLHPeriod As Byte
Dim GLHBatchNumber As Long
Dim GLHAmount As Currency
Dim GLHReference As String
Dim GLHDescription As String
Dim GLHSourceCode As Byte
Dim GLHJournalSource As Byte
Dim GLHHistType As String
Dim GLHUpdateFlag As Boolean

Dim RecCount As Long
Dim RefNum As String
Dim Desc As String
Dim TotalAmount As Currency

Dim TotalDebits As Currency
Dim TotalCredits As Currency
Dim RecordCount As Long

Dim IconType As Integer

Dim BatDebits As Currency
Dim BatCredits As Currency
Dim BatBatchNumber As Long
Dim BatRecords As Long


Private Sub cmbFiscalYear_Click()
    
Dim I As Integer
Dim v As Variant
Dim FY As Integer

    Me.cmbPeriod.Clear
    FY = CInt(cmbFiscalYear)
      
    If GLCompany.FirstPeriod = 1 Then
       v = DateSerial(FY, GLCompany.FirstPeriod, 1)
    Else
       v = DateSerial(FY - 1, GLCompany.FirstPeriod, 1)
    End If

    cmbPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    
    For I = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbPeriod.AddItem "Pd. #:" & I + 1 & " - " & Format(v, "mmmm-yyyy")
    Next I
    
    cmbPeriod.ListIndex = 0
    
'    cmbPeriod.Clear
'    Dim ndx, fy As Integer
'
'    fy = CInt(cmbFiscalYear)
'    For ndx = 1 To glcompany.NumberPds
'        cmbPeriod.AddItem glcompany.MonthName(ndx, fy)
'    Next ndx
'    cmbPeriod.ListIndex = bat.period - 1

End Sub

Private Sub CmdExit_Click()
    Response = False
    Me.Hide
End Sub


Private Sub cmdOK_Click()

Dim x As String
Dim trs As New ADODB.Recordset
Dim trs2 As New ADODB.Recordset
Dim SQLStr As String
Dim TID
Dim HRecCount As Long
Dim BudgetFlag As Boolean
    
Dim glPd As Byte
Dim glJS As Byte
    
    BudgetFlag = False
    
'    On Error GoTo glErr
    
    If cmbJournal.ListIndex = -1 Then
       MsgBox "Must pick Journal Source !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       cmbJournal.SetFocus
       Exit Sub
    End If
    
    If Me.cmbPeriod.ListIndex = -1 Then
       MsgBox "Must pick period !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       cmbPeriod.SetFocus
       Exit Sub
    End If
    
    If Me.cmbFiscalYear.ListIndex = -1 Then
       MsgBox "Must pick fiscal year !!!", vbExclamation + vbOKOnly, "Windows GL Data Entry"
       Me.cmbFiscalYear.SetFocus
       Exit Sub
    End If
       
    glPd = cmbPeriod.ListIndex + 1
    glJS = cmbJournal.ItemData(cmbJournal.ListIndex)
       
'    x = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
'    If Not CNOpen(x, Password) Then End
    
    ' open an empty record set for new data
    SQLStr = "SELECT * FROM GLHistory WHERE GLHistory.FiscalYear = 9999"
    rsInit SQLStr, cn, trs2
    
    ' open the batch to copy from
    SQLStr = "SELECT * FROM GLHistory WHERE GLHistory.BatchNumber = " & BatchFrom & _
             " ORDER BY GLHistory.PostDate"
    
    rsInit SQLStr, cn, trs
 
    trs.MoveFirst
    
    Do
    
        trs2.AddNew

        ' assign values
        trs2.Fields("FiscalYear") = CInt(Me.cmbFiscalYear)
        trs2.Fields("Period") = glPd
        trs2.Fields("BatchNumber") = BatchNumberC
        
        ' Budget
        If trs!HisType = "B" Then
            trs2.Fields("JournalSource") = glJS + 100
            BudgetFlag = True
        Else
            trs2.Fields("JournalSource") = glJS
        End If
        
        ' assign the post date
        HRecCount = HRecCount + 1
        trs2.Fields("PostDate") = DateSerial(Year(Now()), Month(Now()), Day(Now())) - 1 + _
                        TimeSerial(0, 0, HRecCount)
        
        ' copy values
        trs2.Fields("Account") = trs!Account
        trs2.Fields("Amount") = trs!Amount
        trs2.Fields("Reference") = trs!Reference
        trs2.Fields("Description") = trs!Description
        trs2.Fields("SourceCode") = trs!SourceCode
        trs2.Fields("HisType") = trs!HisType
        trs2.Fields("UpdateFlag") = trs!UpdateFlag

        trs2.Update
    
        trs.MoveNext
        If trs.EOF Then Exit Do
    
    Loop
    
    trs.Close
    trs2.Close
    
    ' ******** update the batch record
    If GLBatch.GetBatch(BatchNumberC) = False Then
        MsgBox "GL Batch NF?: " & BatchNumberC
        GoBack
    End If
    
    GLBatch.Debits = BatDebits
    GLBatch.Credits = BatCredits
    GLBatch.Records = BatRecords
    GLBatch.Created = Now()
    GLBatch.CreateUser = GLUser.ID
    GLBatch.FiscalYear = CLng(Me.cmbFiscalYear)
    GLBatch.Period = glPd
    
    If Not BudgetFlag Then
        GLBatch.JournalSource = glJS
    Else
        GLBatch.JournalSource = glJS + 100
    End If
    
    GLBatch.Updated = Now()
    GLBatch.UpdateUser = GLUser.ID
    GLBatch.Save (Equate.RecPut)
    
    ' ******** update the batch record
    
    ' update GLAmount
    x = "\Balint\GLUtil.exe " & _
        "SysFile=\Balint\Data\GLSystem.mdb " & _
        "UserID=" & GLUser.ID & " " & _
        "BackName=\Balint\GLEntry.exe " & _
        "ProgName=UpdateB " & _
        "MenuName=" & MenuName & _
        "Batch=" & BatchNumberC

    ' database password if required
    If Password <> "" Then
       x = x & " dbPWd=" & Password
    End If
        
    If Not TestMode Then TID = Shell(x, vbMaximizedFocus)

    Unload Me
    End
    
    Exit Sub
    
glErr:
    MsgBox Error(Err.Number)
End Sub

Public Sub Init()

Dim ndx As Long
Dim CurFY As Integer
    
    userOK = False
    
    If GLBatch.GetBatch(BatchFrom) = False Then
        MsgBox "Batch NF?: " & BatchFrom, vbExclamation
        GoBack
    End If
    
    txtCompanyName = GLCompany.Name
    lblBatchNumber = "Copy From glbatch.ch # " & GLBatch.BatchNumber
    lblCreated = "Created by " & GLBatch.CreateUser & " on " & ShowDate(GLBatch.Created)
    txtRecord = "RECORD COUNT = " & CStr(GLBatch.RecCt)
    txtDebits = "DEBITS = " & Format(GLBatch.Debits, "Currency")
    txtCredits = "CREDITS = " & Format(GLBatch.Credits, "Currency")
    
    ' store the GLglbatch.ch record values
    BatDebits = GLBatch.Debits
    BatCredits = GLBatch.Credits
    BatBatchNumber = GLBatch.BatchNumber
    BatRecords = GLBatch.Records
 
'    For ndx = glcompany.FirstFiscalYear To Year(Now) + 1
'        cmbFiscalYear.AddItem ndx
'    Next ndx
'
'    'if bat.fiscalYear=0 then
'    cmbFiscalYear = bat.fiscalYear
    
    CurFY = Int(GLCompany.LastClose / 10 ^ 4)
    If Int(GLCompany.LastClose / 100) Mod 100 <> 1 Then CurFY = CurFY + 1
    If CurFY < 1990 Or CurFY > 2020 Then CurFY = Year(Now())
    
    For ndx = CurFY + 1 To CurFY - 5 Step -1
        cmbFiscalYear.AddItem ndx
    Next ndx
    cmbFiscalYear.ListIndex = 1
 
'    For ndx = 1 To glcompany.NumberPds
'        cmbPeriod.AddItem glcompany.MonthName(ndx, bat.fiscalYear)
'        cmbPeriod.AddItem glcompany.MonthName(ndx, CurFY)
'    Next ndx
''    cmbPeriod.ListIndex = bat.period - 1
'    cmbPeriod.ListIndex = 0
    
    SQLString = " SELECT * FROM GLJournal ORDER BY JournalSource "
    ndx = 0
    If GLJournal.GetBySQL(SQLString) = True Then
        Do
            ndx = ndx + 1
            cmbJournal.AddItem (GLJournal.JournalSource & " - " & GLJournal.JournalName)
            cmbJournal.ItemData(cmbJournal.NewIndex) = GLJournal.JournalSource
'           ?????
'            If jou.Value(ndx, 0) = bat.JournalSource Then
'                cmbJournal.ListIndex = ndx - 1
'            End If
            
            If GLJournal.GetNext = False Then Exit Do
        Loop
    End If
    
    Response = False

End Sub


Private Sub Form_Load()

'    Set jou = xFactory.GetJournals(FileName)
'    Set JournalList.Array = jou
'    JournalList.Columns(0).Width = 500
'    JournalList.Columns(1).Width = 3500
     
     Response = False

     ' hide the QuickBook fields by default
     ' QBShow False

End Sub



Private Function LastDay(ByVal yr As Integer, ByVal Mo As Byte) As Variant

    ' add one month
    If Mo = 12 Then
       yr = yr + 1
       Mo = 1
    Else
       Mo = Mo + 1
    End If
    
    ' subtract one day
    LastDay = DateAdd("d", -1, DateSerial(yr, Mo, 1))

End Function


Private Function GetNumber(ByVal InString As String) As Long

' return a long from the digits at the beginning of a string

Dim x1, x2 As String
Dim Ln, i1, i2 As Long

    GetNumber = 0
    If IsNull(InString) Then Exit Function
       
    x2 = ""
    Ln = Len(InString)
    If Ln = 0 Then Exit Function
    i1 = 0
    
    Do
       i1 = i1 + 1
       If i1 > Ln Then Exit Do
       x1 = Mid(InString, i1, 1)
       If InStr(1, "0123456789", x1, vbTextCompare) = 0 Then Exit Do
       x2 = x2 & Mid(InString, i1, 1)
    Loop
    
    If IsNumeric(x2) Then GetNumber = CLng(x2)

End Function


