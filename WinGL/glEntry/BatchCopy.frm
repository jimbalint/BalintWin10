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
      Left            =   3030
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

Dim bat As New rBatch
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

Dim i, j, k, l As Long
Dim x, y, z As String

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
    
Dim i As Integer
Dim v As Variant
Dim fy As Integer

    Me.cmbPeriod.Clear
    fy = CInt(cmbFiscalYear)
      
    If com.FirstPeriod = 1 Then
       v = DateSerial(fy, com.FirstPeriod, 1)
    Else
       v = DateSerial(fy - 1, com.FirstPeriod, 1)
    End If

    cmbPeriod.AddItem "Pd. #:1" & " - " & Format(v, "mmmm-yyyy")
    
    For i = 1 To 11
        v = DateSerial(Year(v), Month(v) + 1, 1)
        cmbPeriod.AddItem "Pd. #:" & i + 1 & " - " & Format(v, "mmmm-yyyy")
    Next i
    
    cmbPeriod.ListIndex = 0
    
'    cmbPeriod.Clear
'    Dim ndx, fy As Integer
'
'    fy = CInt(cmbFiscalYear)
'    For ndx = 1 To com.NumberPds
'        cmbPeriod.AddItem com.MonthName(ndx, fy)
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
    
    ' On Error GoTo glErr
    
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
    glJS = jou.Value(cmbJournal.ListIndex + 1, 0)
       
    If BalintFolder = "" Then
        x = Mid(App.Path, 1, 2) & Mid(com.FileName, 3, Len(com.FileName) - 2)
    Else
        x = BalintFolder & "\Data\" & mdbName(com.FileName)
    End If
    If Not CNOpen(x, Password) Then End
    
    ' open an empty record set for new data
    SQLStr = "SELECT * FROM GLHistory WHERE GLHistory.FiscalYear = 9999"
    rsInit SQLStr, Cn, trs2
    
    ' open the batch to copy from
    SQLStr = "SELECT * FROM GLHistory WHERE GLHistory.BatchNumber = " & BatchFrom & _
             " ORDER BY GLHistory.PostDate"
    
    rsInit SQLStr, Cn, trs
 
    If trs.RecordCount = 0 Then
        MsgBox "There are no history records in Batch #" & BatchFrom, vbExclamation
        Unload Me
    End If
    
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
    Cn.Close
    
    ' ******** update the batch record
    bat.GetBatch BatchNumberC, FileName
    
    bat.debits = BatDebits
    bat.credits = BatCredits
    bat.nRecords = BatRecords
    bat.Created = Now()
    bat.createUser = curUser
    bat.fiscalYear = CLng(Me.cmbFiscalYear)
    bat.period = glPd
    
    If Not BudgetFlag Then
        bat.JournalSource = glJS
    Else
        bat.JournalSource = glJS + 100
    End If
    
    bat.Updated = Now()
    bat.updateUser = curUser
    
    bat.PutRecord BatchNumberC, FileName
    
    ' ******** update the batch record
    
    ' update GLAmount
    If BalintFolder = "" Then
         x = "\Balint\GLUtil.exe " & _
             "SysFile=\Balint\Data\GLSystem.mdb " & _
             "UserID=" & curUser & " " & _
             "BackName=\Balint\GLEntry.exe " & _
             "ProgName=UpdateB " & _
             "Batch=" & bat.BatchNumber
    Else
         x = "c:\Balint\GLUtil.exe " & _
            "SysFile=" & BalintFolder & "\Data\GLSystem.mdb " & _
            "UserID=" & curUser & " " & _
            "BackName=" & "c:\Balint\GLEntry.exe " & _
            "ProgName=UpdateB " & _
            "Batch=" & bat.BatchNumber & _
            " BalintFolder=" & BalintFolder
    End If

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
    bat.GetBatch BatchFrom, FileName
    txtCompanyName = com.Name
    lblBatchNumber = "Copy From Batch # " & bat.BatchNumber
    lblCreated = "Created by " & UserName(bat.createUser) & " on " & ShowDate(bat.Created)
    txtRecord = "RECORD COUNT = " & CStr(bat.nRecords)
    txtDebits = "DEBITS = " & Format(bat.debits, "Currency")
    txtCredits = "CREDITS = " & Format(bat.credits, "Currency")
    
    ' store the GLBatch record values
    BatDebits = bat.debits
    BatCredits = bat.credits
    BatBatchNumber = bat.BatchNumber
    BatRecords = bat.nRecords
 
'    For ndx = com.FirstFiscalYear To Year(Now) + 1
'        cmbFiscalYear.AddItem ndx
'    Next ndx
'
'    'if bat.fiscalYear=0 then
'    cmbFiscalYear = bat.fiscalYear
    
    CurFY = Int(com.lastClose / 10 ^ 4)
    If Int(com.lastClose / 100) Mod 100 <> 1 Then CurFY = CurFY + 1
    If CurFY < 1990 Or CurFY > 2020 Then CurFY = Year(Now())
    
    For ndx = CurFY + 1 To CurFY - 5 Step -1
        cmbFiscalYear.AddItem ndx
    Next ndx
    cmbFiscalYear.ListIndex = 1
 
'    For ndx = 1 To com.NumberPds
'        cmbPeriod.AddItem com.MonthName(ndx, bat.fiscalYear)
'        cmbPeriod.AddItem com.MonthName(ndx, CurFY)
'    Next ndx
''    cmbPeriod.ListIndex = bat.period - 1
'    cmbPeriod.ListIndex = 0
    
    Set jou = xFactory.GetJournals(FileName)
    
    For ndx = 1 To jou.UpperBound(1)
        cmbJournal.AddItem (CStr(jou.Value(ndx, 0)) & " - " & jou.Value(ndx, 1))
        If jou.Value(ndx, 0) = bat.JournalSource Then
            cmbJournal.ListIndex = ndx - 1
        End If
    Next ndx

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
Dim ln, i1, i2 As Long

    GetNumber = 0
    If IsNull(InString) Then Exit Function
       
    x2 = ""
    ln = Len(InString)
    If ln = 0 Then Exit Function
    i1 = 0
    
    Do
       i1 = i1 + 1
       If i1 > ln Then Exit Do
       x1 = Mid(InString, i1, 1)
       If InStr(1, "0123456789", x1, vbTextCompare) = 0 Then Exit Do
       x2 = x2 & Mid(InString, i1, 1)
    Loop
    
    If IsNumeric(x2) Then GetNumber = CLng(x2)

End Function


