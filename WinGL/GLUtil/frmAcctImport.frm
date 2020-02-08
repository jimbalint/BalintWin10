VERSION 5.00
Begin VB.Form frmAcctImport 
   Caption         =   "Import Chart of Accounts"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcctImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   7185
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   5280
      Width           =   1815
   End
   Begin VB.ComboBox cmbGLCompany 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Width           =   10215
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1815
      Left            =   1665
      TabIndex        =   5
      Top             =   2760
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Select GL Company to import chart of accounts from:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   4815
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
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmAcctImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x, Y As String
Dim GLIDFrom, GLIDTo As Long
Dim cn2 As New ADODB.Connection
Dim rsFrom As New ADODB.Recordset
Dim rsTo As New ADODB.Recordset
Dim Ct As Long

Dim FileName As String
Dim Address1 As String
Dim Address2 As String
Dim Address3 As String
Dim City As String
Dim FirstPAcct As Long
Dim FirstPeriod As Byte
Dim LastClose As Long
Dim LastUpdate As Long
Dim NetProfitAcct As Long
Dim ID As Long
Dim NumberPds As Byte
Dim PctBaseAcct As Long
Dim RetEarnAcct As Long
Dim State As String
Dim SubDigits As Byte
Dim SuspAcct As Long
Dim ZipCode As Long
Dim FirstFiscalYear As Integer
Dim LastBatch As Long

Dim LowBranch As Long
Dim HiBranch As Long
Dim LowConsolidated As Long
Dim HiConsolidated As Long


Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    
    Me.lblWarning.Caption = " * * * W A R N I N G * * *" & vbCr & vbCr & _
                            "ALL G/L Data for " & GLCompany.Name & vbCr & _
                            "will be deleted and the chart of accounts" & vbCr & _
                            "will be replaced!!!"
    
    ' store the ID for the current company
    GLIDTo = GLCompany.ID
    
    ' init the dropdown
    If GLCompany.GetBySQL("SELECT * FROM GLCompany ORDER BY NAME") = False Then
        MsgBox "No GL Company data found!", vbExclamation
        GoBack
    End If
    With Me.cmbGLCompany
        Do
            If GLCompany.ID <> GLIDTo Then
                Y = Trim(GLCompany.Name)
                If Len(Y) >= 30 Then
                    x = Mid(Y, 1, 30) & " " & GLCompany.FileName
                Else
                    x = Y & Space(31 - Len(Y)) & GLCompany.FileName
                End If
                .AddItem x
                .ItemData(.NewIndex) = GLCompany.ID
            End If
            If GLCompany.GetNext = False Then Exit Do
        Loop
    End With
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()

    If Me.cmbGLCompany.ListIndex = -1 Then
        MsgBox "Select a company to import the chart of accounts from!", vbInformation
        Exit Sub
    End If
    
    x = "OK to replace delete ALL G/L data for: " & Me.lblCompanyName & vbCr & _
        "and import chart of accounts from: " & Me.cmbGLCompany & "?"
    If MsgBox(x, vbQuestion + vbYesNo) = vbNo Then Exit Sub

    ' get the copy from GL company
    With Me.cmbGLCompany
        GLIDFrom = .ItemData(.ListIndex)
    End With
    
    If GLCompany.GetData(GLIDFrom) = False Then
        MsgBox "Can not open copy from GL Company File: " & GLIDFrom, vbExclamation
        GoBack
    End If

    ' variables to store company info
    Address1 = GLCompany.Address1
    Address2 = GLCompany.Address2
    Address3 = GLCompany.Address3
    City = GLCompany.City
    FirstPAcct = GLCompany.FirstPAcct
    FirstPeriod = GLCompany.FirstPeriod
    NetProfitAcct = GLCompany.NetProfitAcct
    NumberPds = GLCompany.NumberPds
    PctBaseAcct = GLCompany.PctBaseAcct
    RetEarnAcct = GLCompany.RetEarnAcct
    State = GLCompany.State
    SubDigits = GLCompany.SubDigits
    ZipCode = GLCompany.ZipCode
    FirstFiscalYear = GLCompany.FirstFiscalYear
    LowBranch = GLCompany.LowBranch
    HiBranch = GLCompany.HiBranch
    LowConsolidated = GLCompany.LowConsolidated
    HiConsolidated = GLCompany.HiConsolidated
    
    ' connect to the copy from db
    Set cn2 = New ADODB.Connection
    cn2.Provider = "Microsoft.Jet.OLEDB.4.0"
    x = Mid(App.Path, 1, 2) & Mid(GLCompany.FileName, 3, Len(GLCompany.FileName) - 2)
    cn2.ConnectionString = x
    On Error Resume Next
    cn2.Open
    If Err.Number <> 0 Then
        MsgBox "Error connecting to: " & GLCompany.FileName & vbCr & _
               Err.Number & " " & Err.Description, vbExclamation
        GoBack
    End If
    On Error GoTo 0
    
    frmProgress.Show
    frmProgress.lblMsg1 = Me.lblCompanyName & vbCr & _
                        "Importing GL Chart of Accounts from: " & GLCompany.Name
    frmProgress.Refresh
    
    ' *** remove and copy data ***
    frmProgress.lblMsg2 = "Now Removing GL Account Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLAccount"
    cn.Execute SQLString
    
    frmProgress.lblMsg2 = "Now Copying GL Account Data ..."
    frmProgress.Refresh
    SQLString = "SELECT * FROM GLAccount"
    rsInit SQLString, cn2, rsFrom
    If rsFrom.RecordCount > 0 Then
        
        SQLString = "SELECT * FROM GLAccount"
        rsInit SQLString, cn, rsTo
        
        Ct = 0
        rsFrom.MoveFirst
        Do
            
            If Ct Mod 20 = 1 Then
                frmProgress.lblMsg2 = "Copying GL Account Data " & _
                                      Format(Ct, "#,###,##0") & " of: " & _
                                      Format(rsFrom.RecordCount, "#,###,##0")
                frmProgress.Refresh
            End If
                                     
            rsTo.AddNew
            rsTo!DescNumber = rsFrom!DescNumber
            rsTo!Account = rsFrom!Account
            rsTo!AllSchedules = rsFrom!AllSchedules
            rsTo!AllStatements = rsFrom!AllStatements
            rsTo!BranchAcct = rsFrom!BranchAcct
            rsTo!BSColumn = rsFrom!BSColumn
            rsTo!ConsAcct = rsFrom!ConsAcct
            rsTo!Date1 = rsFrom!Date1
            rsTo!Date2 = rsFrom!Date2
            rsTo!Description = rsFrom!Description
            rsTo!DollarSign = rsFrom!DollarSign
            rsTo!LineFeeds = rsFrom!LineFeeds
            rsTo!PrintTab = rsFrom!PrintTab
            rsTo!SignRevSched = rsFrom!SignRevSched
            rsTo!SignRevStmt = rsFrom!SignRevStmt
            rsTo!TotalLevel = rsFrom!TotalLevel
            rsTo!TotalOnLedger = rsFrom!TotalOnLedger
            rsTo!AcctType = rsFrom!AcctType
            rsTo.Update
            rsFrom.MoveNext
        Loop Until rsFrom.EOF
    End If
    
    If TableExists("GLFFSched", cn) Then
        
        frmProgress.lblMsg2 = "Now Removing GL FF Schedule Data ..."
        frmProgress.Refresh
        SQLString = "DELETE * FROM GLFFSched"
        cn.Execute SQLString
    
        If TableExists("GLFFSched", cn2) Then
            SQLString = "SELECT * FROM GLFFSched"
                        
            rsInit SQLString, cn2, rsFrom
            
            If rsFrom.RecordCount > 0 Then
                rsInit SQLString, cn, rsTo
                rsFrom.MoveFirst
                Do
                    rsTo.AddNew
                    rsTo!GlobalID = rsFrom!GlobalID
                    rsTo!Account = rsFrom!Account
                    rsTo!SortOrder = rsFrom!SortOrder
                    rsTo!PercentBase = rsFrom!PercentBase
                    rsTo!PrintTab = rsFrom!PrintTab
                    rsTo!LineFeeds = rsFrom!LineFeeds
                    rsTo!AltDesc = rsFrom!AltDesc
                    rsTo!ReportID = rsFrom!ReportID
                    rsTo!SignReverse = rsFrom!SignReverse
                    rsTo.Update
                    rsFrom.MoveNext
                Loop Until rsFrom.EOF
            End If
        End If
    End If
    
    frmProgress.lblMsg2 = "Now Removing GL Journal Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLJournal"
    cn.Execute SQLString
    
    SQLString = "SELECT * FROM GLJournal"
    rsInit SQLString, cn2, rsFrom
    If rsFrom.RecordCount > 0 Then
        rsInit SQLString, cn, rsTo
        rsFrom.MoveFirst
        Do
            rsTo.AddNew
            rsTo!JournalSource = rsFrom!JournalSource
            rsTo!JournalName = rsFrom!JournalName
            rsTo.Update
            rsFrom.MoveNext
        Loop Until rsFrom.EOF
    End If
    
    ' copy company information
    frmProgress.lblMsg2 = "Now copying company information ...."
    If GLCompany.GetData(GLIDTo) = False Then
        MsgBox "GL Company Error: " & GLIDFrom, vbExclamation
        GoBack
    End If
    
    GLCompany.Address1 = Address1
    GLCompany.Address2 = Address2
    GLCompany.Address3 = Address3
    GLCompany.City = City
    GLCompany.FirstPAcct = FirstPAcct
    GLCompany.FirstPeriod = FirstPeriod
    GLCompany.NetProfitAcct = NetProfitAcct
    GLCompany.NumberPds = NumberPds
    GLCompany.PctBaseAcct = PctBaseAcct
    GLCompany.RetEarnAcct = RetEarnAcct
    GLCompany.State = State
    GLCompany.SubDigits = SubDigits
    GLCompany.ZipCode = ZipCode
    GLCompany.FirstFiscalYear = FirstFiscalYear
    GLCompany.LowBranch = LowBranch
    GLCompany.HiBranch = HiBranch
    GLCompany.LowConsolidated = LowConsolidated
    GLCompany.HiConsolidated = HiConsolidated
    GLCompany.Save (Equate.RecPut)
    
    ' *** remove only ***
    frmProgress.lblMsg2 = "Now Removing GL Amount Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLAmount"
    cn.Execute SQLString
    
    frmProgress.lblMsg2 = "Now Removing GL Batch Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLBatch"
    cn.Execute SQLString
    
    frmProgress.lblMsg2 = "Now Removing GL History Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLHistory"
    cn.Execute SQLString

    frmProgress.lblMsg2 = "Now Removing GL Print Data ..."
    frmProgress.Refresh
    SQLString = "DELETE * FROM GLPrint"
    cn.Execute SQLString
    
    frmProgress.Hide
    
    MsgBox "Import of Chart of Accounts from: " & Me.cmbGLCompany & vbCr & _
           "is complete.", vbInformation
              
    GoBack

End Sub

