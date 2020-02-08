VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPRHistImport 
   Caption         =   "Paryoll History Import from SuperDOS"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   6660
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlFName 
      Left            =   8400
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TDBNumber6Ctl.TDBNumber tdbDays 
      Height          =   375
      Left            =   2258
      TabIndex        =   1
      Top             =   3000
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
      _ExtentY        =   661
      Calculator      =   "frmPRHistImport.frx":0000
      Caption         =   "frmPRHistImport.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPRHistImport.frx":00B4
      Keys            =   "frmPRHistImport.frx":00D2
      Spin            =   "frmPRHistImport.frx":011C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdLookup 
      Height          =   495
      Left            =   4680
      Picture         =   "frmPRHistImport.frx":0144
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtFileName 
      Height          =   390
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   7815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblMsg3 
      Alignment       =   2  'Center
      Caption         =   "Msg3"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   9135
   End
   Begin VB.Label lblMsg2 
      Alignment       =   2  'Center
      Caption         =   "Msg2"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   9135
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Msg1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label1 
      Caption         =   "SuperDOS History file to import:"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmPRHistImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ASCIIChannel As Variant
Dim X, Y, z As String
Dim i, j, k As Long
Dim LoYM As Long
Dim HiYM As Long
Dim CompName As String
Dim RecCt As Long
Dim rsPE As New ADODB.Recordset
Dim RSEmp As New ADODB.Recordset
Dim PEDate As Date
Dim Ct As Long
Dim FMT As String
Dim LastYM As Long
Dim LastPE As Date
Dim Flg As Boolean
Dim MM, DD, YY As Long
Dim DirDepDed1 As Long
Dim DirDepDed2 As Long

Private Sub Form_Load()

    ' *** TO DO ***
    ' dir dep deductions
    ' other tax logic commented out
    ' PRD import
    ' *****************************
    
    FMT = "#,###,##0"

    Me.lblCompanyName = PRCompany.Name
    
    Me.lblMsg1.Caption = ""
    Me.lblMsg2.Caption = ""
    Me.lblMsg3.Caption = ""
    
    tdbIntegerSet Me.tdbDays
    Me.tdbDays = PRCompany.CheckDays

 Me.txtFileName = "c:\Balint\Data\PRH11901.txt"

    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdLookup_Click()
    Me.txtFileName = GetTxtName("PRH*.txt", Me.cdlFName)
End Sub

Private Sub cmdOK_Click()
    
    ASCIIChannel = FreeFile
    Open Me.txtFileName For Input As ASCIIChannel

    ' first field must be "HIST"
    Input #ASCIIChannel, X
    If X <> "HIST" Then
        MsgBox "Invalid file: " & Me.txtFileName, vbExclamation
        GoBack
    End If
    
    Input #ASCIIChannel, CompName
    Input #ASCIIChannel, Y
    LoYM = CLng(Y)
    Input #ASCIIChannel, Y
    HiYM = CLng(Y)
    
    If MsgBox("OK to import PR history for: " & Trim(CompName) & vbCr & _
              LoYM & " To: " & HiYM & "?", vbQuestion + vbYesNo) = vbNo Then
        GoBack
    End If

    RecCt = 0
    
    rsPE.CursorLocation = adUseClient
    rsPE.Fields.Append "PEDate", adDouble
    rsPE.Open , , adOpenDynamic, adLockOptimistic
    
    RSEmp.CursorLocation = adUseClient
    RSEmp.Fields.Append "EmpNum", adDouble
    RSEmp.Open , , adOpenDynamic, adLockOptimistic
    
    frmProgress.Show
    frmProgress.lblMsg1 = "Scanning PR History Import File: " & Trim(Me.txtFileName)
    frmProgress.Refresh
    
    ' pre-scan the file
    ' warn if PE date already exists
    ' kick out if employee DNE
    Do
        Input #ASCIIChannel, X
        If X = "END" Then
            Exit Do
        ElseIf X = "HIS" Then
            For i = 1 To 53
                Input #ASCIIChannel, X
                If i = 1 Then           ' YYYYMM
                ElseIf i = 2 Then       ' EE#
                    SQLString = "EmpNum = " & Trim(X)
                    RSEmp.Find SQLString, 0, adSearchForward, 1
                    If RSEmp.EOF Then
                        RSEmp.AddNew
                        RSEmp!EmpNum = CDbl(X)
                        RSEmp.Update
                    End If
                ElseIf i = 3 Then       ' ck#
                ElseIf i = 4 Then       ' pe date
                    SQLString = "PEDate = " & Trim(X)
                    rsPE.Find SQLString, 0, adSearchForward, 1
                    If rsPE.EOF Then
                        rsPE.AddNew
                        rsPE!PEDate = CDbl(X)
                        rsPE.Update
                    End If
                End If
            Next i
        Else        ' PRDist
        End If
        RecCt = RecCt + 1
        If RecCt Mod 10 = 1 Then
            frmProgress.lblMsg2 = "Scanning File: " & Format(RecCt, FMT)
            frmProgress.Refresh
        End If
    Loop

    frmProgress.Hide

    ' employee numbers not found?
    If RSEmp.RecordCount > 0 Then
        RSEmp.MoveFirst
        Do
            SQLString = "SELECT * FROM PREmployee WHERE EmployeeNumber = " & RSEmp!EmpNum
            If PREmployee.GetBySQL(SQLString) = False Then
                MsgBox "Employee Number not found in Windows PR: " & RSEmp!EmpNum, vbExclamation
                GoBack
            End If
            RSEmp.MoveNext
        Loop Until RSEmp.EOF
    End If

    ' PE date already in Windows?
    If rsPE.RecordCount > 0 Then
        rsPE.MoveFirst
        Do
            PEDate = DateSerial(Int(rsPE!PEDate / 10 ^ 4), Int(rsPE!PEDate / 100) Mod 100, rsPE!PEDate Mod 100)
            SQLString = "SELECT * FROM PRHist WHERE PEDate = " & CLng(PEDate)
            If PRHist.GetBySQL(SQLString) = True Then
                If MsgBox("PE Date: " & Format(PEDate, "mm/dd/yyyy") & vbCr & _
                          "Already in Windows - OK to import?", vbYesNo + vbQuestion) = vbNo Then
                    GoBack
                End If
            End If
            rsPE.MoveNext
        Loop Until rsPE.EOF
    End If

    ' OK to proceed......
    rsPE.Close
    rsPE.CursorLocation = adUseClient
    rsPE.Fields.Append "YearMonth", adDouble
    rsPE.Fields.Append "PEDate", adDouble
    rsPE.Fields.Append "BatchID", adDouble
    rsPE.Open , , adOpenDynamic, adLockOptimistic
    
    Close #ASCIIChannel
    Open Me.txtFileName For Input As ASCIIChannel
    
    ' first field must be "HIST"
    Input #ASCIIChannel, X
    If X <> "HIST" Then
        MsgBox "Invalid file: " & Me.txtFileName, vbExclamation
        GoBack
    End If
    
    Input #ASCIIChannel, CompName
    Input #ASCIIChannel, Y
    LoYM = CLng(Y)
    Input #ASCIIChannel, Y
    HiYM = CLng(Y)

    Ct = 0
    
    frmProgress.Show
    frmProgress.lblMsg1 = "Now importing History records ..."
    frmProgress.Refresh
    
    Do
        
        Ct = Ct + 1
        If Ct Mod 10 = 1 Then
            frmProgress.lblMsg2 = "On Record: " & Format(Ct, FMT) & " Of: " & Format(RecCt, FMT)
            frmProgress.Refresh
        End If
        Input #ASCIIChannel, X
        If X = "END" Then
            Exit Do
        ElseIf X = "HIS" Then
            ImportHistory
        Else
            MsgBox "Invalid data: " & X, vbExclamation
            GoBack
        End If
    Loop

    ' udpate to PRBatch
    If rsPE.RecordCount > 0 Then
        rsPE.MoveFirst
        Do
            
            frmProgress.lblMsg2 = "Creating Batch records for: " & Format(rsPE!PEDate, "mm/dd/yyyy")
            frmProgress.Refresh
            
            PRBatch.OpenRS
            PRBatch.Clear
            PRBatch.PEDate = rsPE!PEDate
            PRBatch.CheckDate = rsPE!PEDate + Me.tdbDays
            PRBatch.YearMonth = Year(PRBatch.CheckDate) * 100 + Month(PRBatch.CheckDate)
            PRBatch.UserID = User.ID
            PRBatch.CreateDate = Int(Now())
            PRBatch.RecCount = 0
            PRBatch.Save (Equate.RecAdd)
            
            SQLString = "SELECT * FROM PRHist WHERE YearMonth = " & rsPE!YearMonth & _
                        " AND PEDate = " & rsPE!PEDate & _
                        " ORDER BY HistID"
            If PRHist.GetBySQL(SQLString) Then
                Do
                    PRHist.BatchID = PRBatch.BatchID
                    PRHist.Save (Equate.RecPut)
                    PRBatch.RecCount = PRBatch.RecCount + 1
                    If PRHist.GetNext = False Then Exit Do
                Loop
            End If
            
            PRBatch.Save (Equate.RecPut)
            
            SQLString = "SELECT * FROM PRDist WHERE YearMonth = " & rsPE!YearMonth & _
                        " AND PEDate = " & rsPE!PEDate & _
                        " ORDER BY HistID"
            If PRDist.GetBySQL(SQLString) Then
                Do
                    PRDist.BatchID = PRBatch.BatchID
                    PRDist.Save (Equate.RecPut)
                    If PRDist.GetNext = False Then Exit Do
                Loop
            End If
            
            SQLString = "SELECT * FROM PRItemHist WHERE YearMonth = " & rsPE!YearMonth & _
                        " AND PEDate = " & rsPE!PEDate & _
                        " ORDER BY HistID"
            If PRItemHist.GetBySQL(SQLString) Then
                Do
                    PRItemHist.BatchID = PRBatch.BatchID
                    PRItemHist.Save (Equate.RecPut)
                    If PRItemHist.GetNext = False Then Exit Do
                Loop
            End If
            
            rsPE.MoveNext
        Loop Until rsPE.EOF
    
    End If

    frmProgress.Hide

    MsgBox Format(RecCt, FMT) & " history records imported", vbInformation
    
    GoBack

End Sub

Private Sub ImportHistory()

Dim TotalOE As Currency
Dim TotalOEHours As Single
Dim TotalDed As Currency
Dim TotalDirDep As Currency

Dim OEHours(1 To 10) As Single
Dim OEAmount(1 To 10) As Currency
Dim TestDate As Date

Dim CWTTotal, CWTAmount, ERNTotal, ERNAmount As Currency
Dim SplitAmount, SWTTotal, SWTAmount As Currency
Dim DistID As Long
Dim CtyID As Long
Dim OTXAmt As Currency

    ' *** SWT and CWT to be split among all PRDist records ***
    DistID = 0      ' storing the first PRDist item created
                    ' use for SWT and CWT rounding if necessary
    
    CtyID = 0
    
    CtyID = PRCompany.DfltCityID
    
    PRHist.Clear
    
    For i = 1 To 53
        
        Input #ASCIIChannel, X

        If X = "" Or X = "0.00" Then GoTo NextHist

        If i = 1 Then PRHist.YearMonth = CLng(X)

        ' employee number - get EmployeeID
        If i = 2 Then
            SQLString = "SELECT * from PREmployee WHERE PREmployee.EmployeeNumber = " & CStr(X)
            If Not PREmployee.GetBySQL(SQLString) Then
                ' add a new employee
                PREmployee.Clear
                PREmployee.EmployeeNumber = X
                PREmployee.FirstName = "NEW"
                PREmployee.LastName = "EMPLOYEE"
                PREmployee.Save (Equate.RecAdd)
            End If
            PRHist.EmployeeID = PREmployee.EmployeeID
            PRHist.Save (Equate.RecAdd)     ' save it so a PRHistID is generated
        End If

        If i = 3 Then PRHist.CheckNumber = CLng(X)

        If i = 4 Then       ' PE Date yyyymmdd
            MM = Mid(X, 5, 2)
            DD = Mid(X, 7, 2)
            YY = Mid(X, 1, 4)
            PRHist.PEDate = DateSerial(YY, MM, DD)
            PRHist.CheckDate = PRHist.PEDate + tdbDays
        End If

        If i = 5 And X <> "0" Then       ' department number
            SQLString = "SELECT * FROM PRDepartment WHERE PRDepartment.DepartmentNumber = " & CStr(X)
            If Not PRDepartment.GetBySQL(SQLString) Then
            Else
                PRHist.DepartmentID = PRDepartment.DepartmentID
            End If
        End If

        If i = 6 Then PRHist.RegRate = CCur(X)
        ' i = 7 leave state number blank - assume Ohio for conversions
        If i = 8 Then PRHist.RegHours = CSng(X)
        If i = 9 Then PRHist.OTHours = CSng(X)

        If i >= 10 And i <= 19 Then
            TotalOEHours = TotalOEHours + CSng(X)
            OEHours(i - 9) = CSng(X)
        End If

        If i = 20 Then PRHist.RegAmount = CCur(X)
        If i = 21 Then PRHist.OTAmount = CCur(X)

        If i >= 22 And i <= 31 Then
            OEAmount(i - 21) = CCur(X)
        End If

        ' deductions - add to PRItemHist
        If i >= 32 And i <= 41 And X <> "0.00" Then

            ' find the PRItem record
            If DirDepDed1 = i - 31 Or DirDepDed2 = i - 31 Then
                SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeDirDepDed) & _
                            " AND PRItem.SDNumber = " & CStr(i - 31) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            Else
                SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeDED) & _
                            " AND PRItem.SDNumber = " & CStr(i - 31) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            End If
            
            ' not found - get the employer item and add an employee item
            If Not PRItem.GetBySQL(SQLString) Then
                
                SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 AND PRItem.SDNumber = " & CStr(i - 31) & _
                            " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeDED)
                
                If Not PRItem.GetBySQL(SQLString) Then
                    MsgBox "Employer PRItem not found - ded # " & CStr(i - 31) & " " & PREmployee.EmployeeNumber, vbCritical
                    End
                End If
                
                PRItem.UseEmployer = 1
                PRItem.EmployeeID = PREmployee.EmployeeID
                PRItem.EmployerItemID = PRItem.ItemID
                PRItem.SDNumber = i - 31
                PRItem.Save (Equate.RecAdd)
            
            End If

            PRItemHist.OpenRS
            PRItemHist.Clear
            PRItemHist.EmployeeID = PRHist.EmployeeID
            PRItemHist.HistID = PRHist.HistID
            PRItemHist.DepartmentID = PRHist.DepartmentID
            PRItemHist.ItemID = PRItem.ItemID
            PRItemHist.EmployerItemID = PRItem.EmployerItemID
            PRItemHist.Hours = 0
            PRItemHist.Amount = CCur(X)
            PRItemHist.ManualAmount = 1
            PRItemHist.YearMonth = PRHist.YearMonth
            PRItemHist.PEDate = PRHist.PEDate

            If i - 31 = DirDepDed1 Or i - 31 = DirDepDed2 Then
                PRItemHist.ItemType = PREquate.ItemTypeDirDepDed
            Else
                PRItemHist.ItemType = PREquate.ItemTypeDED
            End If
            
            PRItemHist.Save (Equate.RecAdd)

        End If

        If i >= 22 And i <= 31 Then
            TotalOE = TotalOE + CCur(X)
        End If

        If i >= 32 And i <= 41 Then
            If i - 31 = DirDepDed1 Or i - 31 = DirDepDed2 Then
                TotalDirDep = TotalDirDep + CCur(X)
            Else
                TotalDed = TotalDed + CCur(X)
            End If
        End If

        If i = 42 Then PRHist.SSTax = CCur(X)
        If i = 43 Then PRHist.MedTax = CCur(X)
        If i = 44 Then PRHist.FWTTax = CCur(X)
        If i = 45 Then PRHist.SWTTax = CCur(X)
        If i = 46 Then PRHist.CWTTax = CCur(X)

'        '  other tax as city tax?
'        If I >= 47 And I <= 51 And PRHist.CWTTax = 0 And CtyID = 0 Then
'            If I = 47 And Tax6City <> 0 Then
'                PRHist.CWTTax = CCur(X)
'                CtyID = Tax6City
'            End If
'            If I = 48 And Tax7City <> 0 Then
'                PRHist.CWTTax = CCur(X)
'                CtyID = Tax7City
'            End If
'            If I = 49 And Tax8City <> 0 Then
'                PRHist.CWTTax = CCur(X)
'                CtyID = Tax8City
'            End If
'            If I = 50 And Tax9City <> 0 Then
'                PRHist.CWTTax = CCur(X)
'                CtyID = Tax9City
'            End If
'            If I = 51 And Tax0City <> 0 Then
'                PRHist.CWTTax = CCur(X)
'                CtyID = Tax0City
'            End If
'        End If

'        OTXAmt = CCur(X)
'        If I = 47 And SDTax6ID <> 0 And OTXAmt <> 0 Then
'            AddSDTax SDTax6ID, OTXAmt
'        End If
'        If I = 48 And SDTax7ID <> 0 And OTXAmt <> 0 Then
'            AddSDTax SDTax7ID, OTXAmt
'        End If
'        If I = 49 And SDTax8ID <> 0 And OTXAmt <> 0 Then
'            AddSDTax SDTax8ID, OTXAmt
'        End If
'        If I = 50 And SDTax9ID <> 0 And OTXAmt <> 0 Then
'            AddSDTax SDTax9ID, OTXAmt
'        End If
'        If I = 51 And SDTax0ID <> 0 And OTXAmt <> 0 Then
'            AddSDTax SDTax0ID, OTXAmt
'        End If

        If i = 52 Then PRHist.Gross = CCur(X)
        If i = 53 Then PRHist.Net = CCur(X)

NextHist:
    
    Next i

    ' final updates
    PRHist.Deductions = TotalDed
    PRHist.OEAmount = TotalOE
    PRHist.OEHours = TotalOEHours
    PRHist.DirectDeposit = TotalDirDep

    ' set manual flags for all imports
    PRHist.ManualSSTax = 1
    PRHist.ManualMedTax = 1
    PRHist.ManualFWTTax = 1
    PRHist.StateID = DfltStateID

    PRHist.Save (Equate.RecPut)

    ' keep rs of unique PE Date and YearMonth for batch record create
    Flg = False
    If rsPE.RecordCount > 0 Then
        rsPE.MoveFirst
        Do
            If rsPE!PEDate = PRHist.PEDate And rsPE!YearMonth = PRHist.YearMonth Then
                Flg = True
                Exit Do
            End If
            rsPE.MoveNext
        Loop Until rsPE.EOF
    End If
    If Flg = False Then
        rsPE.AddNew
        rsPE!PEDate = PRHist.PEDate
        rsPE!YearMonth = PRHist.YearMonth
        rsPE.Update
    End If

    ' !!!!!!!!!!!!!!!!!!!!
    ' dont write to PRDist for dist companies
    ' **************************
    ' !!!!!!!!!!!!!!!!!!!!
    ' PRDist logic not done yet
    ' !!!!!!!!!!!!!!!!!!!!
    ' If DistFlag = True Then Exit Sub

    ' init the variables
    CWTTotal = PRHist.CWTTax
    CWTAmount = 0
    SWTTotal = PRHist.SWTTax
    SWTAmount = 0
    ERNTotal = PRHist.Gross
    ERNAmount = 0

    ' write regular and overtime to PRDist
    PRDist.OpenRS
    PRDist.Clear
    
    ' ================================================
    ' if using other tax as city tax
    ' assign the cityid now
    ' dont overwrite later
    If CtyID <> 0 Then PRDist.CityID = CtyID
    ' ================================================
    
    PRDist.EmployeeID = PREmployee.EmployeeID
    PRDist.HistID = PRHist.HistID
    PRDist.DepartmentID = PRHist.DepartmentID
    PRDist.YearMonth = PRHist.YearMonth
    PRDist.PEDate = PRHist.PEDate
    PRDist.DistType = PREquate.DistTypeReg
    
'    ' use the employer item id
'    SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeRegPay) & _
'                " AND PRItem.EmployeeID = 0"
'    If Not PRItem.GetBySQL(SQLString) Then
'        MsgBox "Regular pay item nf: ", vbCritical
'        End
'    End If
    
    PRDist.ItemID = PRItem.ItemID
    
    ' zero for Reg Ern
    PRDist.EmployerItemID = 0
    
    PRDist.ItemType = PREquate.ItemTypeRegPay
    PRDist.Amount = PRHist.RegAmount
    PRDist.ManualAmount = 1
    PRDist.Hours = PRHist.RegHours
    PRDist.Rate = PRHist.RegRate
    
    PRDist.StateID = DfltStateID
    PRDist.HistFlag = 1
    
    ' split the state and city tax amounts
    CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
    PRDist.CityTax = CWTAmount
    CWTTotal = CWTTotal - CWTAmount
    PRDist.ManualCityTax = 1
    
    SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
    PRDist.StateTax = SWTAmount
    SWTTotal = SWTTotal - SWTAmount
    PRDist.ManualStateTax = 1
    
    PRDist.Save (Equate.RecAdd)
    
    If PRDist.Amount <> 0 Then DistID = PRDist.DistID
    
    ' Over Time
    If PRHist.OTAmount <> 0 Then
    
        PRDist.Clear
        
        ' ================================================
        ' if using other tax as city tax
        ' assign the cityid now
        ' dont overwrite later
        If CtyID <> 0 Then PRDist.CityID = CtyID
        ' ================================================
        
        PRDist.EmployeeID = PREmployee.EmployeeID
        
        PRDist.EmployerItemID = 0
        
        PRDist.HistID = PRHist.HistID
        PRDist.DepartmentID = PRHist.DepartmentID
        PRDist.YearMonth = PRHist.YearMonth
        PRDist.PEDate = PRHist.PEDate
        PRDist.DistType = PREquate.DistTypeOT
        
        ' use the employer item id
        SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeOvtPay) & _
                    " AND PRItem.EmployeeID = 0"
        If Not PRItem.GetBySQL(SQLString) Then
            MsgBox "Overtime pay item nf: ", vbCritical
            End
        End If
        PRDist.ItemID = PRItem.ItemID
        PRDist.ItemType = PREquate.ItemTypeOvtPay
        PRDist.Amount = PRHist.OTAmount
        PRDist.ManualAmount = 1
        PRDist.Hours = PRHist.OTHours
        PRDist.Rate = PRHist.OTRate
        
        PRDist.StateID = DfltStateID
        PRDist.HistFlag = 1
        
        CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
        PRDist.CityTax = CWTAmount
        CWTTotal = CWTTotal - CWTAmount
        PRDist.ManualCityTax = 1
        
        SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
        PRDist.StateTax = SWTAmount
        SWTTotal = SWTTotal - SWTAmount
        PRDist.ManualStateTax = 1
        
        PRDist.Save (Equate.RecAdd)
    
        If DistID = 0 Then DistID = PRDist.DistID
    
    End If
    
    ' write other earnings to PRDist
    For i = 1 To 10
        If OEHours(i) <> 0 Or OEAmount(i) <> 0 Then
            
            PRDist.Clear
        
            ' ================================================
            ' if using other tax as city tax
            ' assign the cityid now
            ' dont overwrite later
            If CtyID <> 0 Then PRDist.CityID = CtyID
            ' ================================================

            PRDist.EmployeeID = PREmployee.EmployeeID
            PRDist.HistID = PRHist.HistID
            ' state id
            ' city id
            ' job id
            ' customer id
            PRDist.DepartmentID = PRHist.DepartmentID
            
            PRDist.ItemID = PRItem.ItemID
            
            PRDist.YearMonth = PRHist.YearMonth
            PRDist.PEDate = PRHist.PEDate
            PRDist.CheckDate = PRHist.CheckDate
            PRDist.DistType = PREquate.DistTypeItem
            PRDist.ItemType = PREquate.ItemTypeOE
            
            SQLString = "SELECT * FROM PRItem WHERE PRItem.ItemType = " & CStr(PREquate.ItemTypeOE) & _
                        " AND PRItem.SDNumber = " & CStr(i) & " AND PRItem.EmployeeID = " & CStr(PREmployee.EmployeeID)
            
            If Not PRItem.GetBySQL(SQLString) Then
                
                SQLString = "SELECT * FROM PRItem WHERE PRItem.EmployeeID = 0 AND PRItem.SDNumber = " & CStr(i) & _
                            " AND PRItem.ItemType = " & CStr(PREquate.ItemTypeOE)
                
                If Not PRItem.GetBySQL(SQLString) Then
                    MsgBox "Employer PRItem not found - oe # " & CStr(i), vbCritical
                    End
                End If
                
                PRItem.EmployeeID = PREmployee.EmployeeID
                PRItem.EmployerItemID = PRItem.ItemID
                PRItem.ItemID = 0
                PRItem.UseEmployer = 1
                PRItem.Save (Equate.RecAdd)
            
                PRDist.EmployerItemID = PRItem.ItemID
            
            Else
            
                PRDist.EmployerItemID = PRItem.EmployerItemID
            
            End If
            
            
            PRDist.ItemID = PRItem.ItemID
            
            PRDist.Hours = OEHours(i)
            If OEHours(i) = 0 Then
                PRDist.Rate = OEAmount(i)
            Else
                PRDist.Rate = OEAmount(i) / OEHours(i)
            End If
            PRDist.Amount = OEAmount(i)
            PRDist.ManualAmount = 1
            
            ' billing rate
            ' state wage
            ' state tax
            ' city wage
            ' city tax
            
            PRDist.HistFlag = 1
            PRDist.StateID = DfltStateID
        
            CWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.CWTTax)
            PRDist.CityTax = CWTAmount
            CWTTotal = CWTTotal - CWTAmount
            PRDist.ManualCityTax = 1
            
            SWTAmount = SplitCalc(PRDist.Amount, PRHist.Gross, PRHist.SWTTax)
            PRDist.StateTax = SWTAmount
            SWTTotal = SWTTotal - SWTAmount
            PRDist.ManualStateTax = 1
            
            PRDist.Save (Equate.RecAdd)
            
            If DistID = 0 Then DistID = PRDist.DistID
        
        End If
    
    Next i

    ' rounding correction???
    If SWTTotal <> 0 Then
        If DistID = 0 Then
            MsgBox "No amounts ???", vbExclamation
            End
        End If
        If Not PRDist.GetByID(DistID) Then
            MsgBox "PRDist err: " & DistID, vbExclamation
            End
        End If
        PRDist.StateTax = PRDist.StateTax + SWTTotal
        PRDist.Save (Equate.RecPut)
    End If

    If CWTTotal <> 0 Then
        If DistID = 0 Then
            MsgBox "No amounts ???", vbExclamation
            End
        End If
        If Not PRDist.GetByID(DistID) Then
            MsgBox "PRDist err: " & DistID, vbExclamation
            End
        End If
        PRDist.CityTax = PRDist.CityTax + CWTTotal
        PRDist.Save (Equate.RecPut)
    End If

End Sub

Private Function GetTxtName(ByVal WildCard As String, _
                    ByRef cmd As CommonDialog) As String
      
Dim OPath As String
Dim jbName As String

   ' store original path
   OPath = App.Path
      
   cmd.Filter = "Export Files|" & WildCard
   cmd.DefaultExt = ".txt"
   cmd.DialogTitle = "Select File to Import"
   jbName = Left(App.Path, 2) & "\Balint\Data"
   cmd.InitDir = jbName
   cmd.ShowOpen
   GetTxtName = cmd.FileName
   If GetTxtName = "" Then Exit Function

   ' restore original drive and path
   ChDrive (Left(OPath, 2))
   ChDir (OPath)

End Function




