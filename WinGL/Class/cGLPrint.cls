VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public FiscalYear As Long
Public BeginDate As Long
Public EndDate As Long
Public Copies As Integer
Public HiAccount As Long
Public HiBranchAcct As Integer
Public HiConsAcct As Integer
Public ID As Long
Public LowAccount As Long
Public LowBranchAcct As Integer
Public LowConsAcct As Integer
Public LowerCaseDate As Boolean
Public Output As String
Public PrintBIB As Byte
Public PrtAcctNum As Boolean
Public PrtZeroBal As Boolean
Public RegBraCon As Byte
Public RegCmp As Byte
Public ReportDate As Long
Public RoundDollars As Boolean
Public SepPage As Boolean
Public StaSch As Byte
Public SupprCP As Boolean
Public UseMathRec As Boolean
Public User As String
Public ReportName As String
Public WidePrint As Boolean

Dim SQLString As String
Dim rs As New ADODB.Recordset

Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = Equate.RecAdd Then
        rs.AddNew
    End If

    If IsNull(FiscalYear) Then FiscalYear = 0
    If IsNull(BeginDate) Then BeginDate = 0
    If IsNull(EndDate) Then EndDate = 0
    If IsNull(Copies) Then Copies = 1
    If IsNull(HiAccount) Then HiAccount = 0
    If IsNull(HiBranchAcct) Then HiBranchAcct = 0
    If IsNull(HiConsAcct) Then HiConsAcct = 0
    If IsNull(LowAccount) Then LowAccount = 0
    If IsNull(LowBranchAcct) Then LowBranchAcct = 0
    If IsNull(LowConsAcct) Then LowConsAcct = 0
    If IsNull(LowerCaseDate) Then LowerCaseDate = False
    If IsNull(Output) Then Output = ""
    If IsNull(PrintBIB) Then PrintBIB = 0
    If IsNull(PrtAcctNum) Then PrtAcctNum = False
    If IsNull(PrtZeroBal) Then PrtZeroBal = False
    If IsNull(RegBraCon) Then RegBraCon = 0
    If IsNull(RegCmp) Then RegCmp = 0
    If IsNull(ReportDate) Then ReportDate = 0
    If IsNull(RoundDollars) Then RoundDollars = False
    If IsNull(SepPage) Then SepPage = False
    If IsNull(StaSch) Then StaSch = 0
    If IsNull(SupprCP) Then SupprCP = False
    If IsNull(UseMathRec) Then UseMathRec = False
    If IsNull(User) Then User = ""
    If IsNull(WidePrint) Then WidePrint = False

    rs.Fields("FiscalYear") = FiscalYear
    rs.Fields("BeginDate") = BeginDate
    rs.Fields("EndDate") = EndDate
    rs.Fields("Copies") = Copies
    rs.Fields("HiAccount") = HiAccount
    rs.Fields("HiBranchAcct") = HiBranchAcct
    rs.Fields("HiConsAcct") = HiConsAcct
    rs.Fields("LowAccount") = LowAccount
    rs.Fields("LowBranchAcct") = LowBranchAcct
    rs.Fields("LowConsAcct") = LowConsAcct
    rs.Fields("LowerCaseDate") = LowerCaseDate
    
    ' Output or gOutput
    On Error Resume Next
    rs.Fields("Output") = Trim(Output & "")
    If Err.Number = 3265 Then
        rs.Fields("gOutput") = Trim(Output & "")
    End If
    On Error GoTo 0
    
    rs.Fields("PrintBIB") = PrintBIB
    rs.Fields("PrtAcctNum") = PrtAcctNum
    rs.Fields("PrtZeroBal") = PrtZeroBal
    rs.Fields("RegBraCon") = RegBraCon
    rs.Fields("RegCmp") = RegCmp
    rs.Fields("ReportDate") = ReportDate
    rs.Fields("RoundDollars") = RoundDollars
    rs.Fields("SepPage") = SepPage
    rs.Fields("StaSch") = StaSch
    rs.Fields("SupprCP") = SupprCP
    rs.Fields("UseMathRec") = UseMathRec
    rs.Fields("ReportName") = ReportName

    ' User or gUser
    On Error Resume Next
    rs.Fields("User") = User
    If Err.Number = 3265 Then
        rs.Fields("gUser") = User
    End If
    On Error GoTo 0
    
    rs.Fields("WidePrint") = WidePrint
    
    rs.Update
    
    ID = rs!ID
    
'    rs.Close
'    Set rs = Nothing
    
End Sub
    
Public Function GetByID(ByVal PrintID) As Boolean

    SQLString = "SELECT * FROM GLPrint WHERE ID = " & PrintID
    rsInit SQLString, cn, rs
    If rs.RecordCount = 0 Then
        GetByID = False
    Else
        AssignFields
        GetByID = True
    End If

End Function

Public Function GetBySQL(ByVal strSQL As String) As Boolean

    rsInit strSQL, cn, rs
    If rs.RecordCount = 0 Then
        GetBySQL = False
    Else
        AssignFields
        GetBySQL = True
    End If

End Function

Public Function GetData(ByVal strUser As String, ByRef NewRec As Boolean)

    Set rs = New ADODB.Recordset
    
    If strUser = "" Then strUser = "Default"
    
    ' User or gUser ???
    SQLString = "GLPrint"
    rsInit SQLString, cn, rs
        
    On Error Resume Next
    rs!User = ""
    If Err.Number = 3265 Then
        SQLString = "SELECT * FROM GLPrint WHERE gUser = '" & strUser & "'"
    Else
        SQLString = "SELECT * FROM GLPrint WHERE User = '" & strUser & "'"
    End If
    On Error GoTo 0
    
'    rs.Close
    
    NewRec = False
        
    rsInit SQLString, cn, rs
    
    ' Add if not found
    If rs.EOF = True And rs.BOF = True Then
        NewRec = True
        rs.AddNew
       
        ' User or gUser
        On Error Resume Next
        rs.Fields("User") = strUser
        If Err.Number = 3265 Then
            rs.Fields("gUser") = strUser
        End If
        On Error GoTo 0
       
        Clear
        rs.Fields("FiscalYear") = FiscalYear
        rs.Fields("BeginDate") = BeginDate
        rs.Fields("EndDate") = EndDate
        rs.Fields("Copies") = 1
        rs.Fields("HiAccount") = HiAccount
        rs.Fields("HiBranchAcct") = HiBranchAcct
        rs.Fields("HiConsAcct") = HiConsAcct
        rs.Fields("LowAccount") = LowAccount
        rs.Fields("LowBranchAcct") = LowBranchAcct
        rs.Fields("LowConsAcct") = LowConsAcct
        rs.Fields("LowerCaseDate") = LowerCaseDate
               
        ' Output or gOutput
        On Error Resume Next
        rs.Fields("Output") = Output
        If Err.Number = 3265 Then
            rs.Fields("gOutput") = Output
        End If
        On Error GoTo 0
        
        rs.Fields("PrintBIB") = PrintBIB
        rs.Fields("PrtAcctNum") = PrtAcctNum
        rs.Fields("PrtZeroBal") = PrtZeroBal
        rs.Fields("RegBraCon") = RegBraCon
        rs.Fields("RegCmp") = RegCmp
        rs.Fields("ReportDate") = ReportDate
        rs.Fields("RoundDollars") = RoundDollars
        rs.Fields("SepPage") = SepPage
        rs.Fields("StaSch") = StaSch
        rs.Fields("SupprCP") = SupprCP
        rs.Fields("UseMathRec") = UseMathRec
        rs.Fields("WidePrint") = WidePrint
        rs.Update
    End If
    
    ' Populate properties with
    ' results of the query.
    FiscalYear = rs!FiscalYear
    BeginDate = rs!BeginDate
    EndDate = rs!EndDate
    Copies = rs!Copies
    HiAccount = rs!HiAccount
    HiBranchAcct = rs!HiBranchAcct
    HiConsAcct = rs!HiConsAcct
    ID = rs!ID
    LowAccount = rs!LowAccount
    LowBranchAcct = rs!LowBranchAcct
    LowConsAcct = rs!LowConsAcct
    LowerCaseDate = rs!LowerCaseDate
    
    On Error Resume Next
    Output = rs!Output & ""
    If Err.Number = 3265 Then
        Output = rs!gOutput & ""
    End If
    On Error GoTo 0
    
    PrintBIB = rs!PrintBIB
    PrtAcctNum = rs!PrtAcctNum
    PrtZeroBal = rs!PrtZeroBal
    RegBraCon = rs!RegBraCon
    RegCmp = rs!RegCmp
    ReportDate = rs!ReportDate
    RoundDollars = rs!RoundDollars
    SepPage = rs!SepPage
    StaSch = rs!StaSch
    SupprCP = rs!SupprCP
    UseMathRec = rs!UseMathRec
    
    On Error Resume Next
    User = rs!User & ""
    If Err.Number = 3265 Then
        User = rs!gUser & ""
    End If
    On Error GoTo 0
    
    WidePrint = rs!WidePrint
    
'    rs.Close
'    Set rs = Nothing
    
End Function


Public Sub Clear()
        ID = 0
        FiscalYear = 0
        BeginDate = 0
        EndDate = 0
        HiAccount = 0
        HiBranchAcct = 0
        HiConsAcct = 0
        LowAccount = 0
        LowBranchAcct = 0
        LowConsAcct = 0
        LowerCaseDate = 0
        Output = ""
        PrtAcctNum = False
        PrtZeroBal = False
        ReportDate = False
        RoundDollars = False
        SepPage = False
        SupprCP = False
        UseMathRec = False
        User = ""
        WidePrint = False
        Copies = 0
        RegBraCon = 0
        StaSch = 0
        RegCmp = 0
        PrintBIB = 0
        
End Sub

Public Sub OpenRS()
    SQLString = "GLPrint"
    rsInit SQLString, cn, rs
End Sub

Private Sub AssignFields()

    FiscalYear = nNull(rs!FiscalYear)
    BeginDate = nNull(rs!BeginDate)
    EndDate = nNull(rs!EndDate)
    Copies = nNull(rs!Copies)
    HiAccount = nNull(rs!HiAccount)
    HiBranchAcct = nNull(rs!HiBranchAcct)
    HiConsAcct = nNull(rs!HiConsAcct)
    ID = rs!ID
    LowAccount = nNull(rs!LowAccount)
    LowBranchAcct = nNull(rs!LowBranchAcct)
    LowConsAcct = nNull(rs!LowConsAcct)
    LowerCaseDate = rs!LowerCaseDate
    PrintBIB = nNull(rs!PrintBIB)
    PrtAcctNum = rs!PrtAcctNum
    PrtZeroBal = rs!PrtZeroBal
    RegBraCon = nNull(rs!RegBraCon)
    RegCmp = nNull(rs!RegCmp)
    ReportDate = nNull(rs!ReportDate)
    RoundDollars = rs!RoundDollars
    SepPage = rs!SepPage
    StaSch = nNull(rs!StaSch)
    SupprCP = rs!SupprCP
    UseMathRec = rs!UseMathRec
    ReportName = rs!ReportName & ""
    
    ' Output or gOutput
    On Error Resume Next
    Output = Trim(rs!Output & "")
    If Err.Number = 3265 Then
        Output = Trim(rs!gOutput & "")
    End If
    On Error GoTo 0
    
    ' User or gUser
    On Error Resume Next
    User = rs!User
    If Err.Number = 3265 Then
        User = rs!gUser
    End If
    On Error GoTo 0
    
    WidePrint = rs!WidePrint

End Sub
