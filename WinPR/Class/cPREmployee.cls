VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPREmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EmployeeID As Long
Public EmployeeNumber As Long
Public LastName As String
Public FirstName As String
Public MidInit As String
Public AltName As String
Public UseAltName As Byte
Public Address1 As String
Public Address2 As String
Public City As String
Public State As String
Public ZipCode As Long
Public SSN As Long
Public DepartmentID As Long
Public SalaryAmount As Currency
Public HourlyAmount As Currency

Public Inactive As Byte
Public Salaried As Byte

Public PaysPerYear As Byte

Public NoSSTax As Byte
Public NoMedTax As Byte
Public NoFedTax As Byte
Public NoStateTax As Byte
Public NoCityTax As Byte
Public NoFedUnemp As Byte
Public NoStateUnemp As Byte

Public FWTMarried As Byte
Public FWTBasis As Byte
Public FWTAmount As Currency
Public FWTExtraBasis As Byte
Public FWTExtraAmount As Currency
    
Public SWTMarried As Byte
Public SWTBasis As Byte
Public SWTAmount As Currency
Public SWTExtraBasis As Byte
Public SWTExtraAmount As Currency

Public DefaultCityID As Long
Public DefaultJobID As Long
Public CourtesyCityID As Long
Public CourtesyAdd As Byte

Public DateLastPaid As Date
Public DateHired As Date
Public DateLastRaise As Date
Public DateLastReview As Date
Public DateLastLayoff As Date
Public DateLastRecall As Date
Public DateTerminated As Date
Public DateOfBirth As Date

Public TermReason As Long
Public Sex As String
Public RaceCode As Long
Public MaritalStatus As String
    
Public EducationLevel As Long
Public ShiftCode As Long
    
Public x1099Employee As Byte
Public Statutory As Byte
Public EICType As Byte
Public WkcUseDept As Byte
Public WkcCat As Long

Public CheckComment As String

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PREmployee"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngEmpNumber As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngEmpNumber = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PREmployee where EmployeeNumber = " & lngEmpNumber
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT EmployeeNumber from PREmployee"
    
    rsInit SQLString, cn, rs
    
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
       Records = 0
    Else
       rs.MoveLast
       Records = rs.RecordCount
    End If

    rs.Close

End Function

Public Function Find(ByVal lngEmpNumber As Long) As Byte
   
   SQLString = "EmployeeNumber = " & lngEmpNumber
   
   rs.Find SQLString, 0, adSearchForward, 1
   
   If rs.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

Public Function FindFirst() As Byte
   
   rs.MoveFirst
   
   If rs.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
End Function


Public Function GetNext() As Byte

   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignFields
      GetNext = True
   End If

End Function


Public Function Save(ByVal booAddPut As Byte) As String
    
    If booAddPut = True Then      ' Add
       rs.AddNew
    End If

    rs.Fields("EmployeeNumber") = EmployeeNumber
    rs.Fields("LastName") = UCase(LastName)
    rs.Fields("FirstName") = UCase(FirstName)
    rs.Fields("AltName") = UCase(AltName)
    rs.Fields("UseAltName") = UseAltName
    rs.Fields("MidInit") = UCase(MidInit)
    rs.Fields("Address1") = UCase(Address1)
    rs.Fields("Address2") = UCase(Address2)
    rs.Fields("City") = UCase(City)
    rs.Fields("State") = UCase(Mid(State, 1, 2))
    rs.Fields("ZipCode") = ZipCode
    
    rs.Fields("SSN") = RC4Encrypt(CStr(SSN), rc4Key)
    
    rs.Fields("DepartmentID") = DepartmentID
    rs.Fields("SalaryAmount") = SalaryAmount
    rs.Fields("HourlyAmount") = HourlyAmount
    rs.Fields("Inactive") = Inactive
    rs.Fields("Salaried") = Salaried
    rs.Fields("NoSSTax") = NoSSTax
    rs.Fields("NoMedTax") = NoMedTax
    rs.Fields("NoFedTax") = NoFedTax
    rs.Fields("NoStateTax") = NoStateTax
    rs.Fields("NoCityTax") = NoCityTax
    rs.Fields("NoFedUnemp") = NoFedUnemp
    rs.Fields("NoStateUnemp") = NoStateUnemp
    
    rs.Fields("FWTMarried") = FWTMarried
    rs.Fields("FWTBasis") = FWTBasis
    rs.Fields("FWTAmount") = FWTAmount
    rs.Fields("FWTExtraBasis") = FWTExtraBasis
    rs.Fields("FWTExtraAmount") = FWTExtraAmount
    
    rs.Fields("SWTMarried") = SWTMarried
    rs.Fields("SWTBasis") = SWTBasis
    rs.Fields("SWTAmount") = SWTAmount
    rs.Fields("SWTExtraBasis") = SWTExtraBasis
    rs.Fields("SWTExtraAmount") = SWTExtraAmount

    rs.Fields("PaysPerYear") = PaysPerYear

    rs.Fields("DefaultCityID") = DefaultCityID
    rs.Fields("DefaultJobID") = DefaultJobID
    rs.Fields("CourtesyCityID") = CourtesyCityID
    rs.Fields("CourtesyAdd") = CourtesyAdd
 
    rs.Fields("x1099Employee") = x1099Employee
    rs.Fields("Statutory") = Statutory
    rs.Fields("EICType") = EICType
    rs.Fields("WkcUseDept") = WkcUseDept
    rs.Fields("WkcCat") = WkcCat
 
    rs.Fields("DateLastPaid") = DateLastPaid
    rs.Fields("DateHired") = DateHired
    rs.Fields("DateLastRaise") = DateLastRaise
    rs.Fields("DateLastReview") = DateLastReview
    rs.Fields("DateLastLayoff") = DateLastLayoff
    rs.Fields("DateLastRecall") = DateLastRecall
    rs.Fields("DateTerminated") = DateTerminated
    rs.Fields("DateOfBirth") = DateOfBirth

    rs.Fields("TermReason") = TermReason
    rs.Fields("Sex") = Sex
    rs.Fields("RaceCode") = RaceCode
    rs.Fields("MaritalStatus") = MaritalStatus

    rs.Fields("EducationLevel") = EducationLevel
    rs.Fields("ShiftCode") = ShiftCode

    rs.Fields("CheckComment") = CheckComment

    rs.Update
    
    EmployeeID = rs!EmployeeID
    
End Function

Public Sub Clear()
      
    EmployeeID = 0
    EmployeeNumber = 0
    LastName = ""
    FirstName = ""
    MidInit = ""
    AltName = ""
    UseAltName = 0
    Address1 = ""
    Address2 = ""
    City = ""
    State = ""
    ZipCode = 0
    SSN = 0
    DepartmentID = 0
    SalaryAmount = 0
    HourlyAmount = 0
    Inactive = 0
    Salaried = 0
    NoSSTax = 0
    NoMedTax = 0
    NoFedTax = 0
    NoStateTax = 0
    NoCityTax = 0
    NoFedUnemp = 0
    NoStateUnemp = 0
      
    FWTMarried = 0
    FWTBasis = 0
    FWTAmount = 0
    FWTExtraBasis = 0
    FWTExtraAmount = 0
    
    PaysPerYear = 0
    
    SWTMarried = 0
    SWTBasis = 0
    SWTAmount = 0
    SWTExtraBasis = 0
    SWTExtraAmount = 0

    x1099Employee = 0
    Statutory = 0
    EICType = 0
    WkcUseDept = 0
    WkcCat = 0

    DefaultCityID = 0
    CourtesyCityID = 0
    DefaultJobID = 0
    CourtesyAdd = 0

    DateLastPaid = 0
    DateHired = 0
    DateLastRaise = 0
    DateLastReview = 0
    DateLastLayoff = 0
    DateLastRecall = 0
    DateTerminated = 0
    DateOfBirth = 0

    TermReason = 0
    Sex = ""
    RaceCode = 0
    MaritalStatus = ""
    CheckComment = ""

    EducationLevel = 0
    ShiftCode = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PREmployee", cn, rs
    
End Sub

Public Sub AssignFields()
        
    EmployeeID = nNull(rs!EmployeeID)
    EmployeeNumber = nNull(rs!EmployeeNumber)
    LastName = Trim(UCase(rs!LastName)) & ""
    FirstName = Trim(UCase(rs!FirstName)) & ""
    MidInit = Trim(UCase(rs!MidInit)) & ""
    AltName = Trim(UCase(rs!AltName)) & ""
    UseAltName = nNull(rs!UseAltName)
    Address1 = Trim(UCase(rs!Address1)) & ""
    Address2 = Trim(UCase(rs!Address2)) & ""
    City = Trim(UCase(rs!City)) & ""
    State = Trim(UCase(rs!State)) & ""
    ZipCode = nNull(rs!ZipCode)
    
    SSN = RC4Decrypt(CStr(nNull(rs!SSN)), rc4Key)
    
    DepartmentID = nNull(rs!DepartmentID)
    SalaryAmount = nNull(rs!SalaryAmount)
    HourlyAmount = nNull(rs!HourlyAmount)
    Inactive = nNull(rs!Inactive)
    Salaried = nNull(rs!Salaried)
    NoSSTax = nNull(rs!NoSSTax)
    NoMedTax = nNull(rs!NoMedTax)
    NoFedTax = nNull(rs!NoFedTax)
    NoStateTax = nNull(rs!NoStateTax)
    NoCityTax = nNull(rs!NoCityTax)
    NoFedUnemp = nNull(rs!NoFedUnemp)
    NoStateUnemp = nNull(rs!NoStateUnemp)
            
    x1099Employee = nNull(rs!x1099Employee)
    Statutory = nNull(rs!Statutory)
    EICType = nNull(rs!EICType)
    WkcUseDept = nNull(rs!WkcUseDept)
    WkcCat = nNull(rs!WkcCat)
            
    FWTMarried = nNull(rs!FWTMarried)
    FWTBasis = nNull(rs!FWTBasis)
    FWTAmount = nNull(rs!FWTAmount)
    FWTExtraBasis = nNull(rs!FWTExtraBasis)
    FWTExtraAmount = nNull(rs!FWTExtraAmount)
                                            
    SWTMarried = nNull(rs!SWTMarried)
    SWTBasis = nNull(rs!SWTBasis)
    SWTAmount = nNull(rs!SWTAmount)
    SWTExtraBasis = nNull(rs!SWTExtraBasis)
    SWTExtraAmount = nNull(rs!SWTExtraAmount)
    
    PaysPerYear = nNull(rs!PaysPerYear)
                                           
    DefaultCityID = nNull(rs!DefaultCityID)
    DefaultJobID = nNull(rs!DefaultJobID)
    CourtesyCityID = nNull(rs!CourtesyCityID)
    CourtesyAdd = nNull(rs!CourtesyAdd)
                                           
    DateLastPaid = nNull(rs!DateLastPaid)
    DateHired = nNull(rs!DateHired)
    DateLastRaise = nNull(rs!DateLastRaise)
    DateLastReview = nNull(rs!DateLastReview)
    DateLastLayoff = nNull(rs!DateLastLayoff)
    DateLastRecall = nNull(rs!DateLastRecall)
    DateTerminated = nNull(rs!DateTerminated)
    DateOfBirth = nNull(rs!DateOfBirth)
                                           
    TermReason = nNull(rs!TermReason)
    Sex = rs!Sex & ""
    RaceCode = nNull(rs!RaceCode)
    MaritalStatus = rs!MaritalStatus & ""
    CheckComment = Trim(rs!CheckComment) & ""
                                           
    EducationLevel = nNull(rs!EducationLevel)
    ShiftCode = nNull(rs!ShiftCode)
                                           
End Sub

Private Function MakeByte(ByVal BooVal As Byte) As Integer
    
    If BooVal = True Then
        MakeByte = 1
    Else
        MakeByte = 0
    End If
    
End Function

Public Function GetByID(ByVal EEID As Long) As Boolean

    rsInit "SELECT * FROM PREmployee WHERE EmployeeID = " & CStr(EEID), cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cn, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
End Function

Public Function FLName() As String

    If Trim(MidInit) = "" Then
        FLName = Trim(Trim(FirstName) & " " & LastName)
    Else
        FLName = Trim(Trim(FirstName) & " " & Trim(MidInit) & " " & LastName)
    End If

End Function

Public Function LFName() As String

    If IsNull(LastName) Then LastName = " "
    If IsNull(FirstName) Then FirstName = " "

    LFName = Trim(Trim(LastName) & ", " & FirstName)
    
    If Trim(MidInit) <> "" Then
        LFName = Trim(Trim(LFName) & " " & MidInit)
    End If

End Function

Public Function CSZ() As String

    CSZ = Trim(City)
    If State <> "" Then
        CSZ = Trim(CSZ) & ", " & State
    End If
    If ZipCode <> 0 Then
        CSZ = Trim(CSZ) & "  " & Format(ZipCode, "00000")
    End If

End Function
Public Function SSString() As String
    If PREmployee.x1099Employee = PREquate.x1099Inc Then
        SSString = Format(SSN, "##-#######")     ' EIN format for 1099
    Else
        SSString = Format(SSN, "000-00-0000")
    End If
End Function


