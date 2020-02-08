VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    
Public CompanyID As Long
Public Name As String
Public Address1 As String
Public Address2 As String
Public City As String
Public AddrStateID As Long
Public ZipCode As Long
Public PhoneNumber As String
Public StateID As String
Public StateUnempPct As Double
Public StateUnempID As String
Public FederalID As String
Public FederalUnempPct As Double
Public DfltPaysPerYear As Long
Public DfltStateID As Long
Public DfltMinWage As Currency
Public DfltOTRate As Currency
Public DfltRegHrs As Currency
Public FileName As String
       
Public GLAcctSS As Long
Public GLAcctMED As Long
Public GLAcctFWT As Long
Public GLAcctSWT As Long
Public GLAcctCWT As Long
Public GLAcctGross As Long
Public GLAcctNet As Long
    
Public BankName As String
Public BankABA As String
Public BankAccount As String
Public BankAddr1 As String
Public BankAddr2 As String
Public BankFraction As String
    
Public LastCheckNum As Long
Public DfltCityID As Long
Public CheckDays As Long
    
Public WkcPolicyNum As String
Public GLCompanyID As Long
    
Public DfltSortOrder As Byte
Public DirDepBalanced As Byte
    
Public DirDepUseAltID As Byte
Public DirDepAltID As Long
Public DirDepID1 As Byte
    
Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRCompany"
   rsInit SQLString, cnDes, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("Name") = UCase(Name)
    rs.Fields("Address1") = UCase(Address1)
    rs.Fields("Address2") = UCase(Address2)
    rs.Fields("City") = UCase(City)
    rs.Fields("AddrStateID") = AddrStateID
    rs.Fields("ZipCode") = ZipCode
    rs.Fields("PhoneNumber") = PhoneNumber
    rs.Fields("StateID") = StateID
    rs.Fields("StateUnempPct") = StateUnempPct
    rs.Fields("StateUnempID") = StateUnempID
    rs.Fields("FederalID") = FederalID
    rs.Fields("FederalUnempPct") = FederalUnempPct
    rs.Fields("DfltPaysPerYear") = DfltPaysPerYear
    rs.Fields("DfltStateID") = DfltStateID
    rs.Fields("DfltMinWage") = DfltMinWage
    rs.Fields("DfltOTRate") = DfltOTRate
    rs.Fields("DfltRegHrs") = DfltRegHrs
    rs.Fields("FileName") = FileName
    
    rs.Fields("BankName") = BankName
    rs.Fields("BankABA") = BankABA
    rs.Fields("BankAccount") = BankAccount
    rs.Fields("BankAddr1") = BankAddr1
    rs.Fields("BankAddr2") = BankAddr2
    rs.Fields("BankFraction") = BankFraction
    
    rs.Fields("GLAcctSS") = GLAcctSS
    rs.Fields("GLAcctMED") = GLAcctMED
    rs.Fields("GLAcctFWT") = GLAcctFWT
    rs.Fields("GLAcctSWT") = GLAcctSWT
    rs.Fields("GLAcctCWT") = GLAcctCWT
    rs.Fields("GLAcctGross") = GLAcctGross
    rs.Fields("GLAcctNet") = GLAcctNet
    
    rs.Fields("LastCheckNum") = LastCheckNum
    rs.Fields("DfltCityID") = DfltCityID
    rs.Fields("CheckDays") = CheckDays
    
    rs.Fields("WkcPolicyNum") = WkcPolicyNum
    rs.Fields("GLCompanyID") = GLCompanyID
    rs.Fields("DfltSortOrder") = DfltSortOrder
    rs.Fields("DirDepBalanced") = DirDepBalanced
    
    rs.Fields("DirDepUseAltID") = DirDepUseAltID
    rs.Fields("DirDepAltID") = DirDepAltID
    rs.Fields("DirDepID1") = DirDepID1
    
    rs.Update
    
    CompanyID = rs!CompanyID
    
End Function

Public Sub Clear()
      
    CompanyID = 0
    Name = ""
    Address1 = ""
    Address2 = ""
    City = ""
    AddrStateID = 0
    ZipCode = 0
    PhoneNumber = ""
    StateID = ""
    StateUnempPct = 0
    StateUnempID = ""
    FederalID = ""
    FederalUnempPct = 0
    DfltPaysPerYear = 0
    DfltStateID = 0
    DfltMinWage = 0
    DfltOTRate = 0
    DfltRegHrs = 0
    FileName = ""
      
    BankName = ""
    BankABA = ""
    BankAccount = ""
    BankAddr1 = ""
    BankAddr2 = ""
    BankFraction = ""
      
    GLAcctSS = 0
    GLAcctMED = 0
    GLAcctFWT = 0
    GLAcctSWT = 0
    GLAcctCWT = 0
    GLAcctGross = 0
    GLAcctNet = 0
      
    LastCheckNum = 0
    DfltCityID = 0
    CheckDays = 0
      
    WkcPolicyNum = ""
    GLCompanyID = 0
    DfltSortOrder = 0
    DirDepBalanced = 0
      
    DirDepUseAltID = 0
    DirDepAltID = 0
    DirDepID1 = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRCompany", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    CompanyID = nNull(rs!CompanyID)
    Name = Trim(UCase(rs!Name)) & ""
    Address1 = Trim(UCase(rs!Address1)) & ""
    Address2 = Trim(UCase(rs!Address2)) & ""
    City = Trim(UCase(rs!City)) & ""
    AddrStateID = nNull(rs!AddrStateID)
    ZipCode = nNull(rs!ZipCode)
    PhoneNumber = rs!PhoneNumber & ""
    StateID = Trim(rs!StateID) & ""
    StateUnempPct = nNull(rs!StateUnempPct)
    StateUnempID = rs!StateUnempID & ""
    FederalID = Trim(rs!FederalID) & ""
    FederalUnempPct = nNull(rs!FederalUnempPct)
    DfltPaysPerYear = nNull(rs!DfltPaysPerYear)
    DfltStateID = nNull(rs!DfltStateID)
        
    DfltMinWage = nNull(rs!DfltMinWage)
    DfltOTRate = nNull(rs!DfltOTRate)
    DfltRegHrs = nNull(rs!DfltRegHrs)
    
    FileName = Trim(rs!FileName) & ""

    BankName = Trim(UCase(rs!BankName)) & ""
    BankABA = Trim(UCase(rs!BankABA)) & ""
    BankAccount = Trim(UCase(rs!BankAccount)) & ""
    BankAddr1 = Trim(UCase(rs!BankAddr1)) & ""
    BankAddr2 = Trim(UCase(rs!BankAddr2)) & ""
    BankFraction = Trim(UCase(rs!BankFraction)) & ""

    GLAcctSS = nNull(rs!GLAcctSS)
    GLAcctMED = nNull(rs!GLAcctMED)
    GLAcctFWT = nNull(rs!GLAcctFWT)
    GLAcctSWT = nNull(rs!GLAcctSWT)
    GLAcctCWT = nNull(rs!GLAcctCWT)
    GLAcctGross = nNull(rs!GLAcctGross)
    GLAcctNet = nNull(rs!GLAcctNet)

    LastCheckNum = nNull(rs!LastCheckNum)
    DfltCityID = nNull(rs!DfltCityID)
    CheckDays = nNull(rs!CheckDays)

    WkcPolicyNum = rs!WkcPolicyNum & ""
    GLCompanyID = nNull(rs!GLCompanyID)
    DfltSortOrder = nNull(rs!DfltSortOrder)
    DirDepBalanced = nNull(rs!DirDepBalanced)

    DirDepUseAltID = nNull(rs!DirDepUseAltID)
    DirDepAltID = nNull(rs!DirDepAltID)
    DirDepID1 = nNull(rs!DirDepID1)

End Sub

Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cnDes, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
End Function

Public Sub UpdateBatch()

    ' update disconnect record set
    rsSave rs, cnDes

End Sub
Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRCompany WHERE CompanyID = " & CStr(ID)
    
    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function


Public Function GetByFileName(ByVal FName As String) As Long
            
Dim GLC, PRC As String
            
    GetByFileName = 0
            
    ' logic takes the drive letter out of consideration
    GLC = UCase(Mid(FName, 3, Len(Trim(FName)) - 2))
    SQLString = "SELECT * FROM PRCompany"
    rsInit SQLString, cnDes, rs
    If rs.RecordCount = 0 Then
        Exit Function
    End If
    rs.MoveFirst
    Do
        If Not (IsNull(rs!FileName)) And Len(Trim(rs!FileName)) > 3 Then
            PRC = UCase(Mid(rs!FileName, 3, Len(Trim(rs!FileName)) - 2))
            If PRC = GLC Then
                GetByFileName = rs!CompanyID
                Exit Function
            End If
        End If
        rs.MoveNext
    Loop Until rs.EOF
    
End Function


Public Function CSZ() As String

Dim StateAbbrv As String
        
    If PRState.GetByID(AddrStateID) Then
        StateAbbrv = PRState.StateAbbrev
    Else
        StateAbbrv = ""
    End If
        
    CSZ = Trim(City) & ", " & StateAbbrv & "  " & ZipCode

End Function
