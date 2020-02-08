VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public FileName As String
Public Address1 As String
Public Address2 As String
Public Address3 As String
Public City As String
Public FirstPAcct As Long
Public FirstPeriod As Byte
Public LastClose As Long
Public LastUpdate As Long
Public Name As String
Public NetProfitAcct As Long
Public ID As Long
Public NumberPds As Byte
Public PctBaseAcct As Long
Public RetEarnAcct As Long
Public State As String
Public SubDigits As Byte
Public SuspAcct As Long
Public ZipCode As Long
Public FirstFiscalYear As Integer
Public LastBatch As Long

Public LowBranch As Long
Public HiBranch As Long
Public LowConsolidated As Long
Public HiConsolidated As Long

Public CurFiscalYear As Integer
Public CurPeriod As Byte

Public FederalID As String
Public SSN As String
Public CSZ As String

Dim rs As ADODB.Recordset
Dim X As String
Dim Str As String


Public Function Save(ByVal booAddPut As Boolean) As Boolean

    Save = False

    If booAddPut = True Then      ' Add
        X = "SELECT * FROM GLCompany"
        rsInit X, cnDes, rs
        rs.AddNew
    Else                           ' Put
'        x = "SELECT * FROM GLCompany WHERE GLCompany.ID = " & ID
'        rsInit x, cnDes, rs
'        If rs.BOF And rs.EOF Then
'            MsgBox "ID not found: " & ID
'            End
'        End If
    End If
                
    rs.Fields("Address1") = Address1 & ""
    rs.Fields("Address2") = Address2 & ""
    rs.Fields("Address3") = Address3 & ""
    rs.Fields("City") = City & ""
    rs.Fields("FirstPAcct") = nNull(FirstPAcct)
    rs.Fields("FirstPeriod") = nNull(FirstPeriod)
    rs.Fields("LastClose") = nNull(LastClose)
    rs.Fields("LastUpdate") = nNull(LastUpdate)
    rs.Fields("Name") = Name & ""
    rs.Fields("NetProfitAcct") = nNull(NetProfitAcct)
    rs.Fields("NumberPds") = nNull(NumberPds)
    rs.Fields("PctBaseAcct") = nNull(PctBaseAcct)
    rs.Fields("RetEarnAcct") = nNull(RetEarnAcct)
    rs.Fields("State") = State & ""
    rs.Fields("SubDigits") = nNull(SubDigits)
    rs.Fields("SuspAcct") = nNull(SuspAcct)
    rs.Fields("ZipCode") = ZipCode & ""
    rs.Fields("FirstFiscalYear") = nNull(FirstFiscalYear)
   
    rs.Fields("LowBranch") = nNull(LowBranch)
    rs.Fields("HiBranch") = nNull(HiBranch)
    rs.Fields("LowConsolidated") = nNull(LowConsolidated)
    rs.Fields("HiConsolidated") = nNull(HiConsolidated)
    rs.Fields("FileName") = FileName & ""
    rs.Fields("LastBatch") = nNull(LastBatch)
   
    rs.Fields("CurFiscalYear") = nNull(CurFiscalYear)
    rs.Fields("CurPeriod") = nNull(CurPeriod)
   
    rs.Fields("FederalID") = FederalID & ""
    rs.Fields("SSN") = SSN & ""
   
    rs.Update
    
    ID = rs!ID
    
''    rs.Close
''    Set rs = Nothing
    
    Save = True
    
End Function

Public Function GetData(ByVal cID As Long) As Boolean
    
    X = "SELECT * FROM GLCompany WHERE GLCompany.ID = " & cID
    rsInit X, cnDes, rs
          
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
        GetData = False
        Exit Function
    End If
      
    AssignValues
    
    GetData = True
    
End Function
Public Function GetByName(ByVal FName As String) As Boolean
    
    X = "SELECT * FROM GLCompany WHERE GLCompany.FileName = '" & FName & "'"
      
    rsInit X, cnDes, rs
      
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
        GetByName = False
        Exit Function
    End If
      
    AssignValues
      
    GetByName = True
    
End Function

Public Function GetBySQL(ByVal Str As String) As Boolean

    rsInit Str, cnDes, rs
    If rs.RecordCount = 0 Then
        GetBySQL = False
        Exit Function
    End If
    AssignValues
    GetBySQL = True

End Function
Public Function GetByID(ByVal m_ID As Long) As Boolean

    Str = " SELECT * FROM GLCompany WHERE ID = " & m_ID
    rsInit Str, cnDes, rs
    If rs.RecordCount = 0 Then
        GetByID = False
        Exit Function
    End If
    AssignValues
    GetByID = True

End Function

Public Function GetNext() As Boolean

    rs.MoveNext
    If rs.EOF = True Then
        GetNext = False
        Exit Function
    End If
    GetNext = True
    AssignValues

End Function


Public Sub Clear()
      ID = 0
      Address1 = ""
      Address2 = ""
      Address3 = ""
      City = ""
      FirstPAcct = 0
      FirstPeriod = 0
      LastClose = 0
      LastUpdate = 0
      Name = ""
      NetProfitAcct = 0
      NumberPds = 0
      PctBaseAcct = 0
      RetEarnAcct = 0
      State = ""
      SubDigits = 0
      SuspAcct = 0
      ZipCode = 0
      LowBranch = 0
      HiBranch = 0
      LowConsolidated = 0
      HiConsolidated = 0
      FileName = ""
      FirstFiscalYear = 0
      LastBatch = 0
    
      CurPeriod = 0
      CurFiscalYear = 0
    
      FederalID = ""
      SSN = ""
      CSZ = ""

End Sub

Public Sub DeleteAll()

    X = "Delete * from GLCompany"
    rsInit X, cnDes, rs
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub AssignValues()
      
      ' Populate properties with
      ' results of the query.
       ID = rs!ID
       
       Address1 = rs!Address1 & ""
       Address2 = rs!Address2 & ""
       Address3 = rs!Address3 & ""
       City = rs!City & ""
       FirstPAcct = nNull(rs!FirstPAcct)
       FirstPeriod = nNull(rs!FirstPeriod)
       LastClose = nNull(rs!LastClose)
       LastUpdate = nNull(rs!LastUpdate)
       Name = rs!Name & ""
       NetProfitAcct = nNull(rs!NetProfitAcct)
       NumberPds = nNull(rs!NumberPds)
       PctBaseAcct = nNull(rs!PctBaseAcct)
       RetEarnAcct = nNull(rs!RetEarnAcct)
       State = rs!State & ""
       SubDigits = nNull(rs!SubDigits)
       SuspAcct = nNull(rs!SuspAcct)
       ZipCode = nNull(rs!ZipCode)
    
       LowBranch = nNull(rs!LowBranch)
       HiBranch = nNull(rs!HiBranch)
       LowConsolidated = nNull(rs!LowConsolidated)
       HiConsolidated = nNull(rs!HiConsolidated)
       FirstFiscalYear = nNull(rs!FirstFiscalYear)
       LastBatch = nNull(rs!LastBatch)
    
       FileName = rs!FileName & ""
    
       CurPeriod = nNull(rs!CurPeriod)
       CurFiscalYear = nNull(rs!CurFiscalYear)
       
       FederalID = rs!FederalID & ""
       SSN = rs!SSN & ""
       
        CSZ = Trim(City) & ", " & Trim(State) & "  " & Trim(ZipCode)
       
'       rs.Close
'       Set rs = Nothing

End Sub

Public Function MonthName(ByVal mon As Byte, ByVal yr As Integer) As String
    If FirstPeriod <> 1 Then
       If mon <= 12 - FirstPeriod + 1 Then yr = yr - 1
    End If
    mon = mon + FirstPeriod - 1
    If mon > 12 Then mon = mon - 12
    MonthName = Format(CStr(mon) & "/1/" & yr, "mmmm") & ", " & CStr(yr)
End Function

