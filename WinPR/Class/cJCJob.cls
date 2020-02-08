VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cJCJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public JobID As Long
Public Name As String
Public FullName As String
Public CompanyName As String
Public QBID As String
Public QBParentID As String
Public ParentID As Long
Public CityID As Long
Public FirstName As String
Public LastName As String
Public MidInit As String
Public BillAddr1 As String
Public BillAddr2 As String
Public BillAddr3 As String
Public BillAddr4 As String
Public BillCity As String
Public BillState As String
Public BillZip As String
Public ShipAddr1 As String
Public ShipAddr2 As String
Public ShipAddr3 As String
Public ShipAddr4 As String
Public ShipCity As String
Public ShipState As String
Public ShipZip As String

Public Status As String
Public StartDate As Date
Public EndDate As Date
Public Description As String
Public TypeName As String
Public TypeListID As String
Public JobStatus As Byte
    
Public QBTaxCode As String
Public Active As Byte
Public Terms As String
    
Public Records As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM JCJob"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngJobID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngJobID = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from JCJob where JobID = " & lngJobID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

'Public Function Records() As Long
'
'    SQLString = "SELECT JobID from JCJob"
'
'    rsInit SQLString, cn, rs
'
'    ' Check to make sure a record
'    ' actually came back
'    If rs.EOF = True And rs.BOF = True Then
'       Records = 0
'    Else
'       rs.MoveLast
'       Records = rs.RecordCount
'    End If
'
'    rs.Close
'
'End Function

Public Function Find(ByVal lngJobID As Long) As Boolean
   
   SQLString = "JobID = " & lngJobID
   
   rs.Find SQLString, 0, adSearchForward, 1
   
   If rs.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

Public Function FindFirst() As Boolean
   
   rs.MoveFirst
   
   If rs.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
End Function

Public Function GetNext() As Boolean

   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignFields
      GetNext = True
   End If

End Function
Public Function GetFirst() As Boolean

    If rs.RecordCount = 0 Then
        GetFirst = False
    Else
        rs.MoveFirst
        AssignFields
        GetFirst = True
    End If

End Function


Public Function Save(ByVal booAddPut As Byte) As Boolean

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs!Name = Trim(Name) & ""
    rs!FullName = Trim(FullName) & ""
    rs!CompanyName = Trim(CompanyName) & ""
    rs!QBID = Trim(QBID) & ""
    rs!QBParentID = Trim(QBParentID) & ""
    rs!ParentID = nNull(ParentID)
    rs!CityID = nNull(CityID)
    rs!FirstName = Trim(FirstName) & ""
    rs!LastName = Trim(LastName) & ""
    rs!MidInit = Trim(MidInit) & ""
    rs!BillAddr1 = Trim(BillAddr1) & ""
    rs!BillAddr2 = Trim(BillAddr2) & ""
    rs!BillAddr3 = Trim(BillAddr3) & ""
    rs!BillAddr4 = Trim(BillAddr4) & ""
    rs!BillCity = Trim(BillCity) & ""
    rs!BillState = Trim(BillState) & ""
    rs!BillZip = Trim(BillZip) & ""
    rs!ShipAddr1 = Trim(ShipAddr1) & ""
    rs!ShipAddr2 = Trim(ShipAddr2) & ""
    rs!ShipAddr3 = Trim(ShipAddr3) & ""
    rs!ShipAddr4 = Trim(ShipAddr4) & ""
    rs!ShipCity = Trim(ShipCity) & ""
    rs!ShipState = Trim(ShipState) & ""
    rs!ShipZip = Trim(ShipZip) & ""
    rs!Status = Trim(Status) & ""
    rs!JobStatus = nNull(JobStatus)
    rs!StartDate = nNull(StartDate)
    rs!EndDate = nNull(EndDate)
    rs!Description = Trim(Description) & ""
    rs!TypeName = Trim(TypeName) & ""
    rs!TypeListID = Trim(TypeListID) & ""
    
    rs!QBTaxCode = Trim(QBTaxCode) & ""
    rs!Active = nNull(Active)
    rs!Terms = Trim(Terms & "")
    
    rs.Update
    
    JobID = rs!JobID
    
End Function

Public Sub Clear()
      
    JobID = 0
    Name = ""
    FullName = ""
    CompanyName = ""
    QBID = ""
    QBParentID = ""
    ParentID = 0
    CityID = 0
    FirstName = ""
    LastName = ""
    MidInit = ""
    BillAddr1 = ""
    BillAddr2 = ""
    BillAddr3 = ""
    BillAddr4 = ""
    BillCity = ""
    BillState = ""
    BillZip = ""
    ShipAddr1 = ""
    ShipAddr2 = ""
    ShipAddr3 = ""
    ShipAddr4 = ""
    ShipCity = ""
    ShipState = ""
    ShipZip = ""
      
    Status = ""
    JobStatus = 0
    StartDate = 0
    EndDate = 0
    Description = ""
    TypeName = ""
    TypeListID = ""
    Active = 0
    
    QBTaxCode = ""
    Terms = ""
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM JCJob", cn, rs
    
End Sub

Public Sub AssignFields()
        
    JobID = nNull(rs!JobID)
    Name = Trim(rs!Name & "")
    FullName = Trim(rs!FullName & "")
    CompanyName = Trim(rs!CompanyName & "")
    QBID = Trim(rs!QBID & "")
    QBParentID = Trim(rs!QBParentID & "")
    ParentID = nNull(rs!ParentID)
    CityID = nNull(rs!CityID)
    FirstName = Trim(rs!FirstName) & ""
    LastName = Trim(rs!LastName & "")
    MidInit = Trim(rs!MidInit & "")
    BillAddr1 = Trim(rs!BillAddr1 & "")
    BillAddr2 = Trim(rs!BillAddr2 & "")
    BillAddr3 = Trim(rs!BillAddr3 & "")
    BillAddr4 = Trim(rs!BillAddr4 & "")
    BillCity = Trim(rs!BillCity & "")
    BillState = Trim(rs!BillState & "")
    BillZip = Trim(rs!BillZip & "")
    ShipAddr1 = Trim(rs!ShipAddr1 & "")
    ShipAddr2 = Trim(rs!ShipAddr2 & "")
    ShipAddr3 = Trim(rs!ShipAddr3 & "")
    ShipAddr4 = Trim(rs!ShipAddr4 & "")
    ShipCity = Trim(rs!ShipCity & "")
    ShipState = Trim(rs!ShipState & "")
    ShipZip = Trim(rs!ShipZip & "")
                                           
    Status = Trim(rs!Status) & ""
    JobStatus = nNull(rs!JobStatus)
    StartDate = nNull(rs!StartDate)
    EndDate = nNull(rs!EndDate)
    Description = Trim(rs!Description & "")
    TypeName = Trim(rs!TypeName & "")
    TypeListID = Trim(rs!TypeListID & "")
    QBTaxCode = Trim(rs!QBTaxCode & "")
                                           
    Active = nNull(rs!Active)
    Terms = Trim(rs!Terms & "")

End Sub

Public Function GetByID(ByVal JobID As Long) As Boolean

    rsInit "SELECT * FROM JCJob WHERE JobID = " & JobID, cn, rs
    
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
   
   If rs.RecordCount = 0 Then
      GetBySQL = False
      Exit Function
   End If
   
   Records = rs.RecordCount
       
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
End Function

Public Function GetByQBID(ByVal strQBID As String) As Boolean

    rsInit "SELECT * FROM JCJob WHERE QBID = '" & strQBID & "'", cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByQBID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByQBID = True

End Function

Public Sub rsClose()

    On Error Resume Next
    rs.Close
    On Error GoTo 0

End Sub

