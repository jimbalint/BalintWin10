VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRTimeSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TimeSheetID As Long
Public EmployeeID As Long
Public JobID As Long
Public CityID As Long
Public DepartmentID As Long
Public ItemID As Long
Public Note As String
Public SunHours As Single
Public MonHours As Single
Public TueHours As Single
Public WedHours As Single
Public ThuHours As Single
Public FriHours As Single
Public SatHours As Single
Public TotalHours As Single
Public HistID As Long
Public BatchID As Long
Public PEDate As Date
Public CheckDate As Date
Public WEDate As Date
Public CustomerID As Long
Public BillingRate As Currency
Public QBInvID As String

Public PWCraftID As Long
Public PWUnionID As Long
Public PWRegRate As Currency
Public PWOvtRate As Currency
Public PWFringeAmt As Currency

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRTimeSheet"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngTimeSheetID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngTimeSheetID = 0 Then
       ErrMessage = "TimeSheetID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRTimeSheet where TimeSheetID = " & lngTimeSheetID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT TimeSheetID from PRTimeSheet"
    
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

Public Function Find(ByVal lngTimeSheetID As Long) As Byte
   
   SQLString = "TimeSheetID = " & lngTimeSheetID
   
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


Public Function Save(ByVal booAddPut As Byte) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs!EmployeeID = nNull(EmployeeID)
    rs!JobID = nNull(JobID)
    rs!CityID = nNull(CityID)
    rs!DepartmentID = nNull(DepartmentID)
    rs!ItemID = nNull(ItemID)
    rs!Note = Note & ""
    rs!SunHours = nNull(SunHours)
    rs!MonHours = nNull(MonHours)
    rs!TueHours = nNull(TueHours)
    rs!WedHours = nNull(WedHours)
    rs!ThuHours = nNull(ThuHours)
    rs!FriHours = nNull(FriHours)
    rs!SatHours = nNull(SatHours)
    rs!TotalHours = nNull(TotalHours)
    rs!HistID = nNull(HistID)
    rs!BatchID = nNull(BatchID)
    rs!PEDate = nNull(PEDate)
    rs!CheckDate = nNull(CheckDate)
    rs!WEDate = nNull(WEDate)
    rs!CustomerID = nNull(CustomerID)
    rs!BillingRate = nNull(BillingRate)
    rs!QBInvID = Trim(QBInvID & "")
    
    rs!PWCraftID = PWCraftID
    rs!PWUnionID = PWUnionID
    rs!PWRegRate = PWRegRate
    rs!PWOvtRate = PWOvtRate
    rs!PWFringeAmt = PWFringeAmt
    
    rs.Update
    
    TimeSheetID = rs!TimeSheetID
    
End Function

Public Sub Clear()
      
    EmployeeID = 0
    JobID = 0
    CityID = 0
    DepartmentID = 0
    ItemID = 0
    Note = ""
    SunHours = 0
    MonHours = 0
    TueHours = 0
    WedHours = 0
    ThuHours = 0
    FriHours = 0
    SatHours = 0
    TotalHours = 0
    HistID = 0
    BatchID = 0
    PEDate = 0
    CheckDate = 0
    WEDate = 0
    CustomerID = 0
    BillingRate = 0
    QBInvID = ""
      
    PWCraftID = 0
    PWUnionID = 0
    PWRegRate = 0
    PWOvtRate = 0
    PWFringeAmt = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRTimeSheet", cn, rs
    
End Sub

Public Sub AssignFields()
        
    TimeSheetID = nNull(rs!TimeSheetID)
    EmployeeID = nNull(rs!EmployeeID)
    JobID = nNull(rs!JobID)
    CityID = nNull(rs!CityID)
    DepartmentID = nNull(rs!DepartmentID)
    ItemID = nNull(rs!ItemID)
    Note = Trim(rs!Note & "")
    SunHours = nNull(rs!SunHours)
    MonHours = nNull(rs!MonHours)
    TueHours = nNull(rs!TueHours)
    WedHours = nNull(rs!WedHours)
    ThuHours = nNull(rs!ThuHours)
    FriHours = nNull(rs!FriHours)
    SatHours = nNull(rs!SatHours)
    TotalHours = nNull(rs!TotalHours)
    HistID = nNull(rs!HistID)
    BatchID = nNull(rs!BatchID)
    PEDate = nNull(rs!PEDate)
    CheckDate = nNull(rs!CheckDate)
    WEDate = nNull(rs!WEDate)
    CustomerID = nNull(rs!CustomerID)
    BillingRate = nNull(rs!BillingRate)
    QBInvID = Trim(rs!QBInvID & "")
                                           
    PWCraftID = nNull(rs!PWCraftID)
    PWUnionID = nNull(rs!PWUnionID)
    PWRegRate = nNull(rs!PWRegRate)
    PWOvtRate = nNull(rs!PWOvtRate)
    PWFringeAmt = nNull(rs!PWFringeAmt)
                                           
End Sub

Public Function GetByID(ByVal lngTimeSheetID As Long) As Boolean

    rsInit "SELECT * FROM PRTimeSheet WHERE TimeSheetID = " & lngTimeSheetID, cn, rs
    
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

Public Function FirstDate() As Date

    FirstDate = 0
    
    If SunHours <> 0 Then
        FirstDate = WEDate - 6
        Exit Function
    End If
    
    If MonHours <> 0 Then
        FirstDate = WEDate - 5
        Exit Function
    End If
    
    If TueHours <> 0 Then
        FirstDate = WEDate - 4
        Exit Function
    End If
    
    If WedHours <> 0 Then
        FirstDate = WEDate - 3
        Exit Function
    End If
    
    If ThuHours <> 0 Then
        FirstDate = WEDate - 2
        Exit Function
    End If
    
    If FriHours <> 0 Then
        FirstDate = WEDate - 1
        Exit Function
    End If
    
    If SatHours <> 0 Then
        FirstDate = WEDate
        Exit Function
    End If

End Function

Public Function LastDate() As Date

    LastDate = 0
    
    If SatHours <> 0 Then
        LastDate = WEDate
        Exit Function
    End If
    
    If FriHours <> 0 Then
        LastDate = WEDate - 1
        Exit Function
    End If
    
    If ThuHours <> 0 Then
        LastDate = WEDate - 2
        Exit Function
    End If
    
    If WedHours <> 0 Then
        LastDate = WEDate - 3
        Exit Function
    End If
    
    If TueHours <> 0 Then
        LastDate = WEDate - 4
        Exit Function
    End If
    
    If MonHours <> 0 Then
        LastDate = WEDate - 5
        Exit Function
    End If
    
    If SunHours <> 0 Then
        LastDate = WEDate - 6
        Exit Function
    End If

End Function

