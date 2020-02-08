VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public Account As Long
Public Amount As Currency
Public BatchNumber As Long
Public Description As String
Public ID As Long
Public JournalSource As Byte
Public Period As Byte
Public Reference As String
Public SourceCode As Byte
Public HisType As String
Public UpdateFlag As Boolean
Public FiscalYear As Integer
Public PostDate As Date
Public Records As Long

Dim SQLString As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Dim x As String

Public Sub OpenRS()
   SQLString = "GLHistory"
   rsInit SQLString, cn, rs
End Sub

Public Function GetAllByDate() As Boolean
    
    SQLString = "SELECT * FROM GLHistory ORDER BY FiscalYear, Period, PostDate"
    rsInit SQLString, cn, rs
    If rs.BOF And rs.EOF Then
       GetAllByDate = False
       Exit Function
    End If
       
    rs.MoveFirst
    AssignVars
    GetAllByDate = True
       
End Function

Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("Account") = Account
    rs.Fields("Amount") = Amount
    rs.Fields("BatchNumber") = BatchNumber
    rs.Fields("Description") = Description
    rs.Fields("JournalSource") = JournalSource
    rs.Fields("Period") = Period
    rs.Fields("Reference") = Reference
    rs.Fields("SourceCode") = SourceCode
    rs.Fields("HisType") = HisType
    rs.Fields("UpdateFlag") = UpdateFlag
    rs.Fields("FiscalYear") = FiscalYear
    rs.Fields("PostDate") = PostDate
   
    rs.Update
 
End Sub

Public Function GetData(lngID As Long) As Boolean

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If lngID <> 0 Then
        
        rs.Source = "select * from GLHistory " & _
                    "where ID = " & lngID
                    
        rs.MaxRecords = 1
        
        Set rs.ActiveConnection = cn
        
        rs.Open
        
        ' Check to make sure a record
        ' actually came back
        If rs.EOF = True And rs.BOF = True Then
           GetData = False
           Exit Function
        End If
        
        ' Populate properties with
        ' results of the query.
        Account = rs!Account
        Amount = rs!Amount
        BatchNumber = rs!BatchNumber
        Description = rs!Description & ""
        ID = rs!ID
        JournalSource = rs!JournalSource
        Period = rs!Period
        Reference = rs!Reference & ""
        SourceCode = rs!SourceCode & ""
        HisType = rs!HisType & ""
        UpdateFlag = rs!UpdateFlag
        FiscalYear = rs!FiscalYear
        PostDate = rs!PostDate
    
    Else
        GetData = False
        Exit Function
    End If
    
    GetData = True
    rs.Close
    Set rs = Nothing
    
End Function

Public Sub Clear()
        ID = 0
        Account = 0
        Amount = 0
        BatchNumber = 0
        Description = ""
        ID = 0
        JournalSource = 0
        Period = 0
        Reference = ""
        SourceCode = 0
        HisType = ""
        UpdateFlag = False
        FiscalYear = 0
        PostDate = 0
End Sub
Private Sub AssignVars()
        ID = rs!ID
        Account = rs!Account
        Amount = rs!Amount
        BatchNumber = rs!BatchNumber
        Description = rs!Description
        ID = rs!ID
        JournalSource = rs!JournalSource
        Period = rs!Period
        Reference = rs!Reference
        SourceCode = rs!SourceCode
        HisType = rs!HisType & ""
        UpdateFlag = rs!UpdateFlag
        FiscalYear = rs!FiscalYear
        PostDate = nNull(rs!PostDate)
End Sub


Public Sub DeleteAll()

    x = "DELETE * FROM GLHistory"
    rsInit x, cn, rs
    
End Sub
Public Sub DeleteBatch(ByVal BtchNum As Integer)

    x = "DELETE * FROM GLHistory WHERE BatchNum = " & BtchNum
    cn.Execute x
    
End Sub

Public Function QueryByFiscalYearByPeriod(ByVal FY As Long, ByVal Pd As Byte) As Boolean

    SQLString = "SELECT * FROM GLHistory WHERE (FiscalYear = " & FY & _
                " and Period = " & Pd & ")"
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
       QueryByFiscalYearByPeriod = False
       Exit Function
    End If
       
    rs.MoveFirst
    AssignVars
    QueryByFiscalYearByPeriod = True

    Exit Function

' stored query no longer used
'
'   Set cmd = New ADODB.Command
'   Set cmd.ActiveConnection = cn
'   cmd.CommandType = adCmdStoredProc
'   cmd.CommandText = "qryGLHistoryByFiscalYearByPeriod"
'
'   Set param = cmd.CreateParameter("qFiscalYear", adNumeric, adParamInput, , FY)
'   cmd.Parameters.Append param
'
'   Set param = cmd.CreateParameter("qPeriod", adNumeric, adParamInput, , Pd)
'   cmd.Parameters.Append param
'
'   Set rs = cmd.Execute
'
'   If rs.BOF And rs.EOF Then
'      QueryByFiscalYearByPeriod = False
'      Exit Function
'   End If
'
'   rs.MoveFirst
'
'   AssignVars
'
'   QueryByFiscalYearByPeriod = True

End Function

Public Function GetNext() As Boolean
   
   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   End If
      
   AssignVars
   
   GetNext = True
      
End Function

Public Sub CloseRS()
   rs.Close
   Set rs = Nothing
End Sub


Public Function GetByString(ByVal SString As String) As Boolean
                             
    rsInit SString, cn, rs
   
    If rs.BOF And rs.EOF Then
        GetByString = False
        Exit Function
    End If
   
    rs.MoveFirst
    Records = rs.RecordCount
   
    AssignVars
   
    GetByString = True
                             
End Function


