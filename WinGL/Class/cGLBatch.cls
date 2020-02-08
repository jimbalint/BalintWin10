VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Public ID As Long
Public FiscalYear As Integer
Public Period As Byte
Public BatchNumber As Long
Public Debits As Currency
Public Credits As Currency
Public Created As Date
Public Updated As Date
Public CreateUser As String
Public UpdateUser As String
Public Records As Long
Public JournalSource As Integer
Public RecCt As Integer

Dim SQLString As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim param As ADODB.Parameter

Dim x As String

Public Sub OpenRS()
   SQLString = "GLBatch"
   rsInit SQLString, cn, rs
End Sub

Public Sub Save(ByVal booAddPut As Boolean)

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("FiscalYear") = FiscalYear
    rs.Fields("Period") = Period
    rs.Fields("BatchNumber") = BatchNumber
    rs.Fields("Debits") = Debits
    rs.Fields("Credits") = Credits
    rs.Fields("Created") = Created
    rs.Fields("Updated") = Updated
    rs.Fields("CreateUser") = CreateUser
    rs.Fields("UpdateUser") = UpdateUser
    rs.Fields("Records") = Records
    rs.Fields("JournalSource") = JournalSource
    
    rs.Update
    rs.Close
    Set rs = Nothing
 
End Sub

Public Function GetData(lngID As Long) As Boolean

    If lngID <> 0 Then
        
        rs.Source = "SELECT * FROM GLBatch " & _
                    "WHERE ID = " & lngID
                    
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
        ID = rs!ID
        FiscalYear = rs!FiscalYear
        Period = rs!Period
        BatchNumber = rs!BatchNumber
        Debits = rs!Debits
        Credits = rs!Credits
        Created = rs!Created
        Updated = rs!Updated
        CreateUser = rs!CreateUser
        UpdateUser = rs!UpdateUser
        Records = rs!Records
        JournalSource = rs!JournalSource
        
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
        FiscalYear = 0
        Period = 0
        BatchNumber = 0
        Debits = 0
        Credits = 0
        Created = 0
        Updated = 0
        CreateUser = 0
        UpdateUser = 0
        Records = 0
        JournalSource = 0
End Sub
Private Sub AssignVars()
        ID = rs!ID
        FiscalYear = rs!FiscalYear
        Period = rs!Period
        BatchNumber = rs!BatchNumber
        Debits = rs!Debits
        Credits = rs!Credits
        Created = rs!Created
        Updated = rs!Updated
        CreateUser = rs!CreateUser
        UpdateUser = rs!UpdateUser
        Records = rs!Records
        JournalSource = rs!JournalSource
End Sub


Public Sub DeleteAll()

    x = "DELETE * FROM GLBatch"
    rsInit x, cn, rs
    
End Sub

Public Sub DeleteBatch(ByVal BatchNum As Integer)

    x = " DELETE * FROM GLBatch WHERE BatchNumber = " & BatchNum
    cn.Execute x

End Sub

Public Function GetBatch(ByVal BatchNum As Long) As Boolean

    x = "SELECT * FROM GLBatch WHERE BatchNumber = " & BatchNum
    
    rsInit x, cn, rs
    
    If rs.BOF And rs.EOF Then
       GetBatch = False
       Exit Function
    End If

    rs.MoveFirst
   
    AssignVars
   
    GetBatch = True

End Function


Public Sub CloseRS()
   rs.Close
   Set rs = Nothing
End Sub


Public Function GetByString(ByVal SString As String) As Boolean
                             
   rsInit SString, cn, rs
   RecCt = 0
   
   If rs.BOF And rs.EOF Then
      GetByString = False
      Exit Function
   End If
   
   rs.MoveFirst
   RecCt = rs.RecordCount
   
   AssignVars
   
   GetByString = True
                             
End Function


Public Sub AddBatch(ByVal FY As Integer, _
                         ByVal Pd As Byte)

    x = "SELECT * FROM GLBatch ORDER BY BatchNumber DESC"
    rsInit x, cn, rs
    
    Clear   ' reset variable values
    
    ' first batch
    If rs.BOF And rs.EOF Then
       BatchNumber = 1
    Else
       BatchNumber = rs!BatchNumber + 1
    End If
    
    FiscalYear = FY
    Period = Pd
    
    Save (Equate.RecAdd)

End Sub

Public Function GetNext() As Boolean

   rs.MoveNext
   
   If rs.EOF Then
      GetNext = False
      Exit Function
   Else
      AssignVars
      GetNext = True
   End If

End Function

