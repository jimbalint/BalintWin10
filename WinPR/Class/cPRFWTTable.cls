VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRFWTTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FWTTableID As Long
Public StateID As Long
Public TaxYear As Long
Public TaxMonth As Byte
Public msMarried As Byte
Public msSingle As Byte

Public LowAmount As Currency
Public HiAmount As Currency
Public Amount As Currency
Public Percent As Double
Public ExcessBase As Currency

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRFWTTable"
   rsInit SQLString, cnDes, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngHistID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngHistID = 0 Then
       ErrMessage = "HistID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRFWTTable where RecID = " & lngHistID
    
    rsInit SQLString, cnDes, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT EmployeeID from PRFWTTable"
    
    rsInit SQLString, cnDes, rs
    
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

Public Function Find(ByVal lngHistID As Long) As Boolean
   
   SQLString = "RecID = " & lngHistID
   
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


Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("StateID") = StateID
    rs.Fields("TaxYear") = TaxYear
    rs.Fields("TaxMonth") = TaxMonth
    rs.Fields("msMarried") = msMarried
    rs.Fields("msSingle") = msSingle
    rs.Fields("LowAmount") = LowAmount
    rs.Fields("HiAmount") = HiAmount
    rs.Fields("Amount") = Amount
    rs.Fields("Percent") = Percent
    rs.Fields("ExcessBase") = ExcessBase

    rs.Update
    
    FWTTableID = rs!FWTTableID
    
End Function

Public Sub Clear()
      
    StateID = 0
    TaxYear = 0
    TaxMonth = 0
    msMarried = 0
    msSingle = 0
    LowAmount = 0
    HiAmount = 0
    Amount = 0
    Percent = 0
    ExcessBase = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRFWTTable", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    FWTTableID = nNull(rs!FWTTableID)
    StateID = nNull(rs!StateID)
    TaxYear = nNull(rs!TaxYear)
    TaxMonth = nNull(rs!TaxMonth)
    msMarried = nNull(rs!msMarried)
    msSingle = rs!msSingle
    LowAmount = nNull(rs!LowAmount)
    HiAmount = nNull(rs!HiAmount)
    Amount = nNull(rs!Amount)
    Percent = nNull(rs!Percent)
    ExcessBase = nNull(rs!ExcessBase)
        
End Sub

Private Function MakeByte(ByVal BooVal As Boolean) As Integer
    
    If BooVal = True Then
        MakeByte = 1
    Else
        MakeByte = 0
    End If
    
End Function

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

Public Function GetFWT(ByVal StateID As Long, _
                       ByVal MarSng As String, _
                       ByVal TaxYear As Long, _
                       ByVal TaxMonth As Byte, _
                       ByVal FWTWage As Currency) As Currency
                                              
                                              
Dim rsYM, YM As Long
                                              
    If FWTWage = 0 Or TaxYear = 0 Then
        GetFWT = 0
        Exit Function
    End If
                                              
    If FWTWage > 99999999.99 Then FWTWage = 99999999.99
                                              
    ' loop thru and take the latest table entry
    ' enables entry into next period even if tax table entries
    '    don't exist yet
                                              
    If MarSng = "M" Then
        SQLString = "SELECT * FROM PRFWTTable WHERE StateID = " & CStr(StateID) & _
                    " AND msMarried " & _
                    " AND " & FWTWage & " >= LowAmount " & _
                    " AND " & FWTWage & " <= HiAmount" & _
                    " ORDER BY TaxYear, TaxMonth"
    ElseIf MarSng = "S" Then
        SQLString = "SELECT * FROM PRFWTTable WHERE StateID = " & CStr(StateID) & _
                    " AND msSingle " & _
                    " AND " & FWTWage & " >= LowAmount " & _
                    " AND " & FWTWage & " <= HiAmount" & _
                    " ORDER BY TaxYear, TaxMonth"
    ElseIf MarSng = "X" Then     ' no marital difference for Ohio
        SQLString = "SELECT * FROM PRFWTTable WHERE StateID = " & CStr(StateID) & _
                    " AND " & FWTWage & " >= LowAmount " & _
                    " AND " & FWTWage & " <= HiAmount" & _
                    " ORDER BY TaxYear, TaxMonth"
    End If

    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        MsgBox "FWT Table NF: " & StateID & " " & TaxYear & " " & FWTWage, vbExclamation
        End
    End If
    
    GetFWT = -1
    
    YM = TaxYear * 100 + TaxMonth
    
    Do
        
        rsYM = rs!TaxYear * 100 + rs!TaxMonth
        If rsYM > YM Then Exit Do
        
'        If rs!TaxYear > TaxYear Then Exit Do
'        If rs!TaxMonth > TaxMonth Then Exit Do
        
        If StateID = 36 And TaxYear >= 2019 Then
            ' 2019-12-21 use the OH multiplier
            If TaxYear >= 2019 Then
                Dim OHMultiplier As Double
                OHMultiplier = 1
                SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeOHMultiplier & _
                            " AND Year = " & TaxYear
                If PRGlobal.GetBySQL(SQLString) = True Then
                    OHMultiplier = PRGlobal.Amount
                End If

                GetFWT = Round((rs!Amount + (FWTWage - rs!ExcessBase) * rs!Percent / 100) * OHMultiplier, 2)
            Else
                GetFWT = Round(rs!Amount + (FWTWage - rs!ExcessBase) * rs!Percent / 100, 2)
            End If
        Else
            GetFWT = Round(rs!Amount + (FWTWage - rs!ExcessBase) * rs!Percent / 100, 2)
        End If

        rs.MoveNext
        If rs.EOF Then Exit Do
    Loop
        
    If GetFWT = -1 Then
        MsgBox "FWT Table NF: " & StateID & " " & TaxYear & " " & FWTWage, vbExclamation
        End
    End If

'    If MarSng = "M" Then
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = " & TaxYear & _
'                    " AND StateID = " & CStr(StateID) & _
'                    " AND msMarried " & _
'                    " AND " & FWTWage & " >= LowAmount " & _
'                    " AND " & FWTWage & " <= HiAmount"
'    ElseIf MarSng = "S" Then
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = " & TaxYear & _
'                    " AND StateID = " & CStr(StateID) & _
'                    " AND msSingle " & _
'                    " AND " & FWTWage & " >= LowAmount " & _
'                    " AND " & FWTWage & " <= HiAmount"
'    ElseIf MarSng = "X" Then     ' no marital difference for Ohio
'        SQLString = "SELECT * FROM PRFWTTable WHERE TaxYear = " & TaxYear & _
'                    " AND StateID = " & CStr(StateID) & _
'                    " AND " & FWTWage & " >= LowAmount " & _
'                    " AND " & FWTWage & " <= HiAmount"
'    End If
'
'    rsInit SQLString, cnDes, rs
'
'    If rs.BOF And rs.EOF Then
'        MsgBox "FWT Table NF: " & StateID & " " & TaxYear & " " & FWTWage, vbExclamation
'        End
'    End If
'
'    GetFWT = rs!Amount + (FWTWage - rs!ExcessBase) * rs!Percent / 100


End Function


Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRFWTTable WHERE FWTTableID = " & ID
    
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

