VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GlobalID As Long
Public TypeCode As Byte
Public UserID As Long
Public Description As String
Public Amount As Currency
Public Percent As Double
Public Flag As Byte
Public Year As Long
Public Month As Byte
Public Var1 As String
Public Var2 As String
Public Var3 As String
Public Var4 As String
Public Var5 As String
Public Var6 As String
Public Var7 As String
Public Var8 As String
Public Var9 As String
Public Var10 As String

Public Byte1 As Byte
Public Byte2 As Byte
Public Byte3 As Byte
Public Byte4 As Byte
Public Byte5 As Byte
Public Byte6 As Byte
Public Byte7 As Byte
Public Byte8 As Byte
Public Byte9 As Byte
Public Byte10 As Byte

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRGlobal"
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
    
    SQLString = "Delete * from PRGlobal where RecID = " & lngHistID
    
    rsInit SQLString, cnDes, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    Records = rs.RecordCount

'    SQLString = "SELECT EmployeeID from PRGlobal"
'
'    rsInit SQLString, cnDes, rs
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
    
    rs.Fields("TypeCode") = TypeCode
    rs.Fields("UserID") = UserID
    rs.Fields("Description") = UCase(Description)
    rs.Fields("Amount") = Amount
    rs.Fields("Percent") = Percent
    rs.Fields("Flag") = Flag
    rs.Fields("Year") = Year
    rs.Fields("Month") = Month
    rs.Fields("Var1") = Var1
    rs.Fields("Var2") = Var2
    rs.Fields("Var3") = Var3
    rs.Fields("Var4") = Var4
    rs.Fields("Var5") = Var5
    rs.Fields("Var6") = Var6
    rs.Fields("Var7") = Var7
    rs.Fields("Var8") = Var8
    rs.Fields("Var9") = Var9
    rs.Fields("Var10") = Var10
    
    rs!Byte1 = Byte1
    rs!Byte2 = Byte2
    rs!Byte3 = Byte3
    rs!Byte4 = Byte4
    rs!Byte5 = Byte5
    rs!Byte6 = Byte6
    rs!Byte7 = Byte7
    rs!Byte8 = Byte8
    rs!Byte9 = Byte9
    rs!Byte10 = Byte10
    
    rs.Update
    
    GlobalID = rs!GlobalID
    
End Function

Public Sub Clear()
      
    GlobalID = 0
    TypeCode = 0
    UserID = 0
    Description = ""
    Amount = 0
    Percent = 0
    Flag = 0
    Year = 0
    Month = 0
    Var1 = ""
    Var2 = ""
    Var3 = ""
    Var4 = ""
    Var5 = ""
    Var6 = ""
    Var7 = ""
    Var8 = ""
    Var9 = ""
    Var10 = ""

    Byte1 = 0
    Byte2 = 0
    Byte3 = 0
    Byte4 = 0
    Byte5 = 0
    Byte6 = 0
    Byte7 = 0
    Byte8 = 0
    Byte9 = 0
    Byte10 = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRGlobal", cnDes, rs
    
End Sub

Public Sub AssignFields()
        
    GlobalID = nNull(rs!GlobalID)
    TypeCode = nNull(rs!TypeCode)
    UserID = nNull(rs!UserID)
    Description = UCase(rs!Description) & ""
    Amount = nNull(rs!Amount)
    Flag = nNull(rs!Flag)
    Year = nNull(rs!Year)
    Percent = nNull(rs!Percent)
    Month = nNull(rs!Month)
    Var1 = rs!Var1 & ""
    Var2 = rs!Var2 & ""
    Var3 = rs!Var3 & ""
    Var4 = rs!Var4 & ""
    Var5 = rs!Var5 & ""
    Var6 = rs!Var6 & ""
    Var7 = rs!Var7 & ""
    Var8 = rs!Var8 & ""
    Var9 = rs!Var9 & ""
    Var10 = rs!Var10 & ""

    Byte1 = nNull(rs!Byte1)
    Byte2 = nNull(rs!Byte2)
    Byte3 = nNull(rs!Byte3)
    Byte4 = nNull(rs!Byte4)
    Byte5 = nNull(rs!Byte5)
    Byte6 = nNull(rs!Byte6)
    Byte7 = nNull(rs!Byte7)
    Byte8 = nNull(rs!Byte8)
    Byte9 = nNull(rs!Byte9)
    Byte10 = nNull(rs!Byte10)

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

Public Function GetAmount(ByVal GlobalType As Byte, ByVal TaxYear As Long) As Currency

    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & GlobalType & _
                " AND PRGlobal.Year <= " & TaxYear & _
                " ORDER BY PRGlobal.Year DESC"
                
    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetAmount = 0
        Exit Function
    End If
    
    rs.MoveFirst
    GetAmount = rs!Amount
    
End Function

     

Public Function GetByID(ByVal ID As Long) As Boolean

    SQLString = "SELECT * FROM PRGlobal WHERE GlobalID = " & ID
    
    rsInit SQLString, cnDes, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

