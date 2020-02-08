VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GlobalID As Long
Public CompanyID As Long
Public UserID As Long
Public TypeCode As Byte
Public Description As String
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
Public Var1 As String
Public Var2 As String
Public Var3 As String
Public Var4 As String
Public Var5 As String

Public Records As Long

Dim rs As New ADODB.Recordset

Public Function GetBySQL(ByVal SQLString As String) As Boolean

   rsInit SQLString, cnDes, rs
   
   If rs.RecordCount = 0 Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True

End Function
Public Function GetByID(ByVal ID As Long) As Boolean
    SQLString = "SELECT * FROM InvGlobal WHERE GlobalID = " & ID
    GetByID = GetBySQL(SQLString)
End Function

Public Function OpenRS() As Boolean
    OpenRS = GetBySQL("SELECT * FROM InvGlobal")
End Function

Private Sub AssignFields()

    GlobalID = rs!GlobalID
    CompanyID = nNull(rs!CompanyID)
    UserID = nNull(rs!UserID)
    TypeCode = nNull(rs!TypeCode)
    Description = Trim(rs!Description & "")
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
    Var1 = Trim(rs!Var1 & "")
    Var2 = Trim(rs!Var2 & "")
    Var3 = Trim(rs!Var3 & "")
    Var4 = Trim(rs!Var4 & "")
    Var5 = Trim(rs!Var5 & "")

End Sub

Public Sub Clear()

    GlobalID = 0
    CompanyID = 0
    UserID = 0
    TypeCode = 0
    Description = ""
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
    Var1 = ""
    Var2 = ""
    Var3 = ""
    Var4 = ""
    Var5 = ""

End Sub

Private Sub AssignDataSet()

    rs!CompanyID = CompanyID
    rs!UserID = UserID
    rs!TypeCode = TypeCode
    rs!Description = MaxLen(Description, 255)
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
    rs!Var1 = MaxLen(Var1, 50)
    rs!Var2 = MaxLen(Var2, 50)
    rs!Var3 = MaxLen(Var3, 50)
    rs!Var4 = MaxLen(Var4, 50)
    rs!Var5 = MaxLen(Var5, 50)

End Sub

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
Public Function GetPrev() As Boolean

   rs.MovePrevious
   
   If rs.BOF Then
      GetPrev = False
      Exit Function
   Else
      AssignFields
      GetPrev = True
   End If

End Function

Public Sub rsAdd()
    
    rs.AddNew
    AssignDataSet
    rs.Update
    GlobalID = rs!GlobalID
    
End Sub

Public Sub rsPut()
    
    AssignDataSet
    rs.Update

End Sub



