VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CustomerID As Long
Public CustomerName As String
Public ClientID As Long
Public PRCompanyID As Long
Public Address1 As String
Public Address2 As String
Public Address3 As String
Public Address4 As String
Public Bank1 As String
Public Bank2 As String
Public Bank3 As String
Public Bank4 As String
Public BankFraction As String
Public BankABA As String
Public BankAccount As String
Public TwoSignLines As Byte
Public SignImage1 As String
Public SignImage2 As String
Public LogoImage As String
Public CreateDate As Date
Public ModifyDate As Date

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM Customer"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

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
    
    rs!CustomerName = CustomerName
    rs!ClientID = ClientID
    rs!PRCompanyID = PRCompanyID
    rs!Address1 = Address1
    rs!Address2 = Address2
    rs!Address3 = Address3
    rs!Address4 = Address4
    rs!Bank1 = Bank1
    rs!Bank2 = Bank2
    rs!Bank3 = Bank3
    rs!Bank4 = Bank4
    rs!BankFraction = BankFraction
    rs!BankABA = BankABA
    rs!BankAccount = BankAccount
    rs!TwoSignLines = TwoSignLines
    rs!SignImage1 = SignImage1
    rs!SignImage2 = SignImage2
    rs!LogoImage = LogoImage
    rs!CreateDate = CreateDate
    rs!ModifyDate = ModifyDate
    
    rs.Update
    
    CustomerID = rs!CustomerID
    
End Function

Public Sub Clear()
      
    CustomerID = 0
    CustomerName = ""
    ClientID = 0
    PRCompanyID = 0
    Address1 = ""
    Address2 = ""
    Address3 = ""
    Address4 = ""
    Bank1 = ""
    Bank2 = ""
    Bank3 = ""
    Bank4 = ""
    BankFraction = ""
    BankABA = ""
    BankAccount = ""
    TwoSignLines = 0
    SignImage1 = ""
    SignImage2 = ""
    LogoImage = ""
    CreateDate = 0
    ModifyDate = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM Customer", cn, rs
    
End Sub

Public Sub AssignFields()
        
    CustomerID = nNull(rs!CustomerID)
    CustomerName = rs!CustomerName & ""
    ClientID = nNull(rs!ClientID) & ""
    PRCompanyID = nNull(rs!PRCompanyID)
    Address1 = rs!Address1 & ""
    Address2 = rs!Address2 & ""
    Address3 = rs!Address3 & ""
    Address4 = rs!Address4 & ""
    Bank1 = rs!Bank1 & ""
    Bank2 = rs!Bank2 & ""
    Bank3 = rs!Bank3 & ""
    Bank4 = rs!Bank4 & ""
    BankFraction = rs!BankFraction & ""
    BankABA = rs!BankABA & ""
    BankAccount = rs!BankAccount & ""
    TwoSignLines = nNull(rs!TwoSignLines)
    SignImage1 = rs!SignImage1 & ""
    SignImage2 = rs!SignImage2 & ""
    LogoImage = rs!LogoImage & ""
    CreateDate = nNull(rs!CreateDate)
    ModifyDate = nNull(rs!ModifyDate)
                                           
End Sub

Public Function GetByID(ByVal CustID As Long) As Boolean

    rsInit "SELECT * FROM Customer WHERE CustomerID = " & CStr(CustID), cn, rs
    
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


