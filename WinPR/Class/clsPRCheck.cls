VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CustomerName As String
Public ClientID As Long
Public PRCompanyID As Long
Public Address1 As String
Public Address2 As String
Public Address3 As String
Public Address4 As String
Public Addr1Bold As Byte
Public Addr2Bold As Byte
Public Addr3Bold As Byte
Public Addr4Bold As Byte
Public Bank1 As String
Public Bank2 As String
Public Bank3 As String
Public Bank4 As String
Public BankFraction As String
Public BankABA As String
Public BankAccount As String
Public AccountSpace As Byte
Public TwoSignLines As Byte
Public Sign1Left As Long
Public Sign1Top As Long
Public Sign1Height As Long
Public Sign1Width As Long
Public Sign2Left As Long
Public Sign2Top As Long
Public Sign2Height As Long
Public Sign2Width As Long
Public SignImage1 As String
Public SignImage2 As String
Public LogoImage As String
Public CreateDate As Date
Public ModifyDate As Date

Public BankAccountAdd As String
Public AddressAdjust As Long

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRCheck"
   rsInit SQLString, cnPRCK, rs

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
    rs!CustomerName = Mid(CustomerName, 1, 40)
    rs!ClientID = ClientID
    rs!PRCompanyID = PRCompanyID
    rs!Address1 = Mid(Address1, 1, 40)
    rs!Address2 = Mid(Address2, 1, 40)
    rs!Address3 = Mid(Address3, 1, 40)
    rs!Address4 = Mid(Address4, 1, 40)
    rs!Bank1 = Mid(Bank1, 1, 40)
    rs!Bank2 = Mid(Bank2, 1, 40)
    rs!Bank3 = Mid(Bank3, 1, 40)
    rs!Bank4 = Mid(Bank4, 1, 40)
    rs!BankFraction = Mid(BankFraction, 1, 40)
    rs!BankABA = Mid(BankABA, 1, 40)
    rs!BankAccount = Mid(BankAccount, 1, 40)
    rs!AccountSpace = Mid(AccountSpace, 1, 40)
    rs!TwoSignLines = Mid(TwoSignLines, 1, 40)
    rs!SignImage1 = Mid(SignImage1, 1, 40)
    rs!Sign1Left = Mid(Sign1Left, 1, 40)
    rs!Sign1Top = Mid(Sign1Top, 1, 40)
    rs!Sign1Height = Mid(Sign1Height, 1, 40)
    rs!Sign1Width = Mid(Sign1Width, 1, 40)
    rs!SignImage2 = Mid(SignImage2, 1, 40)
    rs!Sign2Left = Mid(Sign2Left, 1, 40)
    rs!Sign2Top = Mid(Sign2Top, 1, 40)
    rs!Sign2Height = Mid(Sign2Height, 1, 40)
    rs!Sign2Width = Mid(Sign2Width, 1, 40)
    rs!LogoImage = Mid(LogoImage, 1, 40)
    rs!CreateDate = CreateDate
    rs!ModifyDate = ModifyDate
    
    rs!Addr1Bold = Addr1Bold
    rs!Addr2Bold = Addr2Bold
    rs!Addr3Bold = Addr3Bold
    rs!Addr4Bold = Addr4Bold
    
    rs!BankAccountAdd = BankAccountAdd
    rs!AddressAdjust = AddressAdjust
    
    rs.Update
    
End Function

Public Sub Clear()
      
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
    AccountSpace = 0
    TwoSignLines = 0
    SignImage1 = ""
    Sign1Left = 0
    Sign1Top = 0
    Sign1Height = 0
    Sign1Width = 0
    Sign2Left = 0
    Sign2Top = 0
    Sign2Height = 0
    Sign2Width = 0
    SignImage2 = ""
    LogoImage = ""
    CreateDate = 0
    ModifyDate = 0
      
    Addr1Bold = 0
    Addr2Bold = 0
    Addr3Bold = 0
    Addr4Bold = 0
      
    BankAccountAdd = ""
    AddressAdjust = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRCheck", cnPRCK, rs
    
End Sub

Public Sub AssignFields()
        
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
    AccountSpace = nNull(rs!AccountSpace) & ""
    TwoSignLines = nNull(rs!TwoSignLines)
    SignImage1 = rs!SignImage1 & ""
    Sign1Left = nNull(rs!Sign1Left) & ""
    Sign1Top = nNull(rs!Sign1Top) & ""
    Sign1Height = nNull(rs!Sign1Height) & ""
    Sign1Width = nNull(rs!Sign1Width) & ""
    SignImage2 = rs!SignImage2 & ""
    Sign2Left = nNull(rs!Sign2Left) & ""
    Sign2Top = nNull(rs!Sign2Top) & ""
    Sign2Height = nNull(rs!Sign2Height) & ""
    Sign2Width = nNull(rs!Sign2Width) & ""
    LogoImage = rs!LogoImage & ""
    CreateDate = nNull(rs!CreateDate)
    ModifyDate = nNull(rs!ModifyDate)
                                           
    Addr1Bold = nNull(rs!Addr1Bold)
    Addr2Bold = nNull(rs!Addr2Bold)
    Addr3Bold = nNull(rs!Addr3Bold)
    Addr4Bold = nNull(rs!Addr4Bold)
                                           
    BankAccountAdd = rs!BankAccountAdd & ""
    AddressAdjust = nNull(rs!AddressAdjust)
                                           
End Sub

Public Function GetByID(ByVal CheckID As Long) As Boolean

    rsInit "SELECT * FROM PRCheck WHERE PRCheckID = " & CStr(CheckID), cnPRCK, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    
    AssignFields
    
    GetByID = True

End Function

Public Function GetBySQL(ByVal SString As String) As Boolean
                             
   rsInit SString, cnPRCK, rs
   
   If rs.BOF And rs.EOF Then
      GetBySQL = False
      Exit Function
   End If
   
   rs.MoveFirst
   
   AssignFields
   
   GetBySQL = True
                             
End Function


