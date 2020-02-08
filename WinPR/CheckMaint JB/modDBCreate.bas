Attribute VB_Name = "modDBCreate"
Option Explicit
Public Sub ClientCreate()

    If MsgBox("OK to create Client file?", vbExclamation + vbYesNo) = vbNo Then Exit Sub

    SQLString = "CREATE TABLE Client ( " & _
                        "[ClientID] Counter, CONSTRAINT cliIDKey PRIMARY KEY ([ClientID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "Client", "ClientName", "Char (255)", cn
    AddField "Client", "Prefix", "Char (255)", cn
    AddField "Client", "Contact", "Char (255)", cn
    AddField "Client", "Phone", "Char (255)", cn
    AddField "Client", "Message1", "Char (255)", cn
    AddField "Client", "Message2", "Char (255)", cn
    AddField "Client", "Message3", "Char (255)", cn
        
End Sub

Public Sub CustomerCreate()

    If MsgBox("OK to create Customer file?", vbExclamation + vbYesNo) = vbNo Then Exit Sub

    SQLString = "CREATE TABLE Customer ( " & _
                        "[CustomerID] Counter, CONSTRAINT cusIDKey PRIMARY KEY ([CustomerID]) ) "
                        
    cn.Execute SQLString
                        
    AddField "Customer", "CustomerName", "Char (40)", cn
    AddField "Customer", "ClientID", "Long", cn
    AddField "Customer", "PRCompanyID", "Long", cn
    AddField "Customer", "Address1", "Char (40)", cn
    AddField "Customer", "Address2", "Char (40)", cn
    AddField "Customer", "Address3", "Char (40)", cn
    AddField "Customer", "Address4", "Char (40)", cn
    AddField "Customer", "Bank1", "Char (40)", cn
    AddField "Customer", "Bank2", "Char (40)", cn
    AddField "Customer", "Bank3", "Char (40)", cn
    AddField "Customer", "Bank4", "Char (40)", cn
    AddField "Customer", "BankFraction", "Char (40)", cn
    AddField "Customer", "BankABA", "Char (9)", cn
    AddField "Customer", "BankAccount", "Char (40)", cn
    AddField "Customer", "TwoSignLines", "Byte", cn
    AddField "Customer", "SignImage1", "Char (40)", cn
    AddField "Customer", "SignImage2", "Char (40)", cn
    AddField "Customer", "LogoImage", "Char (40)", cn
    AddField "Customer", "CreateDate", "DateTime", cn
    AddField "Customer", "ModifyDate", "DateTime", cn
        
End Sub

Private Sub AddField(ByVal TableName As String, _
                     ByVal FieldName As String, _
                     ByVal FieldType As String, _
                     ByVal acn As ADODB.Connection)
                     
    SQLString = "ALTER TABLE " & TableName & _
              " ADD COLUMN [" & FieldName & "]   " & FieldType
    
    acn.Execute SQLString
                     
End Sub


