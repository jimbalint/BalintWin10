Attribute VB_Name = "modStart"
Option Explicit
    
Private Sub Main()

Dim X As String
Dim fName, DriveLetter As String

    Set Client = New clsClient
    Set Customer = New clsCustomer

    ' set global values
    Portrait = 1
    RecAdd = True
    RecPut = False

    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.ConnectionString = "\Balint\CheckData\CheckData.mdb"
    cn.Open
    
    If AddField("Customer", "AccountSpace", "Byte", cn) Then
    End If
    
    If AddField("Customer", "Sign1Left", "Long", cn) Then
    End If
    
    If AddField("Customer", "Sign1Top", "Long", cn) Then
    End If
    
    If AddField("Customer", "Sign1Height", "Long", cn) Then
    End If

    If AddField("Customer", "Sign1Width", "Long", cn) Then
    End If
    
    If AddField("Customer", "Sign2Left", "Long", cn) Then
    End If
    
    If AddField("Customer", "Sign2Top", "Long", cn) Then
    End If
    
    If AddField("Customer", "Sign2Height", "Long", cn) Then
    End If

    If AddField("Customer", "Sign2Width", "Long", cn) Then
    End If
    
    If AddField("Customer", "BankAccountAdd", "String", cn) Then
    End If
    
    If AddField("Customer", "AddressAdjust", "Long", cn) Then
    End If
    
    ' ************************************
'    ClientCreate
'    CustomerCreate
'    End
    ' ************************************

'    SQLString = "SELECT * FROM PRCompany WHERE PRCompany.GLCompanyID = " & User.LastCompany
'    If Not PRCompany.GetBySQL(SQLString) Then
'        MsgBox "PRCompany record NF: " & User.LastCompany, vbCritical
'        End
'    End If
'
'    ' open the company database
'    X = Mid(App.Path, 1, 2) & Mid(PRCompany.FileName, 3, Len(PRCompany.FileName) - 2)
'    CNOpen X, dbPwd
'    CompanyID = PRCompany.CompanyID
    
    frmCheckMaintMenu.Show
    
End Sub
