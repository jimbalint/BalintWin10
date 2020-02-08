Attribute VB_Name = "modStart"
Option Explicit

Private Sub Main()

Dim X As String
Dim FName, DriveLetter As String

    Set Client = New clsClient
    Set Customer = New clsCustomer
    
    RecAdd = True
    RecPut = False

    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.ConnectionString = "C:\Balint\CheckData\CheckData.mdb"
    cn.Open

    ' ************************************
    ' ClientCreate
    ' CustomerCreate
    ' ************************************

    frmTest.Show
    
End Sub
