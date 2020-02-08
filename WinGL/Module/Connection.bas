Attribute VB_Name = "modConnection"
Option Explicit
Public cn As ADODB.Connection
   
Public Sub CNOpen()
   Set cn = New ADODB.Connection
   cn.Provider = "Microsoft.Jet.OLEDB.3.51"
   cn.ConnectionString = "c:\jbdata\gltest.mdb"
   cn.Open
End Sub

