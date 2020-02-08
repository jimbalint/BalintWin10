VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, Account As Long
Public fiscalYear As Integer
Public Amount01 As Currency
Public Amount02 As Currency
Public Amount03 As Currency
Public Amount04 As Currency
Public Amount05 As Currency
Public Amount06 As Currency
Public Amount07 As Currency
Public Amount08 As Currency
Public Amount09 As Currency
Public Amount10 As Currency
Public Amount11 As Currency
Public Amount12 As Currency
Public Amount13 As Currency
Public Budget01 As Currency
Public Budget02 As Currency
Public Budget03 As Currency
Public Budget04 As Currency
Public Budget05 As Currency
Public Budget06 As Currency
Public Budget07 As Currency
Public Budget08 As Currency
Public Budget09 As Currency
Public Budget10 As Currency
Public Budget11 As Currency
Public Budget12 As Currency
Public Budget13 As Currency
Public TotalAmount As Currency
Public TotalBudget As Currency

Public Function GetAmount(ByVal Account As Long, ByVal fiscalYear As Integer, ByVal DataFile As String) As Boolean
    GetAmount = False
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("SELECT * FROM glAmount WHERE Account=" & CStr(Account) & " AND FiscalYear=" & CStr(fiscalYear))
    With rs
        If rs.RecordCount = 0 Then
            .AddNew
            !Account = 0
            !fiscalYear = 0
            !Amount01 = 0
            !Amount02 = 0
            !Amount03 = 0
            !Amount04 = 0
            !Amount05 = 0
            !Amount06 = 0
            !Amount07 = 0
            !Amount08 = 0
            !Amount09 = 0
            !Amount10 = 0
            !Amount11 = 0
            !Amount12 = 0
            !Amount13 = 0
            !Budget01 = 0
            !Budget02 = 0
            !Budget03 = 0
            !Budget04 = 0
            !Budget05 = 0
            !Budget06 = 0
            !Budget07 = 0
            !Budget08 = 0
            !Budget09 = 0
            !Budget10 = 0
            !Budget11 = 0
            !Budget12 = 0
            !Budget13 = 0
            .Update
        End If
        rs.MoveFirst
        ID = !ID
        Account = !Account
        fiscalYear = !fiscalYear
        Amount01 = !Amount01
        Amount02 = !Amount02
        Amount03 = !Amount03
        Amount04 = !Amount04
        Amount05 = !Amount05
        Amount06 = !Amount06
        Amount07 = !Amount07
        Amount08 = !Amount08
        Amount09 = !Amount09
        Amount10 = !Amount10
        Amount11 = !Amount11
        Amount12 = !Amount12
        Amount13 = !Amount13
        Budget01 = !Budget01
        Budget02 = !Budget02
        Budget03 = !Budget03
        Budget04 = !Budget04
        Budget05 = !Budget05
        Budget06 = !Budget06
        Budget07 = !Budget07
        Budget08 = !Budget08
        Budget09 = !Budget09
        Budget10 = !Budget10
        Budget11 = !Budget11
        Budget12 = !Budget12
        Budget13 = !Budget13
    End With
    TotalAmount = Amount01 + Amount02 + Amount03 + Amount04 + Amount05 + Amount06 + Amount07 + Amount08 + Amount09 + Amount10 + Amount11 + Amount12 + Amount13
    TotalBudget = Budget01 + Budget02 + Budget03 + Budget04 + Budget05 + Budget06 + Budget07 + Budget08 + Budget09 + Budget10 + Budget11 + Budget12 + Budget13
    GetAmount = True
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function PutAmount(ByVal Account As Long, ByVal fiscalYear As Integer, ByVal DataFile As String) As Boolean
    PutAmount = False
    On Error GoTo glErr
    Dim db As Database
    Dim rs As Recordset
    Set db = OpenDatabase(FileName)
    Set rs = db.OpenRecordset("SELECT * FROM glAmount WHERE Account=" & CStr(Account) & " AND FiscalYear=" & CStr(fiscalYear))
    rs.Edit
    rs!Amount01 = Amount01
    rs!Amount02 = Amount02
    rs!Amount03 = Amount03
    rs!Amount04 = Amount04
    rs!Amount05 = Amount05
    rs!Amount06 = Amount06
    rs!Amount07 = Amount07
    rs!Amount08 = Amount08
    rs!Amount09 = Amount09
    rs!Amount10 = Amount10
    rs!Amount11 = Amount11
    rs!Amount12 = Amount12
    rs!Amount13 = Amount13
    rs!Budget01 = Budget01
    rs!Budget02 = Budget02
    rs!Budget03 = Budget03
    rs!Budget04 = Budget04
    rs!Budget05 = Budget05
    rs!Budget06 = Budget06
    rs!Budget07 = Budget07
    rs!Budget08 = Budget08
    rs!Budget09 = Budget09
    rs!Budget10 = Budget10
    rs!Budget11 = Budget11
    rs!Budget12 = Budget12
    rs!Budget13 = Budget13
    rs.Update
    PutAmount = True
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function
