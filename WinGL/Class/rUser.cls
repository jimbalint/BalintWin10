VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID, LastCompany As Long
Public LoadLastCompany As Boolean
Public Logon, Name, Password As String
Dim db As DAO.Database
Dim rs As DAO.Recordset

Public Function PutRecord(ByVal ID As Long) As Long
    On Error GoTo glErr
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    Set rs = db.OpenRecordset("Users")
    With rs
        If ID = 0 Then
            If rs.RecordCount > 0 Then .MoveLast
            .AddNew
            LastCompany = 0
            LoadLastCompany = False
            Logon = " "
            Name = " "
            Password = " "
        Else
            .index = "PrimaryKey"
            .Seek "=", ID
            .Edit
        End If
        !Name = Name
        !Logon = Logon
        !Password = Password
        !LastCompany = LastCompany
        !LoadLastCompany = LoadLastCompany
        PutRecord = !ID
        .Update
    End With
    Exit Function
glErr:
    MsgBox Error(Err.Number)
End Function

Public Function GetRecord(ByVal ID As Long) As Boolean
    
    On Error GoTo glErr
    
    GetRecord = False
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    Set rs = db.OpenRecordset("Users")
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        ID = !ID
        Name = !Name & ""
        Logon = !Logon & ""
        Password = !Password & ""
        LastCompany = !LastCompany + 0
        LoadLastCompany = !LoadLastCompany + 0
        GetRecord = True
    End With
    Exit Function

glErr:
    
    Name = ""
    Logon = ""
    Password = ""
    LastCompany = 0
    LoadLastCompany = False
    GetRecord = False

End Function

Public Function OpenDB() As Boolean
    
    On Error GoTo glErr
    
    If BalintFolder = "" Then
        Set db = OpenDatabase("\balint\data\glSystem.mdb")
    Else
        Set db = OpenDatabase(BalintFolder & "\Data\GLSystem.mdb")
    End If
    
    Set rs = db.OpenRecordset("Users")
    Exit Function

glErr:
    OpenDB = False

End Function

Public Function FindByID(ByVal ID As Long) As Boolean
    
    On Error GoTo glErr
    
    With rs
        .index = "PrimaryKey"
        .Seek "=", ID
        ID = !ID
        Name = !Name & ""
        Logon = !Logon & ""
        Password = !Password & ""
        LastCompany = !LastCompany + 0
        LoadLastCompany = !LoadLastCompany + 0
        FindByID = True
    End With

    Exit Function

glErr:
    FindByID = False

End Function

Public Function CloseDB() As Boolean
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
End Function
