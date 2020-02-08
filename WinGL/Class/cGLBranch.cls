VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGLBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const srCouldNotGetData = vbObjectError + 512 + 2
Const srCouldNotSaveData = vbObjectError + 512 + 3

Private mlngID As Long
Private mintBranchNumber As Integer
Private mtxtName As String
Dim rs As New ADODB.Recordset


Public Property Let ID(ByVal lngNew As Long)
    mlngID = lngNew
End Property

Public Property Get ID() As Long
    ID = mlngID
End Property



Public Property Let BranchNumber(ByVal intNew As Integer)
    mintBranchNumber = intNew
End Property

Public Property Get BranchNumber() As Integer
    BranchNumber = mintBranchNumber
End Property



Public Property Let Name(ByVal txtNew As String)
    mtxtName = txtNew
End Property

Public Property Get Name() As String
    Name = mtxtName
End Property
Public Sub Save(ByVal booAddPut As Boolean)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    Set rs.ActiveConnection = cn
 
    If booAddPut = True Then      ' Add
    
         rs.Source = "GLBranch"
         rs.Open
         rs.AddNew
    
    Else                           ' Put
    
         rs.Source = "select * " & _
                     "from GLBranch " & _
                     "where BranchNumber = " & BranchNumber
         rs.Open
    
    End If
    
    rs.Fields("BranchNumber") = BranchNumber
    rs.Fields("Name") = Name
    
    rs.Update
    
    rs.Close
    
    Set rs = Nothing
    
End Sub
    
Public Sub GetData(lngBranchNumber As Long)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If lngBranchNumber <> 0 Then
        
        rs.Source = "select * from GLBranch " & _
                    "where BranchNumber = " & lngBranchNumber
        rs.MaxRecords = 1
        
        Set rs.ActiveConnection = cn
        
        rs.Open
        
        ' Check to make sure a record
        ' actually came back
        If rs.EOF = True And rs.BOF = True Then
            Err.Raise srCouldNotGetData, _
                      "GLBranch - GetData", _
                      "Branch not found."
        End If
        
        ' Populate properties with
        ' results of the query.
        BranchNumber = rs!BranchNumber
        Name = rs!Name & ""
    
    Else
            Err.Raise srCouldNotGetData, _
                      "GLBranch - GetData", _
                      "Branch not found."
        Exit Sub
    End If
    
    rs.Close
    Set rs = Nothing
    
End Sub


Public Sub Clear()
      ID = 0
      BranchNumber = 0
      Name = ""
End Sub

Public Sub DeleteAll()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Source = "Delete * from GLBranch"
    Set rs.ActiveConnection = cn
    rs.Open
    Set rs = Nothing
End Sub

Public Sub OpenRS()
    SQLString = "GLBranch"
    rsInit SQLString, cn, rs
End Sub

