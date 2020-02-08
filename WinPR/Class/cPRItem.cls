VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ItemID As Long
Public EmployeeID As Long
Public Title As String
Public Abbreviation As String
Public Active As Byte

Public ItemType As Byte

Public NoSSTax As Byte
Public NoMedTax As Byte
Public NoFWTTax As Byte
Public NoSWTTax As Byte
Public NoCWTTax As Byte
Public NoSUNTax As Byte
Public NoFUNTax As Byte
Public Tips As Byte
Public NotInNet As Byte

Public DirDepType As Byte
Public DirDepBank As String
Public DirDepABA As String
Public DirDepAccount As String
Public DirDepBasis As Byte
Public DirDepAmtPct As Currency

Public GLAccount As Long
Public Basis As Long
Public MatchPct As Double
Public MaxPct As Double
Public MaxAmount As Currency
Public AmtPct As Currency

Public W2Box12Code As String
Public W2Box14Code As String
Public Pension As Byte
Public SickPay As Byte
Public SDNumber As Byte
Public EmployerItemID As Long
Public UseEmployer As Byte
Public Escrow As Byte
Public Comment As String
Public DirDepRpt As Byte
Public RateDifference As Byte
Public PWFringe As Byte
Public CityID As Long

Dim rs As New ADODB.Recordset
    
Public Sub OpenRS()

   SQLString = "SELECT * FROM PRItem"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngItemID As Long) As Boolean

    Set rs = New ADODB.Recordset
    
    If lngItemID = 0 Then
       ErrMessage = "ItemID not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRItem where RecID = " & lngItemID
    
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT EmployeeID from PRItem"
    
    rsInit SQLString, cn, rs
    
    ' Check to make sure a record
    ' actually came back
    If rs.EOF = True And rs.BOF = True Then
       Records = 0
    Else
       rs.MoveLast
       Records = rs.RecordCount
    End If

    rs.Close

End Function

Public Function Find(ByVal lngItemID As Long) As Boolean
   
   SQLString = "RecID = " & lngItemID
   
   rs.Find SQLString, 0, adSearchForward, 1
   
   If rs.EOF Then
      Find = False
      Exit Function
   End If
      
   Find = True
   
   AssignFields
   
End Function

Public Function FindFirst() As Boolean
   
   rs.MoveFirst
   
   If rs.EOF Then
      FindFirst = False
      Exit Function
   End If
      
   FindFirst = True
   
   AssignFields
   
End Function


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


Public Function Save(ByVal booAddPut As Boolean) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If
    
    rs.Fields("EmployeeID") = EmployeeID
    rs.Fields("Title") = UCase(Title)
    rs.Fields("Abbreviation") = UCase(Abbreviation)
    rs.Fields("Active") = Active
    
    rs.Fields("ItemType") = ItemType
    rs.Fields("NoSSTax") = NoSSTax
    rs.Fields("NoMedTax") = NoMedTax
    rs.Fields("NoFWTTax") = NoFWTTax
    rs.Fields("NoSWTTax") = NoSWTTax
    rs.Fields("NoCWTTax") = NoCWTTax
    rs.Fields("NoSUNTax") = NoSUNTax
    rs.Fields("NoFUNTax") = NoFUNTax
    rs.Fields("Tips") = Tips
    rs.Fields("NotInNet") = NotInNet
    
    rs.Fields("DirDepType") = DirDepType
    rs.Fields("DirDepBank") = Trim(DirDepBank)
    rs.Fields("DirDepABA") = Trim(DirDepABA)
    rs.Fields("DirDepAccount") = Trim(DirDepAccount)
    rs.Fields("DirDepBasis") = DirDepBasis
    rs.Fields("DirDepAmtPct") = DirDepAmtPct
    rs.Fields("GLAccount") = GLAccount
    rs.Fields("Basis") = Basis
    rs.Fields("MatchPct") = MatchPct
    rs.Fields("MaxPct") = MaxPct
    rs.Fields("MaxAmount") = MaxAmount
    rs.Fields("AmtPct") = AmtPct

    rs.Fields("W2Box12Code") = DirDepType
    rs.Fields("W2Box14Code") = DirDepType
    rs.Fields("Pension") = Pension
    rs.Fields("SickPay") = SickPay
    rs.Fields("SDNumber") = SDNumber
    rs.Fields("EmployerItemID") = EmployerItemID
    rs.Fields("UseEmployer") = UseEmployer
    rs.Fields("Escrow") = Escrow
    rs.Fields("Comment") = Comment
    rs.Fields("DirDepRpt") = DirDepRpt
        
    rs.Fields("RateDifference") = RateDifference
    ' rs.Fields("PWFring") = PWFringe
    rs.Fields("CityID") = CityID
        
    rs.Update
    
    ItemID = rs!ItemID
    
End Function

Public Sub Clear()
      
    ItemID = 0
    EmployeeID = 0
    Title = ""
    Abbreviation = ""
    Active = 0
    
    ItemType = 0
    
    NoSSTax = 0
    NoMedTax = 0
    NoFWTTax = 0
    NoSWTTax = 0
    NoCWTTax = 0
    NoSUNTax = 0
    NoFUNTax = 0
    Tips = 0
    NotInNet = 0

    DirDepType = 0
    DirDepBank = ""
    DirDepABA = ""
    DirDepAccount = 0
    DirDepBasis = 0
    DirDepAmtPct = 0

    GLAccount = 0
    Basis = 0
    MatchPct = 0
    MaxPct = 0
    MaxAmount = 0
    AmtPct = 0

    W2Box12Code = ""
    W2Box14Code = ""
    Pension = 0
    SickPay = 0

    SDNumber = 0
    EmployerItemID = 0
    UseEmployer = 0
    Escrow = 0

    Comment = ""
    DirDepRpt = 0
    RateDifference = 0
    PWFringe = 0
    CityID = 0

End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRItem", cn, rs
    
End Sub

Public Sub AssignFields()
        
    ItemID = nNull(rs!ItemID)
    EmployeeID = nNull(rs!EmployeeID)
    Title = UCase(Trim(rs!Title)) & ""
    
    Abbreviation = UCase(Trim(rs!Abbreviation)) & ""
    If Abbreviation = "" Then
        Abbreviation = UCase(Trim(rs!Title)) & ""
    End If
    
    Active = nNull(rs!Active)

    ItemType = nNull(rs!ItemType)
    NoSSTax = nNull(rs!NoSSTax)
    NoMedTax = nNull(rs!NoMedTax)
    NoFWTTax = nNull(rs!NoFWTTax)
    NoSWTTax = nNull(rs!NoSWTTax)
    NoCWTTax = nNull(rs!NoCWTTax)
    NoSUNTax = nNull(rs!NoSUNTax)
    NoFUNTax = nNull(rs!NoFUNTax)
    Tips = nNull(rs!Tips)
    NotInNet = nNull(rs!NotInNet)

    DirDepType = nNull(rs!DirDepType)
    DirDepBank = UCase(Trim(rs!DirDepBank)) & ""
    DirDepABA = UCase(Trim(rs!DirDepABA)) & ""
    DirDepAccount = UCase(Trim(rs!DirDepAccount)) & ""
    DirDepBasis = nNull(rs!DirDepBasis)
    DirDepAmtPct = nNull(rs!DirDepAmtPct)

    GLAccount = nNull(rs!GLAccount)
    Basis = nNull(rs!Basis)
    MatchPct = nNull(rs!MatchPct)
    MaxPct = nNull(rs!MaxPct)
    MaxAmount = nNull(rs!MaxAmount)
    AmtPct = nNull(rs!AmtPct)

    W2Box12Code = nNull(rs!W2Box12Code)
    W2Box14Code = nNull(rs!W2Box14Code)
    Pension = nNull(rs!Pension)
    SickPay = nNull(rs!SickPay)
        
    SDNumber = nNull(rs!SDNumber)
    EmployerItemID = nNull(rs!EmployerItemID)
    UseEmployer = nNull(rs!UseEmployer)
    Escrow = nNull(rs!Escrow)
        
    Comment = rs!Comment & ""
    DirDepRpt = nNull(rs!DirDepRpt)
    RateDifference = nNull(rs!RateDifference)
    PWFringe = nNull(rs!PWFringe)
    CityID = nNull(rs!CityID)

End Sub

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

Public Function GetByID(ByVal ItemID As Long) As Boolean

    rsInit "SELECT * FROM PRItem WHERE PRItem.ItemID = " & CStr(ItemID), cn, rs
    
    If rs.BOF And rs.EOF Then
        GetByID = False
        Exit Function
    End If
    
    rs.MoveFirst
    AssignFields
    GetByID = True
    
End Function

Public Sub UpdateBatch()

    ' update disconnect record set
    rsSave rs, cn

End Sub

