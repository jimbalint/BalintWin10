VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPRW2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public W2ID As Long
Public TaxYear As Long
Public EmployeeID As Long
Public EmployeeNumber As Long
Public BoxA_SSNumber As Long
Public BoxB_FedID As String
Public BoxC_ERName As String
Public BoxC_ERAddr1 As String
Public BoxC_ERAddr2 As String
Public BoxC_ERCity As String
Public BoxC_ERState As String
Public BoxC_ERZip As String
Public BoxD_ControlNumber As Long
Public BoxE_EEFirstName As String
Public BoxE_EELastName As String
Public BoxE_EEMidInit As String
Public BoxE_EEAddr1 As String
Public BoxE_EEAddr2 As String
Public BoxE_EECity As String
Public BoxE_EEState As String
Public BoxE_EEZip As String
Public Box1_Wages As Currency
Public Box2_FedTax As Currency
Public Box3_SSWages As Currency
Public Box4_SSTax As Currency
Public Box5_MedWages As Currency
Public Box6_MedTax As Currency
Public Box7_SSTips As Currency
Public Box8_AllocTips As Currency
Public Box9_EIC As Currency
Public Box10_DCBen As Currency
Public Box11_NQPlans As Currency
Public Box12A_ID As Long
Public Box12A_Code As String
Public Box12A_Amount As Currency
Public Box12B_ID As Long
Public Box12B_Code As String
Public Box12B_Amount As Currency
Public Box12C_ID As Long
Public Box12C_Code As String
Public Box12C_Amount As Currency
Public Box12D_ID As Long
Public Box12D_Code As String
Public Box12D_Amount As Currency
Public Box13_StatEmp As Byte
Public Box13_RetirePlan As Byte
Public Box13_3rdParty As Byte
Public Box14A_ID As Long
Public Box14A_Desc As String
Public Box14A_Amount As Currency
Public Box14B_ID As Long
Public Box14B_Desc As String
Public Box14B_Amount As Currency
Public Box14C_ID As Long
Public Box14C_Desc As String
Public Box14C_Amount As Currency
Public Box14D_ID As Long
Public Box14D_Desc As String
Public Box14D_Amount As Currency
Public Box15A_State As String
Public Box15A_StateID As String
Public Box16A_StateWages As Currency
Public Box17A_StateTax As Currency
Public Box15B_State As String
Public Box15B_StateID As String
Public Box16B_StateWages As Currency
Public Box17B_StateTax As Currency
Public Void As Byte
Public Skip As Byte

Dim rs As New ADODB.Recordset

Public Sub OpenRS()

   SQLString = "SELECT * FROM PRW2"
   rsInit SQLString, cn, rs

   If rs.BOF And rs.EOF Then
      Clear
      Exit Sub
   End If

   rs.MoveFirst
   
   AssignFields

End Sub

Public Function DeleteRecord(ByVal lngPRW2ID As Long) As Byte

    Set rs = New ADODB.Recordset
    
    If lngPRW2ID = 0 Then
       ErrMessage = "Employee number not specified!"
       DeleteRecord = False
       Exit Function
    End If
    
    SQLString = "Delete * from PRW2 where W2ID = " & lngPRW2ID
    rsInit SQLString, cn, rs
    
    DeleteRecord = True
    
End Function

Public Function Records() As Long

    SQLString = "SELECT W2ID from PRW2"
    
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
Public Function GetPrev() As Byte

    If rs.BOF = True Then
        GetPrev = 0
        Exit Function
    End If

    rs.MovePrevious
    AssignFields
    GetPrev = True

End Function

Public Function Save(ByVal booAddPut As Byte) As String

    If booAddPut = True Then      ' Add
       rs.AddNew
    End If

    rs!TaxYear = nNull(TaxYear)
    rs!EmployeeID = nNull(EmployeeID)
    rs!EmployeeNumber = nNull(EmployeeNumber)
    rs!BoxA_SSNumber = nNull(BoxA_SSNumber)
    rs!BoxB_FedID = BoxB_FedID & ""
    rs!BoxC_ERName = BoxC_ERName & ""
    rs!BoxC_ERAddr1 = BoxC_ERAddr1 & ""
    rs!BoxC_ERAddr2 = BoxC_ERAddr2 & ""
    rs!BoxC_ERCity = BoxC_ERCity & ""
    rs!BoxC_ERState = BoxC_ERState & ""
    rs!BoxC_ERZip = BoxC_ERZip & ""
    rs!BoxD_ControlNumber = nNull(BoxD_ControlNumber)
    rs!BoxE_EEFirstName = BoxE_EEFirstName & ""
    rs!BoxE_EELastName = BoxE_EELastName & ""
    rs!BoxE_EEMidInit = BoxE_EEMidInit & ""
    rs!BoxE_EEAddr1 = BoxE_EEAddr1 & ""
    rs!BoxE_EEAddr2 = BoxE_EEAddr2 & ""
    rs!BoxE_EECity = BoxE_EECity & ""
    rs!BoxE_EEState = BoxE_EEState & ""
    rs!BoxE_EEZip = nNull(BoxE_EEZip)
    rs!Box1_Wages = nNull(Box1_Wages)
    rs!Box2_FedTax = nNull(Box2_FedTax)
    rs!Box3_SSWages = nNull(Box3_SSWages)
    rs!Box4_SSTax = nNull(Box4_SSTax)
    rs!Box5_MedWages = nNull(Box5_MedWages)
    rs!Box6_MedTax = nNull(Box6_MedTax)
    rs!Box7_SSTips = nNull(Box7_SSTips)
    rs!Box8_AllocTips = nNull(Box8_AllocTips)
    rs!Box9_EIC = nNull(Box9_EIC)
    rs!Box10_DCBen = nNull(Box10_DCBen)
    rs!Box11_NQPlans = nNull(Box11_NQPlans)
    rs!Box12A_ID = nNull(Box12A_ID)
    rs!Box12A_Code = Box12A_Code & ""
    rs!Box12A_Amount = nNull(Box12A_Amount)
    rs!Box12B_ID = nNull(Box12B_ID)
    rs!Box12B_Code = Box12B_Code & ""
    rs!Box12B_Amount = nNull(Box12B_Amount)
    rs!Box12C_ID = nNull(Box12C_ID)
    rs!Box12C_Code = Box12C_Code & ""
    rs!Box12C_Amount = nNull(Box12C_Amount)
    rs!Box12D_ID = nNull(Box12D_ID)
    rs!Box12D_Code = Box12D_Code & ""
    rs!Box12D_Amount = nNull(Box12D_Amount)
    rs!Box13_StatEmp = nNull(Box13_StatEmp)
    rs!Box13_RetirePlan = nNull(Box13_RetirePlan)
    rs!Box13_3rdParty = nNull(Box13_3rdParty)
    rs!Box14A_ID = nNull(Box14A_ID)
    rs!Box14A_Desc = Box14A_Desc & ""
    rs!Box14A_Amount = nNull(Box14A_Amount)
    rs!Box14B_ID = nNull(Box14B_ID)
    rs!Box14B_Desc = Box14B_Desc & ""
    rs!Box14B_Amount = nNull(Box14B_Amount)
    rs!Box14C_ID = nNull(Box14C_ID)
    rs!Box14C_Desc = Box14C_Desc & ""
    rs!Box14C_Amount = nNull(Box14C_Amount)
    rs!Box14D_ID = nNull(Box14D_ID)
    rs!Box14D_Desc = Box14D_Desc & ""
    rs!Box14D_Amount = nNull(Box14D_Amount)
    rs!Box15A_State = Box15A_State & ""
    rs!Box15A_StateID = Box15A_StateID & ""
    rs!Box16A_StateWages = nNull(Box16A_StateWages)
    rs!Box17A_StateTax = nNull(Box17A_StateTax)
    rs!Box15B_State = Box15B_State & ""
    rs!Box15B_StateID = Box15B_StateID & ""
    rs!Box16B_StateWages = nNull(Box16B_StateWages)
    rs!Box17B_StateTax = nNull(Box17B_StateTax)
    rs!Void = nNull(Void)
    rs!Skip = nNull(Skip)

    rs.Update
    
    W2ID = rs!W2ID
    
End Function

Public Sub Clear()
      
    W2ID = 0
    TaxYear = 0
    EmployeeID = 0
    EmployeeNumber = 0
    BoxA_SSNumber = 0
    BoxB_FedID = ""
    BoxC_ERName = ""
    BoxC_ERAddr1 = ""
    BoxC_ERAddr2 = ""
    BoxC_ERCity = ""
    BoxC_ERState = ""
    BoxC_ERZip = ""
    BoxD_ControlNumber = 0
    BoxE_EEFirstName = ""
    BoxE_EELastName = ""
    BoxE_EEMidInit = ""
    BoxE_EEAddr1 = ""
    BoxE_EEAddr2 = ""
    BoxE_EECity = ""
    BoxE_EEState = ""
    BoxE_EEZip = ""
    Box1_Wages = 0
    Box2_FedTax = 0
    Box3_SSWages = 0
    Box4_SSTax = 0
    Box5_MedWages = 0
    Box6_MedTax = 0
    Box7_SSTips = 0
    Box8_AllocTips = 0
    Box9_EIC = 0
    Box10_DCBen = 0
    Box11_NQPlans = 0
    Box12A_ID = 0
    Box12A_Code = ""
    Box12A_Amount = 0
    Box12B_ID = 0
    Box12B_Code = ""
    Box12B_Amount = 0
    Box12C_ID = 0
    Box12C_Code = ""
    Box12C_Amount = 0
    Box12D_ID = 0
    Box12D_Code = ""
    Box12D_Amount = 0
    Box13_StatEmp = 0
    Box13_RetirePlan = 0
    Box13_3rdParty = 0
    Box14A_ID = 0
    Box14A_Desc = ""
    Box14A_Amount = 0
    Box14B_ID = 0
    Box14B_Desc = ""
    Box14B_Amount = 0
    Box14C_ID = 0
    Box14C_Desc = ""
    Box14C_Amount = 0
    Box14D_ID = 0
    Box14D_Desc = ""
    Box14D_Amount = 0
    Box15A_State = ""
    Box15A_StateID = ""
    Box16A_StateWages = 0
    Box17A_StateTax = 0
    Box15B_State = ""
    Box15B_StateID = ""
    Box16B_StateWages = 0
    Box17B_StateTax = 0
    Void = 0
    Skip = 0
      
End Sub

Public Sub DeleteAll()

    Set rs = New ADODB.Recordset
    rsInit "DELETE * FROM PRW2", cn, rs
    
End Sub

Public Sub AssignFields()
        
    W2ID = nNull(rs!W2ID)
    TaxYear = nNull(rs!TaxYear)
    BoxA_SSNumber = nNull(rs!BoxA_SSNumber)
    EmployeeID = nNull(rs!EmployeeID)
    EmployeeNumber = nNull(rs!EmployeeNumber)
    BoxB_FedID = TextSet(rs!BoxB_FedID)
    BoxC_ERName = TextSet(rs!BoxC_ERName)
    BoxC_ERAddr1 = TextSet(rs!BoxC_ERAddr1)
    BoxC_ERAddr2 = TextSet(rs!BoxC_ERAddr2)
    BoxC_ERCity = TextSet(rs!BoxC_ERCity)
    BoxC_ERState = TextSet(rs!BoxC_ERState)
    BoxC_ERZip = TextSet(rs!BoxC_ERZip)
    BoxD_ControlNumber = nNull(rs!BoxD_ControlNumber)
    BoxE_EEFirstName = TextSet(rs!BoxE_EEFirstName)
    BoxE_EELastName = TextSet(rs!BoxE_EELastName)
    BoxE_EEMidInit = TextSet(rs!BoxE_EEMidInit)
    BoxE_EEAddr1 = TextSet(rs!BoxE_EEAddr1)
    BoxE_EEAddr2 = TextSet(rs!BoxE_EEAddr2)
    BoxE_EECity = TextSet(rs!BoxE_EECity)
    BoxE_EEState = TextSet(rs!BoxE_EEState)
    BoxE_EEZip = TextSet(rs!BoxE_EEZip)
    Box1_Wages = nNull(rs!Box1_Wages)
    Box2_FedTax = nNull(rs!Box2_FedTax)
    Box3_SSWages = nNull(rs!Box3_SSWages)
    Box4_SSTax = nNull(rs!Box4_SSTax)
    Box5_MedWages = nNull(rs!Box5_MedWages)
    Box6_MedTax = nNull(rs!Box6_MedTax)
    Box7_SSTips = nNull(rs!Box7_SSTips)
    Box8_AllocTips = nNull(rs!Box8_AllocTips)
    Box9_EIC = nNull(rs!Box9_EIC)
    Box10_DCBen = nNull(rs!Box10_DCBen)
    Box11_NQPlans = nNull(rs!Box11_NQPlans)
    Box12A_ID = nNull(rs!Box12A_ID)
    Box12A_Code = TextSet(rs!Box12A_Code)
    Box12A_Amount = nNull(rs!Box12A_Amount)
    Box12B_ID = nNull(rs!Box12B_ID)
    Box12B_Code = TextSet(rs!Box12B_Code)
    Box12B_Amount = nNull(rs!Box12B_Amount)
    Box12C_ID = nNull(rs!Box12C_ID)
    Box12C_Code = TextSet(rs!Box12C_Code)
    Box12C_Amount = nNull(rs!Box12C_Amount)
    Box12D_ID = nNull(rs!Box12D_ID)
    Box12D_Code = TextSet(rs!Box12D_Code)
    Box12D_Amount = nNull(rs!Box12D_Amount)
    Box13_StatEmp = nNull(rs!Box13_StatEmp)
    Box13_RetirePlan = nNull(rs!Box13_RetirePlan)
    Box13_3rdParty = nNull(rs!Box13_3rdParty)
    Box14A_ID = nNull(rs!Box14A_ID)
    Box14A_Desc = TextSet(rs!Box14A_Desc)
    Box14A_Amount = nNull(rs!Box14A_Amount)
    Box14B_ID = nNull(rs!Box14B_ID)
    Box14B_Desc = TextSet(rs!Box14B_Desc)
    Box14B_Amount = nNull(rs!Box14B_Amount)
    Box14C_ID = nNull(rs!Box14C_ID)
    Box14C_Desc = TextSet(rs!Box14C_Desc)
    Box14C_Amount = nNull(rs!Box14C_Amount)
    Box14D_ID = nNull(rs!Box14D_ID)
    Box14D_Desc = TextSet(rs!Box14D_Desc)
    Box14D_Amount = nNull(rs!Box14D_Amount)
    Box15A_State = TextSet(rs!Box15A_State)
    Box15A_StateID = TextSet(rs!Box15A_StateID)
    Box16A_StateWages = nNull(rs!Box16A_StateWages)
    Box17A_StateTax = nNull(rs!Box17A_StateTax)
    Box15B_State = TextSet(rs!Box15B_State)
    Box15B_StateID = TextSet(rs!Box15B_StateID)
    Box16B_StateWages = nNull(rs!Box16B_StateWages)
    Box17B_StateTax = nNull(rs!Box17B_StateTax)
    Void = nNull(rs!Void)
    Skip = nNull(rs!Skip)
                                           
End Sub

Public Function GetByID(ByVal lngPRW2ID As Long) As Boolean

    rsInit "SELECT * FROM PRW2 WHERE W2ID = " & lngPRW2ID, cn, rs
    
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

Public Function GetFirst() As Boolean
    If rs.RecordCount = 0 Then
        GetFirst = False
        Exit Function
    End If
    rs.MoveFirst
    AssignFields
    GetFirst = True
End Function
