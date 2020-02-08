VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
Dim trs As New ADODB.Recordset
Dim p1 As Currency
Dim p2 As Currency
Dim x As String
Dim CCreate As Boolean
    
Dim i As Integer
Dim j As Integer
    
    ' !!!!!!!!!!!!!!!!!!!!!!!
    ' !!! GLSystem files - create the tables if they don't exist
    ' !!!!!!!!!!!!!!!!!!!!!!!
    
    If Not TableExists("PRCompany", cnDes) Then CompanyCreate
    If Not TableExists("PRCity", cnDes) Then CityCreate
    If Not TableExists("PRState", cnDes) Then
        StateCreate
        ImportState
    End If
    If Not TableExists("PRGlobal", cnDes) Then GlobalCreate
    If Not TableExists("PRFWTTable", cnDes) Then FWTCreate

    ' !!!!!!!!!!!!!!!!!!!!!
    ' !!! PRMas Files - remove and create the tables
    ' !!!!!!!!!!!!!!!!!!!!!
    DropTable "PREmployee", cn
    EmployeeCreate
    
    DropTable "PRDepartment", cn
    DepartmentCreate
    
    DropTable "PRItem", cn
    ItemCreate

    DropTable "PRHist", cn
    HistCreate
    
    DropTable "PRItemHist", cn
    ItemHistCreate
    
    DropTable "PRDist", cn
    DistCreate
    
    DropTable "PRAdjust", cn
    AdjustCreate
    
    DropTable "PRBatch", cn
    BatchCreate
    
    DropTable "PREELists", cn
    EEListsCreate

    aPRImport


' ===========================================================================================

'    trs.CursorLocation = adUseClient
'
'    trs.Fields.Append "EmployeeNumber", adInteger
'    trs.Fields.Append "CityNumber", adSingle
'    trs.Fields.Append "MTDGross", adCurrency
'    trs.Fields.Append "MTDTax", adCurrency
'    trs.Fields.Append "QTDGross", adCurrency
'    trs.Fields.Append "QTDTax", adCurrency
'    trs.Fields.Append "YTDGross", adCurrency
'    trs.Fields.Append "YTDTax", adCurrency
'
'    trs.Open , , adOpenDynamic, adLockOptimistic
'
'    If PREmployee.GetBySQL("SELECT * FROM PREmployee WHERE PREmployee.Inactive = false") Then
'
'        Do
'
'            For i = 1 To 3
'                trs.AddNew Array("EmployeeNumber", "CityNumber", "MTDGRoss", "MTDTax", "QTDGross", "QTDTax", "YTDGross", "YTDTax"), _
'                           Array(PREmployee.EmployeeNumber, i, 100, 5, 300, 15, 1200, 60)
'                trs.UpdateBatch
'            Next i
'
'            If Not PREmployee.GetNext Then Exit Do
'
'        Loop
'
'    End If
'
'    trs.Sort = "EmployeeNumber, CityNumber"
'
'    trs.MoveFirst
'
'    Do
'
'        i = CInt(trs!EmployeeNumber)
'        j = CInt(trs!CityNumber)
'        MsgBox i & " " & j
'
'        trs.MoveNext
'        If trs.EOF Then Exit Do
'
'    Loop

' ===========================================================================================

    ' ******* test sweeps *********
    '
    ' get subtotals by dept

'    trs.CursorLocation = adUseClient
'
'    trs.Fields.Append "Dept", adDouble
'    trs.Fields.Append "Total", adCurrency
'
'    trs.Open , , adOpenDynamic, adLockOptimistic
'
'    If PREmployee.GetBySQL("SELECT * FROM PREmployee") Then
'        Do
'            p1 = GetPRAmount(1, 1, 1, 1, 1, 1, 1, 1)
'
'            X = "Dept=" & CStr(PREmployee.DepartmentNumber)
'            trs.Find X, 0, adSearchForward, 1
'            If trs.EOF Then
'                trs.AddNew Array("Dept", "Total"), _
'                           Array(PREmployee.DepartmentNumber, 0)
'                trs.UpdateBatch
'            End If
'
'            p2 = trs!Total
'            p2 = p2 + p1
'            trs.Fields("Total") = p2
'            trs.Update
'
'            If Not PREmployee.GetNext Then Exit Do
'
'        Loop
'    End If
'
'    trs.Sort = "Dept"
'    trs.MoveFirst
'
'    Do
'
'        MsgBox trs!Dept & " " & trs!Total
'
'        trs.MoveNext
'        If trs.EOF Then Exit Do
'
'    Loop

    End

End Sub
