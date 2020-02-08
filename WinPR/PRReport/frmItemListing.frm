VERSION 5.00
Begin VB.Form frmItemListing 
   Caption         =   "Employee Item Listing"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemListing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActiveEmployees 
      Caption         =   "Show Active Employees only?"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   1373
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CheckBox chkActiveItems 
      Caption         =   "Show Active Items only?"
      Height          =   495
      Left            =   2633
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   4853
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   6975
   End
End
Attribute VB_Name = "frmItemListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j, k As Long
Dim X, Y, Z As String
Dim C1, C2, C3 As Currency
Dim boo As Boolean
Dim StringOut As String

Dim rs As New ADODB.Recordset

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name

    Me.chkActiveEmployees = 1
    Me.chkActiveItems = 1
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()

    SQLString = "SELECT * FROM PREmployee "
    If Me.chkActiveEmployees = 1 Then SQLString = SQLString & "WHERE Inactive = 0 "
    SQLString = SQLString & " ORDER BY EmployeeNumber"
    
    If PREmployee.GetBySQL(SQLString) = False Then
        MsgBox "No Employees Found!", vbExclamation
        GoBack
    End If
        
    SQLString = "SELECT * FROM PRItem WHERE EmployeeID = 0"
    If PRItem.GetBySQL(SQLString) = False Then
        MsgBox "No Items defined for this client!", vbExclamation
        GoBack
    End If
    
    ' record set for item titles ....
    rs.CursorLocation = adUseClient
    rs.Fields.Append "ItemID", adDouble
    rs.Fields.Append "Title", adVarChar, 40, adFldIsNullable
    rs.Fields.Append "RateDifference", adInteger
    rs.Fields.Append "AmtPct", adCurrency
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    Do
        rs.AddNew
        rs!ItemID = PRItem.ItemID
        rs!Title = PRItem.Abbreviation
        rs!RateDifference = PRItem.RateDifference
        rs!AmtPct = PRItem.AmtPct
        rs.Update
        If PRItem.GetNext = False Then Exit Do
    Loop
    
    PrtInit "Port"
    SetFont 9, Equate.Portrait
    
    PageHeader "Payroll Item Listing"
    
    ' loop thru the employees
    Do
        
        ' show the employee name and rate
        PrintValue(1) = PREmployee.EmployeeNumber:          FormatString(1) = "n9"
        PrintValue(2) = " " & PREmployee.LFName:            FormatString(2) = "a40"
        If PREmployee.Salaried = 1 Then
            X = " Salaried"
            C1 = PREmployee.SalaryAmount
        Else
            X = " Hourly"
            C1 = PREmployee.HourlyAmount
        End If
        PrintValue(3) = X:                                  FormatString(3) = "a10"
        PrintValue(4) = C1:                                 FormatString(4) = "d12"
        PrintValue(5) = " ":                                FormatString(5) = "~"
        FormatPrint
        Ln = Ln + 1
        
        ' loop thru the items of the employee
        SQLString = "SELECT * FROM PRItem WHERE EmployeeID = " & PREmployee.EmployeeID
        If Me.chkActiveItems = 1 Then
            SQLString = SQLString & " AND Active = 1"
        End If
        SQLString = SQLString & " ORDER BY ItemType, EmployerItemID"
        If PRItem.GetBySQL(SQLString) = True Then
            k = 0
            StringOut = ""
            Do
                
                ' get the title
                Z = ""
                If PRItem.ItemType = PREquate.ItemTypeDirDepDed Then
                    Z = PRItem.Title
                    Z = PRItem.DirDepBank
                Else
                    rs.Find "ItemID = " & PRItem.EmployerItemID, 0, adSearchForward, 1
                    If rs.EOF = False Then
                        Z = rs!Title
                    End If
                End If
                AddString Z, 15
                AddString " ", 2
                
                If PRItem.Active = 1 Then
                    X = "Active  "
                Else
                    X = "Inactive"
                End If
                AddString X, 10

                If PRItem.ItemType = PREquate.ItemTypeDirDepDed Then
                    Select Case PRItem.DirDepBasis
                        Case PREquate.DirDepBasisNet:   AddString "Net", 10
                        Case PREquate.DirDepBasisAmt:   AddString "Amount", 10
                        Case PREquate.DirDepBasisPct:   AddString "Percent", 10
                        Case Else
                            AddString " ", 10
                    End Select
                Else
                    Select Case PRItem.Basis
                        Case PREquate.BasisAmount:      AddString "Amount", 10
                        Case PREquate.BasisExemptions:  AddString "Exemptions", 10
                        Case PREquate.BasisHourly:      AddString "Hourly", 10
                        Case PREquate.BasisNet:         AddString "Net", 10
                        Case PREquate.BasisPercent:     AddString "Percent", 10
                        Case Else
                            AddString " ", 10
                    End Select
                End If

                C1 = PRItem.AmtPct
                If PREmployee.Salaried = 0 Then
                    If PRItem.ItemType = PREquate.ItemTypeOE And PRItem.Basis = PREquate.BasisHourly Then
                        If PRItem.UseEmployer = 1 Then
                            j = rs!RateDifference
                            C2 = rs!AmtPct
                        Else
                            j = PRItem.RateDifference
                            C2 = PRItem.AmtPct
                        End If
                        If j <> 0 Then
                            Select Case j
                                Case PREquate.BasisAmount
                                    C1 = PREmployee.HourlyAmount + C2
                                Case PREquate.BasisPercent
                                    C1 = PREmployee.HourlyAmount + C2 / 100 * PREmployee.HourlyAmount
                            End Select
                        End If
                    End If
                End If
                    
                AddString Format(C1, "0.00"), 10
                
                k = k + 1
                If k = 2 Then
                    PrintValue(1) = StringOut:  FormatString(1) = "a90"
                    PrintValue(2) = " ":        FormatString(2) = "~"
                    FormatPrint
                    Ln = Ln + 1
                    If Ln > MaxLines Then
                        FormFeed
                        PageHeader "Payroll Item Listing"
                    End If
                    k = 0
                    StringOut = ""
                End If
                
                If PRItem.GetNext = False Then Exit Do
            Loop
        
            ' print the last one
            If StringOut <> "" Then
                PrintValue(1) = StringOut:  FormatString(1) = "a90"
                PrintValue(2) = " ":        FormatString(2) = "~"
                FormatPrint
                Ln = Ln + 1
            End If
        
        End If
        
        If PREmployee.GetNext = False Then Exit Do
        Ln = Ln + 1
    
    Loop
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

    GoBack

End Sub

Private Sub AddString(ByVal InString As String, _
                      ByVal sLen As Long)

    InString = Trim(InString & "")
    
    If Len(InString) > sLen Then
        StringOut = StringOut & Mid(InString, 1, sLen)
    Else
        StringOut = StringOut & InString & Space(sLen - Len(InString))
    End If

End Sub


