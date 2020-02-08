VERSION 5.00
Begin VB.Form frmCompany 
   Caption         =   " COMPANY INFORMATION"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtSubDigits 
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtNetProfitAcct 
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtNumberPds 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtRetEarnAcct 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtPctBaseAcct 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtSuspAcct 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtLastClose 
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtLastUpdate 
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtFirstPd 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtFirstPAcct 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtZipCode 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtState 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtCity 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtAddress3 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox txtAddress2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtAddress1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "glado"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   6240
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Last Close"
      Height          =   255
      Left            =   3120
      TabIndex        =   35
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Last Update"
      Height          =   255
      Left            =   4560
      TabIndex        =   34
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Sub Digits"
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Ret Earn Acct"
      Height          =   255
      Left            =   3120
      TabIndex        =   32
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Susp Acct"
      Height          =   255
      Left            =   4560
      TabIndex        =   31
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Net Profit Acct"
      Height          =   255
      Left            =   6000
      TabIndex        =   30
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Pct Base Acct"
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Number Periods"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "First Period"
      Height          =   255
      Left            =   1680
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "First P Acct"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "State"
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "City"
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Address3"
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address2"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address1"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrs As New ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    'save form fields
    GLCompany.Name = txtName
    GLCompany.Address1 = txtAddress1
    GLCompany.Address2 = txtAddress2
    GLCompany.Address3 = txtAddress3
    GLCompany.City = txtCity
    GLCompany.State = txtState
    GLCompany.ZipCode = txtZipCode
    GLCompany.FirstPAcct = CVar(txtFirstPAcct)
    GLCompany.FirstPeriod = CInt(txtFirstPd)
    GLCompany.LastClose = CVar(txtLastClose)
    GLCompany.LastUpdate = CVar(txtLastUpdate)
    GLCompany.NetProfitAcct = CVar(txtNetProfitAcct)
    GLCompany.NumberPds = CVar(txtNumberPds)
    GLCompany.PctBaseAcct = CVar(txtPctBaseAcct)
    GLCompany.RetEarnAcct = CVar(txtRetEarnAcct)
    GLCompany.SubDigits = CVar(txtSubDigits)
    GLCompany.SuspAcct = CVar(txtSuspAcct)
    
    GLCompany.Save GLCompany.ID, Equate.RecPut
    
    cmdExit_Click
    
End Sub

Private Sub Form_Load()
    
    frmCompany.Caption = " Company Info for " & GLCompany.Name

'    SetAdo cn, mrs, "select * from GLCompany"
    
'    'prime form fields
'    txtName = mrs!Name
'    txtAddress1 = mrs!Address1
'    txtAddress2 = mrs!Address2
'    txtAddress3 = mrs!Address3
'    txtCity = mrs!City
'    txtState = mrs!State
'    txtZipCode = mrs!ZipCode
'    txtFirstPAcct = mrs!FirstPAcct
'    txtFirstPd = mrs!FirstPd
'    txtLastClose = mrs!LastClose
'    txtLastUpdate = mrs!LastUpdate
'    txtNetProfitAcct = mrs!NetProfitAcct
'    txtNumberPds = mrs!NumberPds
'    txtPctBaseAcct = mrs!PctBaseAcct
'    txtRetEarnAcct = mrs!RetEarnAcct
'    txtSubDigits = mrs!SubDigits
'    txtSuspAcct = mrs!SuspAcct

    'prime form fields
    txtName = GLCompany.Name
    txtAddress1 = GLCompany.Address1
    txtAddress2 = GLCompany.Address2
    txtAddress3 = GLCompany.Address3
    txtCity = GLCompany.City
    txtState = GLCompany.State
    txtZipCode = GLCompany.ZipCode
    txtFirstPAcct = GLCompany.FirstPAcct
    txtFirstPd = GLCompany.FirstPeriod
    txtLastClose = GLCompany.LastClose
    txtLastUpdate = GLCompany.LastUpdate
    txtNetProfitAcct = GLCompany.NetProfitAcct
    txtNumberPds = GLCompany.NumberPds
    txtPctBaseAcct = GLCompany.PctBaseAcct
    txtRetEarnAcct = GLCompany.RetEarnAcct
    txtSubDigits = GLCompany.SubDigits
    txtSuspAcct = GLCompany.SuspAcct

End Sub

