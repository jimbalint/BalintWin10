VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmAddEmployee 
   Caption         =   "Add an Employee"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin TDBNumber6Ctl.TDBNumber tdbnumEmployeeNumber 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   661
      Calculator      =   "frmAddEmployee.frx":0000
      Caption         =   "frmAddEmployee.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAddEmployee.frx":009C
      Keys            =   "frmAddEmployee.frx":00BA
      Spin            =   "frmAddEmployee.frx":0104
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   7575
   End
   Begin VB.Label Label1 
      Caption         =   "Employee numbers must be unique !!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   4575
   End
End
Attribute VB_Name = "frmAddEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EmployeeID As Long
Dim rs As New ADODB.Recordset
Dim EmpNum As Long

Private Sub Form_Load()
    
    Me.tdbnumEmployeeNumber.Format = "########0"
    Me.tdbnumEmployeeNumber.DisplayFormat = "########0"
    Me.tdbnumEmployeeNumber.HighlightText = True
    Me.tdbnumEmployeeNumber.Key.Clear = ""
    Me.tdbnumEmployeeNumber.MinValue = 0
    Me.tdbnumEmployeeNumber.MaxValue = 999999999
    
    ' trap keyboard strokes before the
    ' controls on the form does
    Me.KeyPreview = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    EmployeeID = -1
    SelReturn
End Sub

Private Sub cmdOk_Click()

    ' make sure the selected number does not exist
    SQLString = "SELECT EmployeeNumber FROM PREmployee WHERE PREmployee.EmployeeNumber = " & Me.tdbnumEmployeeNumber
    
    rsInit SQLString, cn, rs
    
    If Not (rs.BOF And rs.EOF) Then
        MsgBox "That Employee Number already exists!", vbExclamation, "Add Employee"
        Exit Sub
    End If

    ' go ahead and add it
    PREmployee.OpenRS
    PREmployee.Clear
    PREmployee.EmployeeNumber = Me.tdbnumEmployeeNumber
    PREmployee.FirstName = "New"
    PREmployee.LastName = "Employee"
    PREmployee.DefaultCityID = PRCompany.DfltCityID
    PREmployee.PaysPerYear = PRCompany.DfltPaysPerYear
    PREmployee.FWTBasis = PREquate.BasisExemptions
    PREmployee.SWTBasis = PREquate.BasisExemptions
    PREmployee.Save (Equate.RecAdd)
    EmployeeID = PREmployee.EmployeeID
    SelReturn

End Sub

Public Sub Init()

    SQLString = "SELECT EmployeeNumber FROM PREmployee ORDER BY EmployeeNumber DESC"
    rsInit SQLString, cn, rs
    
    If rs.BOF And rs.EOF Then
        Me.tdbnumEmployeeNumber = 100
    Else
        rs.MoveFirst
        Me.tdbnumEmployeeNumber = rs!EmployeeNumber + 1
    End If

End Sub

Private Sub SelReturn()
    Me.Hide
End Sub

