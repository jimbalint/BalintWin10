VERSION 5.00
Begin VB.Form frmYECityTaxRpt 
   Caption         =   "Year End City Tax Report"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   ScaleHeight     =   3660
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4268
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   788
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox cmbYear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3668
      TabIndex        =   0
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Select Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1988
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   68
      TabIndex        =   3
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmYECityTaxRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrYear As Long


Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub Form_Load()

    CurrYear = Year(Now())

    cmbYear.AddItem CurrYear
    cmbYear.AddItem CurrYear - 1
    cmbYear.AddItem CurrYear - 2
    cmbYear.AddItem CurrYear - 3
    cmbYear.AddItem CurrYear - 4
    cmbYear.AddItem CurrYear - 5
    cmbYear.AddItem CurrYear - 6
    cmbYear.AddItem CurrYear - 7
    cmbYear.AddItem CurrYear - 8
    cmbYear.AddItem CurrYear - 9
    cmbYear.AddItem CurrYear - 10
    
    cmbYear.ListIndex = 0
    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True
End Sub

Private Sub cmbYear_Change()
'    qYear = cmbYear
'    StartDate = "01/01/" & qYear
'    EndDate = "12/31/" & qYear
End Sub

Private Sub cmdExit_Click()
    InitFlag = False
    Me.Hide
    GoBack
End Sub

Private Sub cmdOK_Click()
    qYear = cmbYear.Text
    Startdate = "01/01/" & qYear
    EndDate = "12/31/" & qYear
    
    If cmbYear = 0 And Startdate = 0 And EndDate = 0 Then
        MsgBox "PLEASE ENTER A YEAR", vbCritical, "Year End City Tax Report"
    Else
        InitFlag = True
        YECityTax
    End If

    
    
End Sub


