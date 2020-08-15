VERSION 5.00
Begin VB.Form frmInvChangePrinter 
   Caption         =   "Change Invoice Printer"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   6855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblComputerName 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "frmInvChangePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ComputerName As String
Dim i, j, k As Integer
Dim x, y, z As String
Dim strSQL As String


Private Sub Form_Load()

    ComputerName = Environ("computername")
    Me.lblComputerName.Caption = ComputerName
    Me.KeyPreview = True
    LoadPrinters

End Sub

Private Sub LoadPrinters()

    Set Prvw = New frmPreview
    With Me.cmbPrinter
        For i = 0 To Prvw.vsp.NDevices - 1
            .AddItem Prvw.vsp.Devices(i)
        Next i
        If i > 0 Then Me.cmbPrinter.ListIndex = 0
    End With

End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    strSQL = "select * from InvGlobal where CompanyID = " & PRCompany.CompanyID & _
            " and TypeCode = " & InvEquate.GlobalTypeInvPrinter2 & _
            " and Var1 = '" & ComputerName & "'"
    If InvGlobal.GetBySQL(strSQL) = False Then
        InvGlobal.Clear
        InvGlobal.CompanyID = PRCompany.CompanyID
        InvGlobal.TypeCode = InvEquate.GlobalTypeInvPrinter2
        InvGlobal.Var1 = ComputerName
        InvGlobal.Var2 = Me.cmbPrinter.text
        InvGlobal.rsAdd
    Else
        InvGlobal.Var2 = Me.cmbPrinter.text
        InvGlobal.rsPut
    End If
    Me.Hide
        
End Sub

