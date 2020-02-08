VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmInvNow 
   Caption         =   "KP Invoicing"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvNow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin TDBDate6Ctl.TDBDate tdbInvDate 
      Height          =   735
      Left            =   3690
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
      _ExtentY        =   1296
      Calendar        =   "frmInvNow.frx":030A
      Caption         =   "frmInvNow.frx":040A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInvNow.frx":0484
      Keys            =   "frmInvNow.frx":04A2
      Spin            =   "frmInvNow.frx":0500
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm/dd/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "09/03/2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40424
      CenturyMode     =   0
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   855
      Left            =   1650
      Picture         =   "frmInvNow.frx":0528
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&SAVE TO QB"
      Height          =   855
      Left            =   4290
      Picture         =   "frmInvNow.frx":0832
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   855
      Left            =   6810
      Picture         =   "frmInvNow.frx":0B3C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
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
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   9495
   End
End
Attribute VB_Name = "frmInvNow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean

Private Sub Form_Load()

    Me.lblCompanyName = PRCompany.Name
    
    tdbDateSet Me.tdbInvDate, InvHeader.OrderDate
    InvHeader.InvoiceDate = InvHeader.OrderDate
    InvHeader.rsPut
    
    Me.KeyPreview = True

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdCancel_Click
    End Select
End Sub

Private Sub cmdCancel_Click()
    InvHeader.InvoiceDate = 0
    InvHeader.rsPut
    frmInvProcess.OK = False
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    InvHeader.InvoiceDate = Me.tdbInvDate.Value
    InvHeader.rsPut
    
    SQLString = "SELECT * FROM InvGlobal WHERE CompanyID = " & PRCompany.CompanyID & _
                " AND TypeCode = " & InvEquate.GlobalTypeInvPrinter
    If InvGlobal.GetBySQL(SQLString) = False Then
        MsgBox "Use Global Maintenance to select the invoice printer!", vbExclamation
        Exit Sub
    End If
    
    KP_PrintInvoice frmInvProcess.tdbnumInvNum, InvGlobal.Var1
    
End Sub

Private Sub cmdOK_Click()
    
    With frmInvProcess
        .OK = True
        .InvDate = Me.tdbInvDate.Value
        Unload Me
    End With

End Sub


