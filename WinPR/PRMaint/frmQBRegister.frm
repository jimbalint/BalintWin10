VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmQBRegister 
   Caption         =   "Register QuickBooks Company File"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   9090
   ScaleWidth      =   12390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdQBGO 
      Caption         =   "GET &QB INFORMATION"
      Height          =   975
      Left            =   4928
      TabIndex        =   7
      Top             =   6360
      Width           =   2535
   End
   Begin TDBText6Ctl.TDBText tdbtxtQBCompanyName 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   12135
      _Version        =   65536
      _ExtentX        =   21405
      _ExtentY        =   1296
      Caption         =   "frmQBRegister.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBRegister.frx":0084
      Key             =   "frmQBRegister.frx":00A2
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "TDBText1"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   735
      Left            =   6848
      TabIndex        =   2
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&REGISTER"
      Height          =   735
      Left            =   3488
      TabIndex        =   1
      Top             =   7560
      Width           =   2055
   End
   Begin TDBText6Ctl.TDBText tdbtxtQBFileName 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   12135
      _Version        =   65536
      _ExtentX        =   21405
      _ExtentY        =   1296
      Caption         =   "frmQBRegister.frx":00E6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBRegister.frx":0164
      Key             =   "frmQBRegister.frx":0182
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "TDBText1"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBText6Ctl.TDBText tdbtxtQBFedID 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   12135
      _Version        =   65536
      _ExtentX        =   21405
      _ExtentY        =   1296
      Caption         =   "frmQBRegister.frx":01C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmQBRegister.frx":0246
      Key             =   "frmQBRegister.frx":0264
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   -1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   1
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   0
      LengthAsByte    =   0
      Text            =   "TDBText1"
      Furigana        =   0
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   1448
      TabIndex        =   8
      Top             =   8520
      Width           =   9495
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1928
      TabIndex        =   6
      Top             =   5640
      Width           =   8535
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "frmQBRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Answer As VbMsgBoxResult
Dim PreRegistered As Boolean

Dim xmlMgr As New QBXMLRP2Lib.RequestProcessor2

Dim Ticket As String

Private Sub Form_Load()

    ' PRGlobal
    ' TypeCode      = GlobalTypeQB_Register
    ' UserID        = CompanyID
    ' Var1          = "0" = Company Default / UserID
    ' Var2          = FedID from QB
    ' Var3          = Company name from QB
    ' Var4          = QB File Name (w/ full path)
    ' Var5          = QB File Name only

    Me.lblCompanyName = PRCompany.Name
    Me.cmdRegister.Enabled = False
    
    ' already registered?
    PreRegistered = False
    SQLString = "SELECT * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQB_Register & _
                " AND UserID = " & PRCompany.CompanyID & _
                " AND Var1 = '0'"
    If PRGlobal.GetBySQL(SQLString) = True Then
        Answer = MsgBox(PRCompany.Name & " already registered to QuickBooks File:" & vbCr & vbCr & _
                      Trim(PRGlobal.Var4) & vbCr & _
                      Trim(PRGlobal.Var3) & vbCr & vbCr & _
                      "OK to re-register?", vbExclamation + vbOKCancel)
        If Answer = vbCancel Then GoBack
        PreRegistered = True
    End If

    Me.tdbtxtQBCompanyName.Text = ""
    Me.tdbtxtQBFedID.Text = ""
    Me.tdbtxtQBFileName.Text = ""
    Me.lblMsg1.Caption = ""

    Me.lblInfo = "You must have the QuickBooks file open" & vbCr & _
                 "And be logged in under the Admin user ..."

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
Private Sub Command1_Click()
    
    lblMsg1 = "Test QB company get ..."
    Me.Refresh
    SessMgr.OpenConnection "", "Balint Accounting"
    
    lblMsg1 = "Beginning QB Session B"
    Me.Refresh
    SessMgr.BeginSession "", omDontCare
    
    If GetQBCompany = False Then
        MsgBox "no go ...."
        GoBack
    End If
    
    If IsNull(QBFedID) Or QBFedID = "" Then
        MsgBox "Please fill in the Federal ID number in QuickBooks and try again!", vbExclamation
        SessMgr.EndSession
        SessMgr.CloseConnection
        Exit Sub
    End If
    
    Me.tdbtxtQBFedID = QBFedID
    Me.tdbtxtQBCompanyName = QBCompanyName

    Me.lblMsg1 = ""

    Me.Refresh

End Sub


Private Sub cmdQBGO_Click()

    ' ??????????????????????????????
    ' WTF ??? - just in case ???
    On Error Resume Next
    SessMgr.CloseConnection
    xmlMgr.CloseConnection
    On Error GoTo 0
    ' ??????????????????????????????
    
    ' use XML to get the file name
    Me.lblMsg1 = "Opening QB Connection A"
    Me.Refresh

    xmlMgr.OpenConnection2 "", "Balint Accounting", localQBD

    Me.lblMsg1 = "Opening QB Connection A1"
    Me.Refresh

    Delay 3

    Me.lblMsg1 = "Beginning QB Session A2"
    Me.Refresh
    Ticket = xmlMgr.BeginSession("", qbFileOpenDoNotCare)

    Me.tdbtxtQBFileName = xmlMgr.GetCurrentCompanyFileName(Ticket)

    Me.lblMsg1 = "Closing QB Connection A"
    Me.Refresh

    xmlMgr.EndSession (Ticket)
    xmlMgr.CloseConnection
    
    lblMsg1 = "Opening QB Connection B"
    Me.Refresh
    SessMgr.OpenConnection "", "Balint Accounting"
    
    lblMsg1 = "Beginning QB Session B"
    Me.Refresh
    SessMgr.BeginSession "", omDontCare
    
    If GetQBCompany = False Then GoBack
    
    If IsNull(QBFedID) Or QBFedID = "" Then
        MsgBox "Please fill in the Federal ID number in QuickBooks and try again!", vbExclamation
        SessMgr.EndSession
        SessMgr.CloseConnection
        Exit Sub
    End If
    
    Me.tdbtxtQBFedID = QBFedID
    Me.tdbtxtQBCompanyName = QBCompanyName

    Me.lblMsg1 = ""

    Me.Refresh

    Me.cmdRegister.Enabled = True

    SessMgr.EndSession
    SessMgr.CloseConnection

End Sub

Private Sub cmdRegister_Click()

    ' clear out old PRGlobal if registered before
    If PreRegistered = True Then
        SQLString = "DELETE * FROM PRGlobal WHERE TypeCode = " & PREquate.GlobalTypeQB_Register & _
                    " AND UserID = " & PRCompany.CompanyID
        cnDes.Execute SQLString
    End If

    ' write the registration record
    PRGlobal.Clear
    PRGlobal.TypeCode = PREquate.GlobalTypeQB_Register
    PRGlobal.UserID = PRCompany.CompanyID
    PRGlobal.Var1 = "0"
    PRGlobal.Var2 = Me.tdbtxtQBFedID
    PRGlobal.Var3 = Me.tdbtxtQBCompanyName
    PRGlobal.Var4 = Me.tdbtxtQBFileName
    
    ' store just the file name
    PRGlobal.Var5 = GetFileName(Me.tdbtxtQBFileName)
    
    PRGlobal.Save (Equate.RecAdd)

    ' create / clear QBAccount table
    If TableExists("QBAccount", cn) = False Then
        QBAccountCreate
    End If
    SQLString = "DELETE * FROM QBAccount"
    cn.Execute SQLString
    
    MsgBox "Registration of: " & PRCompany.Name & vbCr & vbCr & _
           "To QuickBooks File: " & vbCr & vbCr & _
           Trim(Me.tdbtxtQBFileName) & vbCr & vbCr & _
           "Complete ....", vbInformation

    GoBack

End Sub


