VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmAccountChange 
   Caption         =   "Change GL Account Number"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   795
      Left            =   7133
      TabIndex        =   3
      Top             =   4560
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   795
      Left            =   3113
      TabIndex        =   2
      Top             =   4500
      Width           =   1755
   End
   Begin TDBNumber6Ctl.TDBNumber tdbGLAccount 
      Height          =   975
      Left            =   900
      TabIndex        =   1
      Top             =   2880
      Width           =   5475
      _Version        =   65536
      _ExtentX        =   9657
      _ExtentY        =   1720
      Calculator      =   "frmAccountChange.frx":0000
      Caption         =   "frmAccountChange.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAccountChange.frx":00A6
      Keys            =   "frmAccountChange.frx":00C4
      Spin            =   "frmAccountChange.frx":010E
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
   Begin VB.ComboBox cmbGLAccount 
      Height          =   375
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2160
      Width           =   10275
   End
   Begin VB.Label Label1 
      Caption         =   "Select Account Number to Change:"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   3435
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   11235
   End
End
Attribute VB_Name = "frmAccountChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I, J, K As Long
Dim x, Y, Z As String
Dim rs As New ADODB.Recordset
Dim ChangeFrom, ChangeTo As Long

Private Sub Form_Load()

    ' load accounts to the combo list
    If GLAccount.GetAllAccounts = False Then
        MsgBox "No GL Accounts found!", vbExclamation
        GoBack
    End If
    Me.MousePointer = vbHourglass
    Me.lblCompanyName = GLCompany.Name & vbCr & "Now loading GL Accounts ..."
    Me.Refresh
    Do
        I = I + 1
        If I Mod 50 = 1 Then
            Me.lblCompanyName = GLCompany.Name & vbCr & _
                              "Loading GL Accounts " & GLAccount.Account
            Me.Refresh
        End If
        With Me.cmbGLAccount
            .AddItem GLAccount.Account & " " & GLAccount.FullDesc
            .ItemData(.NewIndex) = GLAccount.Account
        End With
        If GLAccount.GetNext = False Then Exit Do
    Loop
    Me.MousePointer = vbArrow
    Me.lblCompanyName = GLCompany.Name
    Me.Refresh
    Me.cmbGLAccount.ListIndex = 1
    
    With Me.tdbGLAccount
        .Format = "########0"
        .MinValue = 0
        .MaxValue = 999999999
        .DisplayFormat = ""
    End With
    
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: CmdExit_Click
    End Select
End Sub

Private Sub CmdExit_Click()
    GoBack
End Sub

Private Sub cmdOK_Click()
    
    If IsNull(Me.tdbGLAccount) Then Exit Sub
    If Me.tdbGLAccount = 0 Then Exit Sub
    
    ChangeFrom = Me.cmbGLAccount.ItemData(Me.cmbGLAccount.ListIndex)
    ChangeTo = CLng(Me.tdbGLAccount)
    
    x = "OK to change: " & vbCr & Me.cmbGLAccount.Text & vbCr & "To: " & ChangeTo
    If MsgBox(x, vbYesNo + vbQuestion) = vbNo Then Exit Sub

    If GLAccount.GetAccount(ChangeTo) = True Then
        MsgBox Me.tdbGLAccount & " Account already exists!", vbExclamation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    If GLAccount.GetAccount(ChangeFrom) = False Then
        MsgBox "GLAccount Error: " & ChangeFrom, vbExclamation
        GoBack
    End If
    GLAccount.Account = ChangeTo
    GLAccount.Save (Equate.RecPut)
    
    I = 0
    SQLString = "SELECT * FROM GLHistory WHERE Account = " & Me.cmbGLAccount.ItemData(Me.cmbGLAccount.ListIndex)
    rsInit SQLString, cn, rs
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
            I = I + 1
            If I Mod 20 = 1 Then
                Me.lblCompanyName = GLCompany.Name & vbCr & "Changing History: " & I
                Me.Refresh
            End If
            rs!Account = CLng(Me.tdbGLAccount)
            rs.Update
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    I = 0
    rs.Close
    SQLString = "SELECT * FROM GLAmount WHERE Account = " & Me.cmbGLAccount.ItemData(Me.cmbGLAccount.ListIndex)
    rsInit SQLString, cn, rs
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do
            I = I + 1
            If I Mod 20 = 1 Then
                Me.lblCompanyName = GLCompany.Name & vbCr & "Changing History: " & I
                Me.Refresh
            End If
            rs!Account = CLng(Me.tdbGLAccount)
            rs.Update
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Me.MousePointer = vbArrow
    Me.lblCompanyName = GLCompany.Name
    Me.Refresh
    
    MsgBox "Account: " & Me.cmbGLAccount.Text & vbCr & "Change to: " & ChangeTo, vbInformation

    GoBack

End Sub


