VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmMultDiv 
   Caption         =   "Multiply / Divide GL Accounts"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultDiv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraType 
      Height          =   735
      Left            =   2580
      TabIndex        =   19
      Top             =   3120
      Width           =   4935
      Begin VB.OptionButton optBase 
         Caption         =   "&Base Account"
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optWhole 
         Caption         =   "&Whole Account"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdHiLook 
      Height          =   375
      Left            =   9240
      Picture         =   "frmMultDiv.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdLoLook 
      Height          =   375
      Left            =   4800
      Picture         =   "frmMultDiv.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin TDBNumber6Ctl.TDBNumber tdbValue 
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Calculator      =   "frmMultDiv.frx":091E
      Caption         =   "frmMultDiv.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMultDiv.frx":09AA
      Keys            =   "frmMultDiv.frx":09C8
      Spin            =   "frmMultDiv.frx":0A12
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoAccount 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1185
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmMultDiv.frx":0A3A
      Caption         =   "frmMultDiv.frx":0A5A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMultDiv.frx":0AC6
      Keys            =   "frmMultDiv.frx":0AE4
      Spin            =   "frmMultDiv.frx":0B2E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;0;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   33030145
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5700
      TabIndex        =   10
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2700
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2880
      TabIndex        =   15
      Top             =   4200
      Width           =   4335
      Begin VB.OptionButton optGo 
         Caption         =   "&Execute"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optShow 
         Caption         =   "&Display"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1080
      TabIndex        =   13
      Top             =   2040
      Width           =   4335
      Begin VB.OptionButton optDivide 
         Caption         =   "D&ivide"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optMultiply 
         Caption         =   "&Multiply"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiAccount 
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   1185
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   661
      Calculator      =   "frmMultDiv.frx":0B56
      Caption         =   "frmMultDiv.frx":0B76
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmMultDiv.frx":0BE2
      Keys            =   "frmMultDiv.frx":0C00
      Spin            =   "frmMultDiv.frx":0C4A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;0;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   33030145
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblValue 
      Caption         =   "&Value to Apply:"
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblCompanyName 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1860
      TabIndex        =   14
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label lblHiAccount 
      Caption         =   "&Hi Account:"
      Height          =   360
      Left            =   5460
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblLoAccount 
      Caption         =   "&Low Account:"
      Height          =   360
      Left            =   900
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmMultDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X, Y As String
Dim i As Long

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdHiLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbHiAccount = frmAcctLookup.SelAcct
    Me.optMultiply.SetFocus
End Sub

Private Sub cmdLoLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbLoAccount = frmAcctLookup.SelAcct
    Me.tdbHiAccount.SetFocus
End Sub

Private Sub cmdOk_Click()
    
    If Me.tdbValue = 0 Then
       MsgBox "You must supply a value !!!", vbExclamation + vbOKOnly
       Me.tdbValue.SetFocus
       Exit Sub
    End If

    If Me.optMultiply Then
       X = "Mult"
    Else
       X = "Div"
    End If
            
    If Me.optShow Then
       Y = "Show"
    Else
       Y = "Go"
    End If
            
    Set uDB = GLMultDiv(Me.tdbLoAccount.Value, _
                        Me.tdbHiAccount.Value, _
                        X, _
                        Me.tdbValue, _
                        Me.optBase, _
                        Y)
         
    Set frmResults = New frmResults
    frmResults.lblCompanyName = GLCompany.Name
    frmResults.lblMsg1 = "Multiply/Divide Account Numbers"
    frmResults.lblMsg2 = ""
    frmResults.lblMsg3 = ""
    For i = 1 To uDB.UpperBound(1)
        frmResults.List1.AddItem uDB(i, 0)
    Next i
    frmResults.Show vbModal
    If Y = "Show" Then
       i = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
       If i = vbNo Then GoBack
    Else
       GoBack   ' done
    End If

    Me.tdbLoAccount.SetFocus

End Sub

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.tdbLoAccount = 1
    Me.tdbHiAccount = 999999999
    
    If GLCompany.SubDigits = 0 Then
        fraType.Enabled = False
    End If
    
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

