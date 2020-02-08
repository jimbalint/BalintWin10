VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmCopyBB 
   Caption         =   "Copy Branch / Budget"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopyBB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHiLook 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   8760
      Picture         =   "frmCopyBB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cmdLoLook 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4560
      Picture         =   "frmCopyBB.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   2693
      TabIndex        =   16
      Top             =   1800
      Width           =   4335
      Begin VB.OptionButton optMain 
         Caption         =   "&Main (High)"
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optSub 
         Caption         =   "&Sub (Low)"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2693
      TabIndex        =   11
      Top             =   3480
      Width           =   4335
      Begin VB.OptionButton optGo 
         Caption         =   "&Execute"
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optShow 
         Caption         =   "&Display"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2513
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoAccount 
      Height          =   375
      Left            =   2963
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmCopyBB.frx":091E
      Caption         =   "frmCopyBB.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCopyBB.frx":09AC
      Keys            =   "frmCopyBB.frx":09CA
      Spin            =   "frmCopyBB.frx":0A14
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
      ValueVT         =   1376257
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiAccount 
      Height          =   375
      Left            =   7133
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmCopyBB.frx":0A3C
      Caption         =   "frmCopyBB.frx":0A5C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCopyBB.frx":0ACA
      Keys            =   "frmCopyBB.frx":0AE8
      Spin            =   "frmCopyBB.frx":0B32
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
      ValueVT         =   36372481
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoValue 
      Height          =   375
      Left            =   3308
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmCopyBB.frx":0B5A
      Caption         =   "frmCopyBB.frx":0B7A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCopyBB.frx":0BE8
      Keys            =   "frmCopyBB.frx":0C06
      Spin            =   "frmCopyBB.frx":0C50
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   ""
      EditMode        =   2
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "########0"
      HighlightText   =   0
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
      ValueVT         =   2162689
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiValue 
      Height          =   375
      Left            =   7478
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmCopyBB.frx":0C78
      Caption         =   "frmCopyBB.frx":0C98
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCopyBB.frx":0D06
      Keys            =   "frmCopyBB.frx":0D24
      Spin            =   "frmCopyBB.frx":0D6E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   ""
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
      ValueVT         =   33554433
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblFrom 
      Caption         =   "Sub to copy &FROM:"
      Height          =   255
      Left            =   908
      TabIndex        =   15
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblTo 
      Caption         =   "Sub to copy &TO:"
      Height          =   255
      Left            =   5588
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLoAccount 
      Caption         =   "&Low Account:"
      Height          =   255
      Left            =   1253
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblHiAccount 
      Caption         =   "&High Account:"
      Height          =   255
      Left            =   5213
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
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
      Height          =   375
      Left            =   353
      TabIndex        =   10
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmCopyBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X, Y As String
Dim i, j As Long

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdHiLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbHiAccount = frmAcctLookup.SelAcct
    Me.optSub.SetFocus
End Sub

Private Sub cmdLoLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbLoAccount = frmAcctLookup.SelAcct
    Me.tdbHiAccount.SetFocus
End Sub

Private Sub cmdOk_Click()

    If frmCopyBB.optMain Then
       X = "Main"
    Else
       X = "Sub"
    End If
           
    If frmCopyBB.optShow Then
       Y = "Show"
    Else
       Y = "Go"
    End If
            
    Set uDB = CopyBB(frmCopyBB.tdbLoAccount, _
                     frmCopyBB.tdbHiAccount, _
                     frmCopyBB.tdbLoValue, _
                     frmCopyBB.tdbHiValue, _
                     X, _
                     GLCompany.SubDigits, _
                     Y)
                        
    Set frmResults = New frmResults
    frmResults.lblCompanyName = GLCompany.Name
    frmResults.lblMsg1 = "Copy Branch / Budget"
    frmResults.lblMsg2 = ""
    frmResults.lblMsg3 = ""
    For i = 1 To uDB.UpperBound(1)
        frmResults.List1.AddItem uDB(i, 0)
    Next i
    frmResults.Show vbModal
    If Y = "Show" Then
       i = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
       If i = vbNo Then GoBack
       Me.tdbLoAccount.SetFocus
    Else
       GoBack
    End If

End Sub

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    Me.tdbLoAccount = 1
    Me.tdbHiAccount = 999999999

End Sub

Private Sub optSub_Click()
    If Me.optSub = False Then
       Me.lblFrom.Caption = "Main to copy &FROM:"
       Me.lblTo.Caption = "Main to copy &TO:"
    Else
       Me.lblFrom.Caption = "Sub to copy &FROM:"
       Me.lblTo.Caption = "Sub to copy &TO:"
    End If
    Me.lblFrom.Refresh
    Me.lblTo.Refresh
End Sub

Private Sub optMain_Click()
    If Me.optSub = False Then
       Me.lblFrom = "Main to copy &FROM:"
       Me.lblTo = "Main to copy &TO:"
    Else
       Me.lblFrom = "Sub to copy &FROM:"
       Me.lblTo = "Sub to copy &TO:"
    End If
    Me.lblFrom.Refresh
    Me.lblTo.Refresh
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

