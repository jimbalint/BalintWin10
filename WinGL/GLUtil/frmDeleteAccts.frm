VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmDeleteAccts 
   Caption         =   "Delete Accounts"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeleteAccts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHiLook 
      Height          =   375
      Left            =   8760
      Picture         =   "frmDeleteAccts.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdLoLook 
      Height          =   375
      Left            =   4200
      Picture         =   "frmDeleteAccts.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2280
      Width           =   495
   End
   Begin VB.CheckBox chkDelHist 
      Caption         =   "D&elete History Records ?"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3323
      TabIndex        =   6
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5400
      TabIndex        =   8
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   2520
      TabIndex        =   13
      Top             =   3120
      Width           =   4335
      Begin VB.OptionButton optShow 
         Caption         =   "&Display"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optGo 
         Caption         =   "&Execute"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2243
      TabIndex        =   9
      Top             =   1080
      Width           =   4935
      Begin VB.OptionButton optSub 
         Caption         =   "&Sub Accounts"
         Height          =   375
         Left            =   2640
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optRegular 
         Caption         =   "&Regular"
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin TDBNumber6Ctl.TDBNumber tdbLoValue 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmDeleteAccts.frx":091E
      Caption         =   "frmDeleteAccts.frx":093E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDeleteAccts.frx":09AC
      Keys            =   "frmDeleteAccts.frx":09CA
      Spin            =   "frmDeleteAccts.frx":0A14
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
      ValueVT         =   50724865
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHiValue 
      Height          =   375
      Left            =   7163
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   661
      Calculator      =   "frmDeleteAccts.frx":0A3C
      Caption         =   "frmDeleteAccts.frx":0A5C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmDeleteAccts.frx":0ACA
      Keys            =   "frmDeleteAccts.frx":0AE8
      Spin            =   "frmDeleteAccts.frx":0B32
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
      ValueVT         =   34799617
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblHiValue 
      Caption         =   "High Account:"
      Height          =   255
      Left            =   5243
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblLoValue 
      Caption         =   "Low Account:"
      Height          =   255
      Left            =   923
      TabIndex        =   11
      Top             =   2400
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
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "frmDeleteAccts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mb As Integer
Dim X, Y As String
Dim i, j As Long

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Sub cmdHiLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbHiValue = frmAcctLookup.SelAcct
    Me.optShow.SetFocus
End Sub

Private Sub cmdLoLook_Click()
    frmAcctLookup.Show vbModal
    Me.tdbLoValue = frmAcctLookup.SelAcct
    Me.tdbHiValue.SetFocus
End Sub

Private Sub cmdOk_Click()
    
    If Me.optGo Then
       mb = MsgBox("Are you sure you want to delete the indicated accounts ?", vbCritical + vbOKCancel + vbDefaultButton2, "Delete Accounts !!!")
       If mb = vbCancel Then
          MsgBox "No accounts will be deleted !", vbInformation, ""
          GoBack
       End If
    End If

    If frmDeleteAccts.optRegular Then
       X = "Acct"
    Else
       X = "Sub"
    End If
            
    If frmDeleteAccts.optShow Then
       Y = "Show"
    Else
       Y = "Go"
    End If
            
    Set uDB = DeleteAccts(X, _
                          Me.tdbLoValue, _
                          Me.tdbHiValue, _
                          Y, _
                          Me.chkDelHist)
         
    Set frmResults = New frmResults
    frmResults.lblCompanyName = GLCompany.Name
    frmResults.lblMsg1 = "Delete Accounts"
    frmResults.lblMsg2 = ""
    frmResults.lblMsg3 = ""
    For i = 1 To uDB.UpperBound(1)
        frmResults.List1.AddItem uDB(i, 0)
    Next i
    frmResults.Show vbModal
    If Y = "Show" Then
       i = MsgBox("Try Again ?", vbQuestion + vbYesNo + vbDefaultButton1, "Multiply/Divide Accounts")
       If i = vbNo Then GoBack
       Me.tdbLoValue.SetFocus
    Else
       GoBack
    End If

End Sub

Private Sub Form_Load()
    Me.lblCompanyName = GLCompany.Name
    Me.tdbLoValue = 1
    Me.tdbHiValue = 999999999
End Sub

Private Sub Form_Terminate()
    GoBack
End Sub

Private Sub optGo_Click()
    If optGo = False Then
       Me.chkDelHist.Enabled = False
    Else
       Me.chkDelHist.Enabled = True
    End If
    Me.Refresh
End Sub

Private Sub optShow_Click()
    If optGo = False Then
       Me.chkDelHist.Enabled = False
    Else
       Me.chkDelHist.Enabled = True
    End If
    Me.Refresh
End Sub


Private Sub optRegular_Click()
    If optRegular Then
       Me.lblLoValue = "Low Account:"
       Me.lblHiValue = "High Account:"
    Else
       Me.lblLoValue = "Low Sub:"
       Me.lblHiValue = "High Sub:"
    End If
    Me.Refresh
End Sub


Private Sub optSub_Click()
    If Not optSub Then
       Me.lblLoValue = "Low Account:"
       Me.lblHiValue = "High Account:"
    Else
       Me.lblLoValue = "Low Sub:"
       Me.lblHiValue = "High Sub:"
    End If
    Me.Refresh
End Sub


