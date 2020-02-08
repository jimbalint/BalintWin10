VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDirectPrint 
   Caption         =   "Direct Print"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
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
   ScaleHeight     =   3615
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdl 
      Left            =   240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFileLookUp 
      Caption         =   "&Search"
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "<Select a File to Print>"
      Top             =   1440
      Width           =   7095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Default         =   -1  'True
      Height          =   735
      Left            =   5460
      TabIndex        =   4
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   735
      Left            =   1380
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cmbPrinters 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "<Select Printer>"
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "File to Print:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmDirectPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As String
Dim rsPrinters As New ADODB.Recordset
Dim i As Integer
Dim ch As Variant


Private Sub Form_Load()

    LoadPrinters
    Me.KeyPreview = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:   cmdExit_Click
    End Select
End Sub
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdFileLookUp_Click()
        
    cdl.CancelError = True
    
    ' set to current
    cdl.Flags = cdlCFBoth Or cdlCFEffects
    ' cdl.Filter = "Comma Separated Values|*.csv"
    ' cdl.FileName = GLUser.Logon & ".csv"
    cdl.DialogTitle = "Select a file to print"
    cdl.CancelError = True
    ' cdl.InitDir = "\Balint\Data"

    ' call the file dialog
    On Error Resume Next
    cdl.ShowOpen
    Me.txtFile = cdl.FileName
    
End Sub


Private Sub LoadPrinters()

    For i = 0 To frmRptPrvw.vsp.NDevices - 1
        Me.cmbPrinters.AddItem frmRptPrvw.vsp.Devices(i)
    Next i
    
End Sub

Private Sub cmdPrint_Click()
        
    ' assign
    Me.txtFile = cdl.FileName
    ch = FreeFile

    Open Me.txtFile.Text For Input As #ch

    DP_Init Me.cmbPrinters.Text

    Do Until EOF(ch)
        Line Input #ch, x
        DP_PrintLine x
        ' DP_LF
    Loop
    DP_EndDoc
    End
End Sub

