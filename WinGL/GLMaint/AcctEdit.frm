VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmAcctEdit 
   Caption         =   "Account Edit"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDescLook 
      Height          =   375
      Left            =   3360
      Picture         =   "AcctEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2280
      Width           =   375
   End
   Begin TDBText6Ctl.TDBText txtDescription 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   5295
      _Version        =   65536
      _ExtentX        =   9340
      _ExtentY        =   503
      Caption         =   "AcctEdit.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":0376
      Key             =   "AcctEdit.frx":0394
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
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
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin TDBNumber6Ctl.TDBNumber txtBalSheetCol 
      Height          =   285
      Left            =   9600
      TabIndex        =   7
      Top             =   3480
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":03D8
      Caption         =   "AcctEdit.frx":03F8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":0464
      Keys            =   "AcctEdit.frx":0482
      Spin            =   "AcctEdit.frx":04CC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   255
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   6553605
      MinValueVT      =   6619141
   End
   Begin TDBNumber6Ctl.TDBNumber txtLineFeeds 
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":04F4
      Caption         =   "AcctEdit.frx":0514
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":0580
      Keys            =   "AcctEdit.frx":059E
      Spin            =   "AcctEdit.frx":05E8
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
      MaxValue        =   255
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtPrintTab 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":0610
      Caption         =   "AcctEdit.frx":0630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":069C
      Keys            =   "AcctEdit.frx":06BA
      Spin            =   "AcctEdit.frx":0704
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
      MaxValue        =   255
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   7602181
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtTotalLevel 
      Height          =   285
      Left            =   7080
      TabIndex        =   4
      Top             =   2880
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":072C
      Caption         =   "AcctEdit.frx":074C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":07B8
      Keys            =   "AcctEdit.frx":07D6
      Spin            =   "AcctEdit.frx":0820
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99
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
   Begin TDBNumber6Ctl.TDBNumber txtDescNum 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   2295
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":0848
      Caption         =   "AcctEdit.frx":0868
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":08D4
      Keys            =   "AcctEdit.frx":08F2
      Spin            =   "AcctEdit.frx":093C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5242885
      MinValueVT      =   3014661
   End
   Begin TDBNumber6Ctl.TDBNumber txtAcctNumber 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1050
      Width           =   3015
      _Version        =   65536
      _ExtentX        =   5318
      _ExtentY        =   503
      Calculator      =   "AcctEdit.frx":0964
      Caption         =   "AcctEdit.frx":0984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "AcctEdit.frx":09F0
      Keys            =   "AcctEdit.frx":0A0E
      Spin            =   "AcctEdit.frx":0A58
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "########0;;Null"
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
      ValueVT         =   2088828933
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CheckBox chkSignRevSched 
      Caption         =   "Sign Reverse on Schedule"
      Height          =   255
      Left            =   7590
      TabIndex        =   15
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CheckBox chkSignRevStmt 
      Caption         =   "Sign Reverse on Statement"
      Height          =   375
      Left            =   7590
      TabIndex        =   11
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox chkDollarSign 
      Caption         =   "Dollar Sign"
      Height          =   375
      Left            =   5025
      TabIndex        =   14
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CheckBox chkTlOnLed 
      Caption         =   "Total On Ledger"
      Height          =   375
      Left            =   5025
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CheckBox chkConsAcct 
      Caption         =   "Consolidated Account"
      Height          =   375
      Left            =   2700
      TabIndex        =   13
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CheckBox chkBranchAcct 
      Caption         =   "Branch Account"
      Height          =   375
      Left            =   2700
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CheckBox chkAllSched 
      Caption         =   "All Schedules"
      Height          =   375
      Left            =   750
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox chkAllStmts 
      Caption         =   "All Statements"
      Height          =   375
      Left            =   750
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdAmounts 
      Caption         =   "&Amounts"
      Height          =   375
      Left            =   7530
      TabIndex        =   18
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4950
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2370
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox cmbAcctType 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   2895
      Width           =   3015
   End
   Begin VB.Label lblDescText 
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   2325
      Width           =   6735
   End
   Begin VB.Label lblDescNum 
      Caption         =   "Desc #:"
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   2325
      Width           =   735
   End
   Begin VB.Label lblMsg1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
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
      Left            =   2340
      TabIndex        =   26
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label lblBalSheetCol 
      Caption         =   "Balance Sheet Column:"
      Height          =   255
      Left            =   7440
      TabIndex        =   25
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Label lblLineFeeds 
      Caption         =   "Line Feeds:"
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   3510
      Width           =   1215
   End
   Begin VB.Label lblPrintTab 
      Caption         =   "Print Tab:"
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   3510
      Width           =   855
   End
   Begin VB.Label lblTotalLevel 
      Caption         =   "Total Level:"
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   2955
      Width           =   1095
   End
   Begin VB.Label lblType 
      Caption         =   "Type:"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   2955
      Width           =   975
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   1725
      Width           =   1095
   End
   Begin VB.Label lblAcctNumber 
      Caption         =   "Account #:"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmAcctEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAmounts_Click()
    
Dim EditAcct As Long
Dim boo As Boolean
    
    If GLAccount.AcctType <> "0" Then Exit Sub
    
    EditAcct = GLAccount.Account
    
    frmAcctAmounts.DisplayAccount = GLAccount.Account
    frmAcctAmounts.Show vbModal
    
    ' re-find the account
    boo = GLAccount.Find(EditAcct)

End Sub

Private Sub cmdDescLook_Click()
    
    frmDescriptions.LookUp = 1
    frmDescriptions.Show vbModal
    Me.txtDescNum = frmDescriptions.SelectedNumber
    Me.lblDescText = frmDescriptions.SelectedDescription

End Sub

Private Sub CmdExit_Click()
'   Response = MsgBox("Are you sure you want to exit?", vbExclamation + vbOKCancel, "Record will not be saved!")
'   If Response = True Then Unload Me

   Unload Me

End Sub

Private Sub cmdOK_Click()
   
   If FormType = Equate.FormAdd Or FormType = Equate.FormEdit Then
      
      GLAccount.Account = CLng(txtAcctNumber.Text)
      GLAccount.Description = txtDescription.Text
      GLAccount.DescNumber = txtDescNum.Text
      GLAccount.AcctType = Mid(cmbAcctType.Text, 1, 1)
      GLAccount.TotalLevel = CByte(txtTotalLevel.Text)
      GLAccount.PrintTab = CByte(txtPrintTab.Text)
      GLAccount.LineFeeds = CByte(txtLineFeeds.Text)
      GLAccount.BSColumn = CByte(txtBalSheetCol.Text)
      GLAccount.AllStatements = chkAllStmts.Value
      GLAccount.AllSchedules = chkAllSched.Value
      GLAccount.BranchAcct = chkBranchAcct.Value
      GLAccount.ConsAcct = chkConsAcct.Value
      GLAccount.TotalOnLedger = chkTlOnLed.Value
      GLAccount.DollarSign = chkDollarSign.Value
      GLAccount.SignRevStmt = chkSignRevStmt.Value
      GLAccount.SignRevSched = chkSignRevSched.Value
      
      If FormType = Equate.FormAdd Then
         GLAccount.Save (Equate.RecAdd)
      Else
         GLAccount.Save (Equate.RecPut)
      End If
      
   Else      ' delete it
   
      Response = MsgBox("Are you sure you want to DELETE this record?", vbExclamation + vbOKCancel, CStr(GLAccount.Account) & " " & GLAccount.Description)
      Response = GLAccount.DeleteRecord(GLAccount.Account)
      If Response = False Then
         MsgBox "Account NOT deleted ?", vbCritical
      End If
   
   End If
   
   Unload Me
   
End Sub

Private Sub Form_Load()
    
    Dim ll As Byte
    Dim Actt As Byte
 
    ' populate the account type combo box
    ll = 1
    Do Until glTypeChar(ll) = " "
       
       cmbAcctType.AddItem CStr(glTypeChar(ll)) & " " & glTypeName(ll)
       ll = ll + 1
    
       ' if not adding - store type
       If GLAccount.AcctType = glTypeChar(ll) Then
          Actt = ll
       End If
    
    Loop
    
    ' set the form values
    If FormType = Equate.FormAdd Then                      ' set default values
          
          GLAccount.Clear

'          cmbAcctType.Index = 0         ' default to type zero
          cmbAcctType.Text = "0"
          
          txtDescNum = "0"
          txtTotalLevel.Text = "0"
          txtPrintTab.Text = "0"
          txtLineFeeds.Text = "0"
          txtBalSheetCol.Text = "0"
          
          chkAllStmts.Value = 1
          chkAllSched.Value = 1
       
          Me.lblMsg1.Caption = "Adding NEW Account"
       
    Else       ' edit or delete - get values from GLAccount record
    
          txtAcctNumber.Text = CStr(GLAccount.Account)
          
          txtDescription.Text = GLAccount.Description
          
          If GLAccount.DescNumber = 0 Then
             Me.lblDescText = ""
          Else
             If GLDescription.Find(GLAccount.DescNumber) Then
                Me.lblDescText = GLDescription.Description
             Else
                Me.lblDescText = "Not Found !!!"
             End If
          End If
          
          txtDescNum.Text = CStr(GLAccount.DescNumber)
          
          ' account type
          cmbAcctType.Text = GLAccount.AcctType
          
          ' total level up/down limits
          LevelSet
          txtTotalLevel.Text = CStr(GLAccount.TotalLevel)
          
          txtPrintTab.Text = CStr(GLAccount.PrintTab)
          
          txtLineFeeds.Text = CStr(GLAccount.LineFeeds)
          
          txtBalSheetCol.Text = CStr(GLAccount.BSColumn)
          
          If GLAccount.AllStatements = True Then
             chkAllStmts.Value = 1
          Else
             chkAllStmts.Value = 0
          End If
          
          If GLAccount.AllSchedules = True Then
             chkAllSched.Value = 1
          Else
             chkAllSched.Value = 0
          End If
          
          If GLAccount.BranchAcct = True Then
             chkBranchAcct.Value = 1
          Else
             chkBranchAcct.Value = 0
          End If
          
          If GLAccount.ConsAcct = True Then
             chkConsAcct.Value = 1
          Else
             chkConsAcct.Value = 0
          End If
          
          If GLAccount.TotalOnLedger = True Then
             chkTlOnLed.Value = 1
          Else
             chkTlOnLed.Value = 0
          End If
          
          If GLAccount.DollarSign = True Then
             chkDollarSign.Value = 1
          Else
             chkDollarSign.Value = 0
          End If
          
          If GLAccount.SignRevStmt = True Then
             chkSignRevStmt.Value = 1
          Else
             chkSignRevStmt.Value = 0
          End If
          
          If GLAccount.SignRevSched = True Then
             chkSignRevSched.Value = 1
          Else
             chkSignRevSched.Value = 0
          End If
          
          ' if delete - set button captions
          If FormType = Equate.FormDel Then
             cmdOk.Caption = "&Delete"
             lblMsg1.Caption = "Record will be DELETED !!!"
          Else
             lblMsg1.Caption = "Record will be CHANGED"
          End If
       
    End If
    
End Sub

Private Sub LevelSet()
      ' total level - up/down control limits
'      If GLAccount.AcctType = "D" Then
'         updTotalLevel.Min = 1
'         updTotalLevel.Max = 17
'      ElseIf GLAccount.AcctType = "U" Then
'         updTotalLevel.Min = 1
'         updTotalLevel.Max = 3
'      Else
'         updTotalLevel.Min = 0
'         updTotalLevel.Max = 5
'      End If
End Sub

Private Sub txtDescNum_LostFocus()
    
    If Me.txtDescNum <> 0 Then
       If GLDescription.Find(Me.txtDescNum) Then
          Me.lblDescText = GLDescription.Description
       Else
          Me.lblDescText = "Not Found !!!"
       End If
    Else
       Me.lblDescText = ""
    End If
    Me.lblDescText.Refresh

End Sub
