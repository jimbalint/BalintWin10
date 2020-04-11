VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext8.ocx"
Begin VB.Form frmPayeeEdit 
   Caption         =   "1099 Payee Edit"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayeeEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInactive 
      Caption         =   "Inactive"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "&SAVE AND EXIT"
      Height          =   615
      Left            =   4560
      TabIndex        =   9
      Top             =   6000
      Width           =   2415
   End
   Begin VB.CommandButton cmdSaveAdd 
      Caption         =   "SAVE AND &ADD ANOTHER"
      Height          =   615
      Left            =   1440
      TabIndex        =   8
      Top             =   6000
      Width           =   2415
   End
   Begin TDBText6Ctl.TDBText tdbPayeeName 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":0374
      Key             =   "frmPayeeEdit.frx":0392
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
   Begin TDBNumber6Ctl.TDBNumber tdbPayeeNumber 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   661
      Calculator      =   "frmPayeeEdit.frx":03D6
      Caption         =   "frmPayeeEdit.frx":03F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":0464
      Keys            =   "frmPayeeEdit.frx":0482
      Spin            =   "frmPayeeEdit.frx":04CC
      AlignHorizontal =   0
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   7560
      TabIndex        =   10
      Top             =   6000
      Width           =   2415
   End
   Begin TDBText6Ctl.TDBText tdbAddress 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":04F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":0558
      Key             =   "frmPayeeEdit.frx":0576
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
   Begin TDBText6Ctl.TDBText tdbCSZ 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":05BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":062C
      Key             =   "frmPayeeEdit.frx":064A
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
   Begin TDBText6Ctl.TDBText tdbFederalID 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":068E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":06F8
      Key             =   "frmPayeeEdit.frx":0716
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
   Begin TDBText6Ctl.TDBText tdbAccountNumber 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4440
      Width           =   5775
      _Version        =   65536
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":075A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":07C6
      Key             =   "frmPayeeEdit.frx":07E4
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
   Begin TDBText6Ctl.TDBText tdbComment 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   661
      Caption         =   "frmPayeeEdit.frx":0828
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPayeeEdit.frx":088C
      Key             =   "frmPayeeEdit.frx":08AA
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
   Begin VB.Label Label3 
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   6720
      Width           =   375
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
      TabIndex        =   11
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmPayeeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim rs As New ADODB.Recordset

Public ScreenMode As Byte
Public PayeeID As Long
Public PayeeNumber As Long

Private Sub cmdSaveAdd_Click()
    SaveForm
    Payee99.Clear
    ScreenMode = 1
    Form_Load
    tdbPayeeName.SetFocus
End Sub

Private Sub Form_Load()

    ' ScreenMode = 1 - Add
    ' ScreenMode = 2 - Edit

    Me.lblCompanyName = PRCompany.Name
    Me.KeyPreview = True

    ' set field parameters
    With Me
        tdbTextSet .tdbPayeeName, 50
        tdbTextSet .tdbAddress, 50
        tdbTextSet .tdbCSZ, 50
        tdbTextSet .tdbFederalID, 15
        tdbTextSet .tdbAccountNumber, 50
        .tdbComment.text = ""
    End With

    If ScreenMode = 1 Then
    
        SQLString = "SELECT top 1 PayeeNumber FROM Payee99 ORDER BY PayeeNumber DESC"
        rsInit SQLString, cn, rs
        If rs.RecordCount = 0 Then
            PayeeNumber = 101
        Else
            PayeeNumber = rs!PayeeNumber + 1
        End If

        rs.Close
        
        Payee99.OpenRS
        Payee99.Clear
        Payee99.PayeeNumber = PayeeNumber
        Payee99.FederalID = ""
        Payee99.Save (Equate.RecAdd)

        ' get it back
        SQLString = "SELECT * FROM Payee99 WHERE PayeeNumber = " & PayeeNumber
        If Payee99.GetBySQL(SQLString) = False Then
            MsgBox "Payee99 Not Found! " & PayeeNumber
            GoBack
        End If

        PayeeID = Payee99.PayeeID

    End If
    
    SQLString = "SELECT * FROM Payee99 WHERE PayeeID = " & PayeeID
    If Payee99.GetBySQL(SQLString) = False Then
        MsgBox "PayeeID Not Found: " & PayeeID, vbExclamation
        GoBack
    End If
    
    PayeeNumber = Payee99.PayeeNumber
    PayeeID = Payee99.PayeeID
    
    DisplayForm

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
        Case vbKeyF9: cmdSaveAdd_Click
        Case vbKeyF10: cmdSaveExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    
    ' remove if added on this screen
    If ScreenMode = 1 Then
        SQLString = "DELETE * FROM Payee99 WHERE PayeeNumber = " & PayeeNumber
        cn.Execute SQLString
    End If
    
    Unload Me

End Sub

Private Sub cmdSaveExit_Click()

    SaveForm
    Unload Me

End Sub

Private Sub SaveForm()
    
    ' if payeenumber changed .....
    PayeeNumber = Val(Me.tdbPayeeNumber)
    
    Dim rsp As ADODB.Recordset
    SQLString = "select count(1) as pcount" & _
                " from Payee99 " & _
                " where PayeeNumber = " & Me.tdbPayeeNumber & _
                " and PayeeID <> " & Payee99.PayeeID
    rsInit SQLString, cn, rsp
    If CInt(rsp!pcount) <> 0 Then
        MsgBox "This Payee Number already exists: " & Me.tdbPayeeNumber, vbExclamation, "1099 Payee Edit"
        Exit Sub
    End If

    With Me
        Payee99.PayeeNumber = .tdbPayeeNumber
        Payee99.PayeeName = Trim(.tdbPayeeName)
        Payee99.Address = Trim(.tdbAddress)
        Payee99.CSZ = Trim(.tdbCSZ)
        Payee99.AccountNumber = Trim(.tdbAccountNumber)
        
        Payee99.FederalID = Trim(.tdbFederalID)
        
        Payee99.Comment = Trim(.tdbComment)
        If Me.chkInactive Then
            Payee99.Inactive = 1
        Else
            Payee99.Inactive = 0
        End If
        Payee99.Save (Equate.RecPut)
    End With

    frmPayeeList.PayeeID = Payee99.PayeeID

End Sub

Private Sub DisplayForm()
    With Me
        .tdbPayeeNumber = Payee99.PayeeNumber
        .tdbPayeeName = Trim(Payee99.PayeeName)
        .tdbAddress = Payee99.Address
        .tdbCSZ = Payee99.CSZ
        .tdbAccountNumber = Payee99.AccountNumber
        .tdbFederalID = Payee99.FederalID
        .chkInactive = Payee99.Inactive
        .tdbComment = Payee99.Comment
    End With

    Me.Refresh

End Sub

Private Function CheckPayeeNumber() As Boolean

    ' make sure the
    SQLString = "SELECT PayeeID FROM Payee99 WHERE PayeeNumber = " & PayeeNumber
    rsInit SQLString, cn, rs
    If rs!PayeeID <> PayeeID Then
        MsgBox "This Payee Number already taken: " & PayeeNumber, vbExclamation
        CheckPayeeNumber = False
    Else
        CheckPayeeNumber = True
    End If
    
    rs.Close
    
End Function

