VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPrint99 
   Caption         =   "Print 1099 Form"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint99.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCSV 
      Caption         =   "CSV Export"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotals 
      Caption         =   "&TOTALS"
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   9360
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   9360
      Width           =   1455
   End
   Begin TDBNumber6Ctl.TDBNumber tdbHorz 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   9480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      Calculator      =   "frmPrint99.frx":030A
      Caption         =   "frmPrint99.frx":032A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrint99.frx":0394
      Keys            =   "frmPrint99.frx":03B2
      Spin            =   "frmPrint99.frx":03FC
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   9360
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   6855
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   12855
      _cx             =   22675
      _cy             =   12091
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.CommandButton Command1 
         Caption         =   "&PRINT"
         Height          =   615
         Left            =   7320
         TabIndex        =   11
         Top             =   9120
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox cmbForm 
      Height          =   360
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox cmbTaxYear 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      Top             =   9360
      Width           =   1335
   End
   Begin TDBNumber6Ctl.TDBNumber tdbVertical 
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   9480
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      Calculator      =   "frmPrint99.frx":0424
      Caption         =   "frmPrint99.frx":0444
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmPrint99.frx":04AE
      Keys            =   "frmPrint99.frx":04CC
      Spin            =   "frmPrint99.frx":0516
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
   Begin VB.Label Label3 
      Caption         =   "01/06/2024"
      Height          =   255
      Left            =   11040
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "1099 Form:"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Tax Year:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   975
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
      TabIndex        =   1
      Top             =   360
      Width           =   13335
   End
End
Attribute VB_Name = "frmPrint99"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I, J, K As Long
Dim X, Y, Z As String
Dim boo As Boolean
Dim FormID, RowCount, Rw As Long
Dim PRGlobalID As Long
Dim sOut As String
Dim CommonColumns As String

Private Sub cmdCSV_Click()
    
    ' get payer settings
    SQLString = "select *" & _
                " from PRGlobal " & _
                " where UserID = " & User.ID & _
                " and TypeCode = 30"
    If Not PRGlobal.GetBySQL(SQLString) Then
        MsgBox "Payer data not found!!!", vbExclamation
        GoBack
    End If
    
    Const WindowsFolder = 0
    Const SystemFolder = 1
    Const TemporaryFolder = 2
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim tempFolder: tempFolder = fso.GetSpecialFolder(TemporaryFolder)

    J = 0
    If Me.cmbForm.text = "1099-NEC" Then J = 1
    If Me.cmbForm.text = "1099-MISC" Then J = 1
    If Me.cmbForm.text = "1099-INT" Then J = 1
    If Me.cmbForm.text = "1099-DIV" Then J = 1
    If J = 0 Then
        MsgBox "Form " & Me.cmbForm.text & " not supported for export yet"
        Exit Sub
    End If

    Dim TextChannel As Integer
    
' tempFolder = "\\VBOXSVR\VM-Share\Balint_NewADO_EXE"
    TextFileName = tempFolder & "\" & Me.cmbForm.text & ".csv"
    
    TextChannel = FreeFile
    Do
        On Error Resume Next
        Open TextFileName For Output As #TextChannel
        If Err.Number <> 0 Then
            ErrMsg = "Error Opening: " & TextFileName & vbCr & vbCr & _
                " " & Err.Number & " " & Err.Description
            MsgResponse = MsgBox(ErrMsg, vbRetryCancel + vbExclamation, "File Open Error")
            If MsgResponse <> vbRetry Then
                TextChannel = 0
                TextFileName = ""
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop

    ' header line
    sOut = CommonHeader()
    Select Case Me.cmbForm.text
        Case "1099-NEC": sOut = sOut & NECHeader()
        Case "1099-MISC": sOut = sOut & MiscHeader()
        Case "1099-INT": sOut = sOut & IntHeader()
        Case "1099-DIV": sOut = sOut & DivHeader()
    End Select
    Print #TextChannel, sOut
    
    CommonColumns = SetCommonColumns()
    
    SQLString = " SELECT * FROM Payee99 ORDER BY PayeeName"
    If Payee99.GetBySQL(SQLString) = False Then
        MsgBox "No Payee info found!", vbInformation
        GoBack
    End If
    Do
        sOut = CommonColumns & PayeeColumns()
        Select Case Me.cmbForm.text
            Case "1099-NEC": sOut = sOut & NEC_Columns(Payee99.PayeeID)
            Case "1099-MISC": sOut = sOut & MiscColumns(Payee99.PayeeID)
            Case "1099-INT": sOut = sOut & Int_Columns(Payee99.PayeeID)
            Case "1099-DIV": sOut = sOut & Div_Columns(Payee99.PayeeID)
        End Select
        Print #TextChannel, sOut
        If Payee99.GetNext = False Then Exit Do
    Loop
    
    Close #TextChannel
    TaskID = Shell("cmd /c " & TextFileName, vbNormalFocus)
    
End Sub
Private Function SetCommonColumns()
    Dim aa As Integer
    Dim ary(20) As String
    ary(1) = Me.cmbForm.text
    ary(2) = Me.cmbTaxYear.text
    ary(3) = "FedID"      ' Payer TIN Type
    ary(4) = PrepCSV(GLCompany.FederalID)
    ary(5) = "B"         ' PayerNameType
    ary(6) = PrepCSV(GLCompany.Name)
    ary(7) = ""       ' payer name line 2
    ary(8) = ""       ' payer first name
    ary(9) = ""       ' payer MI
    ary(10) = ""      ' payer last name
    ary(11) = ""      ' payer suffix
    ary(12) = "US"      ' payer country
    ary(13) = PrepCSV(GLCompany.Address1)
    ary(14) = PrepCSV(GLCompany.Address2)
    ary(15) = PrepCSV(GLCompany.City)
    ary(16) = PrepCSV(GLCompany.State)
    ary(17) = GLCompany.ZipCode
    ary(18) = "D"      ' phone type
    ary(19) = PrepCSV(PRGlobal.Var3)       ' phone
    ary(20) = PrepCSV(PRGlobal.Var2)      ' email
    SetCommonColumns = Ary2String(ary, False)
End Function
Private Function Get_TIN_Type(ByVal IDString As String) As String
    IDString = Trim(IDString)
    If Len(IDString) - Len(Replace(IDString, "-", "")) = 1 Then
        Get_TIN_Type = "EIN"
    Else
        Get_TIN_Type = "SSN"
    End If
End Function
Private Function ParseCSZ(ByVal CSZString) As String()
    
    Dim CommaPos As Integer
    Dim LastSpacePos As Integer
    
    ' city, state zip
    CSZString = Trim(CSZString)
    Dim arrcsz(3) As String
    
    CommaPos = InStr(1, CSZString, ",", vbTextCompare)
    LastSpacePos = InStrRev(CSZString, " ")
    If CSZString = "" Then
        arrcsz(1) = ""
        arrcsz(2) = ""
        arrcsz(3) = ""
    ElseIf CommaPos = 0 Or LastSpacePos = 0 Or LastSpacePos < CommaPos Then
        arrcsz(1) = CSZString
        arrcsz(2) = ""
        arrcsz(3) = ""
    Else
        arrcsz(1) = Trim(Mid(CSZString, 1, CommaPos - 1))
        arrcsz(2) = Trim(Mid(CSZString, CommaPos + 1, LastSpacePos - CommaPos))
        arrcsz(3) = Trim(Mid(CSZString, LastSpacePos + 1))
    End If
    
    ParseCSZ = arrcsz
End Function
Private Function PayeeColumns() As String
    Dim aa As Integer
    Dim ary(17) As String
    Dim CSZ As Variant
    CSZ = ParseCSZ(Payee99.CSZ)
    ary(1) = Get_TIN_Type(Payee99.FederalID)     ' TIN Type
    ary(2) = PrepCSV(Payee99.FederalID)
    ary(3) = "B"              ' name type
    ary(4) = PrepCSV(Payee99.PayeeName)
    ary(5) = ""       ' biz name 2
    ary(6) = ""       ' firstname
    ary(7) = ""       ' mnm
    ary(8) = ""       ' last name
    ary(9) = ""       ' suffix
    ary(10) = "US"      ' country
    ary(11) = PrepCSV(Payee99.Address)
    ary(12) = ""      ' addr2
    ary(13) = CSZ(1)
    ary(14) = CSZ(2)      ' state
    ary(15) = CSZ(3)      ' zip
    ary(16) = ""      ' office code
    ary(17) = Payee99.AccountNumber
    PayeeColumns = Ary2String(ary, False)
End Function

Private Function NEC_Columns(ByVal PayeeID As Integer) As String
    Dim aa As Integer
    Dim ary(17) As String
    ary(1) = "N"       ' 2nd TIN notice
    
    Dim box1 As String
    Dim box2 As String
    box1 = GetDetailData99(PayeeID, "1")
    box2 = GetDetailData99(PayeeID, "2")
    If box2 = "" Then
        ary(2) = GetDetailData99(PayeeID, "1")
        ary(3) = "N"
    Else
        ary(2) = ""
        ary(3) = "Y"
    End If
    
    ary(4) = AmtString(GetDetailData99(PayeeID, "4"))
    ary(4) = GetDetailData99(PayeeID, "4")
    
    ary(5) = ""   ' combined federal/state filing
    
    If CDec(GetDetailData99(PayeeID, "7a")) <> 0 Then
        ary(6) = ""   ' state 1
        ary(7) = GetDetailData99(PayeeID, "5a")   ' state WH
        ary(8) = GetDetailData99(PayeeID, "6a")   ' state #
        ary(9) = GetDetailData99(PayeeID, "7a")   ' state income
        ary(10) = ""                          ' local inc tax
        ary(11) = ""                          ' special data entries
    Else
        ary(6) = ""   ' state 1
        ary(7) = ""   ' state WH
        ary(8) = ""   ' state #
        ary(9) = ""   ' state income
        ary(10) = ""                          ' local inc tax
        ary(11) = ""                          ' special data entries
    End If
    
    If CDec(GetDetailData99(PayeeID, "7b")) <> 0 Then
        ary(12) = ""   ' state 2
        ary(13) = GetDetailData99(PayeeID, "5b")   ' state WH
        ary(14) = GetDetailData99(PayeeID, "6b")   ' state #
        ary(15) = GetDetailData99(PayeeID, "7b")   ' state income
        ary(16) = ""                          ' local inc tax
        ary(17) = ""                          ' special data entries
    Else
        ary(12) = ""   ' state 2
        ary(13) = ""   ' state WH
        ary(14) = ""   ' state #
        ary(15) = ""   ' state income
        ary(16) = ""                          ' local inc tax
        ary(17) = ""                          ' special data entries
    End If
    
    NEC_Columns = Ary2String(ary, True)
End Function
Private Function MiscColumns(ByVal PayeeID As Integer) As String
    Dim aa As Integer
    Dim ary(29) As String
    
    ary(1) = ""     ' FACTA filing req
    ary(2) = "N"     ' 2nd TIN notice
    
    ary(3) = GetDetailData99(PayeeID, 1)
    ary(4) = GetDetailData99(PayeeID, 2)
    ary(5) = GetDetailData99(PayeeID, 3)
    ary(6) = GetDetailData99(PayeeID, 4)
    ary(7) = GetDetailData99(PayeeID, 5)
    ary(8) = GetDetailData99(PayeeID, 6)
    ary(9) = GetDetailData99(PayeeID, 7)
    ary(10) = GetDetailData99(PayeeID, 8)
    ary(11) = GetDetailData99(PayeeID, 9)
    ary(12) = GetDetailData99(PayeeID, 10)
    ary(13) = ""    ' fish ...
    ary(14) = GetDetailData99(PayeeID, 12)
    ary(15) = GetDetailData99(PayeeID, 14)
    ary(16) = ""    ' nq def comp
    ary(17) = ""    ' combined fed/state filing
    
    ' state income boxes not defined??
    ary(18) = ""   ' state 1
    ary(19) = ""   ' state WH
    ary(20) = ""   ' state #
    ary(21) = ""   ' state income
    ary(22) = ""                          ' local inc tax
    ary(23) = ""                          ' special data entries
    ary(24) = ""   ' state 1
    ary(25) = ""   ' state WH
    ary(26) = ""   ' state #
    ary(27) = ""   ' state income
    ary(28) = ""                          ' local inc tax
    ary(29) = ""                          ' special data entries
    
'    If CDec(GetDetailData99(PayeeID, "18a")) <> 0 Then
'        ary(18) = ""   ' state 1
'        ary(19) = GetDetailData99(PayeeID, "16a")   ' state WH
'        ary(20) = GetDetailData99(PayeeID, "17a")   ' state #
'        ary(21) = GetDetailData99(PayeeID, "18a")   ' state income
'        ary(22) = ""                          ' local inc tax
'        ary(23) = ""                          ' special data entries
'    Else
'        ary(18) = ""   ' state 1
'        ary(19) = ""   ' state WH
'        ary(20) = ""   ' state #
'        ary(21) = ""   ' state income
'        ary(22) = ""                          ' local inc tax
'        ary(23) = ""                          ' special data entries
'    End If
'
'    If CDec(GetDetailData99(PayeeID, "18b")) <> 0 Then
'        ary(24) = ""   ' state 1
'        ary(25) = GetDetailData99(PayeeID, "16b")   ' state WH
'        ary(26) = GetDetailData99(PayeeID, "17b")   ' state #
'        ary(27) = GetDetailData99(PayeeID, "18b")   ' state income
'        ary(28) = ""                          ' local inc tax
'        ary(29) = ""                          ' special data entries
'    Else
'        ary(24) = ""   ' state 1
'        ary(25) = ""   ' state WH
'        ary(26) = ""   ' state #
'        ary(27) = ""   ' state income
'        ary(28) = ""                          ' local inc tax
'        ary(29) = ""                          ' special data entries
'    End If
    
    MiscColumns = Ary2String(ary, True)
End Function
Private Function Int_Columns(ByVal PayeeID As Integer) As String
    Dim aa As Integer
    Dim ary(30) As String
    
    ary(1) = ""         ' FACTA
    ary(2) = "N"       ' 2nd TIN notice
    ary(3) = GetDetailData99(PayeeID, "RTN")
    For I = 1 To 14
        If I <> 14 Then
            ary(I + 3) = GetDetailData99(PayeeID, I)
        Else
            ary(I + 3) = ""
        End If
    Next I
    ary(18) = ""   ' combined federal/state filing
    
    If CDec(GetDetailData99(PayeeID, "17a")) <> 0 Then
        ary(19) = GetDetailData99(PayeeID, "15a")   ' state 1
        ary(20) = GetDetailData99(PayeeID, "17a")   ' state WH
        ary(21) = GetDetailData99(PayeeID, "16a")   ' state #
        ary(22) = "" ' state income not on the form???
        ary(23) = "" ' local inc tax not on the form???
        ary(24) = "" ' special data entries
    Else
        ary(19) = ""   ' state 1
        ary(20) = ""   ' state WH
        ary(21) = ""   ' state #
        ary(22) = "" ' state income not on the form???
        ary(23) = "" ' local inc tax not on the form???
        ary(24) = "" ' special data entries
    End If
    
    If CDec(GetDetailData99(PayeeID, "17b")) <> 0 Then
        ary(25) = GetDetailData99(PayeeID, "15b")   ' state 1
        ary(26) = GetDetailData99(PayeeID, "17b")   ' state WH
        ary(27) = GetDetailData99(PayeeID, "16b")   ' state #
        ary(28) = "" ' state income not on the form???
        ary(29) = "" ' local inc tax not on the form???
        ary(30) = "" ' special data entries
    Else
        ary(25) = ""   ' state 1
        ary(26) = ""   ' state WH
        ary(27) = ""   ' state #
        ary(28) = "" ' state income not on the form???
        ary(29) = "" ' local inc tax not on the form???
        ary(30) = "" ' special data entries
    End If
    Int_Columns = Ary2String(ary, True)
End Function

Private Function Div_Columns(ByVal PayeeID As Integer) As String
    Dim aa As Integer
    Dim ary(33) As String
    
    ary(1) = ""         ' FACTA
    ary(2) = "N"       ' 2nd TIN notice
    ary(3) = GetDetailData99(PayeeID, "1a")
    
    ary(4) = GetDetailData99(PayeeID, "1b")
    ary(5) = GetDetailData99(PayeeID, "2a")
    ary(6) = GetDetailData99(PayeeID, "2b")
    ary(7) = GetDetailData99(PayeeID, "2c")
    ary(8) = GetDetailData99(PayeeID, "2d")
    ary(9) = GetDetailData99(PayeeID, "2e")
    ary(10) = GetDetailData99(PayeeID, "2f")
    
    ary(11) = GetDetailData99(PayeeID, "3")
    ary(12) = GetDetailData99(PayeeID, "4")
    ary(13) = GetDetailData99(PayeeID, "5")
    ary(14) = GetDetailData99(PayeeID, "6")
    ary(15) = GetDetailData99(PayeeID, "7")
    ary(16) = GetDetailData99(PayeeID, "8")
    ary(17) = GetDetailData99(PayeeID, "9")
    ary(18) = GetDetailData99(PayeeID, "10")
    ary(19) = GetDetailData99(PayeeID, "12")
    ary(20) = GetDetailData99(PayeeID, "13")
    
    ary(21) = ""   ' combined federal/state filing
    
    If CDec(GetDetailData99(PayeeID, "14a")) <> 0 Then
        ary(22) = GetDetailData99(PayeeID, "12a")   ' state 1
        
        ary(23) = GetDetailData99(PayeeID, "14a")   ' state WH
        ary(24) = GetDetailData99(PayeeID, "13a")   ' state #
        ary(25) = "" ' state income not on the form???
        ary(26) = "" ' local inc tax not on the form???
        ary(27) = "" ' special data entries
    Else
        ary(22) = ""   ' state 1
        ary(23) = ""   ' state WH
        ary(24) = ""   ' state #
        ary(25) = "" ' state income not on the form???
        ary(26) = "" ' local inc tax not on the form???
        ary(27) = "" ' special data entries
    End If
    
    If CDec(GetDetailData99(PayeeID, "14b")) <> 0 Then
        ary(28) = GetDetailData99(PayeeID, "12b")   ' state 1
        ary(29) = GetDetailData99(PayeeID, "14b")   ' state WH
        ary(30) = GetDetailData99(PayeeID, "13b")   ' state #
        ary(31) = "" ' state income not on the form???
        ary(32) = "" ' local inc tax not on the form???
        ary(33) = "" ' special data entries
    Else
        ary(28) = ""   ' state 1
        ary(29) = ""   ' state WH
        ary(30) = ""   ' state #
        ary(31) = "" ' state income not on the form???
        ary(32) = "" ' local inc tax not on the form???
        ary(33) = "" ' special data entries
    End If
    
    Div_Columns = Ary2String(ary, True)
End Function

Private Function Ary2String(ByVal ary As Variant, ByVal LastFlag As Boolean) As String
    Ary2String = ""
    For I = 1 To UBound(ary)
        Ary2String = Ary2String & Chr(34) & PrepCSV(ary(I)) & Chr(34)
        If I <> UBound(ary) Or (I = UBound(ary) And LastFlag = False) Then
            Ary2String = Ary2String & Chr(44)
        End If
    Next I
End Function
Private Function AmtString(ByVal strAmt) As String
    If IsNumeric(strAmt) Then
        AmtString = strAmt
    Else
        AmtString = "0.00"
    End If
End Function

Private Function NECHeader() As String
    Dim aa As Integer
    Dim ary(17) As String
    ary(1) = "2nd TIN Notice"
    ary(2) = "Box 1 - Nonemployee Compensation"
    ary(3) = "Box 2 - Payer made direct sales totaling $5000 or more of consumer products to a recipient for resale"
    ary(4) = "Box 4 - Federal income tax withheld"
    ary(5) = "Combined Federal/State Filing"
    ary(6) = "State 1"
    ary(7) = "State 1 - State Tax Withheld"
    ary(8) = "State 1 - State/Payer state number"
    ary(9) = "State 1 - State income"
    ary(10) = "State 1 - Local income tax withheld"
    ary(11) = "State 1 - Special Data Entries"
    ary(12) = "State 2"
    ary(13) = "State 2 - State Tax Withheld"
    ary(14) = "State 2 - State/Payer state number"
    ary(15) = "State 2 - State income"
    ary(16) = "State 2 - Local income tax withheld"
    ary(17) = "State 2 - Special Data Entries"
    NECHeader = Ary2String(ary, True)
End Function

Private Function MiscHeader() As String
    Dim aa As Integer
    Dim ary(29) As String
    ary(1) = "FATCA Filing Requirements"
    ary(2) = "2nd TIN Notice"
    ary(3) = "Box 1 - Rents"
    ary(4) = "Box 2 - Royalties"
    ary(5) = "Box 3 - Other Income"
    ary(6) = "Box 4 - Federal income tax withheld"
    ary(7) = "Box 5 - Fishing boat proceeds"
    ary(8) = "Box 6 - Medical and health care payments"
    ary(9) = "Box 7 - Direct sales of $5000 or more of consumer products to a recipient for resale"
    ary(10) = "Box 8 - Subtitute payments in lieu of dividends or interest"
    ary(11) = "Box 9 - Crop insurance proceeds"
    ary(12) = "Box 10 - Gross proceeds paid to an attorney"
    ary(13) = "Box 11 - Fish purchased for resale"
    ary(14) = "Box 12 - Section 409A deferrals"
    ary(15) = "Box 14 - Excess golden parachute payments"
    ary(16) = "Box 15 - Nonqualified deferred compensation"
    ary(17) = "Combined Federal/State Filing"
    ary(18) = "State 1"
    ary(19) = "State 1 - State Tax Withheld"
    ary(20) = "State 1 - State/Payer state number"
    ary(21) = "State 1 - State income"
    ary(22) = "State 1 - Local income tax withheld"
    ary(23) = "State 1 - Special Data Entries"
    ary(24) = "State 2"
    ary(25) = "State 2 - State Tax Withheld"
    ary(26) = "State 2 - State/Payer state number"
    ary(27) = "State 2 - State income"
    ary(28) = "State 2 - Local income tax withheld"
    ary(29) = "State 2 - Special Data Entries"
    MiscHeader = Ary2String(ary, True)
End Function

Private Function IntHeader() As String
    Dim aa As Integer
    Dim ary(30) As String
    ary(1) = "FATCA Filing Requirements"
    ary(2) = "2nd TIN Notice"
    ary(3) = "Payer's RTN"
    ary(4) = "Box 1 - Interest Income"
    ary(5) = "Box 2 - Early withdrawal penalty"
    ary(6) = "Box 3 - Interest on U.S. Savings Bonds and Treasury obligations"
    ary(7) = "Box 4 - Federal income tax withheld"
    ary(8) = "Box 5 - Investment expenses"
    ary(9) = "Box 6 - Foreign tax paid"
    ary(10) = "Box 7 - Foreign country or U.S. possession"
    ary(11) = "Box 8 - Tax-exempt interest"
    ary(12) = "Box 9 - Specified private activity bond interest"
    ary(13) = "Box 10 - Market discount"
    ary(14) = "Box 11 - Bond premium"
    ary(15) = "Box 12 - Bond premium on Treasury obligations"
    ary(16) = "Box 13 - Bond premium on tax-exempt bond"
    ary(17) = "Box 14 - Tax-exempt and tax credit bond CUSIP no."
    ary(18) = "Combined Federal/State Filing"
    ary(19) = "State 1"
    ary(20) = "State 1 - State Tax Withheld"
    ary(21) = "State 1 - State/Payer state number"
    ary(22) = "State 1 - State income"
    ary(23) = "State 1 - Local income tax withheld"
    ary(24) = "State 1 - Special Data Entries"
    ary(25) = "State 2"
    ary(26) = "State 2 - State Tax Withheld"
    ary(27) = "State 2 - State/Payer state number"
    ary(28) = "State 2 - State income"
    ary(29) = "State 2 - Local income tax withheld"
    ary(30) = "State 2 - Special Data Entries"
    IntHeader = Ary2String(ary, True)
End Function
Private Function DivHeader() As String
    Dim aa As Integer
    Dim ary(33) As String
    ary(1) = "FATCA Filing Requirements"
    ary(2) = "2nd TIN Notice"
    ary(3) = "Box 1a - Total ordinary dividends"
    ary(4) = "Box 1b - Qualified dividends"
    ary(5) = "Box 2a - Total capital gain distribution"
    ary(6) = "Box 2b - Unrecap. Sec. 1250 gain"
    ary(7) = "Box 2c - Section 1202 gain"
    ary(8) = "Box 2d - Collectibles (28%) gain"
    ary(9) = "Box 2e - Section 897 ordinary dividends"
    ary(10) = "Box 2f - Section 897 capital gain"
    ary(11) = "Box 3 - Nondividend distributions"
    ary(12) = "Box 4 - Federal income tax withheld"
    ary(13) = "Box 5 - Section 199A dividends"
    ary(14) = "Box 6 - Investment expenses"
    ary(15) = "Box 7 - Foreign tax paid"
    ary(16) = "Box 8 - Foreign country or U.S. possession"
    ary(17) = "Box 9 - Cash liquidation distributions"
    ary(18) = "Box 10 - Noncash liquidation distributions"
    ary(19) = "Box 12 - Exempt-interest dividends"
    ary(10) = "Box 13 - Specified private activity bond interest dividends"
    ary(21) = "Combined Federal/State Filing"
    ary(22) = "State 1"
    ary(23) = "State 1 - State Tax Withheld"
    ary(24) = "State 1 - State/Payer state number"
    ary(25) = "State 1 - State income"
    ary(26) = "State 1 - Local income tax withheld"
    ary(27) = "State 1 - Special Data Entries"
    ary(28) = "State 2"
    ary(29) = "State 2 - State Tax Withheld"
    ary(30) = "State 2 - State/Payer state number"
    ary(31) = "State 2 - State income"
    ary(32) = "State 2 - Local income tax withheld"
    ary(33) = "State 2 - Special Data Entries"
    DivHeader = Ary2String(ary, True)
End Function

Private Function CommonHeader() As String
    Dim aa As Integer
    Dim ary(37) As String
    ary(1) = "Form Type"
    ary(2) = "Tax Year"
    ary(3) = "Payer TIN Type"
    ary(4) = "Payer Taxpayer ID Number"
    ary(5) = "Payer Name Type"
    ary(6) = "Payer Business or Entity Name Line 1"
    ary(7) = "Payer Business or Entity Name Line 2"
    ary(8) = "Payer First Name"
    ary(9) = "Payer Middle Name"
    ary(10) = "Payer Last Name (Surname)"
    ary(11) = "Payer Suffix"
    ary(12) = "Payer Country"
    ary(13) = "Payer Address Line 1"
    ary(14) = "Payer Address Line 2"
    ary(15) = "Payer City/Town"
    ary(16) = "Payer State/Province/Territory"
    ary(17) = "Payer ZIP/Postal Code"
    ary(18) = "Payer Phone Type"
    ary(19) = "Payer Phone"
    ary(20) = "Payer Email Address"
    ary(21) = "Recipient TIN Type"
    ary(22) = "Recipient Taxpayer ID Number"
    ary(23) = "Recipient Name Type"
    ary(24) = "Recipient Business or Entity Name Line 1"
    ary(25) = "Recipient Business or Entity Name Line 2"
    ary(26) = "Recipient First Name"
    ary(27) = "Recipient Middle Name"
    ary(28) = "Recipient Last Name (Surname)"
    ary(29) = "Recipient Suffix"
    ary(30) = "Recipient Country"
    ary(31) = "Recipient Address Line 1"
    ary(32) = "Recipient Address Line 2"
    ary(33) = "Recipient City/Town"
    ary(34) = "Recipient State/Province/Territory"
    ary(35) = "Recipient ZIP/Postal Code"
    ary(36) = "Office Code"
    ary(37) = "Form Account Number"
    CommonHeader = Ary2String(ary, False)
End Function

Private Function PrepCSV(ByVal InString As String) As String
    InString = Trim(InString)
    InString = Replace(InString, ",", " ")
    InString = Replace(InString, """", " ")
    PrepCSV = InString
End Function


Private Sub cmdTotals_Click()
    cmdSave_Click
    frmTotals.TaxYear = Me.cmbTaxYear
    frmTotals.FormType = GetFormType()
    frmTotals.Show vbModal
End Sub

Private Sub Form_Load()

    Me.lblCompanyName = GLCompany.Name
    
    Me.KeyPreview = True

    With Me
        
        .cmbForm.AddItem "1099-NEC"
        .cmbForm.AddItem "1099-MISC"
        .cmbForm.AddItem "1099-R"
        .cmbForm.AddItem "1099-INT"
        .cmbForm.AddItem "1099-DIV"
        .cmbForm.ListIndex = 0
    
        PopTaxYear .cmbTaxYear
    
    End With

    ' If LCase(User.Logon) <> "jim" Then Me.cmdExcelTest.Visible = False
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With Me.fg
        If Col = 1 Then     ' set select off? - delete all detail99 records for the payee
            If .TextMatrix(Row, 1) = False Then
                SQLString = " DELETE * FROM Detail99 WHERE PayeeID = " & .TextMatrix(Row, 0) & _
                            " AND TaxYear = " & Me.cmbTaxYear.text & _
                            " AND FormType = '" & GetFormType() & "'"
                cn.Execute SQLString
                For I = 5 To .Cols - 1
                    .TextMatrix(Row, I) = ""
                Next I
            End If
        End If
        .TextMatrix(Row, 1) = SetSelect(Row)
    End With
End Sub

Private Sub cmdPrint_Click()
    
    cmdSave_Click
    SaveNudge
    
    HorzNudge = Me.tdbHorz.Value
    VertNudge = Me.tdbVertical.Value
    
    With Me
        X = Mid(.cmbForm.text, 6)
        I = Me.cmbTaxYear.text
        PrintForm99 X, I, False
    End With

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 And Col <= 4 Then
        Cancel = True
    End If
End Sub

Private Sub cmdLoad_Click()

Dim ColCt As Integer

    ' free mode grid
    With Me.fg
        
        ' grid paramters
        .DataMode = flexDMFree
        .FixedCols = 5
        .FixedRows = 1
        .Rows = 1
        .BackColorAlternate = RGB(195, 195, 195)
        .ExplorerBar = flexExMove + flexExSort
        .Editable = flexEDKbdMouse

        ' get the form
        SQLString = " SELECT * FROM Form99 WHERE TaxYear = " & Me.cmbTaxYear.text & " " & _
                    " AND FormType = '" & GetFormType() & "'"
        If Form99.GetBySQL(SQLString) = False Then
            MsgBox "Form NF: " & Me.cmbTaxYear.text & " 1099-" & Me.cmbForm.text, vbExclamation
            GoBack
        End If
        
        FormID = Form99.FormID
        
        ' add the initial columns
        ' same for all 1099 forms
        
        .TextMatrix(0, 0) = "PayeeID"
        .ColData(0) = "PayeeID"
        .ColHidden(0) = True
        .ColWidth(0) = 1500
        
        .TextMatrix(0, 1) = "Select"
        .ColData(1) = "Select"
        .ColDataType(1) = flexDTBoolean
        .ColWidth(1) = 750
        
        .TextMatrix(0, 2) = "Payee #"
        .ColData(2) = "PayeeNumber"
        .ColDataType(2) = flexDTDouble
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Payee Name"
        .ColData(3) = "PayeeName"
        .ColDataType(3) = flexDTString
        .ColWidth(3) = 2000
        
        .TextMatrix(0, 4) = "Payee Fed ID"
        .ColData(4) = "FederalID"
        .ColDataType(4) = flexDTString
        .ColWidth(4) = 1500
        
        ColCt = 4
        
        ' get the fields of the form
        ' put to header line of grid
        ' coldata is boxname
        SQLString = " SELECT * FROM Field99 WHERE TaxYear = " & Me.cmbTaxYear.text & _
                    " AND FormType = '" & GetFormType() & "'" & _
                    " AND QuickEntry > 0 ORDER BY QuickEntry"
        If Field99.GetBySQL(SQLString) = False Then
            MsgBox "Fields NF: " & Me.cmbTaxYear.text & " 1099-" & Me.cmbForm.text, vbExclamation
            GoBack
        End If
        
        Do
            
            ColCt = ColCt + 1
            .Cols = ColCt + 1
            
            .TextMatrix(0, ColCt) = Field99.BTitle
            .ColData(ColCt) = Field99.BoxName
            
            If Field99.FieldFormat = Equate.fmtAmount Then
                .ColDataType(ColCt) = flexDTCurrency
                .ColFormat(ColCt) = "Currency"
                .ColWidth(ColCt) = 1300
            ElseIf Field99.FieldFormat = Equate.fmtString Then
                .ColDataType(ColCt) = flexDTString
                .ColWidth(ColCt) = 2000
            End If
            
            If Field99.GetNext = False Then Exit Do
        
        Loop

        
        ' load the payee data
        SQLString = " SELECT * FROM Payee99 ORDER BY PayeeNumber"
        If Payee99.GetBySQL(SQLString) = False Then
            MsgBox "No Payee info found!", vbExclamation
            GoBack
        End If
        
        Rw = 0

        Do
            With Me.fg
                
                ' inactive w/ no detail data - skip it
                If Payee99.Inactive = 1 Then
                    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & Payee99.PayeeID & _
                                " AND FormType = '" & GetFormType() & "'" & _
                                " AND TaxYear = " & Me.cmbTaxYear.text
                    If Detail99.GetBySQL(SQLString) = False Then
                        GoTo NextPayee
                    End If
                End If
                
                Rw = Rw + 1
                .Rows = Rw + 1
                
                ' load the info from Payee99
                .TextMatrix(Rw, 0) = Payee99.PayeeID
                .TextMatrix(Rw, 1) = False
                .TextMatrix(Rw, 2) = Payee99.PayeeNumber
                .TextMatrix(Rw, 3) = Payee99.PayeeName
                .TextMatrix(Rw, 4) = Payee99.FederalID
                
                ' load the detail data
                SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & Payee99.PayeeID & _
                            " AND FormType = '" & GetFormType() & "'" & _
                            " AND TaxYear = " & Me.cmbTaxYear.text
                If Detail99.GetBySQL(SQLString) = True Then

                    Do
                        For J = 5 To .Cols - 1
                            If .ColData(J) = Detail99.BoxName Then
                                If .ColFormat(J) = "Currency" Then
                                    .TextMatrix(Rw, J) = Format(Detail99.FieldValue, "Currency")
                                Else
                                    .TextMatrix(Rw, J) = Detail99.FieldValue
                                End If
                            End If
                        Next J
                        If Detail99.GetNext = False Then Exit Do
                    Loop
                End If
                
                .TextMatrix(Rw, 1) = SetSelect(Rw)
        
NextPayee:
                If Payee99.GetNext = False Then Exit Do
            
            End With
        
        Loop
        
        ' 2022-01-15 causing issue for new clients???
        ' .AutoSize 0, .Cols - 1
        
        .TabBehavior = flexTabCells
    
    End With
    
    ' load nudge
    PRGlobalID = 0
    With Me
        SQLString = " SELECT * FROM PRGlobal WHERE UserID = " & User.ID & _
                    " AND Description = '" & Me.cmbForm.text & "'"
        If PRGlobal.GetBySQL(SQLString) = False Then
            .tdbHorz.Value = 0
            .tdbVertical.Value = 0
        Else
            PRGlobalID = PRGlobal.GlobalID
            .tdbHorz.Value = PRGlobal.Var1
            .tdbVertical.Value = PRGlobal.Var2
        End If
    End With

End Sub

Private Sub SaveNudge()

    If PRGlobalID = 0 Then
        PRGlobal.Clear
        PRGlobal.UserID = User.ID
        PRGlobal.Description = Me.cmbForm.text
        PRGlobal.Save (Equate.RecAdd)
    End If
    
    PRGlobal.Var1 = Me.tdbHorz.Value
    PRGlobal.Var2 = Me.tdbVertical.Value
    PRGlobal.Save (Equate.RecPut)

End Sub

Private Sub cmdSave_Click()

    With Me.fg
    
        For Rw = 1 To .Rows - 1
    
            ' reset the FieldValue for all for payee/form
            SQLString = " UPDATE Detail99 SET FieldValue = 'JimBo' WHERE " & _
                        " PayeeID = " & .TextMatrix(Rw, 0) & _
                        " AND FormType = '" & GetFormType() & "'" & _
                        " AND TaxYear = " & Me.cmbTaxYear.text
            cn.Execute SQLString
    
            For J = 5 To .Cols - 1
                
                X = Trim(.TextMatrix(Rw, J))
                If X <> "" Then
                    ' see if the detail record already exists
                    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & .TextMatrix(Rw, 0) & _
                                " AND FormType = '" & GetFormType() & "'" & _
                                " AND TaxYear = " & Me.cmbTaxYear.text & _
                                " AND BoxName = '" & .ColData(J) & "'"
                                
                    If Detail99.GetBySQL(SQLString) = True Then
                        SQLString = " UPDATE Detail99 SET FieldValue = '" & X & "' " & _
                                    " WHERE PayeeID = " & .TextMatrix(Rw, 0) & _
                                    " AND FormType = '" & GetFormType() & "'" & _
                                    " AND TaxYear = " & Me.cmbTaxYear.text & _
                                    " AND BoxName = '" & .ColData(J) & "'"
                        cn.Execute SQLString
                    Else
                        Detail99.Clear
                        Detail99.PayeeID = .TextMatrix(Rw, 0)
                        Detail99.FormType = GetFormType
                        Detail99.TaxYear = Me.cmbTaxYear.text
                        Detail99.BoxName = .ColData(J)
                        Detail99.FieldValue = X
                        Detail99.Save (Equate.RecAdd)
                    End If
                End If
                
            Next J
        
            ' delete old records not update
            SQLString = " DELETE * FROM Detail99 WHERE " & _
                        " PayeeID = " & .TextMatrix(Rw, 0) & _
                        " AND FormType = '" & GetFormType() & "' " & _
                        " AND TaxYear = " & Me.cmbTaxYear.text & _
                        " AND FieldValue = 'JimBo'"
            cn.Execute SQLString
        
        Next Rw
    
    End With

    SaveNudge

End Sub
Private Function SetSelect(ByVal fgRow As Long) As Boolean
Dim fgCol As Long

    SetSelect = False
    For fgCol = 5 To fg.Cols - 1
        If Trim(fg.TextMatrix(fgRow, fgCol)) <> "" Then
            SetSelect = True
            Exit Function
        End If
    Next fgCol
    
End Function

Private Sub cmdExit_Click()
    GoBack
End Sub

Private Function GetFormType() As String

Dim ii As Integer

    ii = Me.cmbForm.ListIndex
    GetFormType = ""
    If ii = 0 Then GetFormType = "NEC"
    If ii = 1 Then GetFormType = "MISC"
    If ii = 2 Then GetFormType = "R"
    If ii = 3 Then GetFormType = "INT"
    If ii = 4 Then GetFormType = "DIV"

End Function

Private Function GetDetailData99(ByVal PayeeID As Integer, ByVal BoxName As String) As String
    
    SQLString = "select *" & _
                " from Field99" & _
                " where TaxYear = " & Me.cmbTaxYear.text & _
                " and FormType = '" & GetFormType & "'" & _
                " and BoxName = '" & BoxName & "'"
    If Field99.GetBySQL(SQLString) = False Then
        MsgBox "Field99 Not Found!! " & Me.cmbTaxYear.text & vbCrLf & Me.cmbForm.text & vbCrLf & BoxName
        GoBack
    End If
    
    SQLString = " SELECT * FROM Detail99 WHERE PayeeID = " & PayeeID & _
                " AND FormType = '" & Replace(Me.cmbForm.text, "1099-", "") & "' " & _
                " AND TaxYear = " & Me.cmbTaxYear.text & _
                " AND BoxName = '" & BoxName & "'"
    If Detail99.GetBySQL(SQLString) = False Then
        If Field99.FieldFormat = Equate.fmtAmount Then
            GetDetailData99 = FormatAmt("")
        Else
            GetDetailData99 = ""
        End If
    Else
        If Field99.FieldFormat = Equate.fmtAmount Then
            GetDetailData99 = Detail99.FieldValue
            GetDetailData99 = Replace(GetDetailData99, ",", "")
            GetDetailData99 = Replace(GetDetailData99, "$", "")
        Else
            GetDetailData99 = Detail99.FieldValue
        End If
        
    End If
    
End Function


