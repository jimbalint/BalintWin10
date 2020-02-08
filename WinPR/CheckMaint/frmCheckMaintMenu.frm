VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCheckMaintMenu 
   Caption         =   "Check Maintenance Menu"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAPCheck 
      Caption         =   "AP Check"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreateAll 
      Caption         =   "CREATE A&LL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&REATE FILE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   10
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdTestPrint 
      Caption         =   "&Test Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   8
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelCust 
      Caption         =   "&Delete Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdEditCust 
      Caption         =   "&Edit Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddCustomer 
      Caption         =   "Add &Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddClient 
      Caption         =   "&Add Client"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7800
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fgClient 
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      _cx             =   11456
      _cy             =   2566
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   6720
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid fgCust 
      Height          =   5535
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   7095
      _cx             =   12515
      _cy             =   9763
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   End
   Begin VB.Label Label1 
      Caption         =   "CLIENT LISTING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2925
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "CUSTOMER LISTING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2685
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmCheckMaintMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CustRow, rw As Long
Dim LoadFlag As Boolean
Dim GetSQLString As String
Dim rs As New ADODB.Recordset
Dim I, J, K As Long

Private Sub cmdAPCheck_Click()
    
Dim StartNum, CheckCount, CheckNum, M As Long
Dim pLine As Single

    Nudge = 30
    
    Open "\Balint\Data\APNudge.txt" For Input As #1
    Input #1, HorzNudge, VertNudge
    
    StartNum = CLng(InputBox("Starting check number?", rsCustomer!CustomerName))
    If StartNum = 0 Then Exit Sub
    CheckCount = CLng(InputBox("Number of checks to print?", rsCustomer!CustomerName))
    If CheckCount = 0 Then Exit Sub
    
    PrvwReturn = True

    PrtInit "Port":         pLine = 1
    Prvw.vsp.Font.Name = "ARIAL"
    Prvw.vsp.FontBold = True:       '  Turn on Bold Feature
    
    For CheckNum = StartNum To StartNum + CheckCount - 1
    
        pLine = 1
        
        '  Print Bank Name
        If CheckNum <> StartNum Then Prvw.vsp.NewPage
        Prvw.vsp.FontName = "Arial"
        SetFont 8, Portrait:                            Prvw.vsp.FontBold = True
        
        Prt pLine, 75, Trim(rsCustomer!Bank1)
        
        SetFont 13, Portrait:                           Prvw.vsp.FontBold = True:
        Prt pLine, 55, CheckNum                '''''''''''''          TAKE OUT   ''''''''''''''''''''
        pLine = pLine + 1
    
        '  Print Customer Name and Bank Info
        PosPrint 850, 450, Trim(rsCustomer!CustomerName):       Prvw.vsp.FontBold = False  '  Turn off BOLD feature
    
        SetFont 8, Portrait:
        If Trim(rsCustomer!Bank2) <> "" Then
            Prt pLine, 75, Trim(rsCustomer!Bank2):          pLine = pLine + 1
        End If
        
        If Trim(rsCustomer!Bank3) <> "" Then
            Prt pLine, 75, Trim(rsCustomer!Bank3):      pLine = pLine + 1
        End If
        
        If Trim(rsCustomer!Bank4) <> "" Then
            Prt pLine, 75, Trim(rsCustomer!Bank4):      pLine = pLine + 1
        End If
        
        If Trim(rsCustomer!BankFraction) <> "" Then
            Prt pLine, 75, Trim(rsCustomer!BankFraction):   pLine = pLine + 1
        End If
        
        '  Turn off BOLD feature and reduce font size for Customer Address Section
        Prvw.vsp.FontBold = False:                      SetFont 8, Portrait
    
        If Trim(rsCustomer!Address1) <> "" Then
            If rsCustomer!Addr1Bold = 1 Then
                Prvw.vsp.FontBold = True
                SetFont 10, Portrait
    
            Else
                Prvw.vsp.FontBold = False
                SetFont 8, Portrait
            End If
            PosPrint 850, 790, Trim(rsCustomer!Address1)
        End If
        SetFont 8, Portrait
        If Trim(rsCustomer!Address2) <> "" Then
            If rsCustomer!Addr2Bold = 1 Then
                Prvw.vsp.FontBold = True
            Else
                Prvw.vsp.FontBold = False
    
            End If
    
            PosPrint 850, 1040, Trim(rsCustomer!Address2)
        End If
            
        If Trim(rsCustomer!Address3) <> "" Then
            If rsCustomer!Addr3Bold = 1 Then
                Prvw.vsp.FontBold = True
            Else
                Prvw.vsp.FontBold = False
            End If
            PosPrint 850, 1290, Trim(rsCustomer!Address3)
    
        End If
    
        If Trim(rsCustomer!Address4) <> "" Then
            If rsCustomer!Addr4Bold = 1 Then
                Prvw.vsp.FontBold = True
            Else
                Prvw.vsp.FontBold = False
            End If
            PosPrint 850, 1540, Trim(rsCustomer!Address4)
    
        End If
        
    ''''''''''''''''''    SIGNATURE SECTION   '''''''''''''''''''''''''
    
        SetFont 10, Portrait
               
        If Trim(rsCustomer!SignImage1) <> "" Then
            SignPrint rsCustomer!Sign1Left, rsCustomer!Sign1Top - 150, rsCustomer!Sign1Width, rsCustomer!Sign1Height, rsCustomer!SignImage1
        End If
        pLine = 14
        
        ' =============================================================================
        ' AP fields
        SetFont 10, Portrait
        
        PosPrint 9000, 1200, "Date: " & String(18, "_")
        
        PosPrint 400, 1600, "Pay to the"
        PosPrint 400, 1880, "Order of:"
        
        PosPrint 1200, 1900, String(77, "_") & "$ " & String(15, "_")
        PosPrint 500, 2400, String(90, "_") & " Dollars"
        
        PosPrint 400, 3720, "Memo: " & String(40, "_")
        
        If rsCustomer!TwoSignLines = 1 Then
            PosPrint 6960, 3320, String(40, "_")
            PosPrint 6960, 3720, String(40, "_"):        pLine = pLine + 1
        Else
            pLine = pLine + 2
            PosPrint 6960, 3720, String(40, "_"):        pLine = pLine + 1
        End If
        
        PosPrint 400, 5100, rsCustomer!CustomerName
        PosPrint 11000, 5100, CheckNum
        
        PosPrint 400, 10100, rsCustomer!CustomerName
        PosPrint 11000, 10100, CheckNum
        
        ' =============================================================================

        Prvw.vsp.Font.Name = "MICR Encoding"
        Prvw.vsp.Font.Size = 18
        
        ' check number
        PosPrint 1755, 4290, "C" & Format(CheckNum, "000000000") & "C"
        
        ' ABA Number
        PosPrint 3910, 4290, "A" & Trim(rsCustomer!BankABA) & "A"
        
        ' Account Number
    
        If rsCustomer!AccountSpace = 2 Then
            PosPrint 6235, 4290, Trim(rsCustomer!BankAccount) & "C"
        Else
            PosPrint 6055, 4290, Trim(rsCustomer!BankAccount) & "C"
        End If
    
    Next CheckNum
    
    Prvw.vsp.EndDoc
    Prvw.Show

End Sub

Private Sub Form_Load()
    
    LoadFlag = True
    EditSw = False
    Me.KeyPreview = True
    
    ' Populate Client Listing
    SQLString = "SELECT * FROM Client ORDER BY ClientName"
    rsInit SQLString, cn, rsClient
    rsClient.MoveFirst
    SetGrid rsClient, fgClient
    LoadFlag = False
    
    If rsClient.RecordCount = 0 Then
        MsgBox "No Clients found!", vbCritical
    End If
    
    fgClient.ColWidth(0) = 900
    fgClient.ColWidth(1) = 5500
                        
    fgClient.SelectionMode = flexSelectionByRow
    fgClient.Editable = flexEDNone
    fgClient.AllowSelection = False
    
    rsClient.MoveFirst
    LoadCustGrid
        
End Sub

Private Sub cmdTestPrint_Click()
Dim pLine As Single

    PrvwReturn = True

    PrtInit "Port":         pLine = 1
    Prvw.vsp.Font.Name = "ARIAL"
    Prvw.vsp.FontBold = True:       '  Turn on Bold Feature
    pLine = 1
    
    '  Print Bank Name
    SetFont 8, Portrait:                            Prvw.vsp.FontBold = True
    
    Prt pLine, 75, Trim(rsCustomer!Bank1)
    
    SetFont 13, Portrait:                           Prvw.vsp.FontBold = True:
    Prt pLine, 60, "101"                '''''''''''''          TAKE OUT   ''''''''''''''''''''
    pLine = pLine + 1

    '  Print Customer Name and Bank Info
    PosPrint 1150, 300, Trim(rsCustomer!CustomerName):       Prvw.vsp.FontBold = False  '  Turn off BOLD feature

    SetFont 8, Portrait:
    If Trim(rsCustomer!Bank2) <> "" Then
        Prt pLine, 75, Trim(rsCustomer!Bank2):          pLine = pLine + 1
    End If
    
    If Trim(rsCustomer!Bank3) <> "" Then
        Prt pLine, 75, Trim(rsCustomer!Bank3):      pLine = pLine + 1
    End If
    
    If Trim(rsCustomer!Bank4) <> "" Then
        Prt pLine, 75, Trim(rsCustomer!Bank4):      pLine = pLine + 1
    End If
    
    If Trim(rsCustomer!BankFraction) <> "" Then
        Prt pLine, 75, Trim(rsCustomer!BankFraction):   pLine = pLine + 1
    End If
    
    '  Turn off BOLD feature and reduce font size for Customer Address Section
    Prvw.vsp.FontBold = False:                      SetFont 8, Portrait

    If Trim(rsCustomer!Address1) <> "" Then
        If rsCustomer!Addr1Bold = 1 Then
            Prvw.vsp.FontBold = True
            SetFont 10, Portrait

        Else
            Prvw.vsp.FontBold = False
            SetFont 8, Portrait
        End If
        PosPrint 1150, 590, Trim(rsCustomer!Address1)
    End If
    SetFont 8, Portrait
    If Trim(rsCustomer!Address2) <> "" Then
        If rsCustomer!Addr2Bold = 1 Then
            Prvw.vsp.FontBold = True
        Else
            Prvw.vsp.FontBold = False

        End If

        PosPrint 1150, 840, Trim(rsCustomer!Address2)
    End If
        
    If Trim(rsCustomer!Address3) <> "" Then
        If rsCustomer!Addr3Bold = 1 Then
            Prvw.vsp.FontBold = True
        Else
            Prvw.vsp.FontBold = False
        End If
        PosPrint 1150, 1090, Trim(rsCustomer!Address3)

    End If

    If Trim(rsCustomer!Address4) <> "" Then
        If rsCustomer!Addr4Bold = 1 Then
            Prvw.vsp.FontBold = True
        Else
            Prvw.vsp.FontBold = False
        End If
        PosPrint 1150, 1340, Trim(rsCustomer!Address4)

    End If
    
''''''''''''''''''    SIGNATURE SECTION   '''''''''''''''''''''''''

    SetFont 10, Portrait
           
    If Trim(rsCustomer!SignImage1) <> "" Then
        SignPrint rsCustomer!Sign1Left, rsCustomer!Sign1Top - 150, rsCustomer!Sign1Width, rsCustomer!Sign1Height, rsCustomer!SignImage1
    End If
    pLine = 14
    
    If rsCustomer!TwoSignLines = 1 Then
        PosPrint 6600, 3050, String(43, "_")
        PosPrint 6600, 3450, String(43, "_"):        pLine = pLine + 1
    Else
        pLine = pLine + 2
        PosPrint 6600, 3450, String(43, "_"):        pLine = pLine + 1
    End If
        
        
''''''''''''''''''    BOTTOM SECTION HEADERS  '''''''''''''''''''''''''
        
    Prvw.vsp.Font.Name = "Courier New"
    Prt 23, 55, "CHK DATE: ":                       Prt 23, 65, "11/10/2009"
    Prt 23, 78, "CHK #: ":                          Prt 23, 85, "101"
    Prt 24, 1, "- - - - - - - - - -  CURRENT PD - - YR TO DATE"
    Prt 24, 47, "- - - - - - - - -  CURRENT PD  - - - YR TO DATE"
            
    pLine = pLine + 7
    
     '  Turn on BOLD feature and increase FONT SIZE
    Prvw.vsp.FontBold = True:                       SetFont 13, Portrait
    Prvw.vsp.Font.Name = "ARIAL"
    PosPrint 400, 5140, Trim(rsCustomer!CustomerName):       Prvw.vsp.FontBold = False  '  Turn off BOLD feature

                    
    Prvw.vsp.Font.Name = "MICR Encoding"
    Prvw.vsp.Font.Size = 18
    
    ' check number
    PosPrint 1840, 4050, "C" & Format(101, "000000000") & "C"
    
    ' ABA Number
    PosPrint 3995, 4050, "A" & Trim(rsCustomer!BankABA) & "A"
    
    ' Account Number

    If rsCustomer!AccountSpace = 2 Then
        PosPrint 6320, 4050, Trim(rsCustomer!BankAccount) & "C"
    Else
        PosPrint 6140, 4050, Trim(rsCustomer!BankAccount) & "C"
    End If
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal
 
End Sub

Private Sub LoadCustGrid()
    
    ' Populate Customer Listing
    SQLString = "SELECT * FROM Customer WHERE CLIENTID = " & rsClient!ClientID & " ORDER BY CustomerID"
    SQLString = "SELECT * FROM Customer WHERE CLIENTID = " & rsClient!ClientID & " ORDER BY CustomerName"
    rsInit SQLString, cn, rsCustomer
    SetGrid rsCustomer, fgCust

    If rsCustomer.RecordCount > 0 Then
        rsCustomer.MoveFirst
    Else
        Exit Sub
    End If
            
    fgCust.ColWidth(0) = 1200
    fgCust.ColWidth(1) = 4000
    fgCust.ColWidth(3) = 1200
    For I = 2 To 23
        fgCust.ColHidden(I) = True
    Next I
    
    fgCust.SelectionMode = flexSelectionByRow
    fgCust.Editable = flexEDNone
    fgCust.AllowSelection = True

End Sub
Private Sub cmdAddClient_Click()
    frmAddClient.Show vbModal
End Sub

Private Sub cmdAddCustomer_Click()
    EditSw = False
    ClientID = rsClient!ClientID
    ClientName = rsClient!ClientName
    ClearCust
    frmUpdateCustomer.Show vbModal
    rsCustomer.Close
    LoadCustGrid
End Sub

Private Sub cmdEditCust_Click()
Dim CurrID As Long
    
    CurrID = fgCust.TextMatrix(fgCust.Row, 0)
    
    EditSw = True
 
    If rsCustomer.RecordCount = 0 Then
        MsgBox "This Client has no Customers!", vbCritical
    Else
        CustID = rsCustomer!CustomerID
        ClientID = rsClient!ClientID
        ClientName = rsClient!ClientName
        frmUpdateCustomer.Show vbModal

        Unload frmUpdateCustomer
        rsCustomer.Close
        LoadCustGrid

        rsCustomer.Find "CustomerID = " & rsCustomer!CustomerID, 0, adSearchForward, 1
        If rsCustomer!TwoSignLines = 1 Then
            frmUpdateCustomer.chkTwoSigs = 1
        End If
        Set fgCust.DataSource = rsCustomer.DataSource
           
        rw = fgCust.FindRow(CurrID, 0, 0)
           
        fgCust.TopRow = rw
        fgCust.Select rw, 0
        fgCust.SetFocus
    End If


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: cmdExit_Click
    End Select
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub fgcust_DblClick()
    cmdEditCust_Click
End Sub


Private Sub cmdDelCust_Click()
    
Dim DelConfirm As Integer
    
    If fgCust.Rows = 1 Then Exit Sub
    DelConfirm = MsgBox(fgCust.TextMatrix(fgCust.Row, 1) & vbCr & Trim(fgCust.TextMatrix(fgCust.Row, 2)) & ", " & Trim(fgCust.TextMatrix(fgCust.Row, 3)), _
                        vbExclamation + vbYesNo + vbDefaultButton2, _
                        "Are you S U R E you want to delete:")
    
    If DelConfirm = vbNo Then
       fgCust.SetFocus
       Exit Sub
    End If
    
    ' delete record from file
    SQLString = "DELETE * FROM Customer WHERE CUSTOMER.CUSTOMERID = " & fgCust.TextMatrix(fgCust.Row, 0)

    rw = fgCust.Row
    ' delete record from grid
    rsCustomer.Delete
    rsCustomer.Update      ' Record (save to file)
    rsCustomer.MoveLast    ' Move to the last record in the record set
    fgCust.DataRefresh     ' Update the grid data
    fgCust.Col = 0         ' Go to the first column
    fgCust.SetFocus        ' Move from add button to grid
    
    If rw = fgCust.Rows Then rw = fgCust.Rows - 1
    
    fgCust.Select rw, 0
    fgCust.ShowCell rw, 0

End Sub

Private Sub fgClient_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If LoadFlag = True Then Exit Sub
    rsCustomer.Close
    LoadCustGrid
        
End Sub

Public Sub ClearCust()
    frmUpdateCustomer.tdbCompanyID = 0
    frmUpdateCustomer.TDBCustName = ""
    frmUpdateCustomer.TDBAddr1 = ""
    frmUpdateCustomer.TDBAddr2 = ""
    frmUpdateCustomer.TDBAddr3 = ""
    frmUpdateCustomer.TDBAddr4 = ""
    frmUpdateCustomer.chkBoldAddr1 = 0
    frmUpdateCustomer.chkBoldAddr2 = 0
    frmUpdateCustomer.chkBoldAddr3 = 0
    frmUpdateCustomer.chkBoldAddr4 = 0
    frmUpdateCustomer.TDBBank1 = ""
    frmUpdateCustomer.TDBBank2 = ""
    frmUpdateCustomer.TDBBank3 = ""
    frmUpdateCustomer.TDBBank4 = ""
    frmUpdateCustomer.TDBBankFraction = ""
    frmUpdateCustomer.TDBBankAccount = ""
    frmUpdateCustomer.TDBAcctSpaces = 0
    frmUpdateCustomer.TDBBankABA = ""
    frmUpdateCustomer.TDBSignImage1 = ""
    frmUpdateCustomer.TDBSign1Left = 0
    frmUpdateCustomer.TDBSign1Top = 0
    frmUpdateCustomer.TDBSign1Height = 0
    frmUpdateCustomer.TDBSign1Width = 0
    frmUpdateCustomer.TDBSignImage2 = ""
    frmUpdateCustomer.TDBSign2Left = 0
    frmUpdateCustomer.tdbSign2Top = 0
    frmUpdateCustomer.TDBSign2Height = 0
    frmUpdateCustomer.TDBSign2Width = 0
    frmUpdateCustomer.TDBLogo = ""
End Sub

Private Sub SignPrint(ByVal SignLeft As Long, _
                      ByVal SignTop As Long, _
                      ByVal SignWidth As Long, _
                      ByVal SignHeight As Long, _
                      ByVal FileName As String)
    Prvw.Picture1.Picture = LoadPicture(Trim("c:\balint\CheckData\" & FileName))
    Prvw.vsp.DrawPicture Prvw.Picture1, SignLeft, SignTop, SignWidth, SignHeight, 10
End Sub

Private Sub fgCust_AfterEdit(ByVal Row As Long, ByVal Col As Long)

' resort after edit and move to that row

Dim CurrID As Long
    
    CurrID = fgCust.TextMatrix(fgCust.Row, 0)
        
    rsCustomer.Close
    rsInit GetSQLString, cn, rs
    Set fgCust.DataSource = rs.DataSource
       
    rw = fgCust.FindRow(CurrID, 0, 0)
       
    fgCust.TopRow = rw
    fgCust.Select rw, 0
    fgCust.SetFocus

End Sub
Private Sub cmdCreate_Click()
    
Dim PassWord, BlankName, fName As String
Dim rsPRCK As New ADODB.Recordset
Dim db As DAO.Database
    
    If rsCustomer!PRCompanyID = 0 Then
        MsgBox "Customer.PRCompanyID must be assigned!", vbExclamation
        Exit Sub
    End If
    
    BlankName = "\Balint\Blank\BLANK.mdb"
    PassWord = "pobox45"
    
    ' formulate the output file name
    fName = "\Balint\CheckData\PRCK" & Trim(rsClient!Prefix) & _
            Format(rsCustomer!PRCompanyID, "000000") & ".mdb"
       
    ' copy blank MDB to the output MDB
    FileCopy BlankName, fName
    
    ' open the connection
    Set cnPRCK = New ADODB.Connection
    cnPRCK.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnPRCK.ConnectionString = fName
    cnPRCK.Open

    ' create the file structure
    SQLString = "CREATE TABLE PRCheck ( " & _
                        "[PRCheckID] Counter, CONSTRAINT prckIDKey PRIMARY KEY ([PRCheckID]) ) "
                        
    cnPRCK.Execute SQLString
                        
    AddField "PRCheck", "CustomerName", "Char (40)", cnPRCK
    AddField "PRCheck", "ClientID", "Long", cnPRCK
    AddField "PRCheck", "PRCompanyID", "Long", cnPRCK
    AddField "PRCheck", "Address1", "Char (40)", cnPRCK
    AddField "PRCheck", "Address2", "Char (40)", cnPRCK
    AddField "PRCheck", "Address3", "Char (40)", cnPRCK
    AddField "PRCheck", "Address4", "Char (40)", cnPRCK
    AddField "PRCheck", "Addr1Bold", "Byte", cnPRCK
    AddField "PRCheck", "Addr2Bold", "Byte", cnPRCK
    AddField "PRCheck", "Addr3Bold", "Byte", cnPRCK
    AddField "PRCheck", "Addr4Bold", "Byte", cnPRCK
    AddField "PRCheck", "Bank1", "Char (40)", cnPRCK
    AddField "PRCheck", "Bank2", "Char (40)", cnPRCK
    AddField "PRCheck", "Bank3", "Char (40)", cnPRCK
    AddField "PRCheck", "Bank4", "Char (40)", cnPRCK
    AddField "PRCheck", "BankFraction", "Char (40)", cnPRCK
    AddField "PRCheck", "BankABA", "Char (9)", cnPRCK
    AddField "PRCheck", "BankAccount", "Char (40)", cnPRCK
    AddField "PRCheck", "AccountSpace", "Byte", cnPRCK
    AddField "PRCheck", "TwoSignLines", "Byte", cnPRCK
    AddField "PRCheck", "SignImage1", "Char (40)", cnPRCK
    AddField "PRCheck", "Sign1Left", "Long", cnPRCK
    AddField "PRCheck", "Sign1Top", "Long", cnPRCK
    AddField "PRCheck", "Sign1Height", "Long", cnPRCK
    AddField "PRCheck", "Sign1Width", "Long", cnPRCK
    AddField "PRCheck", "SignImage2", "Char (40)", cnPRCK
    AddField "PRCheck", "Sign2Left", "Long", cnPRCK
    AddField "PRCheck", "Sign2Top", "Long", cnPRCK
    AddField "PRCheck", "Sign2Height", "Long", cnPRCK
    AddField "PRCheck", "Sign2Width", "Long", cnPRCK
    AddField "PRCheck", "LogoImage", "Char (40)", cnPRCK
    AddField "PRCheck", "CreateDate", "DateTime", cnPRCK
    AddField "PRCheck", "ModifyDate", "DateTime", cnPRCK
    
    AddField "PRCheck", "BankAccountAdd", "Char (10)", cnPRCK
    AddField "PRCheck", "AddressAdjust", "Long", cnPRCK

    ' update the fields
    SQLString = "SELECT * FROM PRCheck"
    rsInit SQLString, cnPRCK, rsPRCK

    rsPRCK.AddNew
    rsPRCK!CustomerName = rsCustomer!CustomerName
    rsPRCK!ClientID = rsCustomer!ClientID
    rsPRCK!PRCompanyID = rsCustomer!PRCompanyID
    rsPRCK!Address1 = rsCustomer!Address1
    rsPRCK!Address2 = rsCustomer!Address2
    rsPRCK!Address3 = rsCustomer!Address3
    rsPRCK!Address4 = rsCustomer!Address4
    rsPRCK!Addr1Bold = rsCustomer!Addr1Bold
    rsPRCK!Addr2Bold = rsCustomer!Addr2Bold
    rsPRCK!Addr3Bold = rsCustomer!Addr3Bold
    rsPRCK!Addr4Bold = rsCustomer!Addr4Bold
    rsPRCK!Bank1 = rsCustomer!Bank1
    rsPRCK!Bank2 = rsCustomer!Bank2
    rsPRCK!Bank3 = rsCustomer!Bank3
    rsPRCK!Bank4 = rsCustomer!Bank4
    rsPRCK!BankFraction = rsCustomer!BankFraction
    rsPRCK!BankABA = rsCustomer!BankABA
    rsPRCK!BankAccount = rsCustomer!BankAccount
    rsPRCK!AccountSpace = rsCustomer!AccountSpace
    rsPRCK!TwoSignLines = rsCustomer!TwoSignLines
    rsPRCK!SignImage1 = Trim(rsCustomer!SignImage1)
    rsPRCK!Sign1Left = rsCustomer!Sign1Left
    rsPRCK!Sign1Top = rsCustomer!Sign1Top
    rsPRCK!Sign1Height = rsCustomer!Sign1Height
    rsPRCK!Sign1Width = rsCustomer!Sign1Width
    rsPRCK!SignImage2 = Trim(rsCustomer!SignImage2)
    rsPRCK!Sign2Left = rsCustomer!Sign2Left
    rsPRCK!Sign2Top = rsCustomer!Sign2Top
    rsPRCK!Sign2Height = rsCustomer!Sign2Height
    rsPRCK!Sign2Width = rsCustomer!Sign2Width
    rsPRCK!LogoImage = rsCustomer!LogoImage
    rsPRCK!CreateDate = Now()
    rsPRCK!ModifyDate = Now()
    
    rsPRCK!BankAccountAdd = rsCustomer!BankAccountAdd
    rsPRCK!AddressAdjust = rsCustomer!AddressAdjust
    
    rsPRCK.Update

    ' close the connection
    cnPRCK.Close
    Set cnPRCK = Nothing

    ' set the password
    Set db = OpenDatabase(Name:=fName, _
                          Options:=True, _
                          ReadOnly:=False)
    db.NewPassword "", PassWord
    db.Close
    
    MsgBox Trim(fName) & vbCr & vbCr & "Has been created" & vbCr & _
           "Password is: " & vbCr & PassWord, vbInformation

End Sub

Private Sub cmdCreateAll_Click()
    rsCustomer.MoveFirst
    Do
        cmdCreate_Click
        rsCustomer.MoveNext
    Loop Until rsCustomer.EOF
End Sub


