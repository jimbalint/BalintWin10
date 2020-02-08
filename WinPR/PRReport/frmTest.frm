VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8Ctl.VSFlexGrid fg 
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      _cx             =   12303
      _cy             =   5953
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
      DataMode        =   4
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j As Integer
Dim X, Y As String
Dim p1 As Currency
Dim rs As New ADODB.Recordset
Dim YearFlag As Boolean

Private Sub Form_Load()
    
    X = "c:\Balint\Data\JimBo[].txt"
    
    i = InStr(1, X, "[]", vbTextCompare)
    If i <> 0 Then
        Y = Mid(X, 1, i - 1) & "100209" & Mid(X, i + 2, Len(X) - 1 + 2)
    Else
        Y = X
    End If
    
    MsgBox Y
    End
    
    MICRPrint
    
    
'    PicPrint
    
'    fg.FixedCols = 0                   ' see all cols selected by SQL
'    fg.FocusRect = flexFocusSolid      ' Cell apearance when editable & in focus
'    fg.Editable = flexEDKbdMouse       ' Edit by keys or mouse picks
'
'    fg.BackColorAlternate = RGB(192, 192, 192)          ' light gray
'    fg.TabBehavior = flexTabCells                       ' tab moves between cells
'    fg.AllowSelection = False                          ' don't allow selection of ranges of cells
'
'    fg.TextMatrix(0, 1) = "2"
'    fg.TextMatrix(1, 1) = "4"
'    fg.TextMatrix(2, 1) = "10.5"
'
'    fg.ColFormat(1) = "$##,##0.00"

'    PrtInit ("Port")    ' "Port" = Portrait
'    SetFont 12, Equate.Portrait
'
'    Ln = 10
'
'    x = "OCR A Extended"
'    'x = "PrecisionID OCR A1"
'    Prvw.vsp.Font.Name = x
'
'    PrintValue(1) = x:      FormatString(1) = "a50"
'    PrintValue(2) = "":         FormatString(2) = "~"
'    FormatPrint
'
'    Ln = 15
'
'    For i = 1 To 10
'        PrintValue(1) = "123456780":     FormatString(1) = "a20"
'        PrintValue(2) = "":         FormatString(2) = "~"
'        FormatPrint
'        Ln = Ln + 1
'    Next i
'
'    Prvw.vsp.EndDoc
'    Prvw.Show vbModal
'
'    End

End Sub
Private Sub Command1_Click()
    
    For i = 1 To fg.Rows
        X = fg.TextMatrix(i - 1, 1)
        If X <> "" Then
            p1 = p1 + CCur(X)
        End If
    Next i
    MsgBox p1
    
    
    End
End Sub

Private Sub PicPrint()

    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Ln = 0

    For i = 1 To 10
        Ln = Ln + 1
        PrintValue(1) = "XXXXX": FormatString(1) = "a10"
        PrintValue(2) = "": FormatString(2) = "~"
        FormatPrint
    Next i

    X = "c:\asend\petit.jpg"
    X = "c:\asend\balintteal.jpg"
    X = "c:\balint\data\Leone.jpg"
    X = "c:\balint\data\CIC.jpg"
        
    Prvw.Picture1.Picture = LoadPicture(X)
    ' Prvw.vsp.DrawPicture Prvw.Picture1, "5000twips", "5000twips", "15%", "15%"
    ' left / top / width / height
    Prvw.vsp.DrawPicture Prvw.Picture1, "7900", "2100", "3000", "3000", 10
    Prvw.vsp.EndDoc
    Prvw.Show vbModal


    End

End Sub

Private Sub MICRPrint()

    PrtInit ("Port")
    SetFont 10, Equate.Portrait
    Ln = 0

    Prvw.vsp.Font.Name = "Arial"
    Prvw.vsp.Font.Size = 12

    For i = 1 To 10
        Ln = Ln + 1
        If Ln Mod 2 = 0 Then
            Prvw.vsp.Font.Bold = True
        Else
            Prvw.vsp.Font.Bold = False
        End If
        
        PrintValue(1) = "XXXXX": FormatString(1) = "a10"
        PrintValue(2) = "": FormatString(2) = "~"
        FormatPrint
    
    Next i
    
    Prvw.vsp.Font.Bold = False
    
    Prvw.vsp.Font.Name = "MICR Encoding"
    Prvw.vsp.Font.Size = 18
    
    ' check number
    PosPrint 1840, 4390, "C" & Format(101, "000000000") & "C"
    
    ' ABA number
    PosPrint 3995, 4390, "A" & "041200555" & "A"
    
    ' bank account number
    PosPrint 6140, 4390, "C" & "10265919" & "C"
    
    Prvw.vsp.EndDoc
    Prvw.Show vbModal

    End

End Sub
